#!/usr/bin/env python3
"""
design_core.py: BNI experiment design (v0.5, 2026-11-05). No Qt, no Excel.

From a plan (genotypes x cohorts x treatments x reps, the bench, the N2O capacity) it
builds one row per pot, then:

  1. PLANTING GROUPS. Each pot is measured for N2O on fixed days after planting (DAP),
     one pot at a time, and the bench can measure at most `max_measured_per_day` pots a
     day. So planting is staggered: the pots are split into the FEWEST equally sized
     groups, planted `stagger_days` apart, such that no calendar day needs more than
     the capacity, and no day has a flood and an N2O measurement together (they cannot
     be run at the same time). Groups are at least 3 days apart by default: one group's
     24 h and 48 h N2O days, then the next group's flood. The split is random but
     balanced: every cell (genotype x cohort x treatment x role) is dealt across the
     groups as evenly as it divides. Spares come after: `spares_per_genotype_per_group`
     of every genotype in every group, because a spare can only replace a pot planted
     the same day. They are never measured, so they never change the number of groups.
  2. BENCH LAYOUT. Pots are dealt across the available tables, then the zones of each
     table, keeping every cell, and every planting group, as evenly spread as it
     divides; slots within a zone are random (empty slots too). Position text is
     Tn-Zn-nn, the form the dosing station reads from a pot's tag.
     Inside a group, cohorts whose days would collide (one's flood on another's N2O
     day) are not refused or respaced: the later one is PLANTED LATER (cohort_offsets),
     so every cohort keeps its days after planting and nothing is planted before the
     planting date.
  3. SCHEDULE. Per group and cohort: plant (DAP 0), treatment (treatment_dap), flood and
     sacrificial harvest (the first N2O time before the first N2O DAP), N2O at each of the
     cohort's own times (e.g. 24, 48, 72 h after the flood),
     harvest after the last one. Daily N2O load is counted per calendar day. Watering is
     Monday, Wednesday and Friday, after planting until the flood. Treatment, flood, N2O
     and harvest may fall on any day, weekends included.
  4. N2O ORDER. Per group and cohort the measured pots are put in loads of three (one
     pot per N2O column), each load mixing treatments and genotypes, with every cell
     spread over the three columns. The flood is done in the same order, so every pot
     has the same time between its flood and its measurement.
  5. TASKS. One dated task per piece of work, for Asana (no repeat option there, so every
     repeat is its own task): soil, planting and labelling around each planting day,
     the ZeaL germination check on the weekdays of the week after planting, watering,
     treatment, flood, N2O, harvest, then KCl, dry weights and washing a set number of
     working days later (TASK_LAGS).

The same `seed` always gives the same design. Nothing here knows about Excel; the
workbook (design_xlsx.py) writes what this computes, plus formulas so the dates can be
edited afterwards.
"""
from __future__ import annotations

import datetime as dt
import random
from collections import Counter, defaultdict
from dataclasses import dataclass, field, asdict

__version__ = "0.5"

MAX_ZONES = 8              # a table holds at most 8 zones x 25 slots
MAX_SLOTS = 25
MEASURE_HOURS = (24, 48)   # N2O after the flood; the last one is followed by harvest
MAX_STAGGER_DAYS = 30      # search limit for the gap between planting groups
DEFAULT_STAGGER_DAYS = 3   # 24 h N2O, 48 h N2O, then the next group's flood
WATER_WEEKDAYS = (0, 2, 4)  # Monday, Wednesday, Friday (date.weekday())
WATER_DAY_NAMES = "Mon, Wed, Fri"
N2O_COLUMNS = 3            # pots per N2O load (one per chamber)

MEASURED, SACRIFICIAL, SPARE = "MEASURED", "SACRIFICIAL", "SPARE"
# after swaps (swaps.py): a pot that failed and was replaced by a spare, and a spare that
# was not needed and is thrown out. Both stay on record but are no longer grown.
FAILED, DISCARDED = "FAILED", "DISCARDED"
GONE = (FAILED, DISCARDED)

# Tasks off the bench schedule, in working days (Mon-Fri) from the event they follow
# (Excel's WORKDAY, so the workbook recalculates them). Negative = before.
TASK_LAGS = {
    "prefill": -1,          # pre-fill pots: the working day before planting
    "kcl_setup": -1,        # KCl supplies and tubes: the working day before a harvest
    "kcl_extract": 1,       # KCl batch 1: the working day after the LAST harvest; batch k: k
    "kcl_analysis": 1,      # plate reader: working days after each batch's extraction
    "dry_weights": 3,       # oven-dried: 3 working days after a harvest
    "wash": 1,              # wash pots: the working day after the final harvest
}
KCL_BATCH_MAX = 72         # soil samples one KCl extraction can take
GERMINATION_WORKDAYS = 5    # ZeaL germination check: the weekdays of the week after planting


# ======================================================================================
# THE PLAN: what the user types in the first window
# ======================================================================================
@dataclass
class Cohort:
    name: str
    first_n2o_dap: int             # day after planting of the cohort's first N2O sample
    measure_hours: tuple[int, ...] = MEASURE_HOURS   # N2O after the flood, whole days

    @property
    def flood_dap(self) -> int:
        """The flood: the first sample's hours before the first N2O day (24 h: the day
        before)."""
        return self.first_n2o_dap - self.measure_hours[0] // 24

    @property
    def offsets(self) -> tuple[int, ...]:
        """Each sample's day after the first N2O day: 24, 48, 72 h -> 0, 1, 2."""
        return tuple(h // 24 - self.measure_hours[0] // 24 for h in self.measure_hours)

    @property
    def last_offset(self) -> int:
        return self.offsets[-1]


@dataclass
class PlanRow:
    genotype: str
    cohort: str
    treatment: str = ""            # blank: no treatment
    reps: int = 0                  # measured pots
    sacrificial: int = 0           # harvested on flood day, never measured


@dataclass
class Plan:
    name: str = "BNI-P1"           # experiment id; barcodes are <name>-S001...
    seed: int = 1
    planting_date: dt.date = field(default_factory=dt.date.today)
    treatment_dap: int | None = None
    tables: int = 6
    zones_per_table: int = 3
    slots_per_zone: int = 8
    max_measured_per_day: int = 50
    min_stagger_days: int = DEFAULT_STAGGER_DAYS
    planting_groups: int = 0       # 0 = choose automatically
    spares_per_genotype_per_group: int = 0   # a spare replaces a pot of its own age
    # dosing station (the workbook's `admin` tab); None = not known yet, left blank there
    wfps_target: float | None = None         # fraction, e.g. 0.40 (wet-up and maintenance)
    wfps_flood: float | None = None          # fraction, e.g. 0.95 (flood target)
    bulk_density: float | None = None        # g/cm3
    particle_density: float = 2.65           # g/cm3
    air_dry_moisture: float | None = None    # fraction of air-dry soil mass, e.g. 0.04
    air_dry_soil_g: float | None = None      # air-dry soil per pot, g
    container_weight_g: float | None = None  # empty pot, g
    cohorts: list[Cohort] = field(default_factory=list)
    rows: list[PlanRow] = field(default_factory=list)

    # ---- save / load (the app's plan file) ----
    def to_dict(self) -> dict:
        d = asdict(self)
        d["planting_date"] = self.planting_date.isoformat()
        for c in d["cohorts"]:
            c["measure_hours"] = list(c["measure_hours"])
        return d

    @classmethod
    def from_dict(cls, d: dict) -> "Plan":
        d = dict(d)
        d["planting_date"] = dt.date.fromisoformat(d["planting_date"])
        # plan files before v0.5 had one list of N2O times for every cohort
        hours = tuple(d.pop("measure_hours", MEASURE_HOURS))
        d["cohorts"] = [Cohort(**{**c, "measure_hours": tuple(c.get("measure_hours", hours))})
                        for c in d.get("cohorts", [])]
        d["rows"] = [PlanRow(**r) for r in d.get("rows", [])]
        if "spares_per_genotype" in d:                 # v0.1/v0.2 plan files
            d.setdefault("spares_per_genotype_per_group", d.pop("spares_per_genotype"))
        known = set(cls.__dataclass_fields__)
        return cls(**{k: v for k, v in d.items() if k in known})

    @property
    def capacity(self) -> int:
        return self.tables * self.zones_per_table * self.slots_per_zone

    @property
    def max_samples(self) -> int:
        """The most N2O samples any cohort takes."""
        return max((len(c.measure_hours) for c in self.cohorts), default=len(MEASURE_HOURS))

    def cohort(self, name: str) -> Cohort | None:
        return next((c for c in self.cohorts if c.name == name), None)


class PlanError(ValueError):
    """The plan cannot be built; the message says what to fix."""


# The dosing station's admin-tab inputs: (attribute, admin key, plausible range). The
# ranges are the station's own (wfps.py), so a slipped decimal is caught here, not at the
# bench. Blank is allowed: the tab is written with the cell empty, to fill in later.
ADMIN_INPUTS = (("wfps_target", "wfps_target", (0.2, 1.0)),
                ("wfps_flood", "wfps_flood", (0.2, 1.0)),
                ("bulk_density", "bulk_density", (0.5, 2.0)),
                ("particle_density", "particle_density", (2.0, 3.5)),
                ("air_dry_moisture", "air_dry_moisture", (0.0, 0.25)),
                ("air_dry_soil_g", "air_dry_soil_g", (1.0, 5000.0)),
                ("container_weight_g", "container_weight_g", (1.0, 5000.0)))
ADMIN_REQUIRED = ("wfps_target", "bulk_density", "air_dry_moisture", "air_dry_soil_g",
                  "container_weight_g")


def admin_missing(plan: Plan) -> list[str]:
    """Admin inputs the dosing station needs that the plan does not have yet."""
    return [a for a in ADMIN_REQUIRED if getattr(plan, a) is None]


def validate(plan: Plan) -> list[str]:
    """Problems that stop a build (empty list = OK)."""
    p = []
    for attr, key, (lo, hi) in ADMIN_INPUTS:
        v = getattr(plan, attr)
        if v is not None and not lo <= v <= hi:
            p.append(f"{key} {v:g} is outside {lo:g} to {hi:g}"
                     + (" (a fraction: 0.40 for 40 %)" if "wfps" in key or "moisture" in key
                        else ""))
    if (plan.wfps_flood is not None and plan.wfps_target is not None
            and plan.wfps_flood <= plan.wfps_target):
        p.append(f"flood WFPS {plan.wfps_flood:g} must be above the target WFPS "
                 f"{plan.wfps_target:g}")
    if not plan.name.strip():
        p.append("experiment name is empty")
    if not 1 <= plan.tables:
        p.append("at least one table is needed")
    if not 1 <= plan.zones_per_table <= MAX_ZONES:
        p.append(f"zones per table must be 1 to {MAX_ZONES}")
    if not 1 <= plan.slots_per_zone <= MAX_SLOTS:
        p.append(f"slots per zone must be 1 to {MAX_SLOTS}")
    if plan.max_measured_per_day < 1:
        p.append("max N2O measurements per day must be at least 1")
    if plan.min_stagger_days < 2:
        p.append("days between planting groups must be at least 2")
    names = [c.name for c in plan.cohorts]
    if not names:
        p.append("add at least one cohort")
    if len(set(names)) != len(names):
        p.append("cohort names must be different")
    for c in plan.cohorts:
        if not c.name.strip():
            p.append("a cohort has no name")
        h = c.measure_hours
        if not h:
            p.append(f"cohort {c.name}: give at least one N2O time (e.g. 24, 48)")
            continue
        if any(x % 24 or x < 24 for x in h):
            p.append(f"cohort {c.name}: N2O times must be whole days after the flood "
                     f"(24, 48, 72 h ...), not {', '.join(map(str, h))}")
            continue
        if list(h) != sorted(set(h)):
            p.append(f"cohort {c.name}: N2O times must be in order, each once "
                     f"({', '.join(map(str, h))})")
            continue
        if c.flood_dap < 1:
            p.append(f"cohort {c.name}: the flood (first N2O day {c.first_n2o_dap} minus "
                     f"{h[0]} h) would fall before planting")
        if plan.treatment_dap is not None and plan.treatment_dap >= c.flood_dap:
            p.append(f"cohort {c.name}: treatment (DAP {plan.treatment_dap}) must come before "
                     f"the flood (DAP {c.flood_dap})")
    if not any(r.reps or r.sacrificial for r in plan.rows):
        p.append("the plan has no pots: enter reps")
    seen = set()
    for i, r in enumerate(plan.rows, 1):
        if not (r.genotype.strip() or r.reps or r.sacrificial):
            continue                                   # a blank row
        if not r.genotype.strip():
            p.append(f"row {i}: genotype is empty")
        if plan.cohort(r.cohort) is None:
            p.append(f"row {i}: cohort {r.cohort!r} is not in the cohort list")
        if r.reps < 0 or r.sacrificial < 0:
            p.append(f"row {i}: reps cannot be negative")
        key = (r.genotype.strip(), r.cohort, r.treatment.strip())
        if key in seen:
            p.append(f"row {i}: {key[0]} / {key[1]} / {key[2] or 'no treatment'} is listed twice")
        seen.add(key)
    if plan.spares_per_genotype_per_group < 0:
        p.append("spares cannot be negative")
    n = count_pots(plan)                               # spares are checked once groups are known
    if n > plan.capacity:
        p.append(f"{n} pots do not fit: {plan.tables} tables x {plan.zones_per_table} zones x "
                 f"{plan.slots_per_zone} slots = {plan.capacity}")
    return p


def genotypes(plan: Plan) -> list[str]:
    out = []
    for r in plan.rows:
        g = r.genotype.strip()
        if g and g not in out:
            out.append(g)
    return out


def count_pots(plan: Plan, groups: int = 0) -> int:
    """Measured + sacrificial pots, plus the spares for `groups` planting groups (0: the
    spares are left out, the number of groups not being known yet)."""
    rows = [r for r in plan.rows if r.genotype.strip()]
    return (sum(max(0, r.reps) + max(0, r.sacrificial) for r in rows)
            + plan.spares_per_genotype_per_group * len(genotypes(plan)) * groups)


# ======================================================================================
# POTS
# ======================================================================================
@dataclass
class Pot:
    barcode: str = ""
    role: str = MEASURED
    genotype: str = ""
    cohort: str = ""
    treatment: str = ""
    rep: int = 0
    group: int = 0                 # planting group, 1-based
    group_in_tag: bool = False     # spares, when there are several planting groups
    load: int = 0                  # N2O load within its group x cohort (measured pots)
    column: int = 0                # N2O column within the load
    seq: int = 0                   # order within its group x cohort: flood, N2O, harvest
    table: int = 0
    zone: int = 0
    slot: int = 0
    was: str = ""                  # FAILED: the role it had (MEASURED / SACRIFICIAL)
    swap: str = ""                 # the other pot of its swap: the spare that replaced a
                                   # FAILED pot, or the failed pot a promoted spare replaces

    @property
    def active(self) -> bool:
        """Still grown and dosed: not failed, not a discarded spare."""
        return self.role not in GONE

    @property
    def cell(self) -> tuple:
        return (self.genotype, self.cohort, self.treatment, self.role)

    @property
    def position(self) -> str:
        return f"T{self.table}-Z{self.zone}-{self.slot:02d}"

    @property
    def label(self) -> str:
        """Human tag: genotype-cohort-treatment-r03 / -SAC2 / SPARE-01 (G2-SPARE-01 when
        there are several planting groups: a spare only replaces a pot of its own group)."""
        role = self.was if self.role == FAILED else self.role
        if role in (SPARE, DISCARDED):
            g = f"G{self.group}-" if self.group_in_tag else ""
            return (f"{self.genotype}-{g}SPARE-{self.rep:02d}"
                    + ("-DISCARDED" if role == DISCARDED else ""))
        parts = [self.genotype, self.cohort] + ([self.treatment] if self.treatment else [])
        rep = f"r{self.rep:02d}" if role == MEASURED else f"SAC{self.rep}"
        return "-".join(parts + [rep]) + ("-FAILED" if self.role == FAILED else "")


def make_pots(plan: Plan) -> list[Pot]:
    """Measured and sacrificial pots; spares are added per group by add_spares()."""
    pots = []
    for r in plan.rows:
        g = r.genotype.strip()
        if not g:
            continue
        for k in range(1, max(0, r.reps) + 1):
            pots.append(Pot(role=MEASURED, genotype=g, cohort=r.cohort,
                            treatment=r.treatment.strip(), rep=k))
        for k in range(1, max(0, r.sacrificial) + 1):
            pots.append(Pot(role=SACRIFICIAL, genotype=g, cohort=r.cohort,
                            treatment=r.treatment.strip(), rep=k))
    return pots


def add_spares(plan: Plan, pots: list[Pot], groups: int) -> None:
    """`spares_per_genotype_per_group` spares of every genotype in every planting group."""
    for g in range(1, groups + 1):
        for geno in genotypes(plan):
            for k in range(1, plan.spares_per_genotype_per_group + 1):
                pots.append(Pot(role=SPARE, genotype=geno, rep=k, group=g,
                                group_in_tag=groups > 1))


def _deal(pots: list[Pot], bins: list, rng: random.Random, key_extra=None,
          capacity: dict | None = None) -> dict:
    """Assign every pot to a bin so each cell, then each bin's total, is as even as it
    divides. Cells are dealt one after another in random order, pots within a cell in
    random order; ties between bins are broken at random. Returns {id(pot): bin}."""
    by_cell = defaultdict(list)
    for p in pots:
        by_cell[p.cell].append(p)
    cells = list(by_cell)
    rng.shuffle(cells)
    cells.sort(key=lambda c: -len(by_cell[c]))       # big cells first: they set the shape
    cell_n = defaultdict(Counter)
    extra_n = defaultdict(Counter)
    total = Counter()
    out = {}
    for c in cells:
        members = by_cell[c][:]
        rng.shuffle(members)
        for p in members:
            open_bins = [b for b in bins if capacity is None or total[b] < capacity[b]]
            if not open_bins:
                raise PlanError("no room left on the bench")
            ex = key_extra(p) if key_extra else None
            b = min(open_bins, key=lambda b: (cell_n[c][b], extra_n[ex][b], total[b],
                                              rng.random()))
            out[id(p)] = b
            cell_n[c][b] += 1
            extra_n[ex][b] += 1
            total[b] += 1
    return out


# ======================================================================================
# 1. PLANTING GROUPS
# ======================================================================================
def cohort_offsets(plan: Plan) -> dict[str, int]:
    """{cohort: days its planting follows the group's planting date}. Cohorts whose days
    would collide (one's flood on another's N2O day) are planted LATER, never earlier and
    never by changing their days after planting: the cohort with the earliest first N2O
    day keeps the planting date, each next one takes the smallest delay (0 if none is
    needed) that keeps its flood off the others' N2O days and theirs off its own."""
    flood, n2o, out = set(), set(), {}
    order = sorted(range(len(plan.cohorts)), key=lambda i: (plan.cohorts[i].first_n2o_dap, i))
    for i in order:
        c = plan.cohorts[i]
        off = 0
        while True:
            f = off + c.flood_dap
            m = {off + c.first_n2o_dap + o for o in c.offsets}
            if f not in n2o and not (m & flood):
                break
            off += 1
        out[c.name] = off
        flood.add(f); n2o.update(m)
    return out


def daily_load(plan: Plan, pots: list[Pot], stagger_days: int,
               offsets: dict[str, int]) -> Counter:
    """{days after group 1's planting: N2O measurements that day}."""
    load = Counter()
    for p in pots:
        if p.role != MEASURED:
            continue
        c = plan.cohort(p.cohort)
        for off in c.offsets:
            load[(p.group - 1) * stagger_days + offsets[c.name] + c.first_n2o_dap + off] += 1
    return load


def clash_days(plan: Plan, groups: int, stagger_days: int,
               offsets: dict[str, int]) -> list[int]:
    """Days (after group 1's planting) with a flood AND an N2O measurement: they cannot
    be run together."""
    flood, n2o = set(), set()
    for g in range(groups):
        for c in plan.cohorts:
            base = g * stagger_days + offsets[c.name]
            flood.add(base + c.flood_dap)
            n2o.update(base + c.first_n2o_dap + off for off in c.offsets)
    return sorted(flood & n2o)


def choose_groups(plan: Plan, pots: list[Pot], rng: random.Random, offsets: dict[str, int]
                  ) -> tuple[int, int, Counter, list[str]]:
    """Split the pots into the fewest balanced planting groups, planted `stagger` days
    apart, that keep every day's N2O load within capacity and never put a flood on an
    N2O day. Inside a group the cohorts are planted `offsets` days after the group
    (cohort_offsets), so they never collide with each other. Sets pot.group. Returns
    (groups, stagger_days, load, warnings)."""
    warnings = []
    ok_clash = lambda k, s: not clash_days(plan, k, s, offsets)
    n_measured = sum(p.role == MEASURED for p in pots)
    fixed = plan.planting_groups if plan.planting_groups > 0 else None
    tries = [fixed] if fixed else range(1, max(1, n_measured) + 1)
    state = rng.getstate()
    for k in tries:
        rng.setstate(state)                            # same split whatever k was tried
        assign = _deal(pots, list(range(1, k + 1)), rng)
        for p in pots:
            p.group = assign[id(p)]
        for s in range(plan.min_stagger_days, MAX_STAGGER_DAYS + 1):
            load = daily_load(plan, pots, s, offsets)
            if max(load.values(), default=0) <= plan.max_measured_per_day and ok_clash(k, s):
                return k, s, load, warnings
            if k == 1:
                break                                  # one group: the gap changes nothing
    if fixed:
        s = plan.min_stagger_days
        load = daily_load(plan, pots, s, offsets)      # reported, not hidden
        if max(load.values(), default=0) > plan.max_measured_per_day:
            warnings.append(f"{max(load.values())} N2O measurements fall on one day "
                            f"(capacity {plan.max_measured_per_day}) with {fixed} planting "
                            f"group(s) fixed: use Auto or more groups")
        if clash_days(plan, fixed, s, offsets):
            warnings.append(f"with {fixed} planting group(s) {s} days apart a flood falls on "
                            f"an N2O day: use Auto or a longer gap")
        return fixed, s, load, warnings
    raise PlanError("no planting schedule keeps N2O within the daily capacity")


# ======================================================================================
# 2. BENCH LAYOUT
# ======================================================================================
def place(plan: Plan, pots: list[Pot], rng: random.Random) -> None:
    """Sets pot.table, .zone, .slot."""
    tables = list(range(1, plan.tables + 1))
    per_table = plan.zones_per_table * plan.slots_per_zone
    t_of = _deal(pots, tables, rng, key_extra=lambda p: p.group,
                 capacity={t: per_table for t in tables})
    for p in pots:
        p.table = t_of[id(p)]
    zones = list(range(1, plan.zones_per_table + 1))
    for t in tables:
        on_t = [p for p in pots if p.table == t]
        z_of = _deal(on_t, zones, rng, key_extra=lambda p: p.group,
                     capacity={z: plan.slots_per_zone for z in zones})
        for p in on_t:
            p.zone = z_of[id(p)]
        for z in zones:
            in_z = [p for p in on_t if p.zone == z]
            slots = rng.sample(range(1, plan.slots_per_zone + 1), len(in_z))
            for p, s in zip(in_z, slots):
                p.slot = s


# ======================================================================================
# 3. N2O ORDER
# ======================================================================================
def order_n2o(plan: Plan, pots: list[Pot], rng: random.Random) -> None:
    """Sets load, column and seq on the measured pots of every group x cohort, and seq on
    the sacrificial pots (their harvest order on flood day)."""
    from itertools import permutations
    by_gc = defaultdict(list)
    for p in pots:
        if p.role in (MEASURED, SACRIFICIAL):
            by_gc[(p.group, p.cohort, p.role)].append(p)
    for (g, c, role), members in sorted(by_gc.items()):
        if role == SACRIFICIAL:
            for i, p in enumerate(sorted(members, key=lambda p: (p.genotype, p.treatment, p.rep)), 1):
                p.seq = i
            continue
        # interleave the cells: one pot from each in a fresh random order, round after round,
        # so consecutive pots (one load) are different treatments / genotypes
        cells = defaultdict(list)
        for p in members:
            cells[(p.genotype, p.treatment)].append(p)
        for lst in cells.values():
            rng.shuffle(lst)
        seq = []
        while any(cells.values()):
            keys = [k for k, v in cells.items() if v]
            rng.shuffle(keys)
            keys.sort(key=lambda k: -len(cells[k]))     # the fullest first: no tail of one cell
            if seq:                                     # never the same cell twice running
                last = (seq[-1].genotype, seq[-1].treatment)
                if len(keys) > 1 and keys[0] == last:
                    keys[0], keys[1] = keys[1], keys[0]
            seq.extend(cells[k].pop() for k in keys)
        # column per pot: the permutation that keeps each cell spread over the columns
        col_n = defaultdict(Counter)
        for i in range(0, len(seq), N2O_COLUMNS):
            load = seq[i:i + N2O_COLUMNS]
            cols = range(1, N2O_COLUMNS + 1)
            best = min(permutations(cols, len(load)),
                       key=lambda perm: (sum(col_n[(p.genotype, p.treatment)][col]
                                             for p, col in zip(load, perm)), rng.random()))
            for p, col in zip(load, best):
                p.load, p.column = i // N2O_COLUMNS + 1, col
                col_n[(p.genotype, p.treatment)][col] += 1
        # seq follows load, then column (the order the chambers are filled)
        for k, p in enumerate(sorted(seq, key=lambda p: (p.load, p.column)), 1):
            p.seq = k


# ======================================================================================
# 4. WATERING
# ======================================================================================
def water_days(planted: dt.date, until: dt.date,
               weekdays: tuple[int, ...] = WATER_WEEKDAYS) -> list[dt.date]:
    """Every Monday, Wednesday and Friday after planting, strictly before `until` (the
    flood). Planting day itself is not counted."""
    out = []
    day = planted + dt.timedelta(days=1)
    while day < until:
        if day.weekday() in weekdays:
            out.append(day)
        day += dt.timedelta(days=1)
    return out


def workday(day: dt.date, n: int) -> dt.date:
    """`n` working days (Mon-Fri) after `day` (before, if negative). n = 0 keeps a working
    day and moves a weekend day to the next Monday."""
    if n == 0:
        while day.weekday() >= 5:
            day += dt.timedelta(days=1)
        return day
    step = dt.timedelta(days=1 if n > 0 else -1)
    for _ in range(abs(n)):
        day += step
        while day.weekday() >= 5:
            day += step
    return day


@dataclass
class Swap:
    """A failed pot replaced by a spare of its own genotype and planting group (swaps.py)."""
    date: dt.date
    failed: str                    # barcode
    spare: str                     # barcode
    role: str = ""                 # the label the spare takes over, e.g. B73-C-CTRL-r04
    note: str = ""


@dataclass
class Task:
    """A dated piece of work. The date is `lag` working days from `anchor` (lag None: the
    anchor day itself, weekend or not), so the workbook can write it as a formula on the
    timeline. Anchors: ("plant", g, cohort), ("treat", g, cohort), ("flood", g, cohort),
    ("n2o", g, cohort, offset), ("water", n) = the n-th watering day of the experiment."""
    date: dt.date
    name: str
    section: str = ""      # Asana section: what kind of work
    notes: str = ""
    anchor: tuple = ()
    lag: int | None = None


# ======================================================================================
# 6. THE DESIGN
# ======================================================================================
@dataclass
class Design:
    plan: Plan
    pots: list[Pot]
    groups: int
    stagger_days: int
    load: Counter                  # {DAP of group 1: N2O measurements}
    warnings: list[str] = field(default_factory=list)
    offsets: dict[str, int] = field(default_factory=dict)   # cohort -> planting delay, days

    def group_planting(self, g: int) -> dt.date:
        """Group g's planting date: its first cohort (and its spares) are planted then."""
        return self.plan.planting_date + dt.timedelta(days=(g - 1) * self.stagger_days)

    def planting(self, g: int, cohort: str) -> dt.date:
        """When group g's pots of `cohort` are planted (the group date + the cohort's
        offset). Spares (no cohort) are planted on the group date."""
        return self.group_planting(g) + dt.timedelta(days=self.offsets.get(cohort, 0))

    def planting_days(self, g: int) -> list[tuple[dt.date, list[str]]]:
        """Group g's planting days in order, with the cohorts planted on each."""
        days = defaultdict(list)
        for c in self.plan.cohorts:
            days[self.planting(g, c.name)].append(c.name)
        return sorted(days.items())

    def pot_planting(self, p: Pot) -> dt.date:
        return self.planting(p.group, p.cohort)

    def events(self, g: int, cohort: Cohort) -> list[tuple[dt.date, str]]:
        """What happens to group g, cohort `cohort`, in date order."""
        p0 = self.planting(g, cohort.name)
        day = lambda dap: p0 + dt.timedelta(days=dap)
        has_sac = any(p.role == SACRIFICIAL and p.group == g and p.cohort == cohort.name
                      for p in self.pots)
        ev = [(p0, "Plant")]
        if self.plan.treatment_dap is not None:
            ev.append((day(self.plan.treatment_dap), "Treatment"))
        ev.append((day(cohort.flood_dap),
                   "Sacrificial harvest, flood" if has_sac else "Flood"))
        offs = cohort.offsets
        for h, off in zip(cohort.measure_hours, offs):
            last = off == offs[-1]
            ev.append((day(cohort.first_n2o_dap + off),
                       f"{h} h N2O" + (", harvest" if last else "")))
        return ev

    def water(self, g: int, cohort: Cohort) -> list[dt.date]:
        p0 = self.planting(g, cohort.name)
        return water_days(p0, p0 + dt.timedelta(days=cohort.flood_dap))

    def day_text(self, g: int, cohort: Cohort, day: dt.date) -> str:
        """What the timeline cell for group g, cohort `cohort` says on `day`."""
        ev = dict((d, t) for d, t in self.events(g, cohort))
        watered = day in set(self.water(g, cohort))
        if day in ev:
            t = ev[day]
            return f"{t}; water" if t == "Treatment" and watered else t
        return "Water" if watered else ""

    def anchor_date(self, anchor: tuple) -> dt.date:
        kind, *a = anchor
        if kind == "water":
            return self.water_days_all()[a[0] - 1]
        if kind == "last_harvest":
            return max(d for d, _, _ in self.harvests())
        c = self.plan.cohort(a[1])
        p0 = self.planting(a[0], c.name)
        if kind == "plant":
            return p0
        if kind == "treat":
            return p0 + dt.timedelta(days=self.plan.treatment_dap)
        if kind == "flood":
            return p0 + dt.timedelta(days=c.flood_dap)
        return p0 + dt.timedelta(days=c.first_n2o_dap + a[2])          # "n2o"

    def harvests(self) -> list[tuple[dt.date, str, int]]:
        """Every harvest with soil samples, in date order: (date, 'A sacrificial' /
        'A final' (+ ', group g'), samples). Spares are not harvested."""
        many = self.groups > 1
        out = []
        for g in range(1, self.groups + 1):
            for c in self.plan.cohorts:
                tag = f"{c.name}" + (f" G{g}" if many else "")
                for role, kind, anchor in (
                        (SACRIFICIAL, "sacrificial", ("flood", g, c.name)),
                        (MEASURED, "final", ("n2o", g, c.name, c.last_offset))):
                    n = sum(p.role == role and p.group == g and p.cohort == c.name
                            for p in self.pots)
                    if n:
                        out.append((self.anchor_date(anchor), f"{tag} {kind}", n))
        return sorted(out, key=lambda h: h[0])

    def kcl_batches(self) -> list[list[tuple[str, int]]]:
        """The soil samples split into the fewest KCl extractions of at most KCL_BATCH_MAX,
        as equal as they divide (117 -> 59 + 58), filled in harvest order. Each batch is
        [(harvest, samples)], a harvest split over two batches where the sizes need it."""
        hv = [(name, n) for _, name, n in self.harvests()]
        total = sum(n for _, n in hv)
        if not total:
            return []
        b = -(-total // KCL_BATCH_MAX)
        sizes = [total // b + (1 if i < total % b else 0) for i in range(b)]
        out, queue = [], list(hv)
        for size in sizes:
            batch = []
            while size:
                name, n = queue.pop(0)
                take = min(n, size)
                batch.append((name, take))
                size -= take
                if n > take:
                    queue.insert(0, (name, n - take))
            out.append(batch)
        return out

    def water_days_all(self) -> list[dt.date]:
        """Every day anything is watered, any group or cohort."""
        return sorted({d for g in range(1, self.groups + 1) for c in self.plan.cohorts
                       for d in self.water(g, c)})

    def water_who(self, day: dt.date) -> str:
        """Who is watered on `day`: 'cohorts A, B' (group 1 only) or 'G1 A, G1 B, G2 A'."""
        many = self.groups > 1
        who = [f"G{g} {c.name}" if many else c.name for g in range(1, self.groups + 1)
               for c in self.plan.cohorts if day in self.water(g, c)]
        return ("" if many else "cohorts ") + ", ".join(who)

    def tasks(self) -> list[Task]:
        """Every task with its own date, in date order (see TASK_LAGS)."""
        plan, out = self.plan, []
        lag = TASK_LAGS
        many = self.groups > 1
        who = lambda g, c: f"cohort {c}" + (f", group {g}" if many else "")

        def add(anchor, n, name, section, notes=""):
            day = self.anchor_date(anchor)
            day = day if n is None else workday(day, n)
            out.append(Task(day, name, section, notes, anchor, n))

        for g in range(1, self.groups + 1):
            days = self.planting_days(g)
            for day, cs in days:
                # one set of planting tasks per planting day: the cohorts are planted
                # apart when their days would collide (cohort_offsets)
                tag = ([f"group {g}"] if many else []) \
                    + ([f"cohort{'s' if len(cs) > 1 else ''} {', '.join(cs)}"] if len(days) > 1 else [])
                sfx = f" ({', '.join(tag)})" if tag else ""
                n = sum(p.group == g and (p.cohort in cs or (p.role == SPARE and day == days[0][0]))
                        for p in self.pots)
                at = ("plant", g, cs[0])
                add(at, lag["prefill"], f"Pre-fill pots with weighed soil{sfx}", "Setup",
                    f"{n} pots")
                add(at, None, f"Label, plant seed, and initial wet-up of pots{sfx}",
                    "Planting", f"{n} pots")
                add(at, None, f"Label tables with treatment and group information per pot"
                    f"{sfx}", "Planting", "positions and tags: design map, tags sheet")
                # the 7 days after planting always hold 5 weekdays: working days 1 to 5
                for k in range(1, GERMINATION_WORKDAYS + 1):
                    add(at, k, f"If using ZeaL genotypes, check for germination{sfx}",
                        "Planting", f"check {k} of {GERMINATION_WORKDAYS}")
                if plan.treatment_dap is not None:
                    add(("treat", g, cs[0]), None, f"Apply treatments{sfx}", "Treatment")
            for c in plan.cohorts:
                w = who(g, c.name)
                sac = sum(p.role == SACRIFICIAL and p.group == g and p.cohort == c.name
                          for p in self.pots)
                meas = sum(p.role == MEASURED and p.group == g and p.cohort == c.name
                           for p in self.pots)
                last = ("n2o", g, c.name, c.last_offset)
                harvests = [(("flood", g, c.name), f"sacrificial harvest, {sac} pots")] if sac else []
                harvests.append((last, f"final harvest, {meas} pots"))
                add(harvests[0][0], lag["kcl_setup"],
                    f"KCl soil collection setup (gathering supplies, printing & labeling tubes)"
                    f" - {w}", "KCl", f"{sac + meas} tubes"
                    + (" (sacrificial and final harvests)" if sac else ""))
                add(("flood", g, c.name), None, f"Flood - {w}", "N2O",
                    (f"harvest the {sac} sacrificial pots first; " if sac else "")
                    + f"flood {meas} pots in the flood sheet's order")
                for h, off in zip(c.measure_hours, c.offsets):
                    add(("n2o", g, c.name, off), None, f"N2O measurements, {h} h - {w}", "N2O",
                        f"{meas} pots")
                for anchor, what in harvests:
                    add(anchor, None, f"Plant harvest and soil sample collection - {w}",
                        "Harvest", what)
                    add(anchor, lag["dry_weights"], f"Plant dry weights measure & record - {w}",
                        "Harvest", what)
                add(last, lag["wash"], f"Wash pots for reuse - {w}", "Cleanup",
                    f"{meas + sac} pots")
        # KCl: the soil samples wait until the last harvest, then go in a few equal batches
        # on consecutive working days, each read on the plate reader the day after
        batches = self.kcl_batches()
        for k, parts in enumerate(batches, 1):
            n = sum(x for _, x in parts)
            of = f" (batch {k} of {len(batches)})" if len(batches) > 1 else ""
            what = f"{n} samples: " + ", ".join(f"{name} {x}" for name, x in parts)
            add(("last_harvest",), lag["kcl_extract"] + k - 1, f"KCl extractions{of}", "KCl",
                what)
            add(("last_harvest",), lag["kcl_extract"] + k - 1 + lag["kcl_analysis"],
                f"KCl analysis (plate reader){of}", "KCl", what)
        # watering: one task per day, whatever is still being watered that day
        for k, day in enumerate(self.water_days_all(), 1):
            add(("water", k), None, "Maintenance watering", "Watering", self.water_who(day))
        order = {"Setup": 0, "Planting": 1, "Watering": 2, "Treatment": 3, "KCl": 4,
                 "Harvest": 5, "N2O": 6, "Cleanup": 7}
        out.sort(key=lambda t: (t.date, order.get(t.section, 9), t.name))
        return out

    def calendar_load(self) -> dict[dt.date, int]:
        return {self.plan.planting_date + dt.timedelta(days=d): n
                for d, n in sorted(self.load.items())}

    @property
    def max_load(self) -> int:
        return max(self.load.values(), default=0)

    def counts(self) -> Counter:
        return Counter(p.role for p in self.pots)


def build(plan: Plan) -> Design:
    problems = validate(plan)
    if problems:
        raise PlanError("; ".join(problems))
    rng = random.Random(plan.seed)
    pots = make_pots(plan)
    offsets = cohort_offsets(plan)
    groups, stagger, load, warnings = choose_groups(plan, pots, rng, offsets)
    add_spares(plan, pots, groups)
    if len(pots) > plan.capacity:
        n_sp = len(pots) - count_pots(plan)
        raise PlanError(
            f"{len(pots)} pots do not fit: {count_pots(plan)} planned + {n_sp} spares "
            f"({plan.spares_per_genotype_per_group} x {len(genotypes(plan))} genotypes x "
            f"{groups} planting groups); the bench holds {plan.capacity}. Lower the spares "
            f"or add tables, zones or slots.")
    place(plan, pots, rng)
    order_n2o(plan, pots, rng)
    # barcodes in random order, so a barcode says nothing about the pot's treatment
    numbers = rng.sample(range(1, len(pots) + 1), len(pots))
    width = max(3, len(str(len(pots))))
    for p, n in zip(pots, numbers):
        p.barcode = f"{plan.name.strip()}-S{n:0{width}d}"
    pots.sort(key=lambda p: (p.table, p.zone, p.slot))
    return Design(plan, pots, groups, stagger, load, warnings, offsets)
