#!/usr/bin/env python3
"""
design_core.py: BNI experiment design (v0.2, 2026-10-28). No Qt, no Excel.

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
  3. SCHEDULE. Per group and cohort: plant (DAP 0), treatment (treatment_dap), flood and
     sacrificial harvest (first N2O DAP - 1), N2O at each measure offset (24 h, 48 h),
     harvest after the last one. Daily N2O load is counted per calendar day. Watering is
     Monday, Wednesday and Friday, after planting until the flood. Treatment, flood, N2O
     and harvest may fall on any day, weekends included.
  4. N2O ORDER. Per group and cohort the measured pots are put in loads of three (one
     pot per N2O column), each load mixing treatments and genotypes, with every cell
     spread over the three columns. The flood is done in the same order, so every pot
     has the same time between its flood and its measurement.

The same `seed` always gives the same design. Nothing here knows about Excel; the
workbook (design_xlsx.py) writes what this computes, plus formulas so the dates can be
edited afterwards.
"""
from __future__ import annotations

import datetime as dt
import random
from collections import Counter, defaultdict
from dataclasses import dataclass, field, asdict

__version__ = "0.2"

MAX_ZONES = 8              # a table holds at most 8 zones x 25 slots
MAX_SLOTS = 25
MEASURE_HOURS = (24, 48)   # N2O after the flood; the last one is followed by harvest
MAX_STAGGER_DAYS = 30      # search limit for the gap between planting groups
DEFAULT_STAGGER_DAYS = 3   # 24 h N2O, 48 h N2O, then the next group's flood
WATER_WEEKDAYS = (0, 2, 4)  # Monday, Wednesday, Friday (date.weekday())
WATER_DAY_NAMES = "Mon, Wed, Fri"
N2O_COLUMNS = 3            # pots per N2O load (one per chamber)

MEASURED, SACRIFICIAL, SPARE = "MEASURED", "SACRIFICIAL", "SPARE"


# ======================================================================================
# THE PLAN: what the user types in the first window
# ======================================================================================
@dataclass
class Cohort:
    name: str
    first_n2o_dap: int             # day after planting of the first (24 h) N2O sample


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
    measure_hours: tuple[int, ...] = MEASURE_HOURS
    cohorts: list[Cohort] = field(default_factory=list)
    rows: list[PlanRow] = field(default_factory=list)

    # ---- save / load (the app's plan file) ----
    def to_dict(self) -> dict:
        d = asdict(self)
        d["planting_date"] = self.planting_date.isoformat()
        d["measure_hours"] = list(self.measure_hours)
        return d

    @classmethod
    def from_dict(cls, d: dict) -> "Plan":
        d = dict(d)
        d["planting_date"] = dt.date.fromisoformat(d["planting_date"])
        d["measure_hours"] = tuple(d.get("measure_hours", MEASURE_HOURS))
        d["cohorts"] = [Cohort(**c) for c in d.get("cohorts", [])]
        d["rows"] = [PlanRow(**r) for r in d.get("rows", [])]
        if "spares_per_genotype" in d:                 # v0.1/v0.2 plan files
            d.setdefault("spares_per_genotype_per_group", d.pop("spares_per_genotype"))
        known = set(cls.__dataclass_fields__)
        return cls(**{k: v for k, v in d.items() if k in known})

    @property
    def capacity(self) -> int:
        return self.tables * self.zones_per_table * self.slots_per_zone

    @property
    def measure_offsets(self) -> tuple[int, ...]:
        """Days after the first N2O DAP: 24 h -> 0, 48 h -> 1."""
        return tuple(h // 24 - 1 for h in self.measure_hours)

    def cohort(self, name: str) -> Cohort | None:
        return next((c for c in self.cohorts if c.name == name), None)


class PlanError(ValueError):
    """The plan cannot be built; the message says what to fix."""


def validate(plan: Plan) -> list[str]:
    """Problems that stop a build (empty list = OK)."""
    p = []
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
    if plan.min_stagger_days < len(plan.measure_hours):
        p.append(f"days between planting groups must be at least {len(plan.measure_hours)} "
                 f"so one group's N2O days never overlap the next group's")
    if any(h % 24 or h < 24 for h in plan.measure_hours):
        p.append("N2O sampling times must be whole days after the flood (24 h, 48 h, ...)")
    names = [c.name for c in plan.cohorts]
    if not names:
        p.append("add at least one cohort")
    if len(set(names)) != len(names):
        p.append("cohort names must be different")
    for c in plan.cohorts:
        if not c.name.strip():
            p.append("a cohort has no name")
        if c.first_n2o_dap < 2:
            p.append(f"cohort {c.name}: first N2O must be at least 2 days after planting "
                     f"(the flood is the day before)")
        if plan.treatment_dap is not None and plan.treatment_dap >= c.first_n2o_dap - 1:
            p.append(f"cohort {c.name}: treatment (DAP {plan.treatment_dap}) must come before "
                     f"the flood (DAP {c.first_n2o_dap - 1})")
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
        if self.role == SPARE:
            g = f"G{self.group}-" if self.group_in_tag else ""
            return f"{self.genotype}-{g}SPARE-{self.rep:02d}"
        parts = [self.genotype, self.cohort] + ([self.treatment] if self.treatment else [])
        rep = f"r{self.rep:02d}" if self.role == MEASURED else f"SAC{self.rep}"
        return "-".join(parts + [rep])


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
def daily_load(plan: Plan, pots: list[Pot], stagger_days: int) -> Counter:
    """{days after group 1's planting: N2O measurements that day}."""
    load = Counter()
    for p in pots:
        if p.role != MEASURED:
            continue
        c = plan.cohort(p.cohort)
        for off in plan.measure_offsets:
            load[(p.group - 1) * stagger_days + c.first_n2o_dap + off] += 1
    return load


def clash_days(plan: Plan, groups: int, stagger_days: int) -> list[int]:
    """Days (after group 1's planting) with a flood AND an N2O measurement: they cannot
    be run together."""
    flood, n2o = set(), set()
    for g in range(groups):
        for c in plan.cohorts:
            base = g * stagger_days + c.first_n2o_dap
            flood.add(base - 1)
            n2o.update(base + off for off in plan.measure_offsets)
    return sorted(flood & n2o)


def choose_groups(plan: Plan, pots: list[Pot], rng: random.Random
                  ) -> tuple[int, int, Counter, list[str]]:
    """Split the pots into the fewest balanced planting groups, planted `stagger` days
    apart, that keep every day's N2O load within capacity and never put a flood on an
    N2O day. Sets pot.group. Returns (groups, stagger_days, load, warnings)."""
    warnings = []
    inherent = clash_days(plan, 1, 0)                  # the cohorts alone already clash
    if inherent:
        warnings.append(
            f"cohorts' days collide inside one planting group: a flood falls on an N2O day "
            f"(DAP {', '.join(map(str, inherent))}). Space the cohorts' first N2O days at "
            f"least {len(plan.measure_hours) + 1} days apart.")
    ok_clash = lambda k, s: inherent or not clash_days(plan, k, s)
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
            load = daily_load(plan, pots, s)
            if max(load.values(), default=0) <= plan.max_measured_per_day and ok_clash(k, s):
                return k, s, load, warnings
            if k == 1:
                break                                  # one group: the gap changes nothing
    if fixed:
        s = plan.min_stagger_days
        load = daily_load(plan, pots, s)               # reported, not hidden
        if max(load.values(), default=0) > plan.max_measured_per_day:
            warnings.append(f"{max(load.values())} N2O measurements fall on one day "
                            f"(capacity {plan.max_measured_per_day}) with {fixed} planting "
                            f"group(s) fixed: use Auto or more groups")
        if not inherent and clash_days(plan, fixed, s):
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


# ======================================================================================
# 5. THE DESIGN
# ======================================================================================
@dataclass
class Design:
    plan: Plan
    pots: list[Pot]
    groups: int
    stagger_days: int
    load: Counter                  # {DAP of group 1: N2O measurements}
    warnings: list[str] = field(default_factory=list)

    def group_planting(self, g: int) -> dt.date:
        return self.plan.planting_date + dt.timedelta(days=(g - 1) * self.stagger_days)

    def events(self, g: int, cohort: Cohort) -> list[tuple[dt.date, str]]:
        """What happens to group g, cohort `cohort`, in date order."""
        p0 = self.group_planting(g)
        day = lambda dap: p0 + dt.timedelta(days=dap)
        has_sac = any(p.role == SACRIFICIAL and p.group == g and p.cohort == cohort.name
                      for p in self.pots)
        ev = [(p0, "Plant")]
        if self.plan.treatment_dap is not None:
            ev.append((day(self.plan.treatment_dap), "Treatment"))
        ev.append((day(cohort.first_n2o_dap - 1),
                   "Sacrificial harvest, flood" if has_sac else "Flood"))
        offs = self.plan.measure_offsets
        for h, off in zip(self.plan.measure_hours, offs):
            last = off == offs[-1]
            ev.append((day(cohort.first_n2o_dap + off),
                       f"{h} h N2O" + (", harvest" if last else "")))
        return ev

    def water(self, g: int, cohort: Cohort) -> list[dt.date]:
        p0 = self.group_planting(g)
        return water_days(p0, p0 + dt.timedelta(days=cohort.first_n2o_dap - 1))

    def day_text(self, g: int, cohort: Cohort, day: dt.date) -> str:
        """What the timeline cell for group g, cohort `cohort` says on `day`."""
        ev = dict((d, t) for d, t in self.events(g, cohort))
        watered = day in set(self.water(g, cohort))
        if day in ev:
            t = ev[day]
            return f"{t}; water" if t == "Treatment" and watered else t
        return "Water" if watered else ""

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
    groups, stagger, load, warnings = choose_groups(plan, pots, rng)
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
    return Design(plan, pots, groups, stagger, load, warnings)
