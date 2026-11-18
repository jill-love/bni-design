"""Experiment design: groups, layout, schedule, workbook, window. No hardware.
Run:  python -m pytest -q tests"""
import datetime as dt
import json
import os
import sys
from collections import Counter

import pytest

HERE = os.path.dirname(os.path.abspath(__file__)); R = os.path.dirname(HERE)
sys.path.insert(0, R)
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
import design_core as dc
import design_xlsx as dx

MON = dt.date(2026, 11, 2)


def p0_plan(**kw):
    """The P0 shape: one genotype, 3 cohorts 3 days apart, 3 treatments, 10 + 3 SAC."""
    base = dict(name="B73-P0", seed=7, planting_date=MON, treatment_dap=14,
                cohorts=[dc.Cohort("A", 21), dc.Cohort("B", 24), dc.Cohort("C", 27)],
                rows=[dc.PlanRow("B73", c, t, 10, 3) for c in "ABC" for t in ("CTRL", "SN", "INH")],
                spares_per_genotype_per_group=8)
    base.update(kw)
    return dc.Plan(**base)


def big_plan(**kw):
    """3 genotypes x 2 cohorts x 3 treatments x 12 reps + 2 SAC = 252 pots, 216 measured."""
    base = dict(name="BNI-P1", seed=3, planting_date=MON, treatment_dap=14,
                zones_per_table=5, slots_per_zone=12,
                cohorts=[dc.Cohort("A", 21), dc.Cohort("B", 24)],
                rows=[dc.PlanRow(g, c, t, 12, 2) for g in ("B73", "Mo17", "W22")
                      for c in "AB" for t in ("CTRL", "SN", "INH")])
    base.update(kw)
    return dc.Plan(**base)


# ---------------------------------------------------------------- pots
def test_pot_counts_and_labels():
    d = dc.build(p0_plan())
    c = d.counts()
    assert (c[dc.MEASURED], c[dc.SACRIFICIAL], c[dc.SPARE]) == (90, 27, 8)
    labels = {p.label for p in d.pots}
    assert "B73-A-CTRL-r01" in labels and "B73-C-INH-SAC3" in labels and "B73-SPARE-08" in labels
    assert len({p.barcode for p in d.pots}) == len(d.pots)
    assert all(p.barcode.startswith("B73-P0-S") for p in d.pots)


def test_spares_per_genotype_per_planting_group():
    d = dc.build(big_plan(spares_per_genotype_per_group=2))
    assert d.groups == 3                                    # spares never change the groups
    spares = Counter((p.genotype, p.group) for p in d.pots if p.role == dc.SPARE)
    assert spares == {(g, k): 2 for g in ("B73", "Mo17", "W22") for k in (1, 2, 3)}
    labels = {p.label for p in d.pots if p.role == dc.SPARE}
    assert "Mo17-G3-SPARE-02" in labels and len(labels) == 18
    one = dc.build(p0_plan())                               # one group: no G in the tag
    assert {p.label for p in one.pots if p.role == dc.SPARE} == {
        f"B73-SPARE-{k:02d}" for k in range(1, 9)}


def test_spares_that_do_not_fit_say_why():
    with pytest.raises(dc.PlanError, match="spares"):
        dc.build(big_plan(spares_per_genotype_per_group=20))   # 252 + 180 > 360


def test_old_plan_files_still_open():
    d = big_plan().to_dict()
    d.pop("spares_per_genotype_per_group"); d["spares_per_genotype"] = 4
    assert dc.Plan.from_dict(d).spares_per_genotype_per_group == 4


def test_barcodes_do_not_follow_the_treatment_order():
    d = dc.build(p0_plan())
    by_label = sorted(d.pots, key=lambda p: (p.cohort, p.treatment, p.role, p.rep))
    nums = [int(p.barcode.rsplit("S", 1)[1]) for p in by_label]
    assert nums != sorted(nums)


def test_same_seed_same_design_other_seed_other_design():
    a = [(p.barcode, p.label, p.position, p.group) for p in dc.build(big_plan()).pots]
    b = [(p.barcode, p.label, p.position, p.group) for p in dc.build(big_plan()).pots]
    c = [(p.barcode, p.label, p.position, p.group) for p in dc.build(big_plan(seed=4)).pots]
    assert a == b and a != c


def test_blank_treatment_is_allowed():
    plan = p0_plan(rows=[dc.PlanRow("B73", "A", "", 6), dc.PlanRow("W22", "A", "", 6)],
                   spares_per_genotype_per_group=0)
    d = dc.build(plan)
    assert {p.label for p in d.pots} >= {"B73-A-r01", "W22-A-r06"}


# ---------------------------------------------------------------- planting groups
def test_small_experiment_is_one_group():
    d = dc.build(p0_plan())
    assert d.groups == 1 and d.max_load == 30


def test_fewest_equal_groups_that_fit_capacity():
    d = dc.build(big_plan())
    # 2 groups would put 54 pots of one cohort on one day (> 50); 3 groups fit
    assert d.groups == 3 and d.max_load <= 50
    sizes = Counter(p.group for p in d.pots).values()
    assert max(sizes) - min(sizes) <= 1
    # every cell is split as evenly as it divides
    for cell, n in Counter(p.cell for p in d.pots).items():
        per = Counter(p.group for p in d.pots if p.cell == cell)
        assert max(per.values()) - min(per.get(g, 0) for g in range(1, 4)) <= 1, cell


def test_default_gap_is_3_days():
    assert dc.Plan().min_stagger_days == 3
    d = dc.build(big_plan(cohorts=[dc.Cohort("A", 21)], rows=[
        dc.PlanRow(g, "A", t, 12) for g in ("B73", "Mo17", "W22") for t in ("CTRL", "SN")]))
    assert d.groups == 2 and d.stagger_days == 3            # 72 measured: 2 x 36


def test_gap_widens_when_cohorts_would_collide():
    d = dc.build(big_plan())
    # cohorts are 3 days apart: a 3-day gap stacks group 2's A on group 1's B (72 > 50),
    # 4 and 5 put a flood on an N2O day, 6 is the first that works
    assert d.stagger_days == 6 and d.offsets == {"A": 0, "B": 0}
    assert dc.daily_load(d.plan, d.pots, 3, d.offsets)[24] > 50
    assert dc.clash_days(d.plan, 3, 4, d.offsets) and dc.clash_days(d.plan, 3, 5, d.offsets)
    assert not dc.clash_days(d.plan, 3, 6, d.offsets) and not d.warnings


def test_never_a_flood_on_an_n2o_day():
    for plan in (p0_plan(), big_plan(), big_plan(max_measured_per_day=20)):
        flood, n2o = _flood_n2o_days(dc.build(plan))
        assert not flood & n2o


def _flood_n2o_days(d):
    flood, n2o = set(), set()
    for g, c in d.cells():
        for day, text in d.events(g, c):
            if "flood" in text.lower():
                flood.add(day)
            if "N2O" in text:
                n2o.add(day)
    return flood, n2o


def test_cohorts_that_collide_are_planted_later():
    """B's flood (DAP 22) would fall on A's 48 h N2O: B is planted one day later, as its
    own planting group (it floods the day after A's 48 h), keeping its days after
    planting; nothing is planted before the planting date, nothing is warned."""
    d = dc.build(p0_plan(cohorts=[dc.Cohort("A", 21), dc.Cohort("B", 23)],
                         rows=[dc.PlanRow("B73", c, "CTRL", 5) for c in "AB"]))
    assert d.offsets == {"A": 0, "B": 1} and not d.warnings
    assert d.groups == 2 and [c.name for c in d.cohorts_of(1)] == ["A"]
    assert [c.name for c in d.cohorts_of(2)] == ["B"]
    b = d.plan.cohort("B")
    assert d.group_planting(1) == MON and d.group_planting(2) == MON + dt.timedelta(days=1)
    assert d.events(2, b)[0] == (d.group_planting(2), "Plant")
    flood_b = d.group_planting(2) + dt.timedelta(days=b.first_n2o_dap - 1)
    assert flood_b in [day for day, t in d.events(2, b) if "lood" in t]
    flood, n2o = _flood_n2o_days(d)
    assert not flood & n2o
    assert all(p.cohort in ("", "B") for p in d.pots if p.group == 2)
    assert all(d.pot_planting(p) >= MON for p in d.pots)


@pytest.mark.parametrize("days", [(21, 22), (21, 21, 22), (21, 23, 22), (20, 21, 22, 23)])
def test_offsets_are_never_earlier_and_never_collide(days):
    plan = p0_plan(cohorts=[dc.Cohort(chr(65 + i), n) for i, n in enumerate(days)],
                   rows=[dc.PlanRow("B73", chr(65 + i), "CTRL", 3) for i in range(len(days))])
    off = dc.cohort_offsets(plan)
    assert all(v >= 0 for v in off.values())
    assert off[min(plan.cohorts, key=lambda c: c.first_n2o_dap).name] == 0
    assert not dc.clash_days(plan, 1, 0, off)
    d = dc.build(plan)
    for g, c in d.cells():                                  # days after planting unchanged
        p0 = d.group_planting(g)
        assert d.events(g, c)[2][0] == p0 + dt.timedelta(days=c.first_n2o_dap - 1)


def test_cohorts_that_do_not_collide_share_the_planting_day():
    assert dc.cohort_offsets(p0_plan()) == {"A": 0, "B": 0, "C": 0}
    t = dc.build(p0_plan()).tasks()
    assert len(_named(t, "Label, plant seed")) == 1


def test_planting_tasks_per_planting_group():
    d = dc.build(p0_plan(cohorts=[dc.Cohort("A", 21), dc.Cohort("B", 22), dc.Cohort("C", 27)],
                         rows=[dc.PlanRow("B73", c, "CTRL", 4, 1) for c in "ABC"],
                         spares_per_genotype_per_group=2))
    assert d.offsets == {"A": 0, "B": 2, "C": 0} and d.groups == 2
    plant = _named(d.tasks(), "Label, plant seed")
    assert [(x.date, x.name) for x in plant] == [
        (MON, "Label, plant seed, and initial wet-up of pots (group 1)"),
        (MON + dt.timedelta(days=2), "Label, plant seed, and initial wet-up of pots (group 2)")]
    assert [x.notes for x in plant] == ["12 pots", "7 pots"]   # each group has its spares
    assert len(_named(d.tasks(), "If using ZeaL")) == 10

def test_lower_capacity_means_more_groups():
    assert dc.build(big_plan(max_measured_per_day=20)).groups > dc.build(big_plan()).groups


def test_fixed_group_count_over_capacity_is_reported():
    d = dc.build(big_plan(planting_groups=1))
    assert d.groups == 1 and d.max_load > 50 and d.warnings


# ---------------------------------------------------------------- layout
def test_only_the_tables_available_are_used():
    plan = big_plan(tables=4, slots_per_zone=13)            # 4 x 5 x 13 = 260 places
    d = dc.build(plan)
    assert {p.table for p in d.pots} == {1, 2, 3, 4}
    assert max(Counter(p.table for p in d.pots).values()) <= 65


def test_positions_are_unique_and_inside_the_bench():
    plan = big_plan()
    d = dc.build(plan)
    pos = [(p.table, p.zone, p.slot) for p in d.pots]
    assert len(set(pos)) == len(pos)
    assert all(1 <= t <= 6 and 1 <= z <= 5 and 1 <= s <= 12 for t, z, s in pos)
    assert all(p.position == f"T{p.table}-Z{p.zone}-S{p.slot}" for p in d.pots)


def test_every_cell_and_group_is_spread_evenly_over_tables():
    d = dc.build(big_plan())
    tables = range(1, 7)
    tot = Counter(p.table for p in d.pots)
    assert max(tot.values()) - min(tot.values()) <= 1
    for cell in {p.cell for p in d.pots}:
        per = Counter(p.table for p in d.pots if p.cell == cell)
        assert max(per.get(t, 0) for t in tables) - min(per.get(t, 0) for t in tables) <= 1, cell
    for g in range(1, d.groups + 1):
        per = Counter(p.table for p in d.pots if p.group == g)
        assert max(per.values()) - min(per.get(t, 0) for t in tables) <= 2, g


def p1_plan(**kw):
    """The P1 shape: one genotype, 3 cohorts flooded 2, 4 and 6 days after one treatment
    day, 3 treatments x 10 reps + 4 SAC, 12 % spares: 144 pots, a full 6 x 3 x 8 bench,
    in 4 planting groups (each cohort flooded twice)."""
    base = dict(name="BNI-P1", seed=39937, planting_date=MON, treatment_dap=14,
                cohorts=[dc.Cohort(c, n, (24, 48, 72)) for c, n in (("A", 17), ("B", 19), ("C", 21))],
                rows=[dc.PlanRow("B73", c, t, 10, 4) for c in "ABC" for t in ("CTRL", "SN", "INH")],
                spare_percent=12)
    base.update(kw)
    return dc.Plan(**base)


@pytest.mark.parametrize("seed", [1, 2, 3, 39937])
def test_no_flood_has_a_treatment_bunched_on_a_table(seed):
    """A flood (planting group x cohort) puts each treatment's pots on different tables
    (5 pots on 6 tables: 5 tables), and each treatment and planting group is spread
    evenly over the tables."""
    for plan in (p1_plan(seed=seed), big_plan(seed=seed), p0_plan(seed=seed)):
        d = dc.build(plan)
        n_tab = plan.tables
        for g, c in d.cells():
            for role in (dc.MEASURED, dc.SACRIFICIAL):
                for t in {p.treatment for p in d.pots}:
                    mine = [p for p in d.pots if p.group == g and p.cohort == c.name
                            and p.role == role and p.treatment == t]
                    for geno in {p.genotype for p in mine}:
                        same = [p.table for p in mine if p.genotype == geno]
                        assert max(Counter(same).values(), default=0) <= -(-len(same) // n_tab), \
                            (g, c.name, role, t, geno, sorted(same))
        if plan.name == "BNI-P1":
            for t in ("CTRL", "SN", "INH"):
                per = Counter(p.table for p in d.pots if p.treatment == t and p.role == dc.MEASURED)
                assert max(per.values()) - min(per.get(k, 0) for k in range(1, n_tab + 1)) <= 1, t
            for g in range(1, d.groups + 1):
                per = Counter(p.table for p in d.pots if p.group == g)
                assert max(per.values()) - min(per.get(k, 0) for k in range(1, n_tab + 1)) <= 1, g


def test_zones_are_balanced_within_a_table():
    d = dc.build(big_plan())
    for t in range(1, 7):
        per = Counter(p.zone for p in d.pots if p.table == t)
        assert max(per.values()) - min(per.get(z, 0) for z in range(1, 6)) <= 1


def test_largest_bench_8_zones_25_slots():
    plan = big_plan(tables=2, zones_per_table=8, slots_per_zone=25)
    d = dc.build(plan)
    assert max(p.zone for p in d.pots) <= 8 and max(p.slot for p in d.pots) <= 25


# ---------------------------------------------------------------- validation
@pytest.mark.parametrize("kw, words", [
    (dict(tables=2), "do not fit"),
    (dict(zones_per_table=9), "zones per table"),
    (dict(slots_per_zone=26), "slots per zone"),
    (dict(min_stagger_days=1), "at least 2"),
    (dict(cohorts=[]), "at least one cohort"),
    (dict(treatment_dap=21), "before the flood"),
    (dict(rows=[dc.PlanRow("B73", "Z", "CTRL", 5)]), "not in the cohort list"),
    (dict(rows=[dc.PlanRow("B73", "A", "CTRL", 5), dc.PlanRow("B73", "A", "CTRL", 3)]), "twice"),
])
def test_validation_says_what_to_fix(kw, words):
    problems = dc.validate(p0_plan(**kw))
    assert any(words in p for p in problems), problems
    with pytest.raises(dc.PlanError):
        dc.build(p0_plan(**kw))


# ---------------------------------------------------------------- schedule
def test_events_follow_days_after_planting():
    d = dc.build(big_plan())
    ev = d.events(2, d.plan.cohort("A"))
    p2 = MON + dt.timedelta(days=6)
    assert ev == [(p2, "Plant"), (p2 + dt.timedelta(14), "Treatment"),
                  (p2 + dt.timedelta(20), "Sacrificial harvest, flood"),
                  (p2 + dt.timedelta(21), "24 h N2O"),
                  (p2 + dt.timedelta(22), "48 h N2O, harvest")]


def test_watering_is_monday_wednesday_friday():
    mon = dc.water_days(MON, MON + dt.timedelta(days=18))
    assert [d.strftime("%a %d") for d in mon] == [
        "Wed 04", "Fri 06", "Mon 09", "Wed 11", "Fri 13", "Mon 16", "Wed 18"]
    tue = dc.water_days(MON + dt.timedelta(days=1), MON + dt.timedelta(days=18))
    assert [d.strftime("%a %d") for d in tue][:3] == ["Wed 04", "Fri 06", "Mon 09"]
    for start in range(7):                                 # any planting weekday
        days = dc.water_days(MON + dt.timedelta(days=start), MON + dt.timedelta(days=40))
        assert {d.strftime("%a") for d in days} == {"Mon", "Wed", "Fri"}
        assert all(MON + dt.timedelta(days=start) < d < MON + dt.timedelta(days=40) for d in days)


def test_weekend_events_are_allowed_but_not_weekend_water():
    d = dc.build(p0_plan(planting_date=MON - dt.timedelta(days=1)))   # a Sunday planting
    days = [day for day, _ in d.events(1, d.plan.cohort("A"))]
    assert days[0].weekday() == 6                                      # planted Sunday
    for c in d.plan.cohorts:
        treat = d.planting(1, c.name) + dt.timedelta(days=d.plan.treat_dap(c))
        assert all(day.weekday() < 5 or day == treat for day in d.water(1, c))


def test_n2o_loads_mix_cells_and_spread_over_columns():
    d = dc.build(big_plan())
    for g in range(1, d.groups + 1):
        for c in d.plan.cohorts:
            m = sorted((p for p in d.pots if p.group == g and p.cohort == c.name
                        and p.role == dc.MEASURED), key=lambda p: p.seq)
            assert [p.seq for p in m] == list(range(1, len(m) + 1))
            loads = Counter(p.load for p in m)
            assert all(n == dc.N2O_COLUMNS for n in list(loads.values())[:-1])
            for ld in loads:
                cols = [p.column for p in m if p.load == ld]
                assert len(set(cols)) == len(cols)          # one pot per column
                cells = [(p.genotype, p.treatment) for p in m if p.load == ld]
                assert len(set(cells)) == len(cells)        # no cell twice in a load
            for cell in {(p.genotype, p.treatment) for p in m}:
                per = Counter(p.column for p in m if (p.genotype, p.treatment) == cell)
                assert max(per.values()) - min(per.get(k, 0) for k in (1, 2, 3)) <= 1


def test_plan_file_round_trip():
    plan = big_plan()
    back = dc.Plan.from_dict(json.loads(json.dumps(plan.to_dict())))
    assert back == plan


# ---------------------------------------------------------------- tasks
def test_workday_skips_weekends():
    fri, sat = dt.date(2026, 11, 6), dt.date(2026, 11, 7)
    assert dc.workday(fri, 1) == dt.date(2026, 11, 9)           # Fri + 1 -> Mon
    assert dc.workday(MON, -1) == dt.date(2026, 10, 30)         # Mon - 1 -> Fri
    assert dc.workday(sat, 0) == dc.workday(sat, 1) == dt.date(2026, 11, 9)
    assert dc.workday(sat, -1) == fri
    assert dc.workday(MON, 3) == dt.date(2026, 11, 5)


def _named(tasks, start):
    return [t for t in tasks if t.name.startswith(start)]


def test_every_task_kind_is_listed_once_per_repeat():
    d = dc.build(p0_plan())                                     # 1 group, cohorts A-C, SAC pots
    t = d.tasks()
    assert [x.date for x in t] == sorted(x.date for x in t)
    one_each = ("Pre-fill pots", "Label, plant seed", "Label tables", "Apply treatments",
                "KCl soil collection setup", "Run swaps", "Prep KCl extracts for shipping",
                "Make the treatment solutions", "Print all labels")
    for start in one_each:
        assert len(_named(t, start)) == 1, start
    per_cohort = ("Flood - ", "Wash pots")
    for start in per_cohort:
        assert len(_named(t, start)) == 3, start
    assert len(_named(t, "N2O measurements")) == 3 * 2          # 24 h and 48 h per cohort
    for start in ("Plant harvest", "Plant dry weights"):
        assert len(_named(t, start)) == 3 * 2, start            # sacrificial + final
    assert len(_named(t, "KCl extractions")) == 2               # 117 samples: 2 batches
    assert not _named(t, "KCl analysis")
    water = _named(t, "Maintenance watering")
    assert [x.date for x in water] == sorted({w for c in d.plan.cohorts for w in d.water(1, c)})


def test_germination_check_is_every_weekday_of_the_first_week():
    for start in range(7):                                      # any planting weekday
        p0 = MON + dt.timedelta(days=start)
        d = dc.build(p0_plan(planting_date=p0))
        days = [x.date for x in _named(d.tasks(), "If using ZeaL genotypes, check for germination")]
        want = [p0 + dt.timedelta(days=k) for k in range(1, 8)
                if (p0 + dt.timedelta(days=k)).weekday() < 5]
        assert days == want and len(days) == 5


def test_lab_tasks_are_on_working_days_after_the_harvest():
    d = dc.build(p0_plan(planting_date=MON - dt.timedelta(days=2)))   # weekend harvests
    t = d.tasks()
    lab = ("Pre-fill", "KCl", "Plant dry weights", "Wash pots", "If using ZeaL", "Maintenance")
    treat = {x.date for x in _named(t, "Apply treatments")}         # watered in that day
    assert all(x.date.weekday() < 5 for x in t if x.name.startswith(lab)
               and not (x.name == "Maintenance watering" and x.date in treat))
    for c in d.plan.cohorts:
        final = d.events(1, c)[-1][0]
        mine = [x for x in t if x.name.endswith(f"cohort {c.name}")]
        harvest = [x.date for x in mine if x.name.startswith("Plant harvest")]
        assert max(harvest) == final
        wash = [x.date for x in mine if x.name.startswith("Wash pots")]
        assert wash[0] > final


@pytest.mark.parametrize("plan, sizes", [
    (p0_plan(), [59, 58]),                                      # 90 measured + 27 SAC = 117
    (big_plan(), [63, 63, 63, 63]),                             # 252
    (p0_plan(rows=[dc.PlanRow("B73", "A", "CTRL", 10, 2)], cohorts=[dc.Cohort("A", 21)]),
     [12])])
def test_kcl_batches_are_few_equal_and_after_the_last_harvest(plan, sizes):
    d = dc.build(plan)
    batches = d.kcl_batches()
    assert [sum(n for _, n in b) for b in batches] == sizes
    assert max(sizes) <= dc.KCL_BATCH_MAX
    samples = sum(p.role in (dc.MEASURED, dc.SACRIFICIAL) for p in d.pots)
    assert sum(sizes) == samples
    last = max(day for day, _, _ in d.harvests())
    t = d.tasks()
    ext = [x.date for x in _named(t, "KCl extractions")]
    ship = [x.date for x in _named(t, "Prep KCl extracts for shipping")]
    assert ext == [dc.workday(last, k) for k in range(1, len(sizes) + 1)]
    assert ship == [dc.workday(ext[-1], 1)]                     # once, after the last batch
    if len(sizes) == 1:
        assert [x.name for x in _named(t, "KCl extractions")] == ["KCl extractions"]


def test_kcl_batches_fill_in_harvest_order():
    d = dc.build(p0_plan())
    b = d.kcl_batches()
    assert b[0][0] == ("A sacrificial", 9) and b[-1][-1] == ("C final", 30)
    split = [name for name, _ in b[0]][-1]                      # one harvest spans the two
    assert split == [name for name, _ in b[1]][0] == "B final"


def test_all_labels_are_printed_once_before_the_first_planting():
    for plan in (p0_plan(spares_per_genotype_per_group=2), big_plan(), late_plan()):
        d = dc.build(plan)
        t = d.tasks()
        lab = _named(t, "Print all labels")
        assert len(lab) == 1 and lab[0].section == "Setup" and lab[0].kind == "print_labels"
        first = min(d.group_planting(g) for g in range(1, d.groups + 1))
        assert lab[0].date == dc.workday(first, -1)
        pots = len(d.pots)
        plants = sum(p.role in (dc.MEASURED, dc.SACRIFICIAL) for p in d.pots)
        assert lab[0].notes.startswith(f"{pots + 2 * plants + 4 * plants} labels: ")
        assert f"pots {pots} " in lab[0].notes and f"biomass bags {2 * plants} " in lab[0].notes
        assert f"KCl {4 * plants} " in lab[0].notes


def test_kcl_tubes_are_prepared_once_with_the_setup():
    for plan in (p0_plan(), big_plan(), late_plan()):
        d = dc.build(plan)
        t = d.tasks()
        setup = _named(t, "KCl soil collection setup")
        prefill = _named(t, "Pre-fill pots")
        assert len(setup) == 1 and setup[0].section == "Setup"
        assert setup[0].date == min(x.date for x in prefill)        # with the first setup
        assert setup[0].date < min(day for day, _, _ in d.harvests())
        assert setup[0].notes.startswith(
            f"{sum(p.role in (dc.MEASURED, dc.SACRIFICIAL) for p in d.pots)} tubes")


def test_a_swaps_task_a_week_after_each_planting_day():
    for plan in (p0_plan(spares_per_genotype_per_group=2), big_plan(), late_plan()):
        d = dc.build(plan)
        t = d.tasks()
        plant = _named(t, "Label, plant seed")
        swaps = _named(t, "Run swaps")
        assert len(swaps) == len(plant)                             # one per planting day
        assert sorted(x.date for x in swaps) == sorted(dc.workday(x.date, 5) for x in plant)
        assert all(x.date.weekday() < 5 for x in swaps)


def test_treatments_are_made_the_working_day_before_the_first():
    for plan in (p0_plan(), big_plan(), late_plan(), p0_plan(planting_date=MON + dt.timedelta(1))):
        d = dc.build(plan)
        t = d.tasks()
        make = _named(t, "Make the treatment solutions")
        apply = [x.date for x in _named(t, "Apply treatments")]
        assert len(make) == 1 and make[0].section == "Treatment"
        assert make[0].date == dc.workday(min(apply), -1) and make[0].date.weekday() < 5
        for trt in ("CTRL", "SN"):
            n = sum(p.treatment == trt and p.role in (dc.MEASURED, dc.SACRIFICIAL)
                    for p in d.pots)
            assert f"{trt} ({n} pots)" in make[0].notes


@pytest.mark.parametrize("kw", [dict(treatment_dap=None),
                                dict(rows=[dc.PlanRow("B73", c, "", 3) for c in "ABC"])])
def test_no_treatments_nothing_to_make(kw):
    t = dc.build(p0_plan(**kw)).tasks()
    assert not _named(t, "Make the treatment solutions")


def test_every_task_has_a_kind():
    for plan in (p0_plan(), big_plan(), late_plan()):
        t = dc.build(plan).tasks()
        assert all(x.kind in dc.TASK_KINDS for x in t)
        assert {x.kind for x in t} == set(dc.TASK_KINDS)


def test_several_groups_name_the_group():
    d = dc.build(big_plan())
    assert d.groups > 1
    t = d.tasks()
    for g in range(1, d.groups + 1):
        assert len(_named(t, f"Label, plant seed, and initial wet-up of pots (group {g})")) == 1
        assert len([x for x in t if x.name.startswith("N2O measurements")
                    and x.name.endswith(f"group {g}")]) == len(d.plan.cohorts) * 2
    days = [x.date for x in _named(t, "Maintenance watering")]
    assert len(days) == len(set(days))                          # one watering task per day


def test_asana_sheet_is_a_plain_import_table(tmp_path):
    import openpyxl
    d = dc.build(p0_plan())
    path = tmp_path / "x.xlsx"
    dx.write(d, str(path))
    ws = openpyxl.load_workbook(path)["asana_tasks"]
    rows = list(ws.iter_rows(values_only=True))
    assert rows[0] == ("Name", "Due Date", "Section/Column", "Description", "Assignee",
                       "Collaborators")
    assert len(rows) - 1 == len(d.tasks()) + dx.WATER_SLACK
    assert all(str(r[1]).startswith("=") for r in rows[1:])        # every date is a formula
    assert ws["B2"].number_format == "mm/dd/yyyy"


def _asana_values(path):
    """The asana_tasks sheet recalculated the way Excel would: [(name, date, notes)]."""
    formulas = pytest.importorskip("formulas")
    import numpy as np
    import openpyxl
    import formulas.functions.date as fdate

    class ExcelWorkday:
        """`formulas` gives #VALUE! for WORKDAY from a weekend day; Excel counts from it
        (WORKDAY(Sat, 1) = Mon, WORKDAY(Sat, -1) = Fri). numpy's roll does the same when
        it rolls against the direction of travel."""
        def __getattr__(self, name):
            return getattr(np, name)

        @staticmethod
        def busday_offset(d, n, **kw):
            return np.busday_offset(d, n, roll="backward" if n > 0 else "forward", **kw)

    real, fdate.np = fdate.np, ExcelWorkday()
    try:
        sol = formulas.ExcelModel().loads(str(path)).finish().calculate()
    finally:
        fdate.np = real
    def val(ref):
        key = next(k for k in sol if k.upper().endswith(ref.upper()))
        v = sol[key].value[0][0]
        return v.item() if hasattr(v, "item") else v
    ws = openpyxl.load_workbook(path)["asana_tasks"]
    cell = lambda ref: (val(f"ASANA_TASKS'!{ref}") if str(ws[ref].value or "").startswith("=")
                        else ws[ref].value)                     # constants are not in `sol`
    excel0 = dt.date(1899, 12, 30)
    out = []
    for r in range(2, ws.max_row + 1):
        name = cell(f"A{r}")
        if name in ("", None):
            continue
        day = excel0 + dt.timedelta(days=int(cell(f"B{r}")))
        out.append((name, day, cell(f"D{r}") or ""))
    return sorted(out, key=str)


def _expected(d):
    return sorted(((t.name, t.date, t.notes) for t in d.tasks()), key=str)


def late_plan(**kw):
    """Cohorts that collide inside one group: B and D are planted later (offsets)."""
    return p0_plan(cohorts=[dc.Cohort("A", 21), dc.Cohort("B", 22), dc.Cohort("C", 27),
                            dc.Cohort("D", 28)],
                   rows=[dc.PlanRow("B73", c, t, 4, 1) for c in "ABCD" for t in ("CTRL", "SN")],
                   **kw)


def test_asana_dates_calculate(tmp_path):
    for plan in (p0_plan(), big_plan(), late_plan()):
        d = dc.build(plan)
        path = tmp_path / "x.xlsx"
        dx.write(d, str(path))
        assert _asana_values(path) == _expected(d)


def test_a_planting_date_typed_over_moves_its_group(tmp_path):
    """Typing a later date over planting group 2 (cohorts B and D) moves only their dates."""
    import openpyxl
    d = dc.build(late_plan())
    assert [c.name for c in d.cohorts_of(2)] == ["B", "D"]
    path = tmp_path / "x.xlsx"
    dx.write(d, str(path))
    wb = openpyxl.load_workbook(path)
    ws = wb["timeline"]
    head = next(r for r in range(1, 40) if ws.cell(r, 5).value == "Planting group")
    ws.cell(head + 2, 6).value = d.group_planting(2) + dt.timedelta(days=3)
    wb.save(path)
    rows = d.tasks()                                        # the rows are fixed at writing
    d.rounds[1] = (d.rounds[1][0], d.rounds[1][1] + 3)       # group 2 three days later
    for c in ("B", "D"):                                     # (its cohorts follow it)
        d.offsets[c] += 3
    moved = [(t.name, d.anchor_date(t.anchor) if t.lag is None
              else dc.workday(d.anchor_date(t.anchor), t.lag), t.notes)
             for t in rows if t.anchor[0] != "water"]
    water = [("Maintenance watering", day, d.water_who(day)) for day in d.water_days_all()]
    strip = lambda rows: sorted(((n, day, "" if n.startswith("KCl") else x)
                                 for n, day, x in rows), key=str)
    assert strip(_asana_values(path)) == strip(moved + water)
    assert [day for name, day, _ in moved if name == "Flood - cohort B, group 2"] == \
           [d.group_planting(2) + dt.timedelta(days=21)]

def test_asana_dates_follow_timeline_edits(tmp_path):
    import openpyxl
    d = dc.build(p0_plan())
    path = tmp_path / "x.xlsx"
    dx.write(d, str(path))
    wb = openpyxl.load_workbook(path)
    ws = wb["timeline"]
    ws["B4"] = MON + dt.timedelta(days=3)                          # plant on Thursday instead
    row_b = next(r for r in range(1, 40) if ws.cell(r, 1).value == "B")
    ws.cell(row_b, 2).value = 26                                   # cohort B 2 days later
    wb.save(path)
    d.plan.planting_date = MON + dt.timedelta(days=3)
    d.plan.cohort("B").first_n2o_dap = 26
    # KCl batch notes (which harvests) are written once; their dates follow the edit
    strip = lambda rows: sorted(((n, day, "" if n.startswith("KCl") else x) for n, day, x in rows),
                                key=str)
    want = strip(_expected(d))
    assert want != strip(_expected(dc.build(p0_plan())))          # the edit moved things
    assert strip(_asana_values(path)) == want


# ---------------------------------------------------------------- workbook
def test_workbook_sheets_and_jars(tmp_path):
    import openpyxl
    d = dc.build(big_plan())
    path = tmp_path / "x.xlsx"
    dx.write(d, str(path))
    wb = openpyxl.load_workbook(path)
    per = [f"{k}_G{g}_{c}" for g in range(1, d.groups + 1) for c in "AB"
           for k in ("flood", "measure", "harvest")]
    assert wb.sheetnames == ["timeline", "design", "layout"] + per + ["tags", "groups",
                                                                       "assignments",
                                                                       "asana_tasks", "jars",
                                                                       "admin", "info",
                                                                       "plan"]
    assert wb["plan"].sheet_state == "hidden"
    jars = wb["jars"]
    assert [c.value for c in jars[1]] == ["barcode", "tag"]
    rows = list(jars.iter_rows(min_row=2, values_only=True))
    assert len(rows) == len(d.pots)
    import re
    assert all(re.search(r"\bT\d+-Z\d+-\d+\b", tag) for _, tag in rows)  # station reads this
    assert jars["A2"].number_format == "@"
    design = wb["design"]
    assert "timeline!$F$" in design["H2"].value                       # linked planting date


def test_one_group_sheets_are_named_by_cohort(tmp_path):
    import openpyxl
    d = dc.build(p0_plan())
    path = tmp_path / "p0.xlsx"
    dx.write(d, str(path))
    names = openpyxl.load_workbook(path).sheetnames
    assert {"flood_A", "measure_A", "harvest_A", "flood_C", "harvest_C", "tags"} <= set(names)


def test_flood_measure_harvest_contents(tmp_path):
    import openpyxl
    d = dc.build(p0_plan())
    path = tmp_path / "p0.xlsx"
    dx.write(d, str(path))
    wb = openpyxl.load_workbook(path)
    meas = sorted((p for p in d.pots if p.cohort == "A" and p.role == dc.MEASURED),
                  key=lambda p: p.seq)
    sac = [p for p in d.pots if p.cohort == "A" and p.role == dc.SACRIFICIAL]
    fl = wb["flood_A"]
    col_a = [fl.cell(r, 1).value for r in range(1, fl.max_row + 1)]
    assert "SACRIFICIAL harvest first" in col_a and "FLOOD in this order" in col_a
    bars = [fl.cell(r, 2).value for r in range(col_a.index("FLOOD in this order") + 2,
                                                fl.max_row + 1)]
    assert bars == [p.barcode for p in meas]               # flood in N2O order
    sac_rows = col_a.index("FLOOD in this order") - col_a.index("SACRIFICIAL harvest first") - 2
    assert sac_rows == len(sac) == 9
    assert "timeline!" in fl["B2"].value                    # the date follows the timeline
    ms = wb["measure_A"]
    heads = [r for r in range(1, ms.max_row + 1) if ms.cell(r, 1).value == "load"]
    assert len(heads) == 2                                 # 24 h and 48 h blocks
    first = [ms.cell(r, 4).value for r in range(heads[0] + 1, heads[0] + 1 + len(meas))]
    assert first == [p.barcode for p in meas]
    hv = wb["harvest_A"]
    assert [hv.cell(r, 2).value for r in range(5, 5 + len(meas))] == [p.barcode for p in meas]
    tags = wb["tags"]
    assert [c.value for c in tags[1]][:3] == ["barcode", "tag_text", "position"]
    assert tags.max_row - 1 == len(d.pots)


def test_layout_colours_cohort_family_and_treatment_shade(tmp_path):
    import openpyxl
    for plan in (p0_plan(), big_plan()):
        d = dc.build(plan)
        path = tmp_path / "l.xlsx"
        dx.write(d, str(path))
        ws = openpyxl.load_workbook(path)["layout"]
        colours = dx.layout_colours(d)
        trts = dx.treatments(d)
        assert set(colours) == {(c.name, t) for c in d.plan.cohorts for t in trts}
        assert len(set(colours.values())) == len(colours)          # every combination differs
        for ci, c in enumerate(d.plan.cohorts):                    # one family per cohort
            ramp = dx.FAMILIES[ci][1]
            shades = [ramp.index(colours[(c.name, t)][0]) for t in trts]
            assert shades == sorted(shades) and len(set(shades)) == len(trts)  # light to dark
        legend = {ws.cell(r, col).fill.start_color.rgb[-6:]
                  for r in range(5, 5 + len(d.plan.cohorts))
                  for col in range(2, 2 + len(trts))}
        assert legend == {bg for bg, _ in colours.values()}
        at = {(p.table, p.zone, p.slot): p for p in d.pots}
        seen = 0
        for row in ws.iter_rows():
            for cell in row:
                v = str(cell.value or "")
                p = next((p for p in d.pots if v.endswith("\n" + p.barcode)), None)
                if p is None:
                    continue
                seen += 1
                want = SPARE if p.role == dc.SPARE else colours[(p.cohort, p.treatment)][0]
                assert cell.fill.start_color.rgb[-6:] == want, (p.barcode, v)
                assert v.startswith(f"G{p.group}  ") == (d.groups > 1 and p.role != dc.SPARE)
        assert seen == len(d.pots)


SPARE = dx.SPARE_BG


def test_timeline_cells_carry_their_colours(tmp_path):
    """The designed colours are plain fills too, so a viewer that does not run conditional
    formatting still shows them; weekends are shaded in the date columns."""
    import openpyxl
    d = dc.build(big_plan())
    path = tmp_path / "x.xlsx"
    dx.write(d, str(path))
    ws = openpyxl.load_workbook(path)["timeline"]
    hdr = next(r for r in range(1, 60) if ws.cell(r, 1).value == "Date")
    cal0 = min(d.group_planting(g) for g in range(1, d.groups + 1))
    cols = d.cells()
    colour = lambda cell: (cell.fill.fgColor.rgb or "")[-6:] if cell.fill.fill_type else ""
    want = {"flood": dx.FLOOD_BG, "N2O": dx.N2O_BG, "Plant": dx.PLANT_BG,
            "Treatment": dx.TREAT_BG, "Water": dx.WATER_BG}
    seen = set()
    for i in range(40):
        r, day = hdr + 2 + i, cal0 + dt.timedelta(days=i)
        assert (colour(ws.cell(r, 1)) == dx.WEEKEND_BG) == (day.weekday() >= 5)
        for j, (g, c) in enumerate(cols):
            text = d.day_text(g, c, day)
            word = next((w for w in want if w.lower() in text.lower()), None)
            assert colour(ws.cell(r, 3 + j)) == (want[word] if word else ""), (day, text)
            seen.add(word)
    assert {"flood", "N2O", "Plant", "Treatment", "Water"} <= seen


def test_timeline_formulas_calculate(tmp_path):
    """Recalculate the workbook the way Excel would (needs the `formulas` package)."""
    formulas = pytest.importorskip("formulas")
    import openpyxl
    for plan in (big_plan(), late_plan()):
        _check_timeline(formulas, openpyxl, dc.build(plan), tmp_path / "x.xlsx")


def _check_timeline(formulas, openpyxl, d, path):
    dx.write(d, str(path))
    sol = formulas.ExcelModel().loads(str(path)).finish().calculate()
    def val(ref):
        key = next(k for k in sol if k.upper().endswith(ref.upper()))
        v = sol[key].value[0][0]
        return v.item() if hasattr(v, "item") else v
    ws = openpyxl.load_workbook(path)["timeline"]
    hdr = next(r for r in range(1, 60) if ws.cell(r, 1).value == "Date")
    start = hdr + 2
    tot_col = next(c for c in range(1, 40) if ws.cell(hdr, c).value == "N2O pots")
    from openpyxl.utils import get_column_letter as L
    excel0 = dt.date(1899, 12, 30)
    got = {}
    last = start
    while str(ws.cell(last + 1, 1).value or "").startswith("="):
        last += 1
    for r in range(start, last + 1):
        day = excel0 + dt.timedelta(days=int(val(f"TIMELINE'!A{r}")))
        got[day] = int(val(f"TIMELINE'!{L(tot_col)}{r}"))
        for j, (g, c) in enumerate(d.cells()):
            text = val(f"TIMELINE'!{L(3 + j)}{r}")
            assert text == d.day_text(g, c, day), (day, g, c.name, text, d.day_text(g, c, day))
    assert {k: v for k, v in got.items() if v} == d.calendar_load()
    flood_col = tot_col - 1
    assert all(int(val(f"TIMELINE'!{L(flood_col)}{r}")) == 0
               or int(val(f"TIMELINE'!{L(tot_col)}{r}")) == 0
               for r in range(start, last + 1))           # never both on one day


# ---------------------------------------------------------------- N2O times per cohort
def mixed_plan(**kw):
    """A samples at 24, 48, 72 h; B at 24, 48 h; C only at 48 and 96 h."""
    return p0_plan(cohorts=[dc.Cohort("A", 21, (24, 48, 72)), dc.Cohort("B", 24, (24, 48)),
                            dc.Cohort("C", 30, (48, 96))],
                   rows=[dc.PlanRow("B73", c, t, 4, 1) for c in "ABC" for t in ("CTRL", "SN")],
                   **kw)


def test_each_cohort_has_its_own_n2o_times():
    d = dc.build(mixed_plan())
    for c in d.plan.cohorts:
        ev = d.events(1, c)
        p0 = d.planting(1, c.name)
        flood = [day for day, t in ev if "flood" in t.lower()]
        n2o = [(day, t) for day, t in ev if "N2O" in t]
        assert flood == [p0 + dt.timedelta(days=c.first_n2o_dap - c.measure_hours[0] // 24)]
        assert [t.split(",")[0] for _, t in n2o] == [f"{h} h N2O" for h in c.measure_hours]
        assert [day for day, _ in n2o] == [flood[0] + dt.timedelta(days=h // 24)
                                           for h in c.measure_hours]
        assert n2o[-1][1].endswith("harvest") and n2o[0][0] == p0 + dt.timedelta(
            days=c.first_n2o_dap)                           # first N2O keeps its meaning
    t = d.tasks()
    for c in "ABC":
        n = [x for x in t if x.name.startswith("N2O measurements") and f"cohort {c}" in x.name]
        assert len(n) == len(d.plan.cohort(c).measure_hours)


def test_load_and_clash_count_every_cohorts_times():
    d = dc.build(mixed_plan())
    meas = Counter(p.cohort for p in d.pots if p.role == dc.MEASURED)
    assert sum(d.load.values()) == sum(meas[c.name] * len(c.measure_hours)
                                       for c in d.plan.cohorts)
    flood, n2o = _flood_n2o_days(d)
    assert not flood & n2o


@pytest.mark.parametrize("hours, words", [((24, 36), "whole days"), ((48, 24), "in order"),
                                          ((24, 24), "in order"), ((), "at least one")])
def test_bad_n2o_times_are_refused(hours, words):
    with pytest.raises(dc.PlanError, match=words):
        dc.build(p0_plan(cohorts=[dc.Cohort("A", 21, hours)],
                         rows=[dc.PlanRow("B73", "A", "CTRL", 3)]))


def test_flood_must_follow_treatment_with_late_first_sample():
    # first sample at 72 h: the flood is 3 days before the first N2O day (DAP 18)
    with pytest.raises(dc.PlanError, match="before the flood"):
        dc.build(p0_plan(treatment_dap=18, cohorts=[dc.Cohort("A", 21, (72,))],
                         rows=[dc.PlanRow("B73", "A", "CTRL", 3)]))


def test_old_plan_files_give_every_cohort_the_shared_times():
    d = p0_plan().to_dict()
    for c in d["cohorts"]:
        c.pop("measure_hours")
    d["measure_hours"] = [24, 48, 72]                          # the v0.4 plan-level list
    plan = dc.Plan.from_dict(d)
    assert all(c.measure_hours == (24, 48, 72) for c in plan.cohorts)
    assert dc.Plan.from_dict(json.loads(json.dumps(mixed_plan().to_dict()))) == mixed_plan()


def test_mixed_times_formulas_calculate(tmp_path):
    formulas = pytest.importorskip("formulas")
    import openpyxl
    d = dc.build(mixed_plan())
    _check_timeline(formulas, openpyxl, d, tmp_path / "x.xlsx")
    assert _asana_values(tmp_path / "x.xlsx") == _expected(d)
    wb = openpyxl.load_workbook(tmp_path / "x.xlsx")
    sheet = lambda kind, c: f"{kind}_{dx._sheet_tag(d, next(g for g, x in d.cells() if x.name == c), c)}"
    assert [c.value for c in wb[sheet("measure", "A")]["A"] if str(c.value or "").endswith(
        ("N2O", "harvest"))] == ["24 h N2O", "48 h N2O", "72 h N2O, then harvest"]
    assert wb[sheet("harvest", "C")]["A1"].value.endswith("after the 96 h N2O")


def test_window_reads_n2o_times():
    from PySide6 import QtWidgets
    import design_app
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    w = design_app.DesignWindow(mixed_plan())
    assert w.cohorts.text(0, 2) == "24, 48, 72"
    w.cohorts.set_text(1, 2, "24 48 72 96")
    w.cohorts.set_text(2, 2, "")
    p = w.read_plan()
    assert [c.measure_hours for c in p.cohorts] == [(24, 48, 72), (24, 48, 72, 96), (24, 48)]
    w.cohorts.set_text(2, 2, "24, two")
    with pytest.raises(dc.PlanError, match="cohort C"):
        w.read_plan()
    w.close()


# ---------------------------------------------------------------- admin tab
SOIL = dict(wfps_target=0.40, wfps_flood=0.95, bulk_density=1.2, air_dry_moisture=0.04,
            air_dry_soil_g=300.0, container_weight_g=35.0)


def _admin(path):
    import openpyxl
    ws = openpyxl.load_workbook(path)["admin"]
    return {r[0].value: (r[1], r[2].value) for r in ws.iter_rows()}


def test_admin_tab_holds_the_dosing_inputs(tmp_path):
    d = dc.build(p0_plan(**SOIL))
    path = tmp_path / "x.xlsx"
    dx.write(d, str(path))
    a = _admin(path)
    want = dict(SOIL, particle_density=2.65, trial_id="B73-P0", profile="wet_up",
                date=MON.isoformat())
    for k, v in want.items():
        assert a[k][0].value == v, k
    assert all(not str(c.value or "").startswith("=") for c, _ in a.values())  # no formulas
    assert "40 %" in a["wfps_target"][1] and "REQUIRED" not in "".join(n for _, n in a.values())


def test_admin_tab_flags_what_is_missing(tmp_path):
    d = dc.build(p0_plan(wfps_target=0.40))
    path = tmp_path / "x.xlsx"
    dx.write(d, str(path))
    a = _admin(path)
    assert dc.admin_missing(d.plan) == ["bulk_density", "air_dry_moisture", "air_dry_soil_g",
                                        "container_weight_g"]
    for k in dc.admin_missing(d.plan):
        assert a[k][0].value is None and a[k][1].startswith("REQUIRED"), k
    assert a["wfps_flood"][0].value is None and not a["wfps_flood"][1].startswith("REQUIRED")


@pytest.mark.parametrize("kw, words", [
    (dict(wfps_target=40), "0.40 for 40 %"),
    (dict(wfps_target=0.4, wfps_flood=0.35), "above the target"),
    (dict(air_dry_moisture=4), "outside"),
    (dict(bulk_density=12), "outside")])
def test_admin_inputs_are_checked(kw, words):
    with pytest.raises(dc.PlanError, match=words):
        dc.build(p0_plan(**kw))


def test_admin_inputs_survive_a_plan_file():
    plan = p0_plan(**SOIL)
    assert dc.Plan.from_dict(json.loads(json.dumps(plan.to_dict()))) == plan
    old = p0_plan().to_dict()                                   # a plan file from before
    for k in SOIL:
        old.pop(k)
    assert dc.Plan.from_dict(old).wfps_target is None


def test_dosing_station_provisions_from_the_design_map(tmp_path):
    """The real station code (../bni-dosing) reads this workbook as a trial: jars, admin,
    and the WFPS targets. Skipped where the dosing repo is not beside this one."""
    dosing = os.path.join(R, "..", "bni-dosing")
    if not os.path.exists(os.path.join(dosing, "provision_jars.py")):
        pytest.skip("bni-dosing not beside bni-design")
    sys.path.insert(0, dosing)
    try:
        import provision_jars as P
    finally:
        sys.path.remove(dosing)
    d = dc.build(p0_plan(**SOIL))
    path = tmp_path / "trial.xlsx"
    dx.write(d, str(path))
    jars, meta, admin = P.read_jars(str(path))
    assert [j.barcode for j in jars] == [p.barcode for p in d.pots]
    res, _, tare = P.compute_from_admin(admin)
    flood = P.flood_from_admin(admin)
    assert tare == SOIL["container_weight_g"] + SOIL["air_dry_soil_g"]
    assert res.target_mass_g > tare and flood.target_mass_g > res.target_mass_g
    assert P.admin_profile(admin) == "wet_up"


def test_window_reads_percent_as_fractions():
    from PySide6 import QtWidgets
    import design_app
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    w = design_app.DesignWindow(p0_plan(**SOIL))
    assert w.admin_fields["wfps_target"][0].text() == "40"
    assert w.read_plan().wfps_target == pytest.approx(0.40)
    w.admin_fields["air_dry_moisture"][0].setText("5")
    w.admin_fields["bulk_density"][0].setText("")
    p = w.read_plan()
    assert p.air_dry_moisture == pytest.approx(0.05) and p.bulk_density is None
    w.admin_fields["bulk_density"][0].setText("abc")
    with pytest.raises(dc.PlanError, match="Bulk density"):
        w.read_plan()
    w.close()


# ---------------------------------------------------------------- window
def test_window_generates_from_the_example_plan():
    from PySide6 import QtWidgets
    import design_app
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    w = design_app.DesignWindow()
    assert not w.save_b.isEnabled()
    w._generate()
    assert w.design is not None and w.save_b.isEnabled()
    assert "PLANTING GROUPS: 1" in w.summary.toPlainText()
    w.tables.setValue(1)                                   # edits make the design stale
    assert w.design is None and not w.save_b.isEnabled()
    w._generate()
    assert w.design is None and "do not fit" in w.summary.toPlainText()
    w.close()


def test_window_delete_key_clears_selected_cells():
    from PySide6 import QtCore, QtGui, QtWidgets
    import design_app
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    w = design_app.DesignWindow()
    t = w.table
    assert t.text(0, 0)                                     # the example plan is filled in
    for key in (QtCore.Qt.Key_Delete, QtCore.Qt.Key_Backspace):
        t.set_text(0, 0, "B73"); t.set_text(0, 3, "5")
        t.clearSelection(); t.setCurrentCell(0, 0)
        t.item(0, 0).setSelected(True); t.item(0, 3).setSelected(True)
        t.keyPressEvent(QtGui.QKeyEvent(QtCore.QEvent.KeyPress, key, QtCore.Qt.NoModifier))
        assert t.text(0, 0) == "" and t.text(0, 3) == "", key
    w.close()


def test_window_paste_from_excel():
    from PySide6 import QtWidgets
    import design_app
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    w = design_app.DesignWindow()
    w.table.setRowCount(0); w.table.setRowCount(3)
    w.table.setCurrentCell(0, 0)
    QtWidgets.QApplication.clipboard().setText("W22\tA\tSN\t4\t1\nW22\tB\tSN\t4\t1\n")
    w.table.paste()
    plan = w.read_plan()
    assert [(r.genotype, r.cohort, r.treatment, r.reps, r.sacrificial) for r in plan.rows] == \
        [("W22", "A", "SN", 4, 1), ("W22", "B", "SN", 4, 1)]
    w.close()


def test_window_spares_as_percent_or_fixed():
    from PySide6 import QtWidgets
    import design_app
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    w = design_app.DesignWindow(p0_plan(spare_percent=10, spares_per_genotype_per_group=0))
    assert w.spare_unit.currentIndex() == 0 and w.spares.value() == 10
    p = w.read_plan()
    assert p.spare_percent == 10 and p.spares_per_genotype_per_group == 0
    w.spare_unit.setCurrentIndex(1); w.spares.setValue(3)
    p = w.read_plan()
    assert p.spare_percent == 0 and p.spares_per_genotype_per_group == 3
    w.load_plan(p0_plan(spares_per_genotype_per_group=5))      # an older plan file
    assert w.spare_unit.currentIndex() == 1 and w.spares.value() == 5
    w.close()


# ---------------------------------------------------------------- treatment day per cohort
def trial_plan(**kw):
    """A treatment time trial: cohort A treated 14 days after planting, B 16, C 18."""
    return p0_plan(cohorts=[dc.Cohort("A", 21, treatment_dap=14),
                            dc.Cohort("B", 24, treatment_dap=16),
                            dc.Cohort("C", 27, treatment_dap=18)], **kw)


def test_each_cohort_is_treated_on_its_own_day():
    d = dc.build(trial_plan())
    for c, dap in zip(d.plan.cohorts, (14, 16, 18)):
        for g, k in d.cells():
            if k.name == c.name:
                treat = [day for day, t in d.events(g, k) if t == "Treatment"]
                assert treat == [d.planting(g, k.name) + dt.timedelta(days=dap)]


def test_a_blank_cohort_treatment_day_is_the_plans():
    d = dc.build(p0_plan(cohorts=[dc.Cohort("A", 21), dc.Cohort("B", 24, treatment_dap=17),
                                  dc.Cohort("C", 27)]))
    assert [d.plan.treat_dap(c) for c in d.plan.cohorts] == [14, 17, 14]
    assert d.plan.treat_per_cohort
    assert not dc.build(p0_plan()).plan.treat_per_cohort


def test_cohort_treatment_without_a_plan_treatment_day():
    d = dc.build(p0_plan(treatment_dap=None,
                         cohorts=[dc.Cohort("A", 21, treatment_dap=10), dc.Cohort("B", 24),
                                  dc.Cohort("C", 27)]))
    assert [d.plan.treat_dap(c) for c in d.plan.cohorts] == [10, None, None]
    treated = {k.name for g, k in d.cells() for _, t in d.events(g, k) if t == "Treatment"}
    assert treated == {"A"}
    assert len(_named(d.tasks(), "Make the treatment solutions")) == 1


def test_one_apply_task_per_treatment_day_naming_its_cohorts():
    d = dc.build(trial_plan())
    apply = _named(d.tasks(), "Apply treatments")
    days = {(g, k.name): d.planting(g, k.name) + dt.timedelta(days=d.plan.treat_dap(k))
            for g, k in d.cells()}
    assert len(apply) == len(days)
    for t in apply:
        assert t.kind == "treatment" and t.section == "Treatment"
        g, c = t.anchor[1], t.anchor[2]
        assert t.date == days[(g, c)] and f"cohort {c}" in t.name
    make = _named(d.tasks(), "Make the treatment solutions")
    assert len(make) == 1 and make[0].date == dc.workday(min(days.values()), -1)
    assert "cohort A day 14, cohort B day 16, cohort C day 18 after planting" in make[0].notes


def test_same_day_cohorts_share_one_apply_task():
    d = dc.build(p0_plan(cohorts=[dc.Cohort("A", 21, treatment_dap=14),
                                  dc.Cohort("B", 24, treatment_dap=14), dc.Cohort("C", 27)]))
    assert not d.plan.treat_per_cohort
    assert [t.name for t in _named(d.tasks(), "Apply treatments")] == ["Apply treatments"] * \
        len(_named(d.tasks(), "Label, plant seed"))


@pytest.mark.parametrize("dap, words", [(20, "before the flood"), (-1, "before planting")])
def test_a_cohort_treatment_day_is_checked(dap, words):
    plan = p0_plan(cohorts=[dc.Cohort("A", 21, treatment_dap=dap), dc.Cohort("B", 24),
                            dc.Cohort("C", 27)])
    assert any(words in p and "cohort A" in p for p in dc.validate(plan))


def test_cohort_treatment_days_are_saved_in_the_plan_file():
    plan = dc.Plan.from_dict(trial_plan().to_dict())
    assert [c.treatment_dap for c in plan.cohorts] == [14, 16, 18]
    old = p0_plan().to_dict()
    for c in old["cohorts"]:
        c.pop("treatment_dap")                                  # plan files before this
    assert all(c.treatment_dap is None for c in dc.Plan.from_dict(old).cohorts)


def test_cohort_treatment_days_in_the_workbook(tmp_path):
    import openpyxl
    d = dc.build(trial_plan())
    path = tmp_path / "x.xlsx"
    dx.write(d, str(path))
    ws = openpyxl.load_workbook(path)["timeline"]
    top = dx.TIMELINE_TOP
    assert ws.cell(top, 4).value == dx.TREAT_HEAD
    assert [ws.cell(top + 1 + i, 4).value for i in range(3)] == [14, 16, 18]
    ws = openpyxl.load_workbook(path)["info"]
    rows = list(ws.iter_rows(values_only=True))
    i = next(k for k, r in enumerate(rows) if r and r[0] == "cohort")
    assert [r[4] for r in rows[i + 1:i + 4]] == [14, 16, 18]


def test_cohort_treatment_days_calculate_in_the_workbook(tmp_path):
    formulas = pytest.importorskip("formulas")
    import openpyxl
    for plan in (trial_plan(), p0_plan(cohorts=[dc.Cohort("A", 21, treatment_dap=12),
                                                dc.Cohort("B", 24), dc.Cohort("C", 27)])):
        d = dc.build(plan)
        _check_timeline(formulas, openpyxl, d, tmp_path / "x.xlsx")
        assert _asana_values(tmp_path / "x.xlsx") == _expected(d)


# ---------------------------------------------------------------- treatments are watered in
def _plain_water(d, g, c):
    p0 = d.planting(g, c.name)
    return dc.water_days(p0, p0 + dt.timedelta(days=c.flood_dap))


@pytest.mark.parametrize("dap", range(1, 20))
def test_a_treatment_is_watered_in(dap):
    """Any weekday: on a watering day nothing changes; on another day the pots are watered
    that day and the next day's watering, if the next day is one, is skipped."""
    d = dc.build(p0_plan(treatment_dap=dap))
    for g, c in d.cells():
        plain = _plain_water(d, g, c)
        treat = d.planting(g, c.name) + dt.timedelta(days=dap)
        nxt = treat + dt.timedelta(days=1)
        want = plain if treat in plain else sorted([x for x in plain if x != nxt] + [treat])
        assert d.water(g, c) == want
        assert d.day_text(g, c, treat) == "Treatment; water"


def test_tuesday_treatment_skips_wednesday_saturday_skips_nothing():
    d = dc.build(p0_plan(planting_date=MON, treatment_dap=15))     # planted Mon: DAP 15 Tue
    g, c = d.cells()[0]
    tue = d.planting(g, c.name) + dt.timedelta(days=15)
    assert tue.strftime("%a") == "Tue"
    w = d.water(g, c)
    assert tue in w and tue + dt.timedelta(1) not in w
    assert d.day_text(g, c, tue + dt.timedelta(1)) == ""
    d = dc.build(p0_plan(planting_date=MON, treatment_dap=12))     # Sat: Sun is never watered
    g, c = d.cells()[0]
    sat = d.planting(g, c.name) + dt.timedelta(days=12)
    assert d.water(g, c) == sorted(_plain_water(d, g, c) + [sat])


def test_no_treatment_watering_unchanged():
    d = dc.build(p0_plan(treatment_dap=None))
    assert all(d.water(g, c) == _plain_water(d, g, c) for g, c in d.cells())
    assert all(x.weekday() in dc.WATER_WEEKDAYS for x in d.water_days_all())


def test_watering_tasks_follow_the_treatment():
    d = dc.build(trial_plan())                         # A 14, B 16, C 18 days after planting
    water = {t.date for t in _named(d.tasks(), "Maintenance watering")}
    assert water == set(d.water_days_all())
    for g, c in d.cells():
        treat = d.planting(g, c.name) + dt.timedelta(days=d.plan.treat_dap(c))
        assert treat in water


def test_treatment_watering_calculates_in_the_workbook(tmp_path):
    formulas = pytest.importorskip("formulas")
    import openpyxl
    for plan in (p0_plan(treatment_dap=15), p0_plan(treatment_dap=12),
                 p0_plan(planting_date=MON + dt.timedelta(days=1), treatment_dap=13)):
        d = dc.build(plan)
        _check_timeline(formulas, openpyxl, d, tmp_path / "x.xlsx")
        assert _asana_values(tmp_path / "x.xlsx") == _expected(d)
