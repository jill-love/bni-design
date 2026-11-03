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
    assert d.stagger_days == 6
    assert dc.daily_load(d.plan, d.pots, 3)[24] > 50
    assert dc.clash_days(d.plan, 3, 4) and dc.clash_days(d.plan, 3, 5)
    assert not dc.clash_days(d.plan, 3, 6) and not d.warnings


def test_never_a_flood_on_an_n2o_day():
    for plan in (p0_plan(), big_plan(), big_plan(max_measured_per_day=20)):
        d = dc.build(plan)
        flood, n2o = set(), set()
        for g in range(1, d.groups + 1):
            for c in plan.cohorts:
                for day, text in d.events(g, c):
                    if "flood" in text.lower():
                        flood.add(day)
                    if "N2O" in text:
                        n2o.add(day)
        assert not flood & n2o


def test_cohorts_too_close_are_warned_not_refused():
    d = dc.build(p0_plan(cohorts=[dc.Cohort("A", 21), dc.Cohort("B", 23)],
                         rows=[dc.PlanRow("B73", c, "CTRL", 5) for c in "AB"]))
    assert d.warnings and "collide" in d.warnings[0]


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
    assert all(p.position == f"T{p.table}-Z{p.zone}-{p.slot:02d}" for p in d.pots)


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
        assert all(day.weekday() < 5 for day in d.water(1, c))


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
    one_each = ("Pre-fill pots", "Label, plant seed", "Label tables", "Apply treatments")
    for start in one_each:
        assert len(_named(t, start)) == 1, start
    per_cohort = ("KCl soil collection setup", "Flood - ", "Wash pots")
    for start in per_cohort:
        assert len(_named(t, start)) == 3, start
    assert len(_named(t, "N2O measurements")) == 3 * 2          # 24 h and 48 h per cohort
    for start in ("Plant harvest", "KCl extractions", "KCl analysis", "Plant dry weights"):
        assert len(_named(t, start)) == 3 * 2, start            # sacrificial + final
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
    assert all(x.date.weekday() < 5 for x in t if x.name.startswith(lab))
    for c in d.plan.cohorts:
        final = d.events(1, c)[-1][0]
        mine = [x for x in t if x.name.endswith(f"cohort {c.name}")]
        ext = [x.date for x in mine if x.name.startswith("KCl extractions")]
        ana = [x.date for x in mine if x.name.startswith("KCl analysis")]
        setup = [x.date for x in mine if x.name.startswith("KCl soil collection setup")]
        harvest = [x.date for x in mine if x.name.startswith("Plant harvest")]
        assert all(a > e for a, e in zip(ana, ext)) and all(e > h for e, h in zip(ext, harvest))
        assert setup[0] < min(harvest) and max(harvest) == final
        wash = [x.date for x in mine if x.name.startswith("Wash pots")]
        assert wash[0] > final


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
    assert rows[0] == ("Name", "Due Date", "Section/Column", "Description")
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


def test_asana_dates_calculate(tmp_path):
    for plan in (p0_plan(), big_plan()):
        d = dc.build(plan)
        path = tmp_path / "x.xlsx"
        dx.write(d, str(path))
        assert _asana_values(path) == _expected(d)


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
    want = _expected(d)
    assert want != _expected(dc.build(p0_plan()))                 # the edit moved things
    assert _asana_values(path) == want


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
                                                                       "asana_tasks", "jars",
                                                                       "info"]
    jars = wb["jars"]
    assert [c.value for c in jars[1]] == ["barcode", "tag"]
    rows = list(jars.iter_rows(min_row=2, values_only=True))
    assert len(rows) == len(d.pots)
    import re
    assert all(re.search(r"\bT\d+-Z\d+-\d+\b", tag) for _, tag in rows)  # station reads this
    assert jars["A2"].number_format == "@"
    design = wb["design"]
    assert design["H2"].value.startswith("=timeline!$F$")              # linked planting date


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
    assert fl["B2"].value.startswith("=timeline!")         # the date follows the timeline
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


def test_layout_is_bright_with_a_legend(tmp_path):
    import openpyxl
    for plan, by in ((p0_plan(), "cohort"), (big_plan(), "group")):
        d = dc.build(plan)
        path = tmp_path / "l.xlsx"
        dx.write(d, str(path))
        ws = openpyxl.load_workbook(path)["layout"]
        assert ws["A3"].value == f"Colour = {'planting group' if by == 'group' else 'cohort'}"
        got_by, colours = dx.layout_colours(d)
        assert got_by == by
        legend = {ws.cell(r, 1).fill.start_color.rgb[-6:]: ws.cell(r, 2).value
                  for r in range(4, 4 + len(colours))}
        assert set(legend) == {h.upper() for h in colours.values()}
        fills = {c.fill.start_color.rgb[-6:] for row in ws.iter_rows(min_row=4 + len(colours))
                 for c in row if c.value and "\n" in str(c.value)}
        assert fills <= set(legend) | {dx.SPARE_BG}


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
    cols = [(g, c) for g in range(1, d.groups + 1) for c in d.plan.cohorts]
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
    d = dc.build(big_plan())
    path = tmp_path / "x.xlsx"
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
        for j, (g, c) in enumerate((g, c) for g in range(1, d.groups + 1) for c in d.plan.cohorts):
            text = val(f"TIMELINE'!{L(3 + j)}{r}")
            assert text == d.day_text(g, c, day), (day, g, c.name, text, d.day_text(g, c, day))
    assert {k: v for k, v in got.items() if v} == d.calendar_load()
    flood_col = tot_col - 1
    assert all(int(val(f"TIMELINE'!{L(flood_col)}{r}")) == 0
               or int(val(f"TIMELINE'!{L(tot_col)}{r}")) == 0
               for r in range(start, last + 1))           # never both on one day


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
