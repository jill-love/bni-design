"""Pot swaps (swaps.py, swap_app.py): read a design map back, swap, rebuild, keep typed data.
Run:  python -m pytest -q tests"""
import datetime as dt
import os
import sys

import openpyxl
import pytest

HERE = os.path.dirname(os.path.abspath(__file__)); R = os.path.dirname(HERE)
sys.path.insert(0, R); sys.path.insert(0, HERE)
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
import design_core as dc
import design_xlsx as dx
import swaps as S
from test_design import SOIL, big_plan, p0_plan

DAY = dt.date(2026, 11, 10)


def _key(d):
    return [(p.barcode, p.role, p.cohort, p.treatment, p.rep, p.group, p.table, p.zone,
             p.slot, p.load, p.column, p.seq) for p in d.pots]


def _map(tmp_path, plan=None):
    d = dc.build(plan or p0_plan(**SOIL))
    path = str(tmp_path / "map.xlsx")
    dx.write(d, path)
    return d, path


def _first(d, **kw):
    return next(p for p in d.pots if all(getattr(p, k) == v for k, v in kw.items()))


def test_a_design_map_reads_back_exactly(tmp_path):
    for plan in (p0_plan(**SOIL), big_plan(spares_per_genotype_per_group=1)):
        d0, path = _map(tmp_path, plan)
        d, swaps = S.read_design_map(path)
        assert _key(d) == _key(d0) and swaps == []
        assert (d.groups, d.stagger_days, d.offsets, d.load) == \
               (d0.groups, d0.stagger_days, d0.offsets, d0.load)
        assert d.plan == d0.plan


def test_older_maps_read_from_info_and_admin(tmp_path):
    d0, path = _map(tmp_path)
    wb = openpyxl.load_workbook(path)
    del wb[dx.PLAN_SHEET]                          # a map from before the hidden plan tab
    wb.save(path)
    d, _ = S.read_design_map(path)
    assert _key(d) == _key(d0) and d.offsets == d0.offsets
    assert [(c.name, c.first_n2o_dap, c.measure_hours) for c in d.plan.cohorts] == \
           [(c.name, c.first_n2o_dap, c.measure_hours) for c in d0.plan.cohorts]
    assert d.plan.wfps_target == SOIL["wfps_target"]


def test_the_spare_takes_over_the_failed_pot(tmp_path):
    d, path = _map(tmp_path)
    f = _first(d, role=dc.MEASURED, cohort="B")
    old = (f.label, f.role, f.cohort, f.treatment, f.rep, f.load, f.column, f.seq)
    s = S.eligible_spares(d, f)[0]
    sw = S.swap(d, f.barcode, s.barcode, DAY, "no germination")
    assert (s.label, s.role, s.cohort, s.treatment, s.rep, s.load, s.column, s.seq) == old
    assert sw == dc.Swap(DAY, f.barcode, s.barcode, old[0], "no germination")
    assert f.role == dc.FAILED and f.was == dc.MEASURED and f.label == old[0] + "-FAILED"
    assert (f.swap, s.swap) == (s.barcode, f.barcode)
    assert (s.table, s.zone, s.slot) != (f.table, f.zone, f.slot)   # nobody moves


@pytest.mark.parametrize("case, words", [
    ("unknown", "not a pot"), ("spare as failed", "not one in the experiment"),
    ("not a spare", "not an unused spare"), ("twice", "already failed"),
    ("spare used", "already replaces"), ("other genotype", "own genotype"),
    ("other group", "own age")])
def test_bad_swaps_are_refused(tmp_path, case, words):
    d, _ = _map(tmp_path, big_plan(spares_per_genotype_per_group=1))
    f = _first(d, role=dc.MEASURED, genotype="B73", group=1)
    spare = S.eligible_spares(d, f)[0]
    other = _first(d, role=dc.MEASURED, genotype="B73", group=2)
    args = {"unknown": ("NOPE", spare.barcode),
            "spare as failed": (spare.barcode, spare.barcode),
            "not a spare": (f.barcode, other.barcode),
            "other genotype": (f.barcode, _first(d, role=dc.SPARE, genotype="W22",
                                                 group=1).barcode),
            "other group": (f.barcode, _first(d, role=dc.SPARE, genotype="B73",
                                              group=2).barcode)}
    if case in ("twice", "spare used"):
        S.swap(d, f.barcode, spare.barcode)
        args[case] = ((f.barcode, _first(d, role=dc.SPARE, genotype="B73",
                                         group=2).barcode) if case == "twice"
                      else (other.barcode, spare.barcode))
    with pytest.raises(S.SwapError, match=words):
        S.swap(d, *args[case])


def test_rebuilt_map_follows_the_swap(tmp_path):
    d, path = _map(tmp_path)
    f = _first(d, role=dc.SACRIFICIAL, cohort="A")
    m = _first(d, role=dc.MEASURED, cohort="C")
    s1, s2 = S.eligible_spares(d, f)[:2]
    swaps = [S.swap(d, f.barcode, s1.barcode, DAY), S.swap(d, m.barcode, s2.barcode, DAY)]
    gone = S.discard_spares(d)
    rep = S.rebuild(path, d, swaps)
    assert os.path.exists(rep["backup"])
    wb = openpyxl.load_workbook(path)
    jars = [r[0] for r in wb["jars"].iter_rows(min_row=2, values_only=True)]
    assert s1.barcode in jars and s2.barcode in jars
    assert not {f.barcode, m.barcode, *gone} & set(jars)            # the station drops them
    assert len(jars) == sum(p.role in (dc.MEASURED, dc.SACRIFICIAL) for p in d.pots)
    flood = [r[1] for r in wb["flood_C"].iter_rows(values_only=True)]
    assert s2.barcode in flood and m.barcode not in flood
    assert s1.barcode in [r[1] for r in wb["flood_A"].iter_rows(values_only=True)]
    measure = [r[3] for r in wb["measure_C"].iter_rows(values_only=True)]
    assert measure.count(s2.barcode) == len(d.plan.cohort("C").measure_hours)
    assert s2.barcode in [r[1] for r in wb["harvest_C"].iter_rows(values_only=True)]
    sw = list(wb["swaps"].iter_rows(min_row=2, values_only=True))
    assert [(r[1], r[2], r[3]) for r in sw] == [(x.failed, x.spare, x.role) for x in swaps]
    layout = {str(c.value).split("\n")[1].split()[0]: c for row in wb["layout"].iter_rows()
              for c in row if c.value and "\n" in str(c.value)}
    assert layout[m.barcode].font.strike and layout[s2.barcode].font.b
    assert f"(for {m.barcode[-4:]})" in layout[s2.barcode].value
    d2, sw2 = S.read_design_map(path)                                # and it reads back
    assert _key(d2) == _key(d) and sw2 == swaps
    assert {p.barcode for p in d2.pots if p.role == dc.DISCARDED} == set(gone)


def test_a_second_session_keeps_the_first(tmp_path):
    d, path = _map(tmp_path)
    f1, f2 = [p for p in d.pots if p.role == dc.MEASURED][:2]
    S.rebuild(path, d, [S.swap(d, f1.barcode, S.eligible_spares(d, f1)[0].barcode, DAY)])
    d, swaps = S.read_design_map(path)
    swaps.append(S.swap(d, f2.barcode, S.eligible_spares(d, f2)[0].barcode, DAY))
    S.rebuild(path, d, swaps)
    d3, sw3 = S.read_design_map(path)
    assert [x.failed for x in sw3] == [f1.barcode, f2.barcode]
    assert sum(p.role == dc.FAILED for p in d3.pots) == 2


def test_typed_values_and_added_tabs_are_kept(tmp_path):
    d, path = _map(tmp_path)
    wb = openpyxl.load_workbook(path)
    fl = wb["flood_A"]
    hdr = [r for r in range(1, 80) if "barcode" in [c.value for c in fl[r]]]
    row = hdr[-1] + 1
    keep_bc = fl.cell(row, 2).value
    fl.cell(row, 6).value = 412.5                                   # mass before g
    ms = wb["measure_B"]
    ms["B4"] = 23.5                                                 # chamber air T line
    tl = wb["timeline"]
    tl["B4"] = dt.datetime(2026, 11, 9)                             # planting date edited
    wb["admin"]["B6"] = 1.31                                        # bulk density
    extra = wb.create_sheet("Dry shoots")
    extra["A1"], extra["B2"] = "Sample", 1.194
    wb.save(path)
    f = _first(d, role=dc.MEASURED, cohort="C")
    rep = S.rebuild(path, d, [S.swap(d, f.barcode, S.eligible_spares(d, f)[0].barcode, DAY)])
    assert rep["sheets"] == ["Dry shoots"] and rep["cells"] >= 2 and rep["inputs"] >= 2
    wb = openpyxl.load_workbook(path)
    fl = wb["flood_A"]
    r = next(r for r in range(1, 80) if fl.cell(r, 2).value == keep_bc)
    assert fl.cell(r, 6).value == 412.5
    assert wb["measure_B"]["B4"].value == 23.5
    assert wb["timeline"]["B4"].value == dt.datetime(2026, 11, 9)
    assert wb["admin"]["B6"].value == 1.31
    assert (wb["Dry shoots"]["A1"].value, wb["Dry shoots"]["B2"].value) == ("Sample", 1.194)


def test_typed_values_follow_the_barcode_not_the_row(tmp_path):
    d, path = _map(tmp_path)
    f = _first(d, role=dc.MEASURED, cohort="A")
    wb = openpyxl.load_workbook(path)
    fl = wb["flood_A"]
    r = next(r for r in range(1, 80) if fl.cell(r, 2).value == f.barcode)
    fl.cell(r, 6).value = 999                                       # typed on the failed pot
    wb.save(path)
    s = S.eligible_spares(d, f)[0]
    S.rebuild(path, d, [S.swap(d, f.barcode, s.barcode, DAY)])
    fl = openpyxl.load_workbook(path)["flood_A"]
    r = next(r for r in range(1, 80) if fl.cell(r, 2).value == s.barcode)
    assert fl.cell(r, 6).value is None                              # not handed to the spare


def test_dosing_station_reloads_the_swapped_map(tmp_path):
    dosing = os.path.join(R, "..", "bni-dosing")
    if not os.path.exists(os.path.join(dosing, "trials.py")):
        pytest.skip("bni-dosing not beside bni-design")
    sys.path.insert(0, dosing)
    try:
        import provision_jars as P
        import run_order as ro
    finally:
        sys.path.remove(dosing)
    d, path = _map(tmp_path)
    f = _first(d, role=dc.MEASURED, cohort="A")
    s = S.eligible_spares(d, f)[0]
    turn = f.seq
    S.rebuild(path, d, [S.swap(d, f.barcode, s.barcode, DAY)])
    jars, _, admin = P.read_jars(path)
    assert s.barcode in {j.barcode for j in jars} and f.barcode not in {j.barcode for j in jars}
    assert P.admin_profile(admin) == "wet_up"
    order = ro.read_experiment(path).order("flood_A")
    codes = [i.barcode for i in order.items]
    assert s.barcode in codes and f.barcode not in codes
    assert codes.index(s.barcode) == turn - 1                       # the failed pot's turn


def test_window_adds_removes_and_saves(tmp_path):
    from PySide6 import QtWidgets
    import swap_app
    d, path = _map(tmp_path)
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    w = swap_app.SwapWindow(path)
    f1, f2 = [p for p in w.design.pots if p.role == dc.MEASURED][:2]
    for f in (f1, f2):
        w.failed.setText(f.barcode); w._failed_entered(focus=True)
        assert w.spare.count() >= 1
        w.spare.setCurrentIndex(0); w._add()
    assert len(w.pending) == 2
    w.table.selectRow(1); w._remove()
    assert [s.failed for s in w.pending] == [f1.barcode]
    lock = os.path.join(str(tmp_path), "~$map.xlsx")
    open(lock, "w").close()                                         # open in Excel
    orig = QtWidgets.QMessageBox.warning
    QtWidgets.QMessageBox.warning = lambda *a, **k: None
    QtWidgets.QMessageBox.information = lambda *a, **k: None
    try:
        w._save()
        assert S.read_design_map(path)[1] == []                     # refused
        os.remove(lock)
        w._save()
    finally:
        QtWidgets.QMessageBox.warning = orig
    assert [s.failed for s in S.read_design_map(path)[1]] == [f1.barcode]
    assert w.saved and not w.pending
    w.close()


# ---------------------------------------------------------------- planting group = date
def _late_map(tmp_path):
    """Cohort B collides with A: it is planted 2 days later, as planting group 2."""
    plan = p0_plan(cohorts=[dc.Cohort("A", 21), dc.Cohort("B", 22)],
                   rows=[dc.PlanRow(g, c, "CTRL", 3, 1) for g in ("B73", "Mo17") for c in "AB"],
                   spares_per_genotype_per_group=2)
    return _map(tmp_path, plan)


def test_a_spare_only_replaces_a_pot_of_its_planting_group(tmp_path):
    d, path = _late_map(tmp_path)
    assert d.offsets == {"A": 0, "B": 2} and d.groups == 2
    on_day = _first(d, role=dc.MEASURED, cohort="A", genotype="B73")
    late = _first(d, role=dc.MEASURED, cohort="B", genotype="B73")
    assert (on_day.group, late.group) == (1, 2)
    early_spare = _first(d, role=dc.SPARE, genotype="B73", group=1)
    late_spare = _first(d, role=dc.SPARE, genotype="B73", group=2)
    assert S.planted(d, early_spare) == S.planted(d, on_day) != S.planted(d, late)
    assert S.planted(d, late_spare) == S.planted(d, late)
    assert {p.group for p in S.eligible_spares(d, late)} == {2}
    assert {p.group for p in S.eligible_spares(d, on_day)} == {1}
    with pytest.raises(S.SwapError, match="planted .* own age"):
        S.swap(d, late.barcode, early_spare.barcode)
    assert late.role == dc.MEASURED and early_spare.role == dc.SPARE   # nothing changed
    S.swap(d, late.barcode, late_spare.barcode)
    S.swap(d, on_day.barcode, early_spare.barcode)


def test_every_planting_group_has_its_own_spares(tmp_path):
    d, path = _late_map(tmp_path)                       # 2 per genotype per group
    sp = [p for p in d.pots if p.role == dc.SPARE]
    assert len(sp) == 2 * 2 * 2                         # genotypes x groups x 2
    assert sorted({p.label.split("-SPARE")[0] for p in sp}) == ["B73-G1", "B73-G2",
                                                                "Mo17-G1", "Mo17-G2"]
    d2, _ = S.read_design_map(path)
    assert {(p.barcode, p.group) for p in d2.pots if p.role == dc.SPARE} == \
           {(p.barcode, p.group) for p in sp}


def test_a_spare_only_replaces_its_own_genotype(tmp_path):
    d, _ = _late_map(tmp_path)
    f = _first(d, role=dc.MEASURED, cohort="A", genotype="B73")
    assert all(p.genotype == "B73" for p in S.eligible_spares(d, f))
    with pytest.raises(S.SwapError, match="own genotype"):
        S.swap(d, f.barcode, _first(d, role=dc.SPARE, genotype="Mo17", group=1).barcode)


def test_older_maps_with_two_planting_days_per_group_are_split(tmp_path):
    """A map made when a group could plant on two days (B later, spares 'planted with')."""
    d, path = _late_map(tmp_path)
    wb = openpyxl.load_workbook(path)
    del wb[dx.PLAN_SHEET]
    info = wb["info"]
    for row in info.iter_rows():                        # the old info layout
        if row[0].value == "planting rounds":
            row[0].value = "planting groups"
        elif row[0].value == "days between rounds":
            row[0].value = "days between groups"
        elif row[0].value == "planting groups (one planting date each)":
            row[0].value = "planting groups (old)"
    ds = wb["design"]
    head = [c.value for c in ds[1]]
    gcol, ccol = head.index("planting_group") + 1, head.index("cohort") + 1
    ds.cell(1, len(head) + 1, "planted_with")
    for r in range(2, ds.max_row + 1):                  # back to one group, two days
        if ds.cell(r, gcol).value == 2:
            ds.cell(r, gcol).value = 1
            if not ds.cell(r, ccol).value:
                ds.cell(r, len(head) + 1, "B")
    wb.save(path)
    d2, _ = S.read_design_map(path)
    assert d2.groups == 2
    assert {(p.barcode, p.group) for p in d2.pots} == {(p.barcode, p.group) for p in d.pots}


def test_window_lists_only_spares_of_that_planting_group(tmp_path):
    from PySide6 import QtWidgets
    import swap_app
    d, path = _late_map(tmp_path)
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    w = swap_app.SwapWindow(path)
    late = _first(w.design, role=dc.MEASURED, cohort="B", genotype="B73")
    w.failed.setText(late.barcode); w._failed_entered(focus=True)
    listed = [w.spare.itemData(i) for i in range(w.spare.count())]
    want = {p.barcode for p in w.design.pots if p.role == dc.SPARE and p.genotype == "B73"
            and p.group == late.group}
    assert set(listed) == want and "planted" in w.failed_info.text()
    early = _first(w.design, role=dc.SPARE, genotype="B73", group=1)
    w.spare.setEditText(early.barcode); w._add()                     # scanned anyway
    assert w.pending == [] and "own age" in w.banner.text()
    w.close()


def test_discard_keeps_spares_of_a_later_planting_group(tmp_path):
    d, _ = _late_map(tmp_path)                       # group 1 day 0, group 2 day 2
    first = min(S.planted(d, p) for p in d.pots)
    gone = S.discard_spares(d, first + dt.timedelta(days=8))       # group 1's: 8 days old
    kept = [p for p in d.pots if p.role == dc.SPARE]
    assert gone and all(p.group == 1 for p in d.pots if p.barcode in gone)
    assert kept and all(p.group == 2 for p in kept)                 # group 2's: 6 days old
    assert S.discard_spares(d, first + dt.timedelta(days=9))        # a day later they go
    assert not [p for p in d.pots if p.role == dc.SPARE]


# ---------------------------------------------------------------- spares scale
def test_spares_scale_with_each_groups_pots():
    import math
    plan = p0_plan(cohorts=[dc.Cohort("A", 21), dc.Cohort("B", 22)],
                   rows=[dc.PlanRow("B73", "A", "CTRL", 30, 9), dc.PlanRow("B73", "B", "CTRL", 15, 5),
                         dc.PlanRow("Mo17", "A", "CTRL", 4)],
                   spares_per_genotype_per_group=0, spare_percent=10)
    d = dc.build(plan)
    from collections import Counter
    n = Counter((p.group, p.genotype) for p in d.pots if p.role in (dc.MEASURED, dc.SACRIFICIAL))
    sp = Counter((p.group, p.genotype) for p in d.pots if p.role == dc.SPARE)
    assert sp == {k: max(1, math.ceil(v / 10)) for k, v in n.items()}
    assert sp[(1, "B73")] == 4 and sp[(2, "B73")] == 2 and sp[(1, "Mo17")] == 1
    assert (2, "Mo17") not in sp                      # no Mo17 in group 2: no Mo17 spare


def test_edited_cohort_treatment_days_are_kept(tmp_path):
    d, path = _map(tmp_path)
    wb = openpyxl.load_workbook(path)
    top = dx.TIMELINE_TOP
    wb["timeline"].cell(top + 2, 4).value = 16                      # cohort B's own day
    wb.save(path)
    f = _first(d, role=dc.MEASURED, cohort="C")
    S.rebuild(path, d, [S.swap(d, f.barcode, S.eligible_spares(d, f)[0].barcode, DAY)])
    tl = openpyxl.load_workbook(path)["timeline"]
    assert tl.cell(top + 2, 4).value == 16
    assert str(tl.cell(top + 1, 4).value).startswith("=")          # A still the plan's day


def test_an_older_maps_planting_delays_are_not_taken_as_treatment_days(tmp_path):
    # maps made before planting groups were planting dates had the cohorts' planting
    # delays, also yellow inputs, where the treatment days are now
    d, path = _map(tmp_path)
    wb = openpyxl.load_workbook(path)
    tl, top = wb["timeline"], dx.TIMELINE_TOP
    tl.cell(top, 4).value = "Planted (days after group)"
    for i in range(3):
        tl.cell(top + 1 + i, 4).value = i
    wb.save(path)
    f = _first(d, role=dc.MEASURED, cohort="C")
    S.rebuild(path, d, [S.swap(d, f.barcode, S.eligible_spares(d, f)[0].barcode, DAY)])
    tl = openpyxl.load_workbook(path)["timeline"]
    assert tl.cell(top, 4).value == dx.TREAT_HEAD
    assert all(str(tl.cell(top + 1 + i, 4).value).startswith("=") for i in range(3))


# ---------------------------------------------------------------- failed, no spare
def test_a_failed_pot_with_no_spare_leaves_the_experiment(tmp_path):
    d, path = _map(tmp_path)
    f = _first(d, role=dc.MEASURED, cohort="A")
    sw = S.fail(d, f.barcode, DAY, "no germination")
    assert f.role == dc.FAILED and f.was == dc.MEASURED and not f.swap and not f.seq
    assert sw.spare == "" and sw.failed == f.barcode
    S.rebuild(path, d, [sw])
    d2, sws = S.read_design_map(path)
    assert [(s.failed, s.spare, s.note) for s in sws] == [(f.barcode, "", "no germination")]
    assert S.pot(d2, f.barcode).role == dc.FAILED
    wb = openpyxl.load_workbook(path)
    assert f.barcode not in {r[0] for r in wb["jars"].iter_rows(values_only=True)}
    import labels
    assert f.barcode not in {p.barcode for p in labels.harvested(d2)}
    with pytest.raises(S.SwapError, match="already failed"):
        S.fail(d2, f.barcode)
    with pytest.raises(S.SwapError, match="only a measured or sacrificial"):
        S.fail(d2, _first(d2, role=dc.SPARE).barcode)


# ---------------------------------------------------------------- spares: 5 working days
@pytest.mark.parametrize("weekday", range(7))
def test_spares_go_at_the_swap_session_whatever_day_they_were_planted(tmp_path, weekday):
    """The 'Run swaps' task is 5 working days after planting: a Sunday planting's session
    is the Friday, 5 days later, and its spares go then."""
    plant = DAY - dt.timedelta(days=DAY.weekday()) + dt.timedelta(days=weekday)
    d = dc.build(p0_plan(planting_date=plant, spares_per_genotype_per_group=2))
    session = dc.workday(plant, 5)
    task = next(t for t in d.tasks() if t.name.startswith("Run swaps"))
    assert task.date == session
    assert not S.discard_spares(d, session - dt.timedelta(days=1))
    assert S.discard_spares(d, session) and not [p for p in d.pots if p.role == dc.SPARE]


# ---------------------------------------------------------------- labels at a session
def _labels_dir(folder):
    return sorted(os.listdir(folder)) if folder else []


def test_a_session_relabels_spares_and_prints_the_groups_it_finished(tmp_path):
    import labels
    d, path = _late_map(tmp_path)                    # group 1 day 0, group 2 day 2
    before, _ = S.read_design_map(path)
    f = _first(d, role=dc.MEASURED, group=1)
    spare = S.eligible_spares(d, f)[0]
    session = [S.swap(d, f.barcode, spare.barcode, DAY)]
    first = min(S.planted(d, p) for p in d.pots)
    S.discard_spares(d, dc.workday(first, 5))        # group 1's session: group 2 not due
    S.rebuild(path, d, session)
    lab = S.write_labels(path, before, d, session, DAY)
    assert lab["groups"] == [1] and lab["relabel"] == 1
    assert _labels_dir(lab["session"]) == sorted(
        [labels.RELABEL] + [f"labels_{n}.zpl" for n, _, _ in labels.SAMPLES])
    with open(os.path.join(lab["session"], labels.RELABEL)) as fh:
        z = fh.read()
    assert z.count("^XA") == 1 and f"^FDQA,{spare.barcode}^FS" in z
    assert f"^FD{f.cohort}\\&^FS" in z and f"^FD#{spare.seq}\\&^FS" in z   # its new role
    with open(os.path.join(lab["session"], "labels_HARVEST_SHOOT.zpl")) as fh:
        shoot = fh.read()
    group1 = [p for p in labels.harvested(d) if p.group == 1]
    assert shoot.count("^XA") == len(group1)
    assert f"{spare.barcode}-SH" in shoot and f"{f.barcode}-SH" not in shoot
    assert all(f"{p.barcode}-SH" not in shoot for p in d.pots if p.group == 2)
    # the main folder follows the design map
    with open(os.path.join(lab["main"], "labels_PLANTING.zpl")) as fh:
        planting = fh.read()
    assert f"^FDQA,{f.barcode}^FS" not in planting
    assert z.split("\n", 2)[2].split("^XZ")[0] in planting          # the spare's new label


def test_a_session_with_spares_left_relabels_but_waits_to_print(tmp_path):
    d, path = _map(tmp_path)
    before, _ = S.read_design_map(path)
    f = _first(d, role=dc.MEASURED)
    session = [S.swap(d, f.barcode, S.eligible_spares(d, f)[0].barcode, DAY)]
    S.rebuild(path, d, session)
    lab = S.write_labels(path, before, d, session, DAY)
    assert lab["groups"] == [] and lab["relabel"] == 1
    import labels
    assert _labels_dir(lab["session"]) == [labels.RELABEL]
    # the next session discards the rest: now the group's labels print, once
    before, _ = S.read_design_map(path)
    S.discard_spares(d, DAY)
    S.rebuild(path, d, session)
    lab2 = S.write_labels(path, before, d, [], DAY)
    assert lab2["groups"] == [1] and lab2["relabel"] == 0 and lab2["session"] != lab["session"]
    assert labels.RELABEL not in _labels_dir(lab2["session"])
    # and after that, a failure with no spare prints nothing new
    before, _ = S.read_design_map(path)
    late = [S.fail(d, _first(d, role=dc.MEASURED).barcode, DAY)]
    S.rebuild(path, d, session + late)
    lab3 = S.write_labels(path, before, d, late, DAY)
    assert lab3["session"] is None and lab3["groups"] == []


def test_window_fails_a_pot_with_no_spare_and_writes_the_labels(tmp_path):
    from PySide6 import QtWidgets
    import swap_app
    d, path = _map(tmp_path)
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    w = swap_app.SwapWindow(path)
    w.day.setDate(DAY)
    f1, f2 = [p for p in w.design.pots if p.role == dc.MEASURED][:2]
    w.failed.setText(f1.barcode); w._failed_entered(focus=True)
    w.spare.setCurrentIndex(0); w._add()
    w.failed.setText(f2.barcode); w._fail()
    assert [s.spare == "" for s in w.pending] == [False, True]
    assert w.table.item(1, 3).text() == "(none)"
    w.discard_cb.setChecked(True)
    said = []
    QtWidgets.QMessageBox.information = lambda *a, **k: said.append(a[2])
    w._save()
    roles = {p.barcode: p.role for p in S.read_design_map(path)[0].pots}
    assert roles[f1.barcode] == roles[f2.barcode] == dc.FAILED
    assert dc.SPARE not in roles.values()
    assert "labels_RELABEL.zpl: 1 new pot label" in said[0] and "planting group 1" in said[0]
    assert any(n.startswith("map_swaps_") or "_swaps_" in n for n in os.listdir(tmp_path))
    w.close()
