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


# ---------------------------------------------------------------- planting date
def _late_map(tmp_path):
    """Cohort B collides with A and is planted 2 days after the group (and its spares)."""
    plan = p0_plan(cohorts=[dc.Cohort("A", 21), dc.Cohort("B", 22)],
                   rows=[dc.PlanRow(g, c, "CTRL", 3, 1) for g in ("B73", "Mo17") for c in "AB"],
                   spares_per_genotype_per_group=2)
    return _map(tmp_path, plan)


def test_a_spare_only_replaces_a_pot_planted_the_same_day(tmp_path):
    d, path = _late_map(tmp_path)
    assert d.offsets == {"A": 0, "B": 2}
    on_day = _first(d, role=dc.MEASURED, cohort="A", genotype="B73")
    late = _first(d, role=dc.MEASURED, cohort="B", genotype="B73")
    early_spare = _first(d, role=dc.SPARE, genotype="B73", planted_with="A")
    late_spare = _first(d, role=dc.SPARE, genotype="B73", planted_with="B")
    assert S.planted(d, early_spare) == S.planted(d, on_day) != S.planted(d, late)
    assert S.planted(d, late_spare) == S.planted(d, late)
    assert {p.planted_with for p in S.eligible_spares(d, late)} == {"B"}
    assert {p.planted_with for p in S.eligible_spares(d, on_day)} == {"A"}
    with pytest.raises(S.SwapError, match="planted .* own age"):
        S.swap(d, late.barcode, early_spare.barcode)
    assert late.role == dc.MEASURED and early_spare.role == dc.SPARE   # nothing changed
    S.swap(d, late.barcode, late_spare.barcode)
    S.swap(d, on_day.barcode, early_spare.barcode)


def test_every_planting_day_has_its_own_spares(tmp_path):
    d, path = _late_map(tmp_path)                       # 2 per genotype per planting day
    sp = [p for p in d.pots if p.role == dc.SPARE]
    assert len(sp) == 2 * 2 * 2                         # genotypes x days x 2
    assert sorted({p.label.split("-SPARE")[0] for p in sp}) == ["B73-A", "B73-B",
                                                                "Mo17-A", "Mo17-B"]
    d2, _ = S.read_design_map(path)
    assert {(p.barcode, p.planted_with) for p in d2.pots if p.role == dc.SPARE} == \
           {(p.barcode, p.planted_with) for p in sp}
    wb = openpyxl.load_workbook(path)
    tags = {r[0]: r[1] for r in wb["jars"].iter_rows(min_row=2, values_only=True)}
    assert all("-A-SPARE-" in tags[p.barcode] or "-B-SPARE-" in tags[p.barcode] for p in sp)


def test_a_spare_only_replaces_its_own_genotype(tmp_path):
    d, _ = _late_map(tmp_path)
    f = _first(d, role=dc.MEASURED, cohort="A", genotype="B73")
    assert all(p.genotype == "B73" for p in S.eligible_spares(d, f))
    with pytest.raises(S.SwapError, match="own genotype"):
        S.swap(d, f.barcode, _first(d, role=dc.SPARE, genotype="Mo17", planted_with="A").barcode)


def test_planting_delay_edited_in_the_timeline_counts(tmp_path):
    d, path = _late_map(tmp_path)
    wb = openpyxl.load_workbook(path)
    tl = wb["timeline"]
    row_b = next(r for r in range(1, 40) if tl.cell(r, 1).value == "B")
    tl.cell(row_b, 4).value = 0                                     # B planted with A after all
    wb.save(path)
    d2, _ = S.read_design_map(path)
    late = _first(d2, role=dc.MEASURED, cohort="B", genotype="B73")
    assert d2.offsets["B"] == 0 and S.eligible_spares(d2, late)


def test_window_lists_only_spares_planted_that_day(tmp_path):
    from PySide6 import QtWidgets
    import swap_app
    d, path = _late_map(tmp_path)
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    w = swap_app.SwapWindow(path)
    late = _first(w.design, role=dc.MEASURED, cohort="B", genotype="B73")
    w.failed.setText(late.barcode); w._failed_entered(focus=True)
    listed = [w.spare.itemData(i) for i in range(w.spare.count())]
    want = {p.barcode for p in w.design.pots if p.role == dc.SPARE and p.genotype == "B73"
            and p.planted_with == "B"}
    assert set(listed) == want and "planted" in w.failed_info.text()
    early = _first(w.design, role=dc.SPARE, genotype="B73", planted_with="A")
    w.spare.setEditText(early.barcode); w._add()                     # scanned anyway
    assert w.pending == [] and "own age" in w.banner.text()
    w.close()


def test_discard_keeps_spares_of_a_later_planting_day(tmp_path):
    d, _ = _late_map(tmp_path)                       # A planted day 0, B day 2 (with spares)
    first = min(S.planted(d, p) for p in d.pots)
    gone = S.discard_spares(d, first + dt.timedelta(days=8))       # A's spares 8 days old
    kept = [p for p in d.pots if p.role == dc.SPARE]
    assert gone and all(S.planted(d, d_p) == first for d_p in d.pots if d_p.barcode in gone)
    assert kept and all(p.planted_with == "B" for p in kept)         # B's: 6 days old
    assert S.discard_spares(d, first + dt.timedelta(days=9))        # a day later they go
    assert not [p for p in d.pots if p.role == dc.SPARE]
