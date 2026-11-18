"""Pot labels (labels.py): QR code labels, written with the design map.
Run:  python -m pytest -q tests"""
import os
import re
import sys

import pytest

HERE = os.path.dirname(os.path.abspath(__file__)); R = os.path.dirname(HERE)
sys.path.insert(0, R); sys.path.insert(0, HERE)
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
import design_core as dc
import labels as LB
from test_design import SOIL, p0_plan


def _texts(z):
    """[(x, height, text, reversed)] of a label's text fields."""
    return [(int(x), int(h), t, bool(rv)) for x, h, rv, t in
            re.findall(r"\^FO(\d+),\d+(?:\^FB\d+,1,0,[LC])?\^A0N,(\d+),\d+(\^FR)?\^FD(.*?)\^FS", z)]


def test_one_label_per_station_pot():
    d = dc.build(p0_plan(**SOIL, spares_per_genotype_per_group=1))
    z = LB.zpl(d)
    pots = [p for p in d.pots if p.active]
    assert z.count("^XA") == z.count("^XZ") == z.count("^BQN") == len(pots)
    assert "^PW900^LL300" in z                                  # 3 x 1 in at 300 dpi
    for p in pots:
        one = LB.zpl_label(p)
        assert f"^FDQA,{p.barcode}^FS" in one                    # the QR code
        texts = {t for _, _, t, _ in _texts(one)}
        assert {p.barcode, p.position, str(p.group), p.genotype} <= texts
        assert {"POSITION", "PG", "GENOTYPE"} <= texts


def test_the_badge_says_cohort_and_flood_order():
    d = dc.build(p0_plan(spares_per_genotype_per_group=1))
    for role, want in ((dc.MEASURED, None), (dc.SACRIFICIAL, "SAC"), (dc.SPARE, "SPARE")):
        p = next(p for p in d.pots if p.role == role)
        badge = [t for _, _, t, rv in _texts(LB.zpl_label(p)) if rv]   # white on black
        if role == dc.MEASURED:
            assert badge == [p.cohort, f"#{p.seq}"]
        elif role == dc.SACRIFICIAL:
            assert badge == [p.cohort, "SAC"]
        else:
            assert badge == ["SPARE"]
        texts = {t for _, _, t, _ in _texts(LB.zpl_label(p))}
        assert ("TREATMENT" in texts) == (role != dc.SPARE)
        if role != dc.SPARE:
            assert p.treatment in texts


def test_flood_order_is_the_flood_sheets():
    d = dc.build(p0_plan())
    for g, c in d.cells():
        mine = sorted((p for p in d.pots if p.group == g and p.cohort == c.name
                       and p.role == dc.MEASURED), key=lambda p: p.seq)
        assert [p.seq for p in mine] == list(range(1, len(mine) + 1))


def test_positions_say_s_for_slot_but_the_station_tag_keeps_its_form(tmp_path):
    import openpyxl
    import design_xlsx as dx
    d = dc.build(p0_plan())
    p = d.pots[0]
    assert p.position == f"T{p.table}-Z{p.zone}-S{p.slot}"
    assert p.station_position == f"T{p.table}-Z{p.zone}-{p.slot:02d}"
    path = str(tmp_path / "x.xlsx")
    dx.write(d, path)
    wb = openpyxl.load_workbook(path)
    jars = {r[0]: r[1] for r in wb["jars"].iter_rows(min_row=2, values_only=True)}
    assert all(re.search(r"\bT\d+-Z\d+-\d+\b", t) for t in jars.values())   # station reads
    design = list(wb["design"].iter_rows(values_only=True))
    col = design[0].index("position")
    assert all(re.fullmatch(r"T\d+-Z\d+-S\d+", r[col]) for r in design[1:])


def test_the_text_stays_inside_its_cell():
    long = dc.Pot(barcode="BNI-P0.1-S001", role=dc.MEASURED, genotype="Mo17xB73xW22-long",
                  cohort="AB", treatment="NITRAPYRIN", group=12, seq=103, table=12, zone=8,
                  slot=25)
    for p in (long,):
        for x, h, t, _ in _texts(LB.zpl_label(p)):
            assert x + LB.CHAR_W * h * len(t) <= 900, t


def test_203_dpi_scales_the_label():
    p = dc.build(p0_plan()).pots[0]
    z = LB.zpl_label(p, dpi=203)
    assert "^PW609^LL203" in z and "^BQN,2,5" in z


def test_written_beside_the_workbook(tmp_path):
    d = dc.build(p0_plan())
    book = tmp_path / "BNI-P0_design_map.xlsx"
    path = LB.path_for(str(book))
    assert path == str(tmp_path / "labels_BNI-P0_design_map.zpl")
    assert LB.write(d, path) == sum(p.active for p in d.pots)
    with open(path, encoding="utf-8") as f:
        assert f.read() == LB.zpl(d)


def test_saving_the_workbook_writes_the_labels(tmp_path, monkeypatch):
    from PySide6 import QtWidgets
    import design_app
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    w = design_app.DesignWindow(p0_plan())
    w._generate()
    book = str(tmp_path / "map.xlsx")
    monkeypatch.setattr(QtWidgets.QFileDialog, "getSaveFileName",
                        staticmethod(lambda *a, **k: (book, "")))
    w._save_workbook()
    assert os.path.exists(book) and os.path.exists(tmp_path / "labels_map.zpl")
    assert "pot labels" in w.banner.text()
    w.close()
