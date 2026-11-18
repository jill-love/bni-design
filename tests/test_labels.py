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
from test_design import SOIL, p0_plan, p1_plan


def _texts(z):
    """[(x, height, text, reversed)] of a label's text fields."""
    return [(int(x), int(h), t.removesuffix("\\&"), bool(rv)) for x, h, rv, t in
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


def test_centred_text_ends_with_a_line_break():
    """Zebra printers only centre a field block's line when it ends with \\&."""
    p = dc.build(p0_plan()).pots[0]
    blocks = re.findall(r"\^FB\d+,1,0,C\^A0N,\d+,\d+(?:\^FR)?\^FD(.*?)\^FS", LB.zpl_label(p))
    assert len(blocks) == 3 and all(t.endswith("\\&") for t in blocks)


def test_the_badge_says_cohort_and_flood_order():
    d = dc.build(p0_plan(spares_per_genotype_per_group=1))
    for role, want in ((dc.MEASURED, None), (dc.SACRIFICIAL, "SAC"), (dc.SPARE, "SPARE")):
        p = next(p for p in d.pots if p.role == role)
        badge = [t for _, _, t, rv in _texts(LB.zpl_label(p)) if rv]   # white on black
        if role == dc.MEASURED:
            assert badge == [p.cohort, f"#{p.seq}"]
        elif role == dc.SACRIFICIAL:
            assert badge == [p.cohort, f"SAC {p.seq}"]           # its harvest order
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


FILES = ["labels_PLANTING.zpl", "labels_HARVEST_SHOOT.zpl", "labels_HARVEST_ROOT.zpl",
         "labels_KCL_COLLECT.zpl", "labels_KCL_WEIGHED.zpl", "labels_KCL_TIN.zpl",
         "labels_KCL_EXTRACT.zpl"]


def _split(z):
    return [x + "^XZ" for x in z.split("^XZ") if x]


def _qr(label):
    return re.search(r"\^BQN,2,\d+\^FDQA,(.*?)\^FS", label)[1]


def test_seven_files_in_a_folder_beside_the_workbook(tmp_path):
    d = dc.build(p1_plan())
    book = tmp_path / "BNI-P1_design_map.xlsx"
    folder = LB.folder_for(str(book))
    assert folder == str(tmp_path / "labels_BNI-P1_design_map")
    made = LB.write(d, folder)
    assert sorted(made) == sorted(FILES) and sorted(os.listdir(folder)) == sorted(FILES)
    harvested = sum(p.role in (dc.MEASURED, dc.SACRIFICIAL) for p in d.pots)
    assert made["labels_PLANTING.zpl"] == len(d.pots) == 144          # spares too
    assert all(made[f] == harvested == 126 for f in FILES[1:])        # no spares
    for name in FILES:
        with open(os.path.join(folder, name), encoding="utf-8") as fh:
            assert fh.read().count("^XA") == made[name]


def test_planting_labels_go_by_planting_group_then_position():
    d = dc.build(p1_plan())
    by_code = {p.barcode: p for p in d.pots}
    order = [by_code[_qr(z)] for z in _split(LB.files(d)["labels_PLANTING.zpl"][0])]
    keys = [(p.group, p.table, p.zone, p.slot) for p in order]
    assert keys == sorted(keys)


def test_harvest_labels_go_by_harvest_date_then_flood_order():
    d = dc.build(p1_plan())
    by_code = {p.barcode: p for p in d.pots}
    fs = LB.files(d)
    orders = {}
    for name, suffix in zip(FILES[1:], ("SH", "RT", "K1", "K2", "K3", "K4")):
        codes = [_qr(z) for z in _split(fs[name][0])]
        assert all(c.endswith(f"-{suffix}") for c in codes)
        orders[name] = [c.rsplit("-", 1)[0] for c in codes]
    assert len({tuple(o) for o in orders.values()}) == 1                # same order in all six
    pots = [by_code[b] for b in orders[FILES[1]]]
    dates = [LB.harvest_date(d, p) for p in pots]
    assert dates == sorted(dates)
    for day in set(dates):                      # within a day and flood: flood order 1, 2, 3...
        for g, c in {(p.group, p.cohort) for p, x in zip(pots, dates) if x == day}:
            same = [p for p, x in zip(pots, dates) if x == day and p.group == g and p.cohort == c]
            assert [p.seq for p in same] == sorted(p.seq for p in same)
    sac = next(p for p in pots if p.role == dc.SACRIFICIAL)
    assert LB.harvest_date(d, sac) == d.anchor_date(("flood", sac.group, sac.cohort))


def test_bag_and_tube_badges():
    d = dc.build(p1_plan())
    fs = LB.files(d)
    want = {"labels_HARVEST_SHOOT.zpl": ["SHOOT", "HARVEST", "date"],
            "labels_HARVEST_ROOT.zpl": ["ROOT", "HARVEST", "date"],
            "labels_KCL_COLLECT.zpl": ["KCl", "COLLECT", "date"],
            "labels_KCL_WEIGHED.zpl": ["KCl", "WEIGHED"],          # no date after collection
            "labels_KCL_TIN.zpl": ["KCl", "TIN"],
            "labels_KCL_EXTRACT.zpl": ["KCl", "EXTRACT"]}
    by_code = {p.barcode: p for p in d.pots}
    for name, lines in want.items():
        for z in _split(fs[name][0])[:5]:
            p = by_code[_qr(z).rsplit("-", 1)[0]]
            day = LB.harvest_date(d, p)
            badge = [t for _, _, t, rv in _texts(z) if rv]
            assert badge == [f"{day.day} {day:%b}".upper() if x == "date" else x for x in lines]
            texts = {t for _, _, t, _ in _texts(z)}
            order = f"SAC {p.seq}" if p.role == dc.SACRIFICIAL else f"#{p.seq}"
            assert {p.barcode, f"{p.cohort} {order}", p.genotype, p.treatment} <= texts
            assert {"POT", "COHORT", "GENOTYPE", "TREATMENT"} <= texts


def test_sacrificial_pots_are_numbered_by_harvest_order():
    """Each flood's sacrificial pots are SAC 1, SAC 2, ... in their harvest order on
    flood day, on the pot label and on every bag and tube label."""
    d = dc.build(p1_plan())
    fs = LB.files(d)
    for g, c in d.cells():
        sac = sorted((p for p in d.pots if p.group == g and p.cohort == c.name
                      and p.role == dc.SACRIFICIAL), key=lambda p: p.seq)
        assert [p.seq for p in sac] == list(range(1, len(sac) + 1)) and len(sac) == 6
        for p in sac:
            assert f"^FDSAC {p.seq}\\&^FS" in LB.zpl_label(p)
            for name in FILES[1:]:
                assert f"^FD{p.cohort} SAC {p.seq}^FS" in fs[name][0]


def test_bag_and_tube_text_stays_inside_its_cell():
    import datetime as dt
    long = dc.Pot(barcode="BNI-P0.1-S001", role=dc.MEASURED, genotype="Mo17xB73xW22-long",
                  cohort="AB", treatment="NITRAPYRIN", group=12, seq=103, table=12, zone=8, slot=25)
    for _, suffix, badge in LB.SAMPLES:
        z = LB.sample_label(long, suffix, badge, dt.date(2026, 12, 22))
        for x, h, t, rv in _texts(z):
            right = 420 if rv else (226 if x < 230 else 900)          # badge, QR text, form
            width = 0.7 if x < 230 else LB.CHAR_W
            assert x + width * h * len(t) <= right + 2, (suffix, t)


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
    assert os.path.exists(book)
    assert sorted(os.listdir(tmp_path / "labels_map")) == sorted(FILES)
    assert "7 label files" in w.banner.text()
    w.close()
