"""Pot labels (labels.py): the dosing station's ZPL labels, written with the design map.
Run:  python -m pytest -q tests"""
import os
import sys

import pytest

HERE = os.path.dirname(os.path.abspath(__file__)); R = os.path.dirname(HERE)
sys.path.insert(0, R); sys.path.insert(0, HERE)
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
import design_core as dc
import labels as LB
from test_design import SOIL, p0_plan

DOSING = os.path.join(os.path.dirname(R), "bni-dosing")     # the station's code, if beside


def test_target_mass_is_the_stations_worked_example():
    # wfps.py's worked check: 500 g air-dry, 3 % moisture, bulk 1.30, particle 2.65,
    # WFPS 0.60, container 95 g -> target 694.6 g
    plan = p0_plan(air_dry_soil_g=500.0, air_dry_moisture=0.03, bulk_density=1.30,
                   particle_density=2.65, wfps_target=0.60, container_weight_g=95.0)
    assert f"{LB.target_mass(plan):.1f}" == "694.6"


@pytest.mark.parametrize("blank", ["air_dry_soil_g", "air_dry_moisture", "bulk_density",
                                   "wfps_target", "container_weight_g"])
def test_no_target_mass_while_a_value_is_blank(blank):
    assert LB.target_mass(p0_plan(**{**SOIL, blank: None})) is None


def test_one_label_per_station_pot():
    d = dc.build(p0_plan(**SOIL))
    z = LB.zpl(d)
    pots = [p for p in d.pots if p.active]
    assert z.count("^XA") == z.count("^XZ") == len(pots)
    for p in pots:
        assert f"^FD{p.barcode}^FS" in z
        assert f"^FD{(p.label + ' ' + p.position)[:24]}^FS" in z
    assert f"^FD{LB.target_mass(d.plan):.1f} g^FS" in z
    assert "^PW900^LL300" in z                                  # 3 x 1 in at 300 dpi


def test_labels_without_the_mass_when_soil_values_are_blank():
    z = LB.zpl(dc.build(p0_plan()))
    assert z.count("^XA") > 0 and " g^FS" not in z


def test_written_beside_the_workbook(tmp_path):
    d = dc.build(p0_plan(**SOIL))
    book = tmp_path / "BNI-P0_design_map.xlsx"
    path = LB.path_for(str(book))
    assert path == str(tmp_path / "labels_BNI-P0_design_map.zpl")
    n, target = LB.write(d, path)
    assert n == sum(p.active for p in d.pots) and target == LB.target_mass(d.plan)
    with open(path, encoding="utf-8") as f:
        assert f.read() == LB.zpl(d)


def test_same_labels_as_start_new_trial(tmp_path):
    """Byte for byte what the dosing station writes for the same design map, when its
    code is beside this folder (bni-dosing)."""
    if not os.path.isdir(DOSING):
        pytest.skip("bni-dosing is not beside bni-design")
    sys.path.insert(0, DOSING)
    try:
        P = pytest.importorskip("provision_jars")
    finally:
        sys.path.remove(DOSING)
    import design_xlsx as dx
    d = dc.build(p0_plan(**SOIL))
    book = str(tmp_path / "map.xlsx")
    dx.write(d, book)
    jars, _, admin = P.read_jars(book)
    res, _, tare = P.compute_from_admin(admin)
    for j in jars:
        j.airdry_tare_g, j.water_to_add_g = tare, res.water_to_add_g
    station = "".join(P.zpl_label(j, 300, 3.0, 1.0) for j in jars)
    assert LB.zpl(d) == station


def test_saving_the_workbook_writes_the_labels(tmp_path, monkeypatch):
    from PySide6 import QtWidgets
    import design_app
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
    for plan, words in ((p0_plan(**SOIL), "target"), (p0_plan(), "no target mass")):
        w = design_app.DesignWindow(plan)
        w._generate()
        book = str(tmp_path / "map.xlsx")
        monkeypatch.setattr(QtWidgets.QFileDialog, "getSaveFileName",
                            staticmethod(lambda *a, **k: (book, "")))
        w._save_workbook()
        assert os.path.exists(book) and os.path.exists(tmp_path / "labels_map.zpl")
        assert words in w.banner.text() and "pot labels" in w.banner.text()
        os.remove(tmp_path / "labels_map.zpl")
        w.close()
