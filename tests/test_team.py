"""Assignee / Collaborators for the Asana import (team.py, the assignments tab).
Run:  python -m pytest -q tests"""
import os
import sys

import openpyxl
import pytest

HERE = os.path.dirname(os.path.abspath(__file__)); R = os.path.dirname(HERE)
sys.path.insert(0, R); sys.path.insert(0, HERE)
import design_core as dc
import design_xlsx as dx
import swaps as S
import team as tm
from test_design import _asana_values, big_plan, p0_plan

PEOPLE = {"Jill": "jill@example.org", "Lina": "lina@example.org",
          "Sarah (Undergrad #1)": "sarah@example.org", "Undergrad #2": ""}


def _team():
    return tm.Team(dict(PEOPLE), dict(tm.DEFAULT_ROLES))


def test_team_file_round_trip(tmp_path):
    path = str(tmp_path / "team.xlsx")
    tm.write(_team(), path)
    back = tm.load(path)
    assert back.people == PEOPLE
    assert {k: v for k, v in back.roles.items() if v[0]} == tm.DEFAULT_ROLES
    assert tm.load(str(tmp_path / "missing.xlsx")) is None


def test_defaults_follow_the_work_breakdown():
    t = tm.defaults()
    assert t.role("prefill") == ("Sarah (Undergrad #1)", ("Lina",))
    assert t.role("wash") == ("Sarah (Undergrad #1)", ("Lina", "Jill"))
    assert t.role("flood") == ("", ())                       # no default: left to fill in
    assert set(tm.DEFAULT_ROLES) <= set(dc.TASK_KINDS)


def _emails(path):
    """{(task name, due): (assignee, collaborators)} with the formulas recalculated."""
    formulas = pytest.importorskip("formulas")
    import datetime as dt
    import formulas.functions.date as fdate
    import numpy as np

    class ExcelWorkday:                                       # as in test_design
        def __getattr__(self, n):
            return getattr(np, n)

        @staticmethod
        def busday_offset(d, n, **kw):
            return np.busday_offset(d, n, roll="backward" if n > 0 else "forward", **kw)

    real, fdate.np = fdate.np, ExcelWorkday()
    try:
        sol = formulas.ExcelModel().loads(str(path)).finish().calculate()
    finally:
        fdate.np = real
    ws = openpyxl.load_workbook(path)["asana_tasks"]
    def val(ref):
        if not str(ws[ref].value or "").startswith("="):
            return ws[ref].value
        v = sol[next(k for k in sol if k.upper().endswith(f"ASANA_TASKS'!{ref}"))].value[0][0]
        return v.item() if hasattr(v, "item") else v
    out = []
    for r in range(2, ws.max_row + 1):
        name = val(f"A{r}")
        if name:
            out.append((name, val(f"E{r}") or "", val(f"F{r}") or ""))
    return out


def test_asana_rows_carry_the_emails(tmp_path):
    d = dc.build(p0_plan())
    path = str(tmp_path / "x.xlsx")
    dx.write(d, path, team=_team())
    rows = _emails(path)
    by = {}
    for name, a, c in rows:
        by.setdefault(name.split(" (")[0].split(" - ")[0].split(",")[0], set()).add((a, c))
    assert by["Pre-fill pots with weighed soil"] == {("sarah@example.org", "lina@example.org")}
    assert by["Wash pots for reuse"] == {("sarah@example.org",
                                          "lina@example.org, jill@example.org")}
    assert by["Plant harvest and soil sample collection"] == {("", "lina@example.org")}  # no email
    assert by["Flood"] == {("", "")}                                                   # no role
    assert by["Maintenance watering"] == {("sarah@example.org", "jill@example.org")}
    assert len(rows) == len(d.tasks())                  # spare watering rows stay blank


def test_assignments_tab_is_the_place_to_edit(tmp_path):
    d = dc.build(p0_plan())
    path = str(tmp_path / "x.xlsx")
    dx.write(d, path, team=_team())
    wb = openpyxl.load_workbook(path)
    ws = wb[dx.ASSIGN_SHEET]
    r = next(r for r in range(1, 40) if ws.cell(r, 1).value == "flood")
    ws.cell(r, 3).value, ws.cell(r, 4).value = "Jill", "Sarah (Undergrad #1)"
    wb.save(path)
    rows = _emails(path)
    assert {(a, c) for n, a, c in rows if n.startswith("Flood")} == \
           {("jill@example.org", "sarah@example.org")}


def test_no_team_file_leaves_them_blank(tmp_path):
    d = dc.build(big_plan())
    path = str(tmp_path / "x.xlsx")
    dx.write(d, path, team=None)
    rows = _emails(path)
    assert all(a == "" and c == "" for _, a, c in rows)
    assert _asana_values(path)                           # the rest still calculates


def test_swaps_keep_the_assignment_edits(tmp_path):
    import datetime as dt
    d = dc.build(p0_plan(spares_per_genotype_per_group=2))
    path = str(tmp_path / "x.xlsx")
    dx.write(d, path, team=_team())
    wb = openpyxl.load_workbook(path)
    ws = wb[dx.ASSIGN_SHEET]
    r = next(r for r in range(1, 40) if ws.cell(r, 1).value == "harvest")
    ws.cell(r, 3).value = "Lina"
    wb.save(path)
    d2, _ = S.read_design_map(path)
    f = next(p for p in d2.pots if p.role == dc.MEASURED)
    S.rebuild(path, d2, [S.swap(d2, f.barcode, S.eligible_spares(d2, f)[0].barcode,
                                dt.date(2026, 11, 10))])
    ws = openpyxl.load_workbook(path)[dx.ASSIGN_SHEET]
    assert ws.cell(r, 3).value == "Lina"
