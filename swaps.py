#!/usr/bin/env python3
"""
swaps.py: replace failed pots with spares in an existing design map (v0.1, 2026-11-05).
No Qt. swap_app.py is the window on top of it.

A spare takes over a failed pot of its own genotype and planting group: its role, cohort,
treatment and rep, and its place in the flood, N2O and harvest order. Neither pot moves
and no barcode changes (the spare gets a hand-written "promoted to ..." label). The failed
pot stays on record as FAILED; spares that are no longer needed can be DISCARDED. Both
leave the `jars` tab, so the dosing station stops expecting them once the trial is
reloaded there.

The design map is then REBUILT from what it already holds, never re-randomised: the pots
come from its `design` tab (positions, groups, N2O loads) and the plan from its hidden
`plan` tab (or, for workbooks made before that tab, from `info` and `admin`). Everything
typed into the old workbook is carried over: values in the flood, measure and harvest
sheets (by barcode), edited yellow timeline inputs, the admin values, and any tab the
program did not make (e.g. dry weights). The old file is kept as a dated backup.
"""
from __future__ import annotations

import datetime as dt
import json
import os
import shutil

import openpyxl

import design_core as dc
import design_xlsx as dx

# tabs the program writes; anything else in a design map was added by hand and is copied
GENERATED = {"timeline", "design", "layout", "tags", "groups", "asana_tasks", "jars",
             "admin", "info", "swaps", dx.PLAN_SHEET}
PER_COHORT = ("flood_", "measure_", "harvest_")


class SwapError(ValueError):
    """The swap cannot be made; the message says why."""


# ======================================================================================
# READ a design map back into a Design
# ======================================================================================
def _num(v, kind=float):
    return None if v in (None, "", "none") else kind(v)


def _plan_from_info(wb) -> tuple[dc.Plan, int, int, dict | None]:
    """Workbooks made before the hidden `plan` tab: the plan from `info` and `admin`."""
    rows = [r for r in wb["info"].iter_rows(values_only=True)]
    kv = {r[0]: r[1] for r in rows if r and r[0] is not None}
    hours0 = tuple(int(h) for h in str(kv.get("N2O after flood (h)", "24, 48")).split(","))
    i = next(k for k, r in enumerate(rows) if r and r[0] == "cohort")
    cohorts, delays, k = [], {}, i + 1
    while k < len(rows) and rows[k] and rows[k][0] is not None:
        r = rows[k]
        hours = (tuple(int(h) for h in str(r[3]).split(",")) if len(r) > 3 and r[3]
                 else hours0)
        cohorts.append(dc.Cohort(str(r[0]), int(r[1]), hours))
        if len(r) > 2 and r[2] is not None:
            delays[str(r[0])] = int(r[2])
        k += 1
    j = next(k for k, r in enumerate(rows) if r and r[0] == "genotype")
    prow = [dc.PlanRow(str(r[0]), str(r[1]), str(r[2] or ""), int(r[3]), int(r[4]))
            for r in rows[j + 1:] if r and r[0] and r[0] != "WARNING" and r[3] is not None]
    adm = {}
    if "admin" in wb.sheetnames:
        adm = {str(r[0]).strip(): r[1] for r in wb["admin"].iter_rows(values_only=True)
               if r and r[0]}
    soil = {a: _num(adm.get(key)) for a, key, _ in dc.ADMIN_INPUTS}
    soil["particle_density"] = soil["particle_density"] or 2.65
    plan = dc.Plan(
        name=str(kv["experiment"]), seed=int(kv["seed"]),
        planting_date=dt.date.fromisoformat(str(kv["planting date, group 1"])[:10]),
        treatment_dap=_num(kv.get("treatment DAP"), int),
        tables=int(kv["tables"]), zones_per_table=int(kv["zones per table"]),
        slots_per_zone=int(kv["slots per zone"]),
        max_measured_per_day=int(kv["max N2O per day (capacity)"]),
        spares_per_genotype_per_group=int(kv.get("spares per genotype per planting group", 0)),
        cohorts=cohorts, rows=prow, **soil)
    return plan, int(kv["planting groups"]), int(kv["days between groups"]), delays or None


def read_design_map(path: str) -> tuple[dc.Design, list[dc.Swap]]:
    """The Design a design map describes (its pots exactly as laid out) and its swaps."""
    wb = openpyxl.load_workbook(path)
    if dx.PLAN_SHEET in wb.sheetnames:
        ws = wb[dx.PLAN_SHEET]
        doc = json.loads("".join(str(r[0]) for r in ws.iter_rows(values_only=True) if r[0]))
        plan = dc.Plan.from_dict(doc["plan"])
        groups, stagger, delays = doc["groups"], doc["stagger_days"], doc["offsets"]
    else:
        plan, groups, stagger, delays = _plan_from_info(wb)
    delays = dict(delays or dc.cohort_offsets(plan))
    if "timeline" in wb.sheetnames:     # planting delays typed over in the timeline win
        tl = wb["timeline"]
        names = {c.name for c in plan.cohorts}
        head = next((r for r in range(1, 40)
                     if str(tl.cell(r, 4).value or "").startswith("Planted")), None)
        if head:
            for r in range(head + 1, head + 1 + len(names)):
                name, v = str(tl.cell(r, 1).value or ""), tl.cell(r, 4).value
                if name in names and isinstance(v, (int, float)):
                    delays[name] = int(v)

    ws = wb["design"]
    head = [str(c.value) for c in ws[1]]
    col = {h: i for i, h in enumerate(head)}
    get = lambda r, h, kind=str: (None if h not in col or r[col[h]] in (None, "")
                                  else kind(r[col[h]]))
    pots = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or not r[0]:
            continue
        role = get(r, "role")
        pots.append(dc.Pot(
            barcode=str(r[0]), role=role, genotype=get(r, "genotype") or "",
            cohort=get(r, "cohort") or "", treatment=get(r, "treatment") or "",
            rep=get(r, "rep", int) or 0, group=get(r, "planting_group", int) or 1,
            group_in_tag=groups > 1, load=get(r, "n2o_load", int) or 0,
            column=get(r, "n2o_column", int) or 0, seq=get(r, "seq", int) or 0,
            table=get(r, "table", int), zone=get(r, "zone", int), slot=get(r, "slot", int),
            was=get(r, "role_before") or "", swap=get(r, "swapped_with") or "",
            planted_with=get(r, "planted_with") or ""))
    load = dc.daily_load(plan, pots, stagger, delays)
    d = dc.Design(plan, pots, groups, stagger, load, [], delays)

    swaps = []
    if "swaps" in wb.sheetnames:
        sh = wb["swaps"]
        if [c.value for c in sh[1]][:3] == ["date", "failed_barcode", "spare_barcode"]:
            for r in sh.iter_rows(min_row=2, values_only=True):
                if r and r[1]:
                    day = r[0].date() if isinstance(r[0], dt.datetime) else r[0]
                    swaps.append(dc.Swap(day, str(r[1]), str(r[2]), str(r[3] or ""),
                                         str(r[4] or "")))
    return d, swaps


# ======================================================================================
# SWAP
# ======================================================================================
def pot(d: dc.Design, barcode: str) -> dc.Pot:
    bc = barcode.strip()
    p = next((p for p in d.pots if p.barcode.casefold() == bc.casefold()), None)
    if p is None:
        raise SwapError(f"{bc}: not a pot of this experiment")
    return p


def planted(d: dc.Design, p: dc.Pot) -> dt.date:
    """The day the pot was planted: its group's date, plus its cohort's planting delay
    (a cohort planted later so its days do not collide). Spares: the group's date."""
    return d.pot_planting(p)


def eligible_spares(d: dc.Design, failed: dc.Pot) -> list[dc.Pot]:
    """Spares that may replace `failed`: same genotype, planted the same day, not used."""
    day = planted(d, failed)
    return [p for p in d.pots if p.role == dc.SPARE and p.genotype == failed.genotype
            and p.group == failed.group and planted(d, p) == day]


def swap(d: dc.Design, failed_bc: str, spare_bc: str, day: dt.date | None = None,
         note: str = "") -> dc.Swap:
    """The spare takes over the failed pot's role, treatment, rep and place in the flood,
    N2O and harvest order; the failed pot becomes FAILED. Changes `d` in place."""
    f, s = pot(d, failed_bc), pot(d, spare_bc)
    if f.role == dc.FAILED:
        raise SwapError(f"{f.barcode} has already failed (replaced by {f.swap})")
    if f.role not in (dc.MEASURED, dc.SACRIFICIAL):
        raise SwapError(f"{f.barcode} is a {f.role.lower()} pot, not one in the experiment: "
                        f"only a measured or sacrificial pot can be replaced")
    if s.role != dc.SPARE:
        what = ("already replaces " + s.swap) if s.swap else f"a {s.role.lower()} pot"
        raise SwapError(f"{s.barcode} is not an unused spare: it is {what}")
    if s.genotype != f.genotype:
        raise SwapError(f"{s.barcode} is {s.genotype}, {f.barcode} is {f.genotype}: a spare "
                        f"only replaces a pot of its own genotype")
    if s.group != f.group or planted(d, s) != planted(d, f):
        raise SwapError(f"{s.barcode} was planted {planted(d, s):%a %d %b} (group {s.group}), "
                        f"{f.barcode} {planted(d, f):%a %d %b} (group {f.group}"
                        + (f", cohort {f.cohort} planted later" if d.offsets.get(f.cohort)
                           else "") + "): a spare only replaces a pot of its own age")
    for a in ("role", "cohort", "treatment", "rep", "load", "column", "seq"):
        setattr(s, a, getattr(f, a))
    s.swap, s.group_in_tag = f.barcode, False
    f.was, f.role, f.swap = f.role, dc.FAILED, s.barcode
    f.load = f.column = f.seq = 0       # its turn in the orders is the spare's now
    return dc.Swap(day or dt.date.today(), f.barcode, s.barcode, s.label, note)


def discard_spares(d: dc.Design) -> list[str]:
    """Every spare still unused becomes DISCARDED (no longer grown). Returns the barcodes."""
    out = []
    for p in d.pots:
        if p.role == dc.SPARE:
            p.role = dc.DISCARDED
            out.append(p.barcode)
    return out


# ======================================================================================
# REBUILD, carrying over what was typed into the old workbook
# ======================================================================================
def _blocks(ws):
    """{(header occurrence, barcode): {column: value}} for every row under a header row
    with a 'barcode' column, and {(row, column): value} for every other row."""
    keyed, loose = {}, {}
    n, bcol = -1, None
    for row in ws.iter_rows():
        vals = [c.value for c in row]
        if "barcode" in vals:
            n, bcol = n + 1, vals.index("barcode")
            continue
        bc = vals[bcol] if bcol is not None and bcol < len(vals) else None
        if bc:
            keyed[(n, str(bc))] = {c.column: c.value for c in row if c.value is not None}
        else:
            for c in row:
                if c.value is not None:
                    loose[(c.row, c.column)] = c.value
    return keyed, loose


def _carry_sheet(old, new) -> int:
    """Typed values from `old` into the empty cells of `new` (same kind of sheet)."""
    ok, ol = _blocks(old)
    copied, n, bcol = 0, -1, None
    for row in new.iter_rows():
        vals = [c.value for c in row]
        if "barcode" in vals:
            n, bcol = n + 1, vals.index("barcode")
            continue
        bc = vals[bcol] if bcol is not None and bcol < len(vals) else None
        src = ok.get((n, str(bc)), {}) if bc else None
        for c in row:
            v = (src.get(c.column) if src is not None else ol.get((c.row, c.column)))
            if c.value is None and v is not None:
                c.value = v
                copied += 1
    return copied


def _carry(old_wb, new_wb) -> dict:
    report = {"cells": 0, "inputs": 0, "sheets": []}
    for name in old_wb.sheetnames:
        old = old_wb[name]
        if name.startswith(PER_COHORT) and name in new_wb.sheetnames:
            report["cells"] += _carry_sheet(old, new_wb[name])
        elif name == "timeline":              # yellow inputs edited in Excel
            new = new_wb[name]
            for row in new.iter_rows():
                for c in row:
                    if (c.fill.fill_type and (c.fill.start_color.rgb or "")[-6:] == dx.INPUT_BG
                            and old[c.coordinate].value is not None
                            and old[c.coordinate].value != c.value):
                        c.value = old[c.coordinate].value
                        report["inputs"] += 1
        elif name == "admin":
            new = new_wb[name]
            keys = {str(r[0].value): r[1] for r in new.iter_rows() if r[0].value}
            for r in old.iter_rows():
                k = str(r[0].value) if r[0].value else None
                if k in keys and len(r) > 1 and r[1].value not in (None, ""):
                    if keys[k].value != r[1].value:
                        keys[k].value = r[1].value
                        report["inputs"] += 1
        elif name not in GENERATED and name not in new_wb.sheetnames:
            new = new_wb.create_sheet(name)   # added by hand: copied as values
            for row in old.iter_rows():
                for c in row:
                    if c.value is not None:
                        new.cell(row=c.row, column=c.column, value=c.value)
            for k, dim in old.column_dimensions.items():
                if dim.width:
                    new.column_dimensions[k].width = dim.width
            report["sheets"].append(name)
    return report


def backup_path(path: str, now: dt.datetime | None = None) -> str:
    stem, ext = os.path.splitext(path)
    return f"{stem}_before_swaps_{(now or dt.datetime.now()):%Y-%m-%d_%H%M%S}{ext}"


def rebuild(path: str, d: dc.Design, swaps: list[dc.Swap], out: str | None = None) -> dict:
    """Write the design map for `d` (with its swaps) over `path` (or to `out`), carrying
    over what was typed into the old one, after copying the old one to a backup.
    Returns {"backup", "cells", "inputs", "sheets"}."""
    out = out or path
    old_wb = openpyxl.load_workbook(path)
    backup = None
    if os.path.abspath(out) == os.path.abspath(path):
        backup = backup_path(path)
        shutil.copy2(path, backup)
    tmp = out + ".tmp.xlsx"
    dx.write(d, tmp, swaps=swaps)
    new_wb = openpyxl.load_workbook(tmp)
    report = _carry(old_wb, new_wb)
    new_wb.save(tmp)
    os.replace(tmp, out)
    return {"backup": backup, **report}
