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
from collections import Counter
import os
import shutil

import openpyxl

import design_core as dc
import design_xlsx as dx
import labels

# tabs the program writes; anything else in a design map was added by hand and is copied
GENERATED = {"timeline", "design", "layout", "tags", "groups", "asana_tasks", "jars",
             "admin", "info", "swaps", dx.PLAN_SHEET, dx.ASSIGN_SHEET, dx.LABEL_LOG}
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
    if "planting rounds" in kv:          # made after groups became planting dates
        return (plan, int(kv["planting rounds"]), int(kv["days between rounds"]),
                delays or None, True)
    return plan, int(kv["planting groups"]), int(kv["days between groups"]), delays or None, False


def _round_order(plan, n_rounds, stagger, delays):
    """[(round, delay)] for planting groups 1, 2, ... as dc.split_groups numbers them."""
    keys = sorted({delays.get(c.name, 0) for c in plan.cohorts}) or [0]
    return sorted(((r, o) for r in range(1, n_rounds + 1) for o in keys),
                  key=lambda ro: ((ro[0] - 1) * stagger + ro[1], ro[0]))


def _split_legacy(plan, pots, n_rounds, stagger, delays):
    """Maps made before planting groups were planting dates: a group (now a round) could
    plant on two days. Each (round, delay) becomes its own group, as dc.split_groups does;
    a spare goes with the cohorts it was planted with."""
    rounds = _round_order(plan, n_rounds, stagger, delays)
    number = {ro: k for k, ro in enumerate(rounds, 1)}
    for p in pots:
        cohort = p.cohort or (p.planted_with.split("+")[0] if p.planted_with else "")
        p.group = number[(p.group, delays.get(cohort, 0))]
        p.group_in_tag = len(rounds) > 1
    return len(rounds), rounds


def _load(d: dc.Design) -> Counter:
    """{days after the first planting: N2O measurements}, from the pots as they are."""
    out = Counter()
    for p in d.pots:
        if p.role == dc.MEASURED:
            c = d.plan.cohort(p.cohort)
            base = (d.group_planting(p.group) - d.plan.planting_date).days + c.first_n2o_dap
            for off in c.offsets:
                out[base + off] += 1
    return out


def read_design_map(path: str) -> tuple[dc.Design, list[dc.Swap]]:
    """The Design a design map describes (its pots exactly as laid out) and its swaps."""
    wb = openpyxl.load_workbook(path)
    if dx.PLAN_SHEET in wb.sheetnames:
        ws = wb[dx.PLAN_SHEET]
        doc = json.loads("".join(str(r[0]) for r in ws.iter_rows(values_only=True) if r[0]))
        plan = dc.Plan.from_dict(doc["plan"])
        groups, stagger, delays = doc["groups"], doc["stagger_days"], doc["offsets"]
        rounds = [tuple(r) for r in doc.get("rounds", [])]
    else:
        plan, groups, stagger, delays, by_date = _plan_from_info(wb)
        rounds = []
        if by_date:                     # pots already carry their planting group
            rounds = _round_order(plan, groups, stagger, delays or dc.cohort_offsets(plan))
            groups = len(rounds)
    delays = dict(delays or dc.cohort_offsets(plan))
    if not rounds and "timeline" in wb.sheetnames:  # older maps: delays typed over win
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
    if not rounds:
        groups, rounds = _split_legacy(plan, pots, groups, stagger, delays)
    d = dc.Design(plan, pots, groups, stagger, Counter(), [], delays, rounds)
    d.load = _load(d)

    swaps = []
    if "swaps" in wb.sheetnames:
        sh = wb["swaps"]
        if [c.value for c in sh[1]][:3] == ["date", "failed_barcode", "spare_barcode"]:
            for r in sh.iter_rows(min_row=2, values_only=True):
                if r and r[1]:
                    day = r[0].date() if isinstance(r[0], dt.datetime) else r[0]
                    swaps.append(dc.Swap(day, str(r[1]), str(r[2] or ""), str(r[3] or ""),
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


def fail(d: dc.Design, failed_bc: str, day: dt.date | None = None, note: str = "") -> dc.Swap:
    """A failed pot with no spare to replace it: it becomes FAILED and leaves the
    experiment (the station, the orders, the harvest labels); its turn in the flood,
    N2O and harvest order is left empty. Changes `d` in place."""
    f = pot(d, failed_bc)
    if f.role == dc.FAILED:
        raise SwapError(f"{f.barcode} has already failed")
    if f.role not in (dc.MEASURED, dc.SACRIFICIAL):
        raise SwapError(f"{f.barcode} is a {f.role.lower()} pot, not one in the experiment: "
                        f"only a measured or sacrificial pot can fail")
    label = f.label
    f.was, f.role, f.swap = f.role, dc.FAILED, ""
    f.load = f.column = f.seq = 0
    return dc.Swap(day or dt.date.today(), f.barcode, "", "", note or f"no spare for {label}")


def record(d: dc.Design, sw: dc.Swap) -> None:
    """Apply a saved or pending swap again: with a spare, or a failure with none."""
    if sw.spare:
        swap(d, sw.failed, sw.spare, sw.date, sw.note)
    else:
        fail(d, sw.failed, sw.date, sw.note)


# a planting day's spares are kept until its swap session: 5 working days after planting
# (the "Run swaps" task, after the 5 germination checks), whatever weekday it was planted
SPARE_WORKDAYS = dc.TASK_LAGS["swaps"]


def spares_due(d: dc.Design, p: dc.Pot, on: dt.date) -> bool:
    """`p`'s planting day has had its swap session by `on`."""
    return on >= dc.workday(planted(d, p), SPARE_WORKDAYS)


def discard_spares(d: dc.Design, on: dt.date | None = None) -> list[str]:
    """Unused spares become DISCARDED (no longer grown). With `on` (the day of the
    swaps), only spares planted at least SPARE_WORKDAYS working days before it: a later
    planting day keeps its spares until its own swap session. Returns the barcodes."""
    out = []
    for p in d.pots:
        if p.role == dc.SPARE and (on is None or spares_due(d, p, on)):
            p.role = dc.DISCARDED
            out.append(p.barcode)
    return out


# ======================================================================================
# LABELS: a planting group's harvest and KCl labels, once its swaps are done
# ======================================================================================
def label_log(path: str) -> list[tuple]:
    """[(date, planting group, label folder)]: the groups whose harvest and KCl labels
    the swap tool has written (the design map's label_log tab)."""
    wb = openpyxl.load_workbook(path, read_only=True)
    if dx.LABEL_LOG not in wb.sheetnames:
        return []
    out = []
    for r in list(wb[dx.LABEL_LOG].iter_rows(values_only=True))[1:]:
        if r and r[1] is not None:
            day = r[0].date() if isinstance(r[0], dt.datetime) else r[0]
            out.append((day, int(r[1]), str(r[2] or "")))
    return out


def ready_groups(d: dc.Design, day: dt.date, printed=()) -> set[int]:
    """Planting groups whose harvest and KCl labels can be written on `day`: their swap
    session is due (SPARE_WORKDAYS after planting), they have no unused spare left (so
    their harvested pots are final), and theirs have not been written yet."""
    left = {p.group for p in d.pots if p.role == dc.SPARE}
    return {g for g in range(1, d.groups + 1)
            if g not in printed and g not in left
            and day >= dc.workday(d.group_planting(g), SPARE_WORKDAYS)}


def save_session(path: str, d: dc.Design, swaps: list[dc.Swap], session: list[dc.Swap],
                 day: dt.date) -> dict:
    """Save a swap session: rebuild the design map (`swaps`: every swap so far, `session`:
    this session's), refresh its pot labels (labels_<map>/), and write this session's
    labels to labels_<map>_swaps_<date>/: a new pot label for each spare promoted
    (labels_RELABEL.zpl) and the harvest and KCl labels of every group now ready
    (ready_groups), which are logged so they are written once. Returns rebuild's report
    plus {"main", "session" (None: nothing to print), "relabel", "groups", "files"}."""
    log = label_log(path)
    groups = ready_groups(d, day, {g for _, g, _ in log})
    promoted = [pot(d, sw.spare) for sw in session if sw.spare]
    folder = labels.session_folder_for(path, day) if (promoted or groups) else None
    log += [(day, g, os.path.basename(folder)) for g in sorted(groups)]
    rep = rebuild(path, d, swaps, log=log)
    main = labels.folder_for(path)
    labels.write(d, main)                                 # the pot labels, for reprints
    files = (labels.write(d, folder, planting=False, groups=groups, relabel=promoted)
             if folder else {})
    return {**rep, "main": main, "session": folder, "relabel": len(promoted),
            "groups": sorted(groups), "files": files}


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


def _same_input(name, old, new, c) -> bool:
    """The old map's cell at c's place is the same input: on the timeline, its table header
    (or, above the tables, its label) says the same. An older map's column D held the
    cohorts' planting delays where the treatment days are now."""
    if name != "timeline":
        return True
    at = ((dx.TIMELINE_TOP, c.column) if c.row > dx.TIMELINE_TOP else (c.row, 1))
    return old.cell(*at).value == new.cell(*at).value


def _carry(old_wb, new_wb) -> dict:
    report = {"cells": 0, "inputs": 0, "sheets": []}
    for name in old_wb.sheetnames:
        old = old_wb[name]
        if name.startswith(PER_COHORT) and name in new_wb.sheetnames:
            report["cells"] += _carry_sheet(old, new_wb[name])
        elif name in ("timeline", dx.ASSIGN_SHEET):   # yellow inputs edited in Excel
            new = new_wb[name]
            for row in new.iter_rows():
                for c in row:
                    if (c.fill.fill_type and (c.fill.start_color.rgb or "")[-6:] == dx.INPUT_BG
                            and old[c.coordinate].value is not None
                            and old[c.coordinate].value != c.value
                            and _same_input(name, old, new, c)):
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


def rebuild(path: str, d: dc.Design, swaps: list[dc.Swap], out: str | None = None,
            log: list | None = None) -> dict:
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
    if log is None:                       # keep the old map's label log
        log = label_log(path)
    dx.write(d, tmp, swaps=swaps, label_log=log)
    new_wb = openpyxl.load_workbook(tmp)
    report = _carry(old_wb, new_wb)
    new_wb.save(tmp)
    os.replace(tmp, out)
    return {"backup": backup, **report}
