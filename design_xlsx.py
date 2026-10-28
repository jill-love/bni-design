#!/usr/bin/env python3
"""
design_xlsx.py: write a Design (design_core.py) to an .xlsx workbook.

Sheets
  timeline  the schedule as a day-by-day calendar, one column per planting group x cohort,
            in the style of the P0 design map. The yellow cells at the top (planting dates,
            days between groups, treatment day, first N2O day per cohort, capacity) are
            inputs: every date in the calendar is a FORMULA on them, so editing one redraws
            the calendar when the workbook is opened in Excel. The last column counts the
            N2O measurements each day and turns red above capacity.
  design    one row per pot: barcode, role, genotype, cohort, treatment, rep, planting
            group, planting date (linked to the timeline), table, zone, slot, position, tag.
  layout    each table as a zone x slot grid: tag and barcode per cell.
  groups    planting groups: date, pots, measured pots per cohort; N2O load per day.
  jars      barcode + tag (tag ends with the Tn-Zn-nn position): the sheet the dosing
            station's provisioning reads (add its `admin` sheet before provisioning).
  info      the plan as entered, seed and version, so the design can be rebuilt exactly.
"""
from __future__ import annotations

import datetime as dt

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

import design_core as dc

HEAD_BG = "374151"         # the station windows' slate
INPUT_BG = "FEF9C3"        # yellow: an input the timeline recalculates from
FLOOD_BG, FLOOD_FG = "4A86E8", "FFFFFF"
N2O_BG = "DBEAFE"
PLANT_BG = "DCFCE7"
TREAT_BG = "FDE68A"
WEEKEND_BG = "F8D9BE"
OVER_BG = "FCA5A5"
GROUP_TINTS = ("FFFFFF", "F1F5F9", "ECFDF5", "FEF3C7", "EDE9FE", "FCE7F3", "E0F2FE", "FEF2F2")

BOLD = Font(bold=True)
HEAD = Font(bold=True, color="FFFFFF")
THIN = Side(style="thin", color="D1D5DB")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
WRAP = Alignment(vertical="top", wrap_text=True)


def _fill(hex_):
    return PatternFill("solid", start_color=hex_, end_color=hex_)


def _header(ws, row, values, col=1):
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=col + i, value=v)
        c.font = HEAD; c.fill = _fill(HEAD_BG); c.alignment = CENTER; c.border = BOX


def _input(cell):
    cell.fill = _fill(INPUT_BG); cell.border = BOX


# ======================================================================================
# TIMELINE
# ======================================================================================
def _timeline(wb, d: dc.Design) -> dict:
    """Returns {group: absolute ref of its planting-date cell} for the design sheet."""
    plan = d.plan
    ws = wb.create_sheet("timeline")
    ws["A1"] = f"{plan.name}: timeline"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = ("Yellow cells are inputs: change a date or a day count and the calendar below "
                "follows. Days are days after planting (DAP) of each pot's planting group.")
    ws["A2"].font = Font(italic=True, color="6B7280")

    params = [("Planting date, group 1", plan.planting_date, "DD MMM YYYY"),
              ("Days between planting groups", d.stagger_days, "0"),
              ("Treatment (days after planting)", plan.treatment_dap, "0"),
              ("Max N2O measurements per day", plan.max_measured_per_day, "0")]
    for i, (label, val, fmt) in enumerate(params):
        ws.cell(row=4 + i, column=1, value=label).font = BOLD
        c = ws.cell(row=4 + i, column=2, value=val)
        c.number_format = fmt; _input(c)
    P1, GAP, TREAT, CAP = "$B$4", "$B$5", "$B$6", "$B$7"

    top = 9
    _header(ws, top, ["Cohort", "First N2O (DAP)", "Flood (DAP)"])
    coh_row = {}
    for i, c in enumerate(plan.cohorts):
        r = top + 1 + i
        coh_row[c.name] = r
        ws.cell(row=r, column=1, value=c.name).font = BOLD
        _input(ws.cell(row=r, column=2, value=c.first_n2o_dap))
        ws.cell(row=r, column=3, value=f"=B{r}-1").border = BOX

    _header(ws, top, ["Planting group", "Planting date", "Pots", "Measured"], col=5)
    grp_row = {}
    for g in range(1, d.groups + 1):
        r = top + g
        grp_row[g] = r
        ws.cell(row=r, column=5, value=g).font = BOLD
        c = ws.cell(row=r, column=6, value=f"={P1}+({g}-1)*{GAP}")
        c.number_format = "DD MMM YYYY"; _input(c)
        ws.cell(row=r, column=7, value=sum(p.group == g for p in d.pots)).border = BOX
        ws.cell(row=r, column=8, value=sum(p.group == g and p.role == dc.MEASURED
                                           for p in d.pots)).border = BOX
    ws.cell(row=top + d.groups + 1, column=5,
            value="(a planting date may be typed over; the others keep following group 1)"
            ).font = Font(italic=True, color="6B7280")

    # ---- the calendar ----
    hdr = max(top + len(plan.cohorts), top + d.groups + 1) + 3
    cols = [(g, c) for g in range(1, d.groups + 1) for c in plan.cohorts]
    first_col = 3
    _header(ws, hdr, ["Date", "Day"] + [f"Group {g} · Cohort {c.name}" for g, c in cols]
            + ["N2O pots"])
    n_row = hdr + 1                                     # measured pots per column (helper)
    ws.cell(row=n_row, column=1, value="measured pots").font = Font(italic=True, color="6B7280")
    for j, (g, c) in enumerate(cols):
        n = sum(p.group == g and p.cohort == c.name and p.role == dc.MEASURED for p in d.pots)
        cell = ws.cell(row=n_row, column=first_col + j, value=n)
        cell.alignment = CENTER; cell.font = Font(italic=True, color="6B7280")

    span_end = max(ev for g, c in cols for ev, _ in d.events(g, c))
    n_days = (span_end - plan.planting_date).days + 1 + 14          # room to move dates later
    start = hdr + 2
    last_col = first_col + len(cols) - 1
    tot_col = last_col + 1
    for i in range(n_days):
        r = start + i
        a = ws.cell(row=r, column=1,
                    value=f"=MIN($F${top + 1}:$F${top + d.groups})" if i == 0 else f"=A{r - 1}+1")
        a.number_format = "D MMM YYYY"; a.border = BOX
        b = ws.cell(row=r, column=2, value=f'=TEXT(A{r},"ddd")'); b.border = BOX
        for j, (g, c) in enumerate(cols):
            pr = f"$F${grp_row[g]}"
            n2o = f"$B${coh_row[c.name]}"
            flood_txt = ("Sacrificial harvest, flood"
                         if any(p.role == dc.SACRIFICIAL and p.group == g and p.cohort == c.name
                                for p in d.pots) else "Flood")
            # innermost first: water between planting and flood, else blank
            f = (f'IF(AND($A{r}>{pr},$A{r}<{pr}+{n2o}-1),'
                 f'IF(WEEKDAY($A{r},2)>5,"Water (weekend)","Water"),"")')
            offs = plan.measure_offsets
            for h, off in reversed(list(zip(plan.measure_hours, offs))):
                txt = f"{h} h N2O" + (", harvest" if off == offs[-1] else "")
                f = f'IF($A{r}={pr}+{n2o}+{off},"{txt}",{f})'
            f = f'IF($A{r}={pr}+{n2o}-1,"{flood_txt}",{f})'
            f = f'IF(IF({TREAT}="",FALSE,$A{r}={pr}+{TREAT}),"Treatment; water",{f})'
            f = f'=IF($A{r}={pr},"Plant",{f})'
            cell = ws.cell(row=r, column=first_col + j, value=f)
            cell.border = BOX; cell.alignment = Alignment(vertical="center")
        rng = f"{L(first_col)}{r}:{L(last_col)}{r}"
        helper = f"{L(first_col)}${n_row}:{L(last_col)}${n_row}"
        t = ws.cell(row=r, column=tot_col,
                    value=f'=SUMPRODUCT(ISNUMBER(SEARCH("N2O",{rng}))*{helper})')
        t.border = BOX; t.alignment = CENTER; t.font = BOLD

    end = start + n_days - 1
    grid = f"{L(first_col)}{start}:{L(last_col)}{end}"
    tl = f"{L(first_col)}{start}"
    for word, bg, fg in (("flood", FLOOD_BG, FLOOD_FG), ("N2O", N2O_BG, None),
                         ("Plant", PLANT_BG, None), ("Treatment", TREAT_BG, None),
                         ("weekend", INPUT_BG, None)):
        ws.conditional_formatting.add(grid, FormulaRule(
            formula=[f'ISNUMBER(SEARCH("{word}",{tl}))'], fill=_fill(bg),
            font=Font(bold=True, color=fg) if fg else None, stopIfTrue=True))
    ws.conditional_formatting.add(f"A{start}:B{end}", FormulaRule(
        formula=[f"WEEKDAY($A{start},2)>5"], fill=_fill(WEEKEND_BG)))
    tot = f"{L(tot_col)}{start}:{L(tot_col)}{end}"
    ws.conditional_formatting.add(tot, FormulaRule(
        formula=[f"{L(tot_col)}{start}>{CAP}"], fill=_fill(OVER_BG), font=Font(bold=True)))
    ws.conditional_formatting.add(tot, FormulaRule(
        formula=[f"{L(tot_col)}{start}=0"], font=Font(color="D1D5DB")))

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 14
    for col in range(first_col, tot_col):
        ws.column_dimensions[L(col)].width = 26
    for col, w in ((5, 15), (6, 14), (7, 8), (8, 10)):
        ws.column_dimensions[L(col)].width = max(ws.column_dimensions[L(col)].width or 0, w)
    ws.column_dimensions[L(tot_col)].width = 10
    ws.freeze_panes = ws.cell(row=start, column=first_col)
    return {g: f"timeline!$F${r}" for g, r in grp_row.items()}


# ======================================================================================
# THE OTHER SHEETS
# ======================================================================================
def _design(wb, d: dc.Design, plant_ref: dict) -> None:
    ws = wb.create_sheet("design")
    cols = ["barcode", "role", "genotype", "cohort", "treatment", "rep", "planting_group",
            "planting_date", "table", "zone", "slot", "position", "tag"]
    _header(ws, 1, cols)
    for r, p in enumerate(d.pots, 2):
        vals = [p.barcode, p.role, p.genotype, p.cohort or None, p.treatment or None, p.rep,
                p.group, f"={plant_ref[p.group]}", p.table, p.zone, p.slot, p.position, p.label]
        for c, v in enumerate(vals, 1):
            ws.cell(row=r, column=c, value=v)
        ws.cell(row=r, column=1).number_format = "@"
        ws.cell(row=r, column=8).number_format = "DD MMM YYYY"
    for c, w in zip(range(1, 14), (16, 13, 12, 8, 11, 6, 9, 13, 6, 6, 6, 11, 24)):
        ws.column_dimensions[L(c)].width = w
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:{L(len(cols))}{len(d.pots) + 1}"


def _layout(wb, d: dc.Design) -> None:
    plan = d.plan
    ws = wb.create_sheet("layout")
    ws["A1"] = (f"Bench layout: {plan.tables} tables x {plan.zones_per_table} zones x "
                f"{plan.slots_per_zone} slots. Zone 1 at the taped end; slots count from the "
                f"left. Cell: tag / barcode; the fill shows the planting group.")
    ws["A1"].font = Font(italic=True, color="6B7280")
    at = {(p.table, p.zone, p.slot): p for p in d.pots}
    r = 3
    for t in range(1, plan.tables + 1):
        n = sum(p.table == t for p in d.pots)
        _header(ws, r, [f"Table {t} ({n} pots)"]
                + [f"slot {s}" for s in range(1, plan.slots_per_zone + 1)])
        for z in range(1, plan.zones_per_table + 1):
            r += 1
            ws.cell(row=r, column=1, value=f"zone {z}").font = BOLD
            for s in range(1, plan.slots_per_zone + 1):
                p = at.get((t, z, s))
                c = ws.cell(row=r, column=1 + s,
                            value=f"{p.label}\n{p.barcode}" if p else "(empty)")
                c.alignment = CENTER; c.border = BOX
                if p:
                    c.fill = _fill(GROUP_TINTS[(p.group - 1) % len(GROUP_TINTS)])
                else:
                    c.font = Font(color="9CA3AF")
            ws.row_dimensions[r].height = 32
        r += 3
    ws.column_dimensions["A"].width = 18
    for s in range(1, plan.slots_per_zone + 1):
        ws.column_dimensions[L(1 + s)].width = 19


def _groups(wb, d: dc.Design, plant_ref: dict) -> None:
    plan = d.plan
    ws = wb.create_sheet("groups")
    _header(ws, 1, ["Planting group", "Planting date", "Pots"]
            + [f"Measured, cohort {c.name}" for c in plan.cohorts])
    for g in range(1, d.groups + 1):
        r = g + 1
        ws.cell(row=r, column=1, value=g).font = BOLD
        ws.cell(row=r, column=2, value=f"={plant_ref[g]}").number_format = "DD MMM YYYY"
        ws.cell(row=r, column=3, value=sum(p.group == g for p in d.pots))
        for j, c in enumerate(plan.cohorts):
            ws.cell(row=r, column=4 + j, value=sum(
                p.group == g and p.cohort == c.name and p.role == dc.MEASURED for p in d.pots))
    r = d.groups + 4
    ws.cell(row=r, column=1, value=(
        f"N2O load per day as designed (capacity {plan.max_measured_per_day}; the timeline's "
        f"last column recalculates if dates are edited)")).font = BOLD
    _header(ws, r + 1, ["Date", "N2O measurements"])
    for i, (day, n) in enumerate(d.calendar_load().items()):
        ws.cell(row=r + 2 + i, column=1, value=day).number_format = "DD MMM YYYY (ddd)"
        ws.cell(row=r + 2 + i, column=2, value=n)
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 18
    for j in range(len(plan.cohorts) + 1):
        ws.column_dimensions[L(3 + j)].width = 20


def _jars(wb, d: dc.Design) -> None:
    ws = wb.create_sheet("jars")
    ws.append(["barcode", "tag"])
    for c in ws[1]:
        c.font = BOLD; c.alignment = Alignment(horizontal="center")
    for p in d.pots:
        ws.append([p.barcode, f"{p.label} {p.position}"])
    for row in ws.iter_rows(min_row=2, max_col=1):
        row[0].number_format = "@"                      # Text: keeps leading zeros
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 34
    ws.freeze_panes = "A2"


def _info(wb, d: dc.Design) -> None:
    plan = d.plan
    ws = wb.create_sheet("info")
    c = d.counts()
    rows = [("experiment", plan.name), ("generated", dt.datetime.now().strftime("%Y-%m-%d %H:%M")),
            ("tool version", dc.__version__), ("seed", plan.seed),
            ("pots", len(d.pots)), ("measured", c[dc.MEASURED]),
            ("sacrificial", c[dc.SACRIFICIAL]), ("spare", c[dc.SPARE]),
            ("tables", plan.tables), ("zones per table", plan.zones_per_table),
            ("slots per zone", plan.slots_per_zone), ("bench capacity", plan.capacity),
            ("planting groups", d.groups), ("days between groups", d.stagger_days),
            ("max N2O per day (designed)", d.max_load),
            ("max N2O per day (capacity)", plan.max_measured_per_day),
            ("N2O after flood (h)", ", ".join(str(h) for h in plan.measure_hours)),
            ("planting date, group 1", plan.planting_date.isoformat()),
            ("treatment DAP", plan.treatment_dap if plan.treatment_dap is not None else "none")]
    for k, v in rows:
        ws.append([k, v])
        ws.cell(row=ws.max_row, column=1).font = BOLD
    ws.append([])
    ws.append(["cohort", "first N2O DAP"])
    for cc in plan.cohorts:
        ws.append([cc.name, cc.first_n2o_dap])
    ws.append([])
    ws.append(["genotype", "cohort", "treatment", "reps", "sacrificial"])
    for pr in plan.rows:
        if pr.genotype.strip():
            ws.append([pr.genotype, pr.cohort, pr.treatment or None, pr.reps, pr.sacrificial])
    for w in d.warnings:
        ws.append(["WARNING", w])
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 20


def write(d: dc.Design, path: str) -> None:
    wb = Workbook()
    wb.remove(wb.active)
    plant_ref = _timeline(wb, d)
    _design(wb, d, plant_ref)
    _layout(wb, d)
    _groups(wb, d, plant_ref)
    _jars(wb, d)
    _info(wb, d)
    wb.save(path)
