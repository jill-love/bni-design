#!/usr/bin/env python3
"""
design_xlsx.py: write a Design (design_core.py) to an .xlsx workbook.

Sheets
  timeline  the schedule as a day-by-day calendar, one column per planting group x cohort,
            in the style of the P0 design map. The yellow cells at the top (planting dates,
            days between groups, treatment day, watering days, first N2O day per
            cohort, capacity) are inputs: every date in the calendar is a FORMULA on them,
            so editing one redraws the calendar when the workbook is opened in Excel. Two
            last columns count the pots flooded and measured each day; N2O above capacity,
            or a flood on an N2O day, turns them red.
  design    one row per pot: barcode, role, genotype, cohort, treatment, rep, planting
            group, planting date (linked to the timeline), table, zone, slot, position,
            N2O load / column / seq, tag.
  layout    each table as a zone x slot grid (tag and barcode per cell), coloured by
            planting group (by cohort when there is only one group), with a legend.
  flood_*   per group x cohort: the sacrificial harvest, then the flood in N2O order.
  measure_* per group x cohort: one block per N2O time (24 h, 48 h) in load order.
  harvest_* per group x cohort: the measured pots after the last N2O time.
            (* = the cohort, or G<group>_<cohort> when there are several planting groups.)
            Their dates are formulas on the timeline too.
  tags      barcode, tag text and position for label printing, by planting group.
  groups    planting groups: date, pots, measured pots per cohort; N2O load per day.
  asana_tasks  one row per task with its due date (Name, Due Date, Section/Column,
            Description): save as CSV and import into an Asana project. Every repeat is its
            own row. Due dates are formulas on the timeline, so they follow its edits.
  jars      barcode + tag (tag ends with the Tn-Zn-nn position): the sheet the dosing
            station's provisioning reads.
  admin     the dosing station's WFPS inputs (key in column A, value in B, note in C):
            with jars, the workbook is ready for the station's Start New Trial. Plain
            values only (the station refuses formulas); blank yellow cells still to fill.
  info      the plan as entered, seed and version, so the design can be rebuilt exactly.
"""
from __future__ import annotations

import datetime as dt
import re

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
WATER_BG = "F0F9FF"
WEEKEND_BG = "F8D9BE"
OVER_BG = "FCA5A5"
SPARE_BG = "D1D5DB"
# layout colours: bright enough to tell apart across a bench
BRIGHT = ("60A5FA", "4ADE80", "FACC15", "F87171", "C084FC", "F472B6", "22D3EE", "FB923C",
          "A3E635", "2DD4BF", "818CF8", "FDA4AF")

BOLD = Font(bold=True)
HEAD = Font(bold=True, color="FFFFFF")
NOTE = Font(italic=True, color="6B7280")
THIN = Side(style="thin", color="D1D5DB")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _fill(hex_):
    return PatternFill("solid", start_color=hex_, end_color=hex_)


def _header(ws, row, values, col=1):
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=col + i, value=v)
        c.font = HEAD; c.fill = _fill(HEAD_BG); c.alignment = CENTER; c.border = BOX


CALENDAR_COLOURS = (("flood", FLOOD_BG, FLOOD_FG), ("N2O", N2O_BG, None),
                    ("Plant", PLANT_BG, None), ("Treatment", TREAT_BG, None),
                    ("Water", WATER_BG, None))


def _paint(cell, text):
    """A calendar cell's colour for `text`, first matching word wins (as the rules do)."""
    for word, bg, fg in CALENDAR_COLOURS:
        if word.lower() in text.lower():
            cell.fill = _fill(bg)
            if fg:
                cell.font = Font(bold=True, color=fg)
            return


def _input(cell):
    cell.fill = _fill(INPUT_BG); cell.border = BOX


def _widths(ws, widths, first=1):
    for i, w in enumerate(widths):
        ws.column_dimensions[L(first + i)].width = w


def _sheet_tag(d: dc.Design, g: int, cohort: str) -> str:
    c = re.sub(r"[^A-Za-z0-9]+", "", cohort) or "X"
    return c if d.groups == 1 else f"G{g}_{c}"


class Refs:
    """Absolute references to the timeline's input cells, for formulas elsewhere."""
    def __init__(self):
        self.plant = {}            # group -> planting date cell
        self.n2o = {}              # cohort -> first N2O DAP cell
        self.delay = {}            # cohort -> planting delay cell (days after the group date)
        self.lag = {}              # cohort -> days from the flood to its first N2O (24 h: 1)
        self.treat = ""            # treatment DAP cell
        self.water_days = ""       # calendar column: the date when anything is watered, else ""
        self.water_who = ""        # calendar column: who is watered that day
        self.calendar = ""         # calendar column: the dates

    def planting(self, g, c):
        """Group g's planting date for cohort c (spares, c = "": the group date)."""
        return f"({self.plant[g]}+{self.delay[c]})" if c else self.plant[g]

    def flood_date(self, g, c):
        return f"{self.planting(g, c)}+{self.n2o[c]}-{self.lag[c]}"

    def n2o_date(self, g, c, off):
        return f"{self.planting(g, c)}+{self.n2o[c]}+{off}"

    def anchor(self, a):
        """A Task anchor (design_core.Design.anchor_date) as an Excel expression."""
        kind, *x = a
        if kind == "plant":
            return self.planting(x[0], x[1])
        if kind == "treat":
            return f"{self.planting(x[0], x[1])}+{self.treat}"
        if kind == "flood":
            return self.flood_date(x[0], x[1])
        if kind == "n2o":
            return self.n2o_date(*x)
        return f"SMALL({self.water_days},{x[0]})"                 # "water"


# ======================================================================================
# TIMELINE
# ======================================================================================
def _timeline(wb, d: dc.Design) -> Refs:
    plan = d.plan
    refs = Refs()
    ws = wb.create_sheet("timeline")
    ws["A1"] = f"{plan.name}: timeline"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = ("Yellow cells are inputs: change a date or a day count and the calendar below "
                "follows. Days are days after planting (DAP) of each pot's planting group. "
                "A cohort whose days would collide with another's is planted later "
                "(Planted, days after group), so its days after planting never change.")
    ws["A2"].font = NOTE

    params = [("Planting date, group 1", plan.planting_date, "DD MMM YYYY"),
              ("Days between planting groups", d.stagger_days, "0"),
              ("Treatment (days after planting)", plan.treatment_dap, "0"),
              ("Water on (day names)", dc.WATER_DAY_NAMES, "@"),
              ("Max N2O measurements per day", plan.max_measured_per_day, "0")]
    for i, (label, val, fmt) in enumerate(params):
        ws.cell(row=4 + i, column=1, value=label).font = BOLD
        c = ws.cell(row=4 + i, column=2, value=val)
        c.number_format = fmt; _input(c)
    P1, GAP, TREAT, WATER, CAP = "$B$4", "$B$5", "$B$6", "$B$7", "$B$8"
    refs.treat = f"timeline!{TREAT}"

    top = 10
    _header(ws, top, ["Cohort", "First N2O (DAP)", "Flood (DAP)", "Planted (days after group)"])
    coh_row = {}
    for i, c in enumerate(plan.cohorts):
        r = top + 1 + i
        coh_row[c.name] = r
        refs.n2o[c.name] = f"timeline!$B${r}"
        ws.cell(row=r, column=1, value=c.name).font = BOLD
        _input(ws.cell(row=r, column=2, value=c.first_n2o_dap))
        refs.lag[c.name] = c.measure_hours[0] // 24
        ws.cell(row=r, column=3, value=f"=B{r}-{refs.lag[c.name]}").border = BOX
        refs.delay[c.name] = f"timeline!$D${r}"
        _input(ws.cell(row=r, column=4, value=d.offsets.get(c.name, 0)))

    _header(ws, top, ["Planting group", "Planting date", "Pots", "Measured"], col=5)
    grp_row = {}
    for g in range(1, d.groups + 1):
        r = top + g
        grp_row[g] = r
        refs.plant[g] = f"timeline!$F${r}"
        ws.cell(row=r, column=5, value=g).font = BOLD
        c = ws.cell(row=r, column=6, value=f"={P1}+({g}-1)*{GAP}")
        c.number_format = "DD MMM YYYY"; _input(c)
        ws.cell(row=r, column=7, value=sum(p.group == g for p in d.pots)).border = BOX
        ws.cell(row=r, column=8, value=sum(p.group == g and p.role == dc.MEASURED
                                           for p in d.pots)).border = BOX
    ws.cell(row=top + d.groups + 1, column=5,
            value="(a planting date may be typed over; the others keep following group 1)"
            ).font = NOTE

    # ---- the calendar ----
    hdr = max(top + len(plan.cohorts), top + d.groups + 1) + 3
    cols = [(g, c) for g in range(1, d.groups + 1) for c in plan.cohorts]
    first_col = 3
    last_col = first_col + len(cols) - 1
    flood_col, n2o_col = last_col + 1, last_col + 2
    who_col, wday_col = last_col + 3, last_col + 4          # helpers for the asana_tasks dates
    late = lambda c: f", planted +{d.offsets[c.name]} d" if d.offsets.get(c.name) else ""
    _header(ws, hdr, ["Date", "Day"] + [f"Group {g} · Cohort {c.name}\n"
                                        f"({', '.join(map(str, c.measure_hours))} h{late(c)})"
                                        for g, c in cols]
            + ["Flood pots", "N2O pots", "Watered", "Water day"])
    ws.row_dimensions[hdr].height = 32
    n_row = hdr + 1                                     # measured pots per column (helper)
    ws.cell(row=n_row, column=1, value="measured pots").font = NOTE
    for j, (g, c) in enumerate(cols):
        n = sum(p.group == g and p.cohort == c.name and p.role == dc.MEASURED for p in d.pots)
        cell = ws.cell(row=n_row, column=first_col + j, value=n)
        cell.alignment = CENTER; cell.font = NOTE

    span_end = max(ev for g, c in cols for ev, _ in d.events(g, c))
    n_days = (span_end - plan.planting_date).days + 1 + 14          # room to move dates later
    start = hdr + 2
    helper = f"{L(first_col)}${n_row}:{L(last_col)}${n_row}"
    cal0 = min(d.group_planting(g) for g in range(1, d.groups + 1))
    for i in range(n_days):
        r = start + i
        a = ws.cell(row=r, column=1,
                    value=f"=MIN($F${top + 1}:$F${top + d.groups})" if i == 0 else f"=A{r - 1}+1")
        a.number_format = "D MMM YYYY"; a.border = BOX
        b = ws.cell(row=r, column=2, value=f'=TEXT(A{r},"ddd")'); b.border = BOX
        day = cal0 + dt.timedelta(days=i)                 # the designed date of this row
        if day.weekday() >= 5:
            a.fill = b.fill = _fill(WEEKEND_BG)
        for j, (g, c) in enumerate(cols):
            pr = f"($F${grp_row[g]}+$D${coh_row[c.name]})"
            n2o = f"$B${coh_row[c.name]}"
            lag = refs.lag[c.name]
            watered = (f'AND($A{r}>{pr},$A{r}<{pr}+{n2o}-{lag},'
                       f'ISNUMBER(SEARCH(TEXT($A{r},"ddd"),{WATER})))')
            flood_txt = ("Sacrificial harvest, flood"
                         if any(p.role == dc.SACRIFICIAL and p.group == g and p.cohort == c.name
                                for p in d.pots) else "Flood")
            f = f'IF({watered},"Water","")'               # innermost: watering, else blank
            offs = c.offsets
            for h, off in reversed(list(zip(c.measure_hours, offs))):
                txt = f"{h} h N2O" + (", harvest" if off == offs[-1] else "")
                f = f'IF($A{r}={pr}+{n2o}+{off},"{txt}",{f})'
            f = f'IF($A{r}={pr}+{n2o}-{lag},"{flood_txt}",{f})'
            f = (f'IF(IF({TREAT}="",FALSE,$A{r}={pr}+{TREAT}),'
                 f'IF({watered},"Treatment; water","Treatment"),{f})')
            f = f'=IF($A{r}={pr},"Plant",{f})'
            cell = ws.cell(row=r, column=first_col + j, value=f)
            cell.border = BOX; cell.alignment = Alignment(vertical="center")
            _paint(cell, d.day_text(g, c, day))
        many = d.groups > 1
        parts = "&".join(f'IF(ISNUMBER(SEARCH("water",{L(first_col + j)}{r})),'
                         f'"{f"G{g} " if many else ""}{c.name}, ","")'
                         for j, (g, c) in enumerate(cols))
        w = ws.cell(row=r, column=who_col,
                    value=f'=IF(LEN({parts})=0,"","{"" if many else "cohorts "}"'
                          f'&LEFT({parts},LEN({parts})-2))')
        w.border = BOX; w.font = NOTE
        wd = ws.cell(row=r, column=wday_col, value=f'=IF({L(who_col)}{r}="","",A{r})')
        wd.number_format = "D MMM"; wd.border = BOX; wd.font = NOTE
        rng = f"{L(first_col)}{r}:{L(last_col)}{r}"
        for col, word in ((flood_col, "flood"), (n2o_col, "N2O")):
            t = ws.cell(row=r, column=col,
                        value=f'=SUMPRODUCT(ISNUMBER(SEARCH("{word}",{rng}))*{helper})')
            t.border = BOX; t.alignment = CENTER; t.font = BOLD

    end = start + n_days - 1
    refs.calendar = f"timeline!$A${start}:$A${end}"
    refs.water_who = f"timeline!${L(who_col)}${start}:${L(who_col)}${end}"
    refs.water_days = f"timeline!${L(wday_col)}${start}:${L(wday_col)}${end}"
    grid = f"{L(first_col)}{start}:{L(last_col)}{end}"
    tl = f"{L(first_col)}{start}"
    for word, bg, fg in CALENDAR_COLOURS:
        ws.conditional_formatting.add(grid, FormulaRule(
            formula=[f'ISNUMBER(SEARCH("{word}",{tl}))'], fill=_fill(bg),
            font=Font(bold=True, color=fg) if fg else None, stopIfTrue=True))
    # the cells also carry the designed colours as plain fills (for viewers that do not
    # run conditional formatting); in Excel these last rules clear them where an edit
    # moved the event away
    ws.conditional_formatting.add(grid, FormulaRule(
        formula=["TRUE"], fill=_fill("FFFFFF"), font=Font(bold=False, color="000000")))
    ws.conditional_formatting.add(f"A{start}:B{end}", FormulaRule(
        formula=[f"WEEKDAY($A{start},2)>5"], fill=_fill(WEEKEND_BG), stopIfTrue=True))
    ws.conditional_formatting.add(f"A{start}:B{end}", FormulaRule(
        formula=["TRUE"], fill=_fill("FFFFFF")))
    fl, nc = f"{L(flood_col)}{start}", f"{L(n2o_col)}{start}"
    both = f"{L(flood_col)}{start}:{L(n2o_col)}{end}"
    ws.conditional_formatting.add(both, FormulaRule(            # a flood on an N2O day
        formula=[f"AND(${L(flood_col)}{start}>0,${L(n2o_col)}{start}>0)"],
        fill=_fill(OVER_BG), font=Font(bold=True), stopIfTrue=True))
    ws.conditional_formatting.add(f"{L(n2o_col)}{start}:{L(n2o_col)}{end}", FormulaRule(
        formula=[f"{nc}>{CAP}"], fill=_fill(OVER_BG), font=Font(bold=True), stopIfTrue=True))
    ws.conditional_formatting.add(both, FormulaRule(
        formula=[f"{fl}=0"], font=Font(color="D1D5DB")))

    note_r = end + 2
    ws.cell(row=note_r, column=1, value=(
        "Flood pots / N2O pots turn red when N2O is above capacity or a flood falls on an N2O "
        "day. Watering follows the day names in 'Water on' (e.g. Mon, Wed, Fri) between "
        "planting and the flood.")).font = NOTE

    _widths(ws, (34, 14))
    for col in range(first_col, last_col + 1):
        ws.column_dimensions[L(col)].width = 26
    for col, w in ((5, 15), (6, 14), (7, 8), (8, 10)):
        ws.column_dimensions[L(col)].width = max(ws.column_dimensions[L(col)].width or 0, w)
    ws.column_dimensions[L(flood_col)].width = 10
    ws.column_dimensions[L(n2o_col)].width = 10
    ws.column_dimensions[L(who_col)].width = 24
    ws.column_dimensions[L(wday_col)].width = 10
    # column F also holds the planting dates: never narrower than a date
    ws.column_dimensions["F"].width = max(ws.column_dimensions["F"].width or 0, 14)
    ws.freeze_panes = ws.cell(row=start, column=first_col)
    return refs


# ======================================================================================
# DESIGN, LAYOUT
# ======================================================================================
def _design(wb, d: dc.Design, refs: Refs) -> None:
    ws = wb.create_sheet("design")
    cols = ["barcode", "role", "genotype", "cohort", "treatment", "rep", "planting_group",
            "planting_date", "table", "zone", "slot", "position", "n2o_load", "n2o_column",
            "seq", "tag"]
    _header(ws, 1, cols)
    for r, p in enumerate(d.pots, 2):
        meas = p.role == dc.MEASURED
        vals = [p.barcode, p.role, p.genotype, p.cohort or None, p.treatment or None, p.rep,
                p.group, f"={refs.planting(p.group, p.cohort)}", p.table, p.zone, p.slot, p.position,
                p.load if meas else None, p.column if meas else None, p.seq or None, p.label]
        for c, v in enumerate(vals, 1):
            ws.cell(row=r, column=c, value=v)
        ws.cell(row=r, column=1).number_format = "@"
        ws.cell(row=r, column=8).number_format = "DD MMM YYYY"
    _widths(ws, (16, 13, 12, 8, 11, 6, 9, 13, 6, 6, 6, 11, 9, 10, 6, 24))
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:{L(len(cols))}{len(d.pots) + 1}"


def layout_colours(d: dc.Design) -> tuple[str, dict]:
    """('group' | 'cohort', {key: hex}). By planting group when there are several (that is
    what is planted together), else by cohort (what is flooded and measured together)."""
    if d.groups > 1:
        return "group", {g: BRIGHT[(g - 1) % len(BRIGHT)] for g in range(1, d.groups + 1)}
    return "cohort", {c.name: BRIGHT[i % len(BRIGHT)] for i, c in enumerate(d.plan.cohorts)}


def _layout(wb, d: dc.Design) -> None:
    plan = d.plan
    ws = wb.create_sheet("layout")
    ws["A1"] = (f"Bench layout: {plan.tables} tables x {plan.zones_per_table} zones x "
                f"{plan.slots_per_zone} slots. Zone 1 at the taped end; slots count from the "
                f"left. Cell: tag / barcode.")
    ws["A1"].font = NOTE
    by, colours = layout_colours(d)

    # ---- legend ----
    ws["A3"] = f"Colour = {'planting group' if by == 'group' else 'cohort'}"
    ws["A3"].font = BOLD
    r = 4
    for key, hex_ in colours.items():
        if by == "group":
            n = sum(p.group == key for p in d.pots)
            text = f"Planting group {key}: planted {d.group_planting(key):%a %d %b %Y}, {n} pots"
        else:
            n = sum(p.cohort == key for p in d.pots)
            text = f"Cohort {key}: {n} pots"
        sw = ws.cell(row=r, column=1, value=f"G{key}" if by == "group" else key)
        sw.fill = _fill(hex_); sw.alignment = CENTER; sw.border = BOX; sw.font = BOLD
        ws.cell(row=r, column=2, value=text)
        r += 1
    if any(p.role == dc.SPARE for p in d.pots):
        sw = ws.cell(row=r, column=1, value="SPARE")
        if by == "cohort":
            sw.fill = _fill(SPARE_BG)
        sw.alignment = CENTER; sw.border = BOX; sw.font = BOLD
        ws.cell(row=r, column=2, value=(
            "Spare: no cohort or treatment yet; replaces a failed pot of its own planting "
            "group" + ("" if by == "cohort" else " (fill = that group)")))
        r += 1
    ws.cell(row=r, column=1, value="SAC").font = Font(bold=True, italic=True)
    ws.cell(row=r, column=1).alignment = CENTER
    ws.cell(row=r, column=2, value="italic tag ending SACn: sacrificial, harvested on flood day")
    r += 1
    sw = ws.cell(row=r, column=1, value="(empty)")
    sw.font = Font(color="9CA3AF"); sw.alignment = CENTER; sw.border = BOX
    ws.cell(row=r, column=2, value="empty slot")
    r += 3

    at = {(p.table, p.zone, p.slot): p for p in d.pots}
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
                if p is None:
                    c.font = Font(color="9CA3AF")
                    continue
                key = p.group if by == "group" else p.cohort
                c.fill = _fill(colours.get(key, SPARE_BG))
                if p.role == dc.SACRIFICIAL:
                    c.font = Font(italic=True)
            ws.row_dimensions[r].height = 32
        r += 3
    ws.column_dimensions["A"].width = 18
    for s in range(1, plan.slots_per_zone + 1):
        ws.column_dimensions[L(1 + s)].width = 19
    ws.column_dimensions["B"].width = max(ws.column_dimensions["B"].width, 19)


# ======================================================================================
# FLOOD / MEASURE / HARVEST, per planting group x cohort
# ======================================================================================
def _dated(ws, row, label, formula):
    ws.cell(row=row, column=1, value=label).font = BOLD
    c = ws.cell(row=row, column=2, value=f"={formula}")
    c.number_format = "ddd D MMM YYYY"; c.font = BOLD


def _cohort_sheets(wb, d: dc.Design, refs: Refs) -> None:
    plan = d.plan
    for g in range(1, d.groups + 1):
        for c in plan.cohorts:
            tag = _sheet_tag(d, g, c.name)
            who = f"Cohort {c.name}" + (f", planting group {g}" if d.groups > 1 else "")
            meas = sorted((p for p in d.pots if p.group == g and p.cohort == c.name
                           and p.role == dc.MEASURED), key=lambda p: p.seq)
            sac = sorted((p for p in d.pots if p.group == g and p.cohort == c.name
                          and p.role == dc.SACRIFICIAL), key=lambda p: p.seq)

            # ---- flood ----
            ws = wb.create_sheet(f"flood_{tag}")
            ws["A1"] = (f"{who}: flood morning. "
                        + (f"1) harvest the {len(sac)} sacrificial pots, 2) flood the "
                           f"{len(meas)} measured pots in this order." if sac else
                           f"Flood the {len(meas)} measured pots in this order."))
            ws["A1"].font = Font(bold=True, size=12)
            _dated(ws, 2, "Date", refs.flood_date(g, c.name))
            r = 4
            if sac:
                _header(ws, r, ["SACRIFICIAL harvest first", "barcode", "position", "genotype",
                                "treatment", "fresh shoot g", "fresh root g", "soil depth cm",
                                "KCl tube", "done"])
                for p in sac:
                    r += 1
                    for j, v in enumerate((p.label, p.barcode, p.position, p.genotype,
                                           p.treatment or None)):
                        ws.cell(row=r, column=1 + j, value=v)
                    for j in range(5, 10):
                        ws.cell(row=r, column=1 + j).border = BOX
                r += 2
            _header(ws, r, ["FLOOD in this order", "barcode", "position", "genotype",
                            "treatment", "mass before g", "flood pass 1 time",
                            "flood pass 2 time", "achieved g", "flag"])
            for p in meas:
                r += 1
                for j, v in enumerate((f"{p.seq:02d} (load {p.load}, col {p.column})",
                                       p.barcode, p.position, p.genotype, p.treatment or None)):
                    ws.cell(row=r, column=1 + j, value=v)
                for j in range(5, 10):
                    ws.cell(row=r, column=1 + j).border = BOX
            _widths(ws, (26, 15, 11, 11, 11, 13, 15, 15, 11, 9))
            ws.freeze_panes = "A4"

            # ---- measure: one block per N2O time ----
            ws = wb.create_sheet(f"measure_{tag}")
            ws["A1"] = (f"{who}: N2O measurement. Loads 1-{max((p.load for p in meas), default=0)}"
                        f" in order, {dc.N2O_COLUMNS} pots per load (one per column); each load"
                        f" mixes treatments and genotypes.")
            ws["A1"].font = Font(bold=True, size=12)
            r = 3
            offs = c.offsets
            for h, off in zip(c.measure_hours, offs):
                _dated(ws, r, f"{h} h N2O" + (", then harvest" if off == offs[-1] else ""),
                       refs.n2o_date(g, c.name, off))
                r += 1
                for j, lab in enumerate(("chamber air T at Start (C):", None, "ambient kPa:",
                                         None, "operator:", None, "k entered per column:")):
                    if lab:
                        ws.cell(row=r, column=1 + j, value=lab).font = NOTE
                    else:
                        ws.cell(row=r, column=1 + j).border = BOX
                r += 1
                _header(ws, r, ["load", "column", "seq", "barcode", "position", "genotype",
                                "treatment", "rep", "top-up g", "mass at seat g", "seat time",
                                "window ok?", "mass out g", "note"])
                for p in meas:
                    r += 1
                    for j, v in enumerate((p.load, p.column, p.seq, p.barcode, p.position,
                                           p.genotype, p.treatment or None, p.rep)):
                        ws.cell(row=r, column=1 + j, value=v)
                    for j in range(8, 14):
                        ws.cell(row=r, column=1 + j).border = BOX
                    if p.column == dc.N2O_COLUMNS or p is meas[-1]:
                        for j in range(14):                 # rule under each load
                            ws.cell(row=r, column=1 + j).border = Border(
                                left=THIN, right=THIN, top=THIN,
                                bottom=Side(style="medium", color="6B7280"))
                r += 3
            _widths(ws, (26, 7, 6, 15, 11, 11, 11, 6, 9, 13, 10, 11, 11, 20))

            # ---- harvest ----
            ws = wb.create_sheet(f"harvest_{tag}")
            ws["A1"] = f"{who}: harvest after the {c.measure_hours[-1]} h N2O"
            ws["A1"].font = Font(bold=True, size=12)
            _dated(ws, 2, "Date", refs.n2o_date(g, c.name, offs[-1]))
            _header(ws, 4, ["seq", "barcode", "position", "genotype", "treatment", "rep",
                            "fresh shoot weight (g)", "DRY shoot weight (g)",
                            "fresh root weight (g)", "DRY root weight (g)", "KCl tube?"])
            for i, p in enumerate(meas, 5):
                for j, v in enumerate((p.seq, p.barcode, p.position, p.genotype,
                                       p.treatment or None, p.rep)):
                    ws.cell(row=i, column=1 + j, value=v)
                for j in range(6, 11):
                    ws.cell(row=i, column=1 + j).border = BOX
            _widths(ws, (8, 15, 11, 11, 11, 6, 15, 15, 15, 15, 10))
            ws.freeze_panes = "A5"


# ======================================================================================
# TAGS, GROUPS, JARS, INFO
# ======================================================================================
def _tags(wb, d: dc.Design, refs: Refs) -> None:
    ws = wb.create_sheet("tags")
    _header(ws, 1, ["barcode", "tag_text", "position", "planting_group", "planting_date",
                    "role", "genotype", "cohort", "treatment"])
    for r, p in enumerate(sorted(d.pots, key=lambda p: (p.group, p.barcode)), 2):
        for j, v in enumerate((p.barcode, p.label, p.position, p.group,
                               f"={refs.planting(p.group, p.cohort)}", p.role, p.genotype,
                               p.cohort or None, p.treatment or None)):
            ws.cell(row=r, column=1 + j, value=v)
        ws.cell(row=r, column=1).number_format = "@"
        ws.cell(row=r, column=5).number_format = "DD MMM YYYY"
    _widths(ws, (16, 24, 11, 9, 13, 13, 12, 8, 11))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:I{len(d.pots) + 1}"


def _groups(wb, d: dc.Design, refs: Refs) -> None:
    plan = d.plan
    ws = wb.create_sheet("groups")
    _header(ws, 1, ["Planting group", "Planting date", "Pots"]
            + [f"Measured, cohort {c.name}" for c in plan.cohorts])
    for g in range(1, d.groups + 1):
        r = g + 1
        ws.cell(row=r, column=1, value=g).font = BOLD
        ws.cell(row=r, column=2, value=f"={refs.plant[g]}").number_format = "DD MMM YYYY"
        ws.cell(row=r, column=3, value=sum(p.group == g for p in d.pots))
        for j, c in enumerate(plan.cohorts):
            ws.cell(row=r, column=4 + j, value=sum(
                p.group == g and p.cohort == c.name and p.role == dc.MEASURED for p in d.pots))
    r = d.groups + 4
    ws.cell(row=r, column=1, value=(
        f"N2O load per day as designed (capacity {plan.max_measured_per_day}; the timeline's "
        f"last columns recalculate if dates are edited)")).font = BOLD
    _header(ws, r + 1, ["Date", "N2O measurements"])
    for i, (day, n) in enumerate(d.calendar_load().items()):
        ws.cell(row=r + 2 + i, column=1, value=day).number_format = "DD MMM YYYY (ddd)"
        ws.cell(row=r + 2 + i, column=2, value=n)
    _widths(ws, (20, 18) + (20,) * (len(plan.cohorts) + 1))


WATER_SLACK = 6            # spare watering rows, for dates edited later


def _asana(wb, d: dc.Design, refs: Refs) -> None:
    """Header + one row per task and nothing else, so File > Save As > CSV imports as is.
    Column names are the ones Asana's CSV importer maps by itself. Every due date is a
    formula on the timeline (WORKDAY for the working-day lags), so editing a date there
    moves the tasks. Watering is the n-th watering day of the timeline, with a few spare
    rows that stay blank unless edited dates add watering days."""
    ws = wb.create_sheet("asana_tasks")
    _header(ws, 1, ["Name", "Due Date", "Section/Column", "Description"])
    tasks = d.tasks()
    n_water = sum(t.anchor[0] == "water" for t in tasks)
    extra = [dc.Task(None, "Maintenance watering", "Watering", "", ("water", n_water + k))
             for k in range(1, WATER_SLACK + 1)]
    for r, t in enumerate(tasks + extra, 2):
        x = refs.anchor(t.anchor)
        due = (x if t.lag is None else f"WORKDAY({x},{t.lag})")
        if t.anchor[0] == "water":
            B = f"B{r}"
            ws.cell(row=r, column=2, value=f'=IFERROR({due},"")')
            ws.cell(row=r, column=1, value=f'=IF({B}="","","{t.name}")')
            ws.cell(row=r, column=3, value=f'=IF({B}="","","{t.section}")')
            ws.cell(row=r, column=4, value=f'=IF({B}="","",INDEX({refs.water_who},'
                                           f'MATCH({B},{refs.calendar},0)))')
        else:
            ws.cell(row=r, column=1, value=t.name)
            ws.cell(row=r, column=2, value=f"={due}")
            ws.cell(row=r, column=3, value=t.section)
            ws.cell(row=r, column=4, value=t.notes or None)
        ws.cell(row=r, column=2).number_format = "mm/dd/yyyy"
    _widths(ws, (78, 12, 15, 60))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:D{ws.max_row}"


def _jars(wb, d: dc.Design) -> None:
    ws = wb.create_sheet("jars")
    ws.append(["barcode", "tag"])
    for c in ws[1]:
        c.font = BOLD; c.alignment = Alignment(horizontal="center")
    for p in d.pots:
        ws.append([p.barcode, f"{p.label} {p.position}"])
    for row in ws.iter_rows(min_row=2, max_col=1):
        row[0].number_format = "@"                      # Text: keeps leading zeros
    _widths(ws, (16, 34))
    ws.freeze_panes = "A2"


def _admin(wb, d: dc.Design) -> None:
    """The dosing station's `admin` tab, in the station template's order. Values, never
    formulas: provisioning refuses a formula it cannot see the result of."""
    plan = d.plan
    ws = wb.create_sheet("admin")
    pct = lambda v: f"= {v * 100:g} %" if v is not None else ""
    rows = [
        ("trial_id", plan.name, "experiment id (the barcodes' prefix)", False),
        ("soil_label", None, "optional", False),
        ("operator", None, "optional", False),
        ("date", plan.planting_date.isoformat(), "planting date, group 1", False),
        ("profile", "wet_up", "first watering: wet-up from air-dry soil", False),
        ("bulk_density", plan.bulk_density, "g/cm3", True),
        ("air_dry_moisture", plan.air_dry_moisture,
         f"fraction of air-dry soil mass {pct(plan.air_dry_moisture)}", True),
        ("particle_density", plan.particle_density, "g/cm3 (2.65 if unknown)", False),
        ("wfps_target", plan.wfps_target,
         f"fraction: wet-up and maintenance target {pct(plan.wfps_target)}", True),
        ("wfps_flood", plan.wfps_flood,
         f"fraction: flood target {pct(plan.wfps_flood)} (blank: no flood target)", True),
        ("container_weight_g", plan.container_weight_g, "empty pot, g", True),
        ("air_dry_soil_g", plan.air_dry_soil_g, "air-dry soil per pot, g", True),
    ]
    for r, (key, val, note, yellow) in enumerate(rows, 1):
        ws.cell(row=r, column=1, value=key).font = BOLD
        c = ws.cell(row=r, column=2, value=val)
        if yellow:
            _input(c)
        missing = val is None and key in dc.ADMIN_REQUIRED
        n = ws.cell(row=r, column=3, value=("REQUIRED: fill in before Start New Trial. "
                                            if missing else "") + note)
        n.font = Font(bold=True, color="B91C1C") if missing else NOTE
    _widths(ws, (22, 16, 60))


def _info(wb, d: dc.Design) -> None:
    plan = d.plan
    ws = wb.create_sheet("info")
    c = d.counts()
    rows = [("experiment", plan.name), ("generated", dt.datetime.now().strftime("%Y-%m-%d %H:%M")),
            ("tool version", dc.__version__), ("seed", plan.seed),
            ("pots", len(d.pots)), ("measured", c[dc.MEASURED]),
            ("sacrificial", c[dc.SACRIFICIAL]), ("spare", c[dc.SPARE]),
            ("spares per genotype per planting group", plan.spares_per_genotype_per_group),
            ("tables", plan.tables), ("zones per table", plan.zones_per_table),
            ("slots per zone", plan.slots_per_zone), ("bench capacity", plan.capacity),
            ("planting groups", d.groups), ("days between groups", d.stagger_days),
            ("max N2O per day (designed)", d.max_load),
            ("max N2O per day (capacity)", plan.max_measured_per_day),
            ("N2O columns per load", dc.N2O_COLUMNS),
            ("watering", dc.WATER_DAY_NAMES),
            ("planting date, group 1", plan.planting_date.isoformat()),
            ("treatment DAP", plan.treatment_dap if plan.treatment_dap is not None else "none")]
    for k, v in rows:
        ws.append([k, v])
        ws.cell(row=ws.max_row, column=1).font = BOLD
    ws.append([])
    ws.append(["cohort", "first N2O DAP", "planted, days after group", "N2O after flood (h)"])
    for cc in plan.cohorts:
        ws.append([cc.name, cc.first_n2o_dap, d.offsets.get(cc.name, 0),
                   ", ".join(str(h) for h in cc.measure_hours)])
    ws.append([])
    ws.append(["genotype", "cohort", "treatment", "reps", "sacrificial"])
    for pr in plan.rows:
        if pr.genotype.strip():
            ws.append([pr.genotype, pr.cohort, pr.treatment or None, pr.reps, pr.sacrificial])
    for w in d.warnings:
        ws.append(["WARNING", w])
    _widths(ws, (28, 20))


def write(d: dc.Design, path: str) -> None:
    wb = Workbook()
    wb.remove(wb.active)
    refs = _timeline(wb, d)
    _design(wb, d, refs)
    _layout(wb, d)
    _cohort_sheets(wb, d, refs)
    _tags(wb, d, refs)
    _groups(wb, d, refs)
    _asana(wb, d, refs)
    _jars(wb, d)
    _admin(wb, d)
    _info(wb, d)
    wb.save(path)
