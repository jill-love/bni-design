#!/usr/bin/env python3
"""
design_app.py: BNI experiment setup window (v0.2, 2026-10-28).

    python design_app.py

One window: the experiment settings, the bench, the N2O capacity, a cohort list and a
spreadsheet-style plan (genotype, cohort, treatment, reps, sacrificial). Generate builds
the design (design_core.py) and shows a summary; Save workbook writes the .xlsx
(design_xlsx.py). The plan itself can be saved and reopened as a .json file.

The plan table takes a paste from Excel (Ctrl+V, several rows and columns at once) and
Delete clears the selected cells, so a plan typed in a spreadsheet can be dropped in.
"""
from __future__ import annotations

import datetime as dt
import json
import os
import sys
from collections import Counter

from PySide6 import QtCore, QtGui, QtWidgets
from PySide6.QtCore import Qt

import design_core as dc
import design_xlsx as dx

TITLE = f"BNI Experiment Setup v{dc.__version__}"
SLATE, BLUE, GREEN, RED, MUTED = "#374151", "#1d4ed8", "#15803d", "#b91c1c", "#6b7280"
PLAN_COLS = ("Genotype", "Cohort", "Treatment", "Reps (measured)", "Sacrificial")
BLANK_ROWS = 5             # the plan table always ends with this many empty rows


def example_plan() -> dc.Plan:
    """What the window opens with: the P0 shape, one genotype, three cohorts."""
    return dc.Plan(
        name="BNI-P1", seed=dt.date.today().toordinal() % 100000,
        planting_date=dt.date.today() + dt.timedelta(days=7), treatment_dap=14,
        cohorts=[dc.Cohort("A", 21), dc.Cohort("B", 24), dc.Cohort("C", 27)],
        rows=[dc.PlanRow("B73", c, t, 10, 3) for c in "ABC" for t in ("CTRL", "SN", "INH")])


# ======================================================================================
# A QTableWidget that behaves like a small spreadsheet
# ======================================================================================
class SheetTable(QtWidgets.QTableWidget):
    def __init__(self, headers, numeric_cols=(), parent=None):
        super().__init__(0, len(headers), parent)
        self.setHorizontalHeaderLabels(headers)
        self.numeric_cols = set(numeric_cols)
        self.horizontalHeader().setStretchLastSection(True)
        self.verticalHeader().setDefaultSectionSize(24)
        self.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)

    def text(self, r, c) -> str:
        it = self.item(r, c)
        return it.text().strip() if it else ""

    def set_text(self, r, c, v):
        it = QtWidgets.QTableWidgetItem("" if v is None else str(v))
        if c in self.numeric_cols:
            it.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.setItem(r, c, it)

    def keyPressEvent(self, ev):
        if ev.matches(QtGui.QKeySequence.Paste):
            self.paste(); return
        if ev.matches(QtGui.QKeySequence.Copy):
            self.copy(); return
        if ev.key() in (Qt.Key_Delete, Qt.Key_Backspace) and self.state() != self.EditingState:
            for it in self.selectedItems():
                it.setText("")
            return
        super().keyPressEvent(ev)

    def paste(self):
        text = QtWidgets.QApplication.clipboard().text()
        if not text:
            return
        rows = [line.split("\t") for line in text.rstrip("\r\n").splitlines()]
        r0 = max(0, self.currentRow()); c0 = max(0, self.currentColumn())
        if r0 + len(rows) > self.rowCount():
            self.setRowCount(r0 + len(rows))
        for i, vals in enumerate(rows):
            for j, v in enumerate(vals):
                if c0 + j < self.columnCount():
                    self.set_text(r0 + i, c0 + j, v.strip())
        self.itemChanged.emit(self.item(r0, c0))

    def copy(self):
        rngs = self.selectedRanges()
        if not rngs:
            return
        g = rngs[0]
        lines = ["\t".join(self.text(r, c) for c in range(g.leftColumn(), g.rightColumn() + 1))
                 for r in range(g.topRow(), g.bottomRow() + 1)]
        QtWidgets.QApplication.clipboard().setText("\n".join(lines))


# ======================================================================================
# THE WINDOW
# ======================================================================================
class DesignWindow(QtWidgets.QMainWindow):
    def __init__(self, plan: dc.Plan | None = None):
        super().__init__()
        self.setWindowTitle(TITLE)
        self.resize(1280, 860)
        self.design: dc.Design | None = None
        self.plan_path: str | None = None

        root = QtWidgets.QWidget(); self.setCentralWidget(root)
        v = QtWidgets.QVBoxLayout(root)

        # ---- banner ----
        self.banner = QtWidgets.QLabel()
        self.banner.setAlignment(Qt.AlignCenter)
        self.banner.setFixedHeight(52)
        f = QtGui.QFont(); f.setPointSize(14); f.setBold(True); self.banner.setFont(f)
        v.addWidget(self.banner)

        body = QtWidgets.QHBoxLayout(); v.addLayout(body, 1)

        # ---- left: settings ----
        left = QtWidgets.QVBoxLayout(); body.addLayout(left, 0)
        exp = QtWidgets.QGroupBox("Experiment"); fl = QtWidgets.QFormLayout(exp)
        self.name = QtWidgets.QLineEdit()
        self.name.setToolTip("Experiment id. Barcodes are <id>-S001, -S002, ...")
        fl.addRow("Experiment id", self.name)
        self.plant = QtWidgets.QDateEdit(); self.plant.setCalendarPopup(True)
        self.plant.setDisplayFormat("ddd d MMM yyyy")
        fl.addRow("Planting date (group 1)", self.plant)
        self.treat_on = QtWidgets.QCheckBox("Treatment on day")
        self.treat = QtWidgets.QSpinBox(); self.treat.setRange(0, 365); self.treat.setSuffix(" after planting")
        self.treat_on.toggled.connect(self.treat.setEnabled)
        tr = QtWidgets.QHBoxLayout(); tr.addWidget(self.treat_on); tr.addWidget(self.treat, 1)
        fl.addRow(tr)
        self.spares = QtWidgets.QSpinBox(); self.spares.setRange(0, 100)
        self.spares.setToolTip("Spare pots per genotype, placed at random like the rest")
        fl.addRow("Spares per genotype", self.spares)
        self.seed = QtWidgets.QSpinBox(); self.seed.setRange(0, 999999)
        self.seed.setToolTip("The same seed always gives the same layout and grouping. "
                             "Change it for a different randomisation.")
        fl.addRow("Random seed", self.seed)
        left.addWidget(exp)

        bench = QtWidgets.QGroupBox("Bench"); bl = QtWidgets.QFormLayout(bench)
        self.tables = QtWidgets.QSpinBox(); self.tables.setRange(1, 20)
        self.tables.setToolTip("Tables open for this experiment. Pots are spread over these only.")
        bl.addRow("Tables available", self.tables)
        self.zones = QtWidgets.QSpinBox(); self.zones.setRange(1, dc.MAX_ZONES)
        bl.addRow("Zones per table", self.zones)
        self.slots = QtWidgets.QSpinBox(); self.slots.setRange(1, dc.MAX_SLOTS)
        bl.addRow("Slots per zone", self.slots)
        self.cap_lbl = QtWidgets.QLabel(); self.cap_lbl.setStyleSheet(f"color:{MUTED};")
        bl.addRow("", self.cap_lbl)
        left.addWidget(bench)

        sched = QtWidgets.QGroupBox("N2O schedule"); sl = QtWidgets.QFormLayout(sched)
        self.cap = QtWidgets.QSpinBox(); self.cap.setRange(1, 1000)
        sl.addRow("Max N2O measurements / day", self.cap)
        self.min_gap = QtWidgets.QSpinBox(); self.min_gap.setRange(2, dc.MAX_STAGGER_DAYS)
        self.min_gap.setSuffix(" days")
        self.min_gap.setToolTip("Shortest gap between planting groups. 3 (the default) gives "
                                "one group's 24 h and 48 h N2O days, then the next group's "
                                "flood: a flood and N2O are never on the same day.")
        sl.addRow("Min. gap between groups", self.min_gap)
        self.groups = QtWidgets.QSpinBox(); self.groups.setRange(0, 50)
        self.groups.setSpecialValueText("Auto (fewest that fit)")
        sl.addRow("Planting groups", self.groups)
        note = QtWidgets.QLabel("Flood 24 h before the first N2O sample (never on an N2O day); "
                                "N2O at 24 h and 48 h; harvest after the 48 h sample. Water "
                                "Monday, Wednesday and Friday.")
        note.setWordWrap(True); note.setStyleSheet(f"color:{MUTED};")
        sl.addRow(note)
        left.addWidget(sched)

        coh = QtWidgets.QGroupBox("Cohorts (time trial)"); cl = QtWidgets.QVBoxLayout(coh)
        self.cohorts = SheetTable(("Cohort", "First N2O (days after planting)"), numeric_cols=(1,))
        self.cohorts.setMinimumHeight(140)
        self.cohorts.setColumnWidth(0, 80)
        cl.addWidget(self.cohorts)
        cb = QtWidgets.QHBoxLayout()
        add_c = QtWidgets.QPushButton("Add cohort"); add_c.clicked.connect(self._add_cohort)
        del_c = QtWidgets.QPushButton("Remove"); del_c.clicked.connect(
            lambda: self._remove_rows(self.cohorts))
        cb.addWidget(add_c); cb.addWidget(del_c); cb.addStretch(1)
        cl.addLayout(cb)
        left.addWidget(coh, 1)

        # ---- middle: the plan ----
        mid = QtWidgets.QVBoxLayout(); body.addLayout(mid, 1)
        pg = QtWidgets.QGroupBox("Plan: one row per genotype x cohort x treatment")
        pl = QtWidgets.QVBoxLayout(pg)
        hint = QtWidgets.QLabel("Type, or paste rows from Excel (Ctrl+V). Treatment may be left "
                                "blank. Sacrificial pots are harvested on flood day.")
        hint.setStyleSheet(f"color:{MUTED};"); hint.setWordWrap(True)
        pl.addWidget(hint)
        self.table = SheetTable(PLAN_COLS, numeric_cols=(3, 4))
        for c, w in enumerate((110, 60, 100, 110)):
            self.table.setColumnWidth(c, w)
        self.table.setMinimumWidth(500)
        self.table.itemChanged.connect(self._plan_edited)
        pl.addWidget(self.table, 1)
        pb = QtWidgets.QHBoxLayout()
        fill = QtWidgets.QPushButton("Fill every combination...")
        fill.setToolTip("Write one row for every genotype x cohort x treatment")
        fill.clicked.connect(self._fill_combinations)
        del_r = QtWidgets.QPushButton("Remove rows"); del_r.clicked.connect(
            lambda: self._remove_rows(self.table))
        self.total_lbl = QtWidgets.QLabel(); self.total_lbl.setStyleSheet("font-weight:bold;")
        pb.addWidget(fill); pb.addWidget(del_r); pb.addStretch(1); pb.addWidget(self.total_lbl)
        pl.addLayout(pb)
        mid.addWidget(pg, 1)

        # ---- right: result ----
        right = QtWidgets.QVBoxLayout(); body.addLayout(right, 1)
        rg = QtWidgets.QGroupBox("Design"); rl = QtWidgets.QVBoxLayout(rg)
        self.summary = QtWidgets.QPlainTextEdit(); self.summary.setReadOnly(True)
        mono = QtGui.QFontDatabase.systemFont(QtGui.QFontDatabase.FixedFont)
        self.summary.setFont(mono)
        self.summary.setLineWrapMode(QtWidgets.QPlainTextEdit.NoWrap)
        rl.addWidget(self.summary, 1)
        right.addWidget(rg, 1)

        # ---- buttons ----
        brow = QtWidgets.QHBoxLayout()
        open_b = QtWidgets.QPushButton("Open plan..."); open_b.clicked.connect(self._open_plan)
        save_p = QtWidgets.QPushButton("Save plan..."); save_p.clicked.connect(self._save_plan)
        self.gen_b = QtWidgets.QPushButton("Generate design")
        self.gen_b.setMinimumHeight(38)
        self.gen_b.setStyleSheet(f"QPushButton{{background:{BLUE};color:white;font-weight:bold;}}")
        self.gen_b.clicked.connect(self._generate)
        self.save_b = QtWidgets.QPushButton("Save workbook...")
        self.save_b.setMinimumHeight(38)
        self.save_b.setStyleSheet(
            f"QPushButton{{background:{GREEN};color:white;font-weight:bold;}}"
            "QPushButton:disabled{background:#d1d5db;color:#6b7280;}")
        self.save_b.clicked.connect(self._save_workbook)
        brow.addWidget(open_b); brow.addWidget(save_p); brow.addStretch(1)
        brow.addWidget(self.gen_b, 1); brow.addWidget(self.save_b, 1)
        v.addLayout(brow)

        for w in (self.tables, self.zones, self.slots):
            w.valueChanged.connect(self._refresh_capacity)
        for w in (self.name, ):
            w.textChanged.connect(self._stale)
        for w in (self.tables, self.zones, self.slots, self.cap, self.min_gap, self.groups,
                  self.treat, self.spares, self.seed):
            w.valueChanged.connect(self._stale)
        self.plant.dateChanged.connect(self._stale)
        self.treat_on.toggled.connect(self._stale)
        self.cohorts.itemChanged.connect(self._stale)

        self.load_plan(plan or example_plan())

    # ---- plan <-> widgets ----
    def load_plan(self, p: dc.Plan):
        self.table.blockSignals(True); self.cohorts.blockSignals(True)
        self.name.setText(p.name)
        self.plant.setDate(QtCore.QDate(p.planting_date.year, p.planting_date.month,
                                        p.planting_date.day))
        self.treat_on.setChecked(p.treatment_dap is not None)
        self.treat.setEnabled(p.treatment_dap is not None)
        self.treat.setValue(p.treatment_dap or 0)
        self.spares.setValue(p.spares_per_genotype)
        self.seed.setValue(p.seed)
        self.tables.setValue(p.tables)
        self.zones.setValue(p.zones_per_table)
        self.slots.setValue(p.slots_per_zone)
        self.cap.setValue(p.max_measured_per_day)
        self.min_gap.setValue(p.min_stagger_days)
        self.groups.setValue(p.planting_groups)
        self.cohorts.setRowCount(len(p.cohorts))
        for r, c in enumerate(p.cohorts):
            self.cohorts.set_text(r, 0, c.name); self.cohorts.set_text(r, 1, c.first_n2o_dap)
        self.table.setRowCount(len(p.rows) + BLANK_ROWS)
        for r, row in enumerate(p.rows):
            for c, val in enumerate((row.genotype, row.cohort, row.treatment, row.reps,
                                     row.sacrificial)):
                self.table.set_text(r, c, val)
        self.table.blockSignals(False); self.cohorts.blockSignals(False)
        self._refresh_capacity(); self._plan_edited()

    def read_plan(self) -> dc.Plan:
        """Widgets -> Plan. Raises PlanError for text that is not a number."""
        def num(tbl, r, c, what):
            t = tbl.text(r, c)
            if not t:
                return 0
            try:
                return int(float(t))
            except ValueError:
                raise dc.PlanError(f"{what} in row {r + 1} is not a number: {t!r}")
        cohorts = []
        for r in range(self.cohorts.rowCount()):
            n = self.cohorts.text(r, 0)
            if n:
                cohorts.append(dc.Cohort(n, num(self.cohorts, r, 1, "first N2O day")))
        rows = []
        for r in range(self.table.rowCount()):
            g = self.table.text(r, 0)
            if not any(self.table.text(r, c) for c in range(5)):
                continue
            rows.append(dc.PlanRow(g, self.table.text(r, 1), self.table.text(r, 2),
                                   num(self.table, r, 3, "reps"),
                                   num(self.table, r, 4, "sacrificial")))
        d = self.plant.date()
        return dc.Plan(
            name=self.name.text().strip(), seed=self.seed.value(),
            planting_date=dt.date(d.year(), d.month(), d.day()),
            treatment_dap=self.treat.value() if self.treat_on.isChecked() else None,
            tables=self.tables.value(), zones_per_table=self.zones.value(),
            slots_per_zone=self.slots.value(), max_measured_per_day=self.cap.value(),
            min_stagger_days=self.min_gap.value(), planting_groups=self.groups.value(),
            spares_per_genotype=self.spares.value(), cohorts=cohorts, rows=rows)

    # ---- small helpers ----
    def _paint(self, color, text):
        self.banner.setStyleSheet(f"color:white;background:{color};border-radius:8px;")
        self.banner.setText(text)

    def _refresh_capacity(self, *_):
        cap = self.tables.value() * self.zones.value() * self.slots.value()
        self.cap_lbl.setText(f"room for {cap} pots")
        self._plan_edited()

    def _plan_edited(self, *_):
        if self.table.rowCount() - self._last_used_row() < BLANK_ROWS:
            self.table.blockSignals(True)
            self.table.setRowCount(self._last_used_row() + BLANK_ROWS)
            self.table.blockSignals(False)
        try:
            n = dc.count_pots(self.read_plan())
            cap = self.tables.value() * self.zones.value() * self.slots.value()
            self.total_lbl.setText(f"{n} pots / room for {cap}")
            self.total_lbl.setStyleSheet(f"font-weight:bold;color:{RED if n > cap else SLATE};")
        except dc.PlanError:
            self.total_lbl.setText("")
        self._stale()

    def _last_used_row(self) -> int:
        for r in range(self.table.rowCount() - 1, -1, -1):
            if any(self.table.text(r, c) for c in range(5)):
                return r + 1
        return 0

    def _stale(self, *_):
        """Any edit makes the generated design out of date."""
        self.design = None
        if hasattr(self, "save_b"):
            self.save_b.setEnabled(False)
            self._paint(SLATE, "Enter the plan, then Generate design")

    def _add_cohort(self):
        r = self.cohorts.rowCount()
        self.cohorts.setRowCount(r + 1)
        last = self.cohorts.text(r - 1, 1) if r else ""
        name = chr(ord("A") + r) if r < 26 else f"C{r + 1}"
        self.cohorts.set_text(r, 0, name)
        self.cohorts.set_text(r, 1, int(last) + 3 if last.isdigit() else 21)

    def _remove_rows(self, tbl):
        rows = sorted({i.row() for i in tbl.selectedIndexes()}, reverse=True)
        for r in rows:
            tbl.removeRow(r)
        self._plan_edited()

    def _fill_combinations(self):
        dlg = QtWidgets.QDialog(self); dlg.setWindowTitle("Fill every combination")
        fl = QtWidgets.QFormLayout(dlg)
        genos = QtWidgets.QLineEdit("B73"); genos.setPlaceholderText("B73, Mo17, W22")
        treats = QtWidgets.QLineEdit("CTRL, SN, INH")
        treats.setPlaceholderText("blank: no treatment")
        reps = QtWidgets.QSpinBox(); reps.setRange(0, 999); reps.setValue(10)
        sac = QtWidgets.QSpinBox(); sac.setRange(0, 999); sac.setValue(0)
        cohorts = [self.cohorts.text(r, 0) for r in range(self.cohorts.rowCount())
                   if self.cohorts.text(r, 0)]
        fl.addRow("Genotypes (comma separated)", genos)
        fl.addRow("Treatments (comma separated)", treats)
        fl.addRow("Cohorts", QtWidgets.QLabel(", ".join(cohorts) or "(add cohorts first)"))
        fl.addRow("Reps (measured)", reps)
        fl.addRow("Sacrificial", sac)
        bb = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok
                                        | QtWidgets.QDialogButtonBox.Cancel)
        bb.accepted.connect(dlg.accept); bb.rejected.connect(dlg.reject)
        fl.addRow(bb)
        if dlg.exec() != QtWidgets.QDialog.Accepted:
            return
        gs = [g.strip() for g in genos.text().split(",") if g.strip()]
        ts = [t.strip() for t in treats.text().split(",") if t.strip()] or [""]
        rows = [(g, c, t, reps.value(), sac.value()) for g in gs for c in cohorts for t in ts]
        if self._last_used_row():
            ans = QtWidgets.QMessageBox.question(
                self, "Replace the plan?", f"Replace the {self._last_used_row()} rows in the "
                f"plan with these {len(rows)}?")
            if ans != QtWidgets.QMessageBox.Yes:
                return
        self.table.blockSignals(True)
        self.table.setRowCount(0); self.table.setRowCount(len(rows) + BLANK_ROWS)
        for r, vals in enumerate(rows):
            for c, val in enumerate(vals):
                self.table.set_text(r, c, val)
        self.table.blockSignals(False)
        self._plan_edited()

    # ---- generate / save ----
    def _generate(self):
        try:
            plan = self.read_plan()
            problems = dc.validate(plan)
            if problems:
                raise dc.PlanError("\n".join(problems))
            self.design = dc.build(plan)
        except dc.PlanError as ex:
            self.design = None
            self.save_b.setEnabled(False)
            self._paint(RED, "Cannot build the design: see the list on the right")
            self.summary.setPlainText("Fix these and Generate again:\n\n- "
                                      + str(ex).replace("\n", "\n- "))
            return
        d = self.design
        warn = bool(d.warnings)
        self._paint(RED if warn else GREEN,
                    f"{len(d.pots)} pots, {d.groups} planting group(s) "
                    f"{d.stagger_days} days apart, max {d.max_load} N2O / day"
                    + ("  (over capacity)" if warn else ""))
        self.summary.setPlainText(summary_text(d))
        self.save_b.setEnabled(True)

    def _save_workbook(self):
        if self.design is None:
            return
        default = os.path.join(self._dir(), f"{self.design.plan.name}_design_map.xlsx")
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save workbook", default,
                                                        "Excel workbook (*.xlsx)")
        if not path:
            return
        try:
            dx.write(self.design, path)
        except PermissionError:
            QtWidgets.QMessageBox.warning(self, "Could not save",
                                          f"{path} is open in another program. Close it and "
                                          f"save again.")
            return
        self._paint(GREEN, f"Saved {os.path.basename(path)}")

    def _dir(self):
        return os.path.dirname(self.plan_path) if self.plan_path else os.path.expanduser("~")

    def _save_plan(self):
        try:
            plan = self.read_plan()
        except dc.PlanError as ex:
            QtWidgets.QMessageBox.warning(self, "Plan not saved", str(ex)); return
        default = self.plan_path or os.path.join(self._dir(), f"{plan.name}_plan.json")
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save plan", default,
                                                        "Plan (*.json)")
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            json.dump(plan.to_dict(), f, indent=2)
        self.plan_path = path
        self._paint(SLATE, f"Plan saved: {os.path.basename(path)}")

    def _open_plan(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Open plan", self._dir(),
                                                        "Plan (*.json)")
        if not path:
            return
        try:
            with open(path, encoding="utf-8") as f:
                plan = dc.Plan.from_dict(json.load(f))
        except Exception as ex:
            QtWidgets.QMessageBox.warning(self, "Could not open the plan",
                                          f"{path}\n{type(ex).__name__}: {ex}")
            return
        self.plan_path = path
        self.load_plan(plan)


def summary_text(d: dc.Design) -> str:
    plan = d.plan
    c = d.counts()
    out = [f"{len(d.pots)} pots: {c[dc.MEASURED]} measured, {c[dc.SACRIFICIAL]} sacrificial, "
           f"{c[dc.SPARE]} spare",
           f"bench: {plan.tables} tables x {plan.zones_per_table} zones x "
           f"{plan.slots_per_zone} slots = {plan.capacity} places",
           "", f"PLANTING GROUPS: {d.groups}, {d.stagger_days} days apart"]
    per_g = Counter(p.group for p in d.pots)
    meas = Counter((p.group, p.cohort) for p in d.pots if p.role == dc.MEASURED)
    for g in range(1, d.groups + 1):
        coh = "  ".join(f"{k.name}:{meas[(g, k.name)]}" for k in plan.cohorts)
        out.append(f"  group {g}  plant {d.group_planting(g):%a %d %b %Y}  "
                   f"{per_g[g]:>4} pots   measured {coh}")
    out += ["", "SCHEDULE"]
    out.append(f"  {'':6} {'treatment':>10} {'flood':>10} " + " ".join(
        f"{h} h N2O".rjust(10) for h in plan.measure_hours))
    for g in range(1, d.groups + 1):
        for k in plan.cohorts:
            ev = dict((t.split(",")[0].replace("Sacrificial harvest", "flood").strip(), day)
                      for day, t in d.events(g, k))
            treat = ev.get("Treatment")
            flood = ev.get("flood") or ev.get("Flood")
            n2o = [ev.get(f"{h} h N2O") for h in plan.measure_hours]
            fmt = lambda x, w: (f"{x:%a %d %b}" if x else "-").rjust(w)
            out.append(f"  G{g} {k.name:<3} {fmt(treat, 10)} {fmt(flood, 10)} "
                       + " ".join(fmt(x, 10) for x in n2o))
    out += ["", f"N2O PER DAY (capacity {plan.max_measured_per_day})"]
    for day, n in d.calendar_load().items():
        flag = "  OVER" if n > plan.max_measured_per_day else ""
        out.append(f"  {day:%a %d %b}  {n:>4}  {'#' * min(60, n)}{flag}")
    out += ["", "POTS PER TABLE"]
    per_t = Counter(p.table for p in d.pots)
    out.append("  " + "  ".join(f"T{t}:{per_t[t]}" for t in range(1, plan.tables + 1)))
    for w in d.warnings:
        out += ["", f"WARNING: {w}"]
    return "\n".join(out)


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    w = DesignWindow()
    w.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
