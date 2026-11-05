#!/usr/bin/env python3
"""
swap_app.py: record failed pots and the spares that replace them (v0.1, 2026-11-05).

    python swap_app.py [design_map.xlsx]

Open the experiment's design map, scan (or type) the failed pot, pick the spare, Add swap;
repeat; optionally discard the spares that are left; Save. The design map is rebuilt with
the spares in place (swaps.py): every tab follows, what was typed into the workbook is kept,
and the old file is kept as a dated backup. No barcode changes: write "promoted to <tag>"
on each spare by hand.

On the greenhouse PC, afterwards: in the dosing station pick the trial again in the Trial
dropdown (it reloads the workbook) and Load experiment orders again.
"""
from __future__ import annotations

import datetime as dt
import os
import sys

from PySide6 import QtCore, QtGui, QtWidgets
from PySide6.QtCore import Qt

import design_core as dc
import swaps as S

TITLE = "BNI Pot Swaps"
SLATE, BLUE, GREEN, RED, MUTED = "#374151", "#1d4ed8", "#15803d", "#b91c1c", "#6b7280"


class SwapWindow(QtWidgets.QMainWindow):
    def __init__(self, path: str | None = None):
        super().__init__()
        self.setWindowTitle(TITLE)
        self.resize(980, 720)
        self.path: str | None = None
        self.design: dc.Design | None = None
        self.saved: list[dc.Swap] = []          # swaps already in the workbook
        self.pending: list[dc.Swap] = []        # added here, not saved yet
        self.discard = False

        root = QtWidgets.QWidget(); self.setCentralWidget(root)
        v = QtWidgets.QVBoxLayout(root)
        self.banner = QtWidgets.QLabel()
        self.banner.setAlignment(Qt.AlignCenter); self.banner.setFixedHeight(48)
        f = QtGui.QFont(); f.setPointSize(13); f.setBold(True); self.banner.setFont(f)
        v.addWidget(self.banner)

        top = QtWidgets.QHBoxLayout(); v.addLayout(top)
        open_b = QtWidgets.QPushButton("Open design map..."); open_b.clicked.connect(self._open)
        self.file_lbl = QtWidgets.QLabel("no design map open")
        self.file_lbl.setStyleSheet(f"color:{MUTED};")
        top.addWidget(open_b); top.addWidget(self.file_lbl, 1)

        add = QtWidgets.QGroupBox("Add a swap"); fl = QtWidgets.QFormLayout(add)
        self.failed = QtWidgets.QLineEdit()
        self.failed.setPlaceholderText("scan or type the failed pot's barcode")
        self.failed.returnPressed.connect(lambda: self._failed_entered(focus=True))
        self.failed.textChanged.connect(lambda _: self._failed_entered())
        fl.addRow("Failed pot", self.failed)
        self.failed_info = QtWidgets.QLabel(); self.failed_info.setStyleSheet(f"color:{MUTED};")
        fl.addRow("", self.failed_info)
        self.spare = QtWidgets.QComboBox(); self.spare.setEditable(True)
        self.spare.setToolTip("Only spares of the same genotype, planted the same day, are "
                              "listed; any other is refused. Scan one, or pick.")
        self.spare.lineEdit().returnPressed.connect(self._add)
        fl.addRow("Spare that replaces it", self.spare)
        self.day = QtWidgets.QDateEdit(QtCore.QDate.currentDate()); self.day.setCalendarPopup(True)
        self.day.setDisplayFormat("ddd d MMM yyyy")
        fl.addRow("Date", self.day)
        self.note = QtWidgets.QLineEdit(); self.note.setPlaceholderText("optional, e.g. no germination")
        fl.addRow("Note", self.note)
        self.add_b = QtWidgets.QPushButton("Add swap"); self.add_b.clicked.connect(self._add)
        fl.addRow("", self.add_b)
        v.addWidget(add)

        lg = QtWidgets.QGroupBox("Swaps"); ll = QtWidgets.QVBoxLayout(lg)
        self.table = QtWidgets.QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(
            ("date", "failed pot", "was", "spare", "spare takes over", "note"))
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        ll.addWidget(self.table, 1)
        row = QtWidgets.QHBoxLayout()
        self.undo_b = QtWidgets.QPushButton("Remove selected (not saved yet)")
        self.undo_b.clicked.connect(self._remove)
        self.discard_cb = QtWidgets.QCheckBox("Discard every spare left unused (no longer grown)")
        self.discard_cb.toggled.connect(self._toggle_discard)
        row.addWidget(self.undo_b); row.addStretch(1); row.addWidget(self.discard_cb)
        ll.addLayout(row)
        self.spares_lbl = QtWidgets.QLabel(); self.spares_lbl.setStyleSheet(f"color:{MUTED};")
        self.spares_lbl.setWordWrap(True)
        ll.addWidget(self.spares_lbl)
        v.addWidget(lg, 1)

        self.save_b = QtWidgets.QPushButton("Save design map")
        self.save_b.setMinimumHeight(38)
        self.save_b.setStyleSheet(
            f"QPushButton{{background:{GREEN};color:white;font-weight:bold;}}"
            "QPushButton:disabled{background:#d1d5db;color:#6b7280;}")
        self.save_b.clicked.connect(self._save)
        v.addWidget(self.save_b)

        self._paint(SLATE, "Open the experiment's design map")
        self._refresh()
        if path:
            self.load(path)

    # ---- state ----
    def load(self, path: str):
        try:
            d, saved = S.read_design_map(path)
        except Exception as ex:
            QtWidgets.QMessageBox.critical(self, "Cannot read", f"{path}\n\n{ex}")
            return
        self.path, self.design, self.saved, self.pending = path, d, saved, []
        self.discard = False
        self.discard_cb.blockSignals(True); self.discard_cb.setChecked(False)
        self.discard_cb.blockSignals(False)
        self.file_lbl.setText(path)
        self._paint(SLATE, f"{d.plan.name}: {len(saved)} swap(s) so far")
        self._refresh()
        self.failed.setFocus()

    def _replay(self):
        """The design as saved plus the swaps added here (so one can be removed)."""
        d, _ = S.read_design_map(self.path)
        for sw in self.pending:
            S.swap(d, sw.failed, sw.spare, sw.date, sw.note)
        if self.discard:
            S.discard_spares(d)
        self.design = d

    def _paint(self, color, text):
        self.banner.setStyleSheet(f"color:white;background:{color};border-radius:8px;")
        self.banner.setText(text)

    def _refresh(self):
        have = self.design is not None
        for w in (self.failed, self.spare, self.day, self.note, self.add_b, self.discard_cb):
            w.setEnabled(have)
        self.save_b.setEnabled(have and bool(self.pending or self.discard))
        self.undo_b.setEnabled(bool(self.pending))
        rows = [(s, True) for s in self.saved] + [(s, False) for s in self.pending]
        self.table.setRowCount(len(rows))
        for r, (s, done) in enumerate(rows):
            was = ""
            if have:
                f = next((p for p in self.design.pots if p.barcode == s.failed), None)
                was = f.label.replace("-FAILED", "") if f else ""
            vals = (f"{s.date:%d %b %Y}", s.failed, was, s.spare, s.role,
                    (s.note or "") + ("" if done else "   (not saved yet)"))
            for c, val in enumerate(vals):
                it = QtWidgets.QTableWidgetItem(val)
                if done:
                    it.setForeground(QtGui.QColor(MUTED))
                self.table.setItem(r, c, it)
        self.table.resizeColumnsToContents()
        if have:
            left = {}
            for p in self.design.pots:
                if p.role == dc.SPARE or (self.discard and p.role == dc.DISCARDED
                                          and not p.swap):
                    k = (p.genotype, p.group)
                    left[k] = left.get(k, 0) + 1
            many = self.design.groups > 1
            self.spares_lbl.setText("Spares left: " + (", ".join(
                f"{g}{f' group {gr}' if many else ''}: {n}"
                for (g, gr), n in sorted(left.items())) or "none")
                + ("   (will be discarded on Save)" if self.discard else ""))
        else:
            self.spares_lbl.setText("")

    # ---- actions ----
    def _open(self):
        start = os.path.dirname(self.path) if self.path else os.getcwd()
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Open design map", start,
                                                        "Excel (*.xlsx)")
        if path:
            self.load(path)

    def _failed_entered(self, focus=False):
        self.spare.clear(); self.failed_info.setText("")
        if self.design is None or not self.failed.text().strip():
            return
        try:
            f = S.pot(self.design, self.failed.text())
        except S.SwapError as ex:
            self.failed_info.setText(str(ex)); return
        if f.role not in (dc.MEASURED, dc.SACRIFICIAL):
            self.failed_info.setText(f"{f.barcode}: {f.label} ({f.role.lower()}): cannot be "
                                     "replaced")
            return
        sp = S.eligible_spares(self.design, f)
        day = S.planted(self.design, f)
        self.failed_info.setText(
            f"{f.label} ({f.genotype}) at {f.position}, planted {day:%a %d %b}: "
            + (f"{len(sp)} spare(s) of {f.genotype} planted that day can replace it" if sp
               else f"no unused {f.genotype} spare was planted {day:%a %d %b}"))
        for p in sp:
            self.spare.addItem(f"{p.barcode}   {p.label} at {p.position}, planted "
                               f"{S.planted(self.design, p):%a %d %b}", p.barcode)
        self.spare.setCurrentIndex(-1)
        if focus:                       # Return (or the scanner's) moves on to the spare
            self.spare.setFocus()

    def _add(self):
        if self.design is None:
            return
        text, i = self.spare.currentText().strip(), self.spare.currentIndex()
        if i >= 0 and self.spare.itemText(i) == self.spare.currentText():
            spare = self.spare.itemData(i)                  # picked from the list
        else:
            spare = text.split()[0] if text else ""         # scanned or typed
        if not spare:
            self._paint(RED, "Pick or scan the spare"); return
        try:
            sw = S.swap(self.design, self.failed.text(), spare,
                        self.day.date().toPython(), self.note.text().strip())
        except S.SwapError as ex:
            self._paint(RED, str(ex)); return
        self.pending.append(sw)
        self._paint(BLUE, f"{sw.spare} takes over {sw.role} from {sw.failed}: Save when done")
        self.failed.clear(); self.note.clear(); self.spare.clear(); self.failed_info.clear()
        self._refresh()
        self.failed.setFocus()

    def _remove(self):
        rows = sorted({i.row() for i in self.table.selectedIndexes()})
        idx = [r - len(self.saved) for r in rows if r >= len(self.saved)]
        if not idx:
            return
        self.pending = [s for k, s in enumerate(self.pending) if k not in idx]
        self._replay(); self._refresh()

    def _toggle_discard(self, on):
        self.discard = on
        self._replay(); self._refresh()

    def _save(self):
        lock = os.path.join(os.path.dirname(self.path), "~$" + os.path.basename(self.path))
        if os.path.exists(lock):
            QtWidgets.QMessageBox.warning(self, "Close it in Excel first",
                                          f"{os.path.basename(self.path)} is open in Excel. "
                                          "Close it there, then Save again.")
            return
        try:
            rep = S.rebuild(self.path, self.design, self.saved + self.pending)
        except Exception as ex:
            QtWidgets.QMessageBox.critical(self, "Save failed", str(ex)); return
        n = len(self.pending)
        QtWidgets.QMessageBox.information(self, "Saved", (
            f"{n} swap(s) saved" + (", unused spares discarded" if self.discard else "")
            + f".\n\nBackup of the old file:\n{rep['backup']}\n\nKept from the old file: "
            f"{rep['cells']} typed cell(s), {rep['inputs']} edited input(s)"
            + (f", tabs {', '.join(rep['sheets'])}" if rep["sheets"] else "") + ".\n\n"
            "On the greenhouse PC: write \"promoted to <tag>\" on each spare. In the dosing "
            "station pick the trial again in the Trial dropdown, then Load experiment "
            "orders again."))
        self.load(self.path)
        self._paint(GREEN, f"Saved: {len(self.saved)} swap(s) in {os.path.basename(self.path)}")


def main() -> int:
    app = QtWidgets.QApplication(sys.argv)
    w = SwapWindow(sys.argv[1] if len(sys.argv) > 1 else None)
    w.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
