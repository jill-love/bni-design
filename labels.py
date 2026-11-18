"""Pot labels for a Zebra printer (ZD621, ZPL), written beside the design map.

One 3 x 1 inch label per pot that goes to the station (the `jars` tab):

    +--------------------------------------------------------------+
    | [QR ]  |#######|  POSITION                                    |
    | [QR ]  |#  B  #|  T1-Z1-S1                                    |
    | [QR ]  |#     #|  ---------------------------------------    |
    |        |# #11 #|  PG | GENOTYPE      | TREATMENT              |
    | S014   |#######|  4  | B73           | INH                    |
    +--------------------------------------------------------------+

- a QR code of the pot's barcode, the barcode in small text beneath it (the scanners
  read the QR code as the barcode, so the stations need no change);
- a black badge: the cohort letter and the pot's flood order in its group and cohort
  (#11); SAC for a sacrificial pot (harvested on flood day, never flooded); SPARE for a
  spare (no cohort or treatment until it replaces a pot);
- the bench position, then planting group, genotype and treatment, each under a small
  heading. A value too wide for its cell is printed smaller.

Print the file with Zebra Setup Utilities: Open Communication With Printer > Open >
Send to Printer."""
from __future__ import annotations

import os

import design_core as dc

DPI = 300                  # a 203 dpi ZD621 needs 203 here
LABEL_IN = (3.0, 1.0)      # width x height, inches
QR_ECC = "Q"               # error correction: Q restores 25 % of a scuffed code
CHAR_W = 0.55              # the printer font's widest characters, as a fraction of height


def _fit(text: str, size: int, width: int) -> int:
    """The largest font height (dots) up to `size` at which `text` fits `width` dots."""
    return max(12, min(size, int(width / (CHAR_W * max(1, len(text))))))


def _text(x, y, size, text, w=None, centre=False, reverse=False) -> str:
    """A text field; with `w`, a one-line block `w` dots wide. A centred block's line ends
    with \\& (a line break): Zebra printers only justify a line that ends with one."""
    block = f"^FB{w},1,0,{'C' if centre else 'L'}" if w else ""
    end = "\\&" if w and centre else ""
    return (f"^FO{x},{y}{block}^A0N,{size},{size}{'^FR' if reverse else ''}"
            f"^FD{text}{end}^FS")


def _box(x, y, w, h, t=None) -> str:
    """A rectangle; t = w (or h) fills it."""
    return f"^FO{x},{y}^GB{w},{h},{t if t is not None else min(w, h)}^FS"


def zpl_label(p: dc.Pot, dpi: int = DPI, w_in: float = LABEL_IN[0],
              h_in: float = LABEL_IN[1]) -> str:
    """One pot's label. Laid out on a 900 x 300 grid (3 x 1 in at 300 dpi) and scaled to
    the printer's dpi and label size."""
    sx, sy = w_in * dpi / 900, h_in * dpi / 300
    X = lambda v: int(round(v * sx))
    Y = lambda v: int(round(v * sy))
    S = lambda v: max(1, int(round(v * min(sx, sy))))
    out = [f"^XA^CI28^PW{X(900)}^LL{Y(300)}"]

    # QR code (8 dots a module at 300 dpi; ^BQ draws it a few dots below its ^FO)
    mag = max(1, S(8))
    out.append(f"^FO{X(30)},{Y(24)}^BQN,2,{mag}^FD{QR_ECC}A,{p.barcode}^FS")
    out.append(_text(X(20), Y(226), S(_fit(p.barcode, 30, 188)), p.barcode, w=X(188),
                     centre=True))

    # the badge: cohort letter and flood order, white on black
    bx, by, bw, bh = 230, 18, 190, 264
    out.append(_box(X(bx), Y(by), X(bw), Y(bh), X(bw)))
    if p.role == dc.SPARE or not p.cohort:
        out.append(_text(X(bx), Y(by + 104), S(52), "SPARE", w=X(bw), centre=True,
                         reverse=True))
    else:
        out.append(_text(X(bx), Y(by + 8), S(_fit(p.cohort, 150, bw - 20)), p.cohort,
                         w=X(bw), centre=True, reverse=True))
        order = "SAC" if p.role == dc.SACRIFICIAL else f"#{p.seq}"
        out.append(_text(X(bx), Y(by + 166), S(_fit(order, 76, bw - 20)), order, w=X(bw),
                         centre=True, reverse=True))

    # the form: position, then planting group | genotype | treatment
    x, r = bx + bw + 22, 886
    out.append(_text(X(x), Y(14), S(26), "POSITION"))
    out.append(_text(X(x), Y(44), S(_fit(p.position, 86, r - x)), p.position))
    out.append(_box(X(x), Y(140), X(r - x), Y(3)))
    cells = [("PG", str(p.group), 78), ("GENOTYPE", p.genotype, 212)]
    if p.treatment and p.role != dc.SPARE:
        cells.append(("TREATMENT", p.treatment, 0))
    cells[-1] = cells[-1][:2] + (r - x - sum(w for _, _, w in cells[:-1]),)   # rest of row
    cx = x
    for i, (head, value, w) in enumerate(cells):
        out.append(_text(X(cx), Y(152), S(_fit(head, 24, w - 18)), head))
        out.append(_text(X(cx), Y(184), S(_fit(value, 72, w - 22)), value))
        cx += w
        if i < len(cells) - 1:
            out.append(_box(X(cx - 12), Y(150), X(3), Y(132)))
    return "".join(out) + "^XZ"


def pots(d: dc.Design) -> list[dc.Pot]:
    """The pots that get a label: those on the `jars` tab, in its order."""
    return [p for p in d.pots if p.active]


def zpl(d: dc.Design, dpi: int = DPI) -> str:
    return "".join(zpl_label(p, dpi) for p in pots(d))


def path_for(workbook: str) -> str:
    """labels_<workbook name>.zpl beside the workbook."""
    stem = os.path.splitext(os.path.basename(workbook))[0]
    return os.path.join(os.path.dirname(os.path.abspath(workbook)), f"labels_{stem}.zpl")


def write(d: dc.Design, path: str, dpi: int = DPI) -> int:
    """Write the labels; returns how many."""
    with open(path, "w", encoding="utf-8") as f:
        f.write(zpl(d, dpi))
    return len(pots(d))
