"""Labels for a Zebra printer (ZD621, ZPL), each file in the order its labels are used.

Save workbook (design_app.py) writes only the pot labels, to labels_<workbook name>/:

    labels_PLANTING.zpl        every pot (spares too): planting group, then position

The swap tool (swaps.py) writes the rest, a planting group at a time, once the group's
swaps are done and its spares used or discarded (its harvested pots are then final), to
labels_<workbook name>_swaps_<date>/:

    labels_RELABEL.zpl         a new pot label for each spare promoted that session
    labels_HARVEST_SHOOT.zpl   shoot bags  } every measured and sacrificial pot of the
    labels_HARVEST_ROOT.zpl    root bags   } group(s): harvest date, then planting group
    labels_KCL_COLLECT.zpl     } KCl       } and cohort, then flood order (a sacrificial
    labels_KCL_WEIGHED.zpl     } tubes     } pot: its harvest order on flood day)
    labels_KCL_TIN.zpl         } and tin   }
    labels_KCL_EXTRACT.zpl     }           }

All are 3 x 1 inch, in one layout: a QR code at the left with its text beneath it, a
black badge saying what the label is for, and a small form.

Pot label (planting): QR = the pot's barcode (what the stations scan). Badge: the cohort
letter and the pot's flood order in its group and cohort (#11); for a sacrificial pot
SAC and its harvest order on flood day (SAC 2, its row on the flood sheet's sacrificial
list); SPARE for a spare. Form: position, then planting group, genotype and treatment.

    +--------------------------------------------------------------+
    | [QR ]  |#######|  POSITION                                    |
    | [QR ]  |#  B  #|  T1-Z1-S1                                    |
    |        |# #11 #|  PG | GENOTYPE      | TREATMENT              |
    | S014   |#######|  4  | B73           | INH                    |
    +--------------------------------------------------------------+

Bag and tube labels: QR = the pot's barcode + -SH (shoot), -RT (root), -K1 to -K4 (KCl
collect, weighed, tin, extract), so every label has its own scan. Badge: SHOOT / ROOT +
HARVEST + the harvest date; KCl + the step, with the harvest date on COLLECT only (the
later steps are not done on harvest day). Form: the pot, then cohort and flood order,
genotype and treatment.

    +--------------------------------------------------------------+
    | [QR ]  |#SHOOT#|  POT                                         |
    | [QR ]  |HARVEST|  BNI-P1-S014                                 |
    |        |22 DEC#|  COHORT | GENOTYPE | TREATMENT               |
    | S014-SH|#######|  B #11  | B73      | INH                     |
    +--------------------------------------------------------------+

A value too wide for its cell is printed smaller. Print a file with Zebra Setup
Utilities: Open Communication With Printer > Open > Send to Printer."""
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


# What prints well on the lab's ZD621, as tuned at the printer: everything 20 dots down
# (^LH, the label home), the QR code at the top left at 7 dots a module, its text under it.
HOME_Y = 20
QR_AT = (41, 0)
QR_MODULE = 7
QR_TEXT_Y = 246


def _start(X, Y) -> list[str]:
    return [f"^XA^CI28^PW{X(900)}^LL{Y(300)}", f"^LH0,{Y(HOME_Y)}"]


def _qr(X, Y, S, data: str) -> str:
    return f"^FO{X(QR_AT[0])},{Y(QR_AT[1])}^BQN,2,{max(1, S(QR_MODULE))}^FD{QR_ECC}A,{data}^FS"


def _end(out: list[str]) -> str:
    """One command per line, as the label was tuned at the printer."""
    return "\n".join(out + ["^XZ"]) + "\n"


def zpl_label(p: dc.Pot, dpi: int = DPI, w_in: float = LABEL_IN[0],
              h_in: float = LABEL_IN[1]) -> str:
    """One pot's label. Laid out on a 900 x 300 grid (3 x 1 in at 300 dpi) and scaled to
    the printer's dpi and label size."""
    sx, sy = w_in * dpi / 900, h_in * dpi / 300
    X = lambda v: int(round(v * sx))
    Y = lambda v: int(round(v * sy))
    S = lambda v: max(1, int(round(v * min(sx, sy))))
    out = _start(X, Y)

    # QR code, its text under it (26 dots for BNI-P1-S107; smaller for a longer barcode)
    out.append(_qr(X, Y, S, p.barcode))
    out.append(_text(X(20), Y(QR_TEXT_Y), S(max(12, min(26, int(188 / (0.65 * len(p.barcode)))))),
                     p.barcode, w=X(188), centre=True))

    # the badge: cohort letter and flood order, white on black
    bx, by, bw, bh = 230, 18, 190, 264
    out.append(_box(X(bx), Y(by), X(bw), Y(bh), X(bw)))
    if p.role == dc.SPARE or not p.cohort:
        out.append(_text(X(bx), Y(by + 104), S(52), "SPARE", w=X(bw), centre=True,
                         reverse=True))
    else:
        out.append(_text(X(bx), Y(by + 8), S(_fit(p.cohort, 150, bw - 20)), p.cohort,
                         w=X(bw), centre=True, reverse=True))
        order = f"SAC {p.seq}" if p.role == dc.SACRIFICIAL else f"#{p.seq}"
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
    return _end(out)


# --- bag and tube labels -----------------------------------------------------------
# (file name, QR suffix, badge lines; "{date}" is the harvest date)
SAMPLES = (("HARVEST_SHOOT", "SH", (("SHOOT", 76), ("HARVEST", 46), ("{date}", 56))),
           ("HARVEST_ROOT", "RT", (("ROOT", 76), ("HARVEST", 46), ("{date}", 56))),
           ("KCL_COLLECT", "K1", (("KCl", 64), ("COLLECT", 46), ("{date}", 46))),
           ("KCL_WEIGHED", "K2", (("KCl", 84), ("WEIGHED", 50))),
           ("KCL_TIN", "K3", (("KCl", 84), ("TIN", 50))),
           ("KCL_EXTRACT", "K4", (("KCl", 84), ("EXTRACT", 50))))


def sample_label(p: dc.Pot, suffix: str, badge: tuple, harvest, dpi: int = DPI,
                 w_in: float = LABEL_IN[0], h_in: float = LABEL_IN[1]) -> str:
    """One bag or tube label of pot `p` (see the module docstring)."""
    sx, sy = w_in * dpi / 900, h_in * dpi / 300
    X = lambda v: int(round(v * sx))
    Y = lambda v: int(round(v * sy))
    S = lambda v: max(1, int(round(v * min(sx, sy))))
    code = f"{p.barcode}-{suffix}"
    out = _start(X, Y)
    out.append(_qr(X, Y, S, code))
    # the printer's hyphens are wide: allow 0.7 of the height a character, which Labelary
    # confirms fits (BNI-P1-S050-K4 at 22 dots on 220)
    out.append(_text(X(6), Y(QR_TEXT_Y), S(max(12, min(26, int(220 / (0.7 * len(code)))))),
                     code, w=X(220), centre=True))

    bx, by, bw, bh = 230, 18, 190, 264
    out.append(_box(X(bx), Y(by), X(bw), Y(bh), X(bw)))
    date = f"{harvest.day} {harvest:%b}".upper()
    lines = [(t.format(date=date), size) for t, size in badge]
    lines = [(t, _fit(t, size, bw - 20)) for t, size in lines]
    gap = (bh - sum(size for _, size in lines)) / (len(lines) + 1)
    y = by + gap
    for t, size in lines:
        out.append(_text(X(bx), Y(int(y)), S(size), t, w=X(bw), centre=True, reverse=True))
        y += size + gap

    x, r = bx + bw + 22, 886
    out.append(_text(X(x), Y(12), S(26), "POT"))
    out.append(_text(X(x), Y(40), S(_fit(p.barcode, 84, r - x)), p.barcode))
    out.append(_box(X(x), Y(140), X(r - x), Y(3)))
    order = f"SAC {p.seq}" if p.role == dc.SACRIFICIAL else f"#{p.seq}"
    cells = [("COHORT", f"{p.cohort} {order}", 190), ("GENOTYPE", p.genotype, 120),
             ("TREATMENT", p.treatment or "-", r - x - 310)]
    size = min(_fit(v, 66, w - 22) for _, v, w in cells)        # one size for the row
    cx = x
    for i, (head, value, w) in enumerate(cells):
        out.append(_text(X(cx), Y(152), S(_fit(head, 24, w - 18)), head))
        out.append(_text(X(cx), Y(186), S(size), value))
        cx += w
        if i < len(cells) - 1:
            out.append(_box(X(cx - 12), Y(150), X(3), Y(132)))
    return _end(out)


# --- which pots, in which order ---------------------------------------------------------
def pots(d: dc.Design) -> list[dc.Pot]:
    """The pots that get a pot label (those on the `jars` tab), in planting order:
    planting group, then table, zone and slot."""
    return sorted((p for p in d.pots if p.active),
                  key=lambda p: (p.group, p.table, p.zone, p.slot))


def harvest_date(d: dc.Design, p: dc.Pot):
    """A sacrificial pot is harvested on its flood day, a measured pot after its cohort's
    last N2O sample."""
    c = d.plan.cohort(p.cohort)
    if p.role == dc.SACRIFICIAL:
        return d.anchor_date(("flood", p.group, c.name))
    return d.anchor_date(("n2o", p.group, c.name, c.last_offset))


def harvested(d: dc.Design, groups=None) -> list[dc.Pot]:
    """The pots that are harvested (measured and sacrificial; of `groups` only, if
    given), in harvest order: date, then planting group and cohort, then flood order (a
    sacrificial pot's: its harvest order on flood day)."""
    names = [c.name for c in d.plan.cohorts]
    return sorted((p for p in d.pots if p.active and p.role in (dc.MEASURED, dc.SACRIFICIAL)
                   and (groups is None or p.group in groups)),
                  key=lambda p: (harvest_date(d, p), p.group, names.index(p.cohort),
                                 p.role != dc.SACRIFICIAL, p.seq))


def zpl(d: dc.Design, dpi: int = DPI) -> str:
    """The planting file's labels."""
    return "".join(zpl_label(p, dpi) for p in pots(d))


RELABEL = "labels_RELABEL.zpl"


def files(d: dc.Design, dpi: int = DPI, planting: bool = True, groups=(),
          relabel=()) -> dict[str, tuple[str, int]]:
    """{file name: (ZPL, labels)}: the planting file (`planting`), a new pot label for
    each promoted spare in `relabel` (labels_RELABEL.zpl, by position), and the shoot,
    root and KCl files of the planting groups in `groups` (only the swap tool asks for
    these: a group's harvested pots are final once its swaps are done)."""
    out = {}
    if planting:
        out["labels_PLANTING.zpl"] = (zpl(d, dpi), len(pots(d)))
    new = sorted(relabel, key=lambda p: (p.group, p.table, p.zone, p.slot))
    if new:
        out[RELABEL] = ("".join(zpl_label(p, dpi) for p in new), len(new))
    hv = harvested(d, set(groups)) if groups else []
    if not hv:
        return out
    for name, suffix, badge in SAMPLES:
        out[f"labels_{name}.zpl"] = ("".join(sample_label(p, suffix, badge, harvest_date(d, p), dpi)
                                             for p in hv), len(hv))
    return out


def folder_for(workbook: str) -> str:
    """labels_<workbook name>/ beside the workbook."""
    stem = os.path.splitext(os.path.basename(workbook))[0]
    return os.path.join(os.path.dirname(os.path.abspath(workbook)), f"labels_{stem}")


def session_folder_for(workbook: str, day) -> str:
    """labels_<workbook name>_swaps_<date>/ beside the workbook (_2, _3 ... when a second
    session that day would overwrite one)."""
    base = f"{folder_for(workbook)}_swaps_{day:%Y-%m-%d}"
    folder, k = base, 1
    while os.path.exists(folder):
        k += 1
        folder = f"{base}_{k}"
    return folder


def write(d: dc.Design, folder: str, dpi: int = DPI, planting: bool = True, groups=(),
          relabel=()) -> dict[str, int]:
    """Write the label files (see `files`) into `folder` (made if missing);
    {file name: labels}."""
    os.makedirs(folder, exist_ok=True)
    out = {}
    for name, (text, n) in files(d, dpi, planting, groups, relabel).items():
        with open(os.path.join(folder, name), "w", encoding="utf-8") as f:
            f.write(text)
        out[name] = n
    return out
