"""Pot labels for a Zebra printer (ZD621, ZPL), written beside the design map.

The labels are the dosing station's (bni-dosing provision_jars.zpl_label, as Start New
Trial writes them): the pot's tag and position at the top left, its target mass at the
top right, a Code 128 barcode, and the barcode in text beneath. 300 dpi, 3 x 1 inch
labels, one per pot that goes to the station (the `jars` tab). The target mass is the
station's WFPS chain (bni-dosing wfps.compute) on the plan's dosing-station inputs; when
one of them is still blank the labels are written without it.

Print the file with Zebra Setup Utilities: Open Communication With Printer > Open >
Send to Printer."""
from __future__ import annotations

import os

import design_core as dc

DPI = 300                  # the station writes 300 dpi; a 203 dpi ZD621 needs 203 here
LABEL_IN = (3.0, 1.0)      # width x height, inches
TAG_CHARS = 24             # the station cuts the tag at 24 characters


def target_mass(plan: dc.Plan) -> float | None:
    """The pot's target mass at the WFPS target, g: container + air-dry soil + water to
    add, exactly as the station computes it. None while an input is blank."""
    need = (plan.air_dry_soil_g, plan.air_dry_moisture, plan.bulk_density,
            plan.wfps_target, plan.container_weight_g)
    if any(v is None for v in need):
        return None
    m_dry = plan.air_dry_soil_g / (1.0 + plan.air_dry_moisture)
    volume = m_dry / plan.bulk_density
    porosity = 1.0 - plan.bulk_density / (plan.particle_density or 2.65)
    water_total = plan.wfps_target * porosity * volume
    water_present = plan.air_dry_soil_g - m_dry
    water_to_add = water_total - water_present
    return (plan.container_weight_g + plan.air_dry_soil_g) + water_to_add


def zpl_label(barcode: str, tag: str, target: float | None, dpi: int = DPI,
              w_in: float = LABEL_IN[0], h_in: float = LABEL_IN[1]) -> str:
    """One label, proportional to the media (the station's layout, line for line)."""
    pw = int(round(w_in * dpi))
    ll = int(round(h_in * dpi))
    m = max(8, int(round(0.05 * dpi)))
    f = max(16, int(round(0.08 * ll)))
    bar_h = int(round(0.46 * ll))
    bar_mod = max(2, int(round(dpi / 100)))
    tag = (tag or "")[:TAG_CHARS]
    tgt = f"{target:.1f} g" if target is not None else ""
    y_top = m
    y_bar = y_top + f + int(round(0.04 * ll))
    y_txt = y_bar + bar_h + int(round(0.03 * ll))
    return (
        f"^XA^CI28"
        f"^PW{pw}^LL{ll}"
        f"^FO{m},{y_top}^A0N,{f},{f}^FD{tag}^FS"
        f"^FO{m},{y_top}^FB{pw - 2 * m},1,0,R^A0N,{f},{f}^FD{tgt}^FS"
        f"^FO0,{y_bar}^FB{pw},1,0,C^BY{bar_mod},2.5,{bar_h}^BCN,{bar_h},N,N,N^FD{barcode}^FS"
        f"^FO0,{y_txt}^FB{pw},1,0,C^A0N,{f},{f}^FD{barcode}^FS^XZ"
    )


def pots(d: dc.Design) -> list[dc.Pot]:
    """The pots that get a label: those on the `jars` tab, in its order."""
    return [p for p in d.pots if p.active]


def zpl(d: dc.Design, dpi: int = DPI) -> str:
    t = target_mass(d.plan)
    return "".join(zpl_label(p.barcode, f"{p.label} {p.position}", t, dpi) for p in pots(d))


def path_for(workbook: str) -> str:
    """labels_<workbook name>.zpl beside the workbook (the station's naming)."""
    stem = os.path.splitext(os.path.basename(workbook))[0]
    return os.path.join(os.path.dirname(os.path.abspath(workbook)), f"labels_{stem}.zpl")


def write(d: dc.Design, path: str, dpi: int = DPI) -> tuple[int, float | None]:
    """Write the labels; returns (labels, target mass or None)."""
    with open(path, "w", encoding="utf-8") as f:
        f.write(zpl(d, dpi))
    return len(pots(d)), target_mass(d.plan)
