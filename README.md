# BNI experiment setup

Plans a BNI pot experiment and writes its design map workbook: where every pot goes on the bench, when each group is planted, and what happens each day.

## What it does

1. **Plan**: in one window, enter the genotypes, cohorts (time-trial groups), treatments and reps, plus the bench and the N2O capacity.
2. **Planting groups**: N2O is measured one pot at a time, at most *max N2O measurements / day* (50 by default). Planting is staggered so that no day needs more than that. The pots are split into the **fewest equally sized groups** that fit, planted at least 2 days apart (more if cohorts would otherwise land on the same day). The split is random but balanced: every genotype x cohort x treatment cell is spread over the groups as evenly as it divides.
3. **Bench layout**: pots are spread over the **tables you say are available**, then the zones of each table. Each cell and each planting group is spread as evenly as it divides, and slots are random. A table holds up to 8 zones x 25 slots. Positions are written `T1-Z2-05`.
4. **Timeline**: per planting group and cohort, all counted in days after planting (DAP):

   | Event | When |
   |---|---|
   | Plant | DAP 0 |
   | Treatment | on its DAP |
   | Sacrificial harvest and flood | the day before the first N2O sample |
   | N2O | 24 h and 48 h after the flood |
   | Harvest | after the 48 h sample |

The same random seed always gives the same design.

## The workbook

| Sheet | What it holds |
|---|---|
| `timeline` | A day-by-day calendar, one column per planting group x cohort. **The yellow cells are inputs**: planting date, days between groups, treatment day, first N2O day per cohort, each group's planting date, and capacity. Edit one in Excel and the calendar redraws. The last column counts N2O pots per day and turns red above capacity. |
| `design` | One row per pot: barcode, role, genotype, cohort, treatment, rep, planting group, planting date (linked to the timeline), table, zone, slot, position, tag. |
| `layout` | Each table as a zone x slot grid. |
| `groups` | Planting groups and the N2O load per day as designed. |
| `jars` | `barcode` + `tag` (the tag ends with the position). This is the sheet the dosing station's provisioning reads; add its `admin` sheet before provisioning. |
| `info` | The plan as entered, the seed and the version. |

## Run it

- **Windows:** double-click `run_design.bat`.
- **Anywhere:** `python design_app.py`.

The plan can be saved and reopened (**Save plan / Open plan**, a `.json` file).

## Test it

```
pip install -r requirements.txt pytest formulas
python -m pytest -q tests
```

## Not in v0.1 yet

- Moisture levels by table pair (P0: tables 1-2 dry, 3-4 middle, 5-6 wet).
- Per-cohort flood, measure and harvest sheets, including the N2O load order.
- A swaps sheet.
- Keeping N2O days off weekends.
