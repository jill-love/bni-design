# BNI experiment setup

Plans a BNI pot experiment and writes its design map workbook: where every pot goes on the bench, when each group is planted, and what happens each day.

## What it does

1. **Plan**: in one window, enter the genotypes, cohorts (time-trial groups), treatments and reps, plus the bench and the N2O capacity. **Spares are per genotype per planting group**: a spare can only replace a pot planted the same day. They are added once the groups are known, and never change the number of groups. Spare tags carry their group, for example `B73-G2-SPARE-01`.
2. **Planting groups**: N2O is measured one pot at a time, at most *max N2O measurements / day* (50 by default). Planting is staggered so that no day needs more than that, and **a flood never falls on an N2O day**. The pots are split into the **fewest equally sized groups** that fit, planted at least 3 days apart by default (one group's 24 h and 48 h N2O days, then the next group's flood), or further apart if needed. The split is random but balanced: every genotype x cohort x treatment cell is spread over the groups as evenly as it divides.
3. **Bench layout**: pots are spread over the **tables you say are available**, then the zones of each table. Each cell and each planting group is spread as evenly as it divides, and slots are random. A table holds up to 8 zones x 25 slots. Positions are written `T1-Z2-05`.
4. **Timeline**: per planting group and cohort, all counted in days after planting (DAP):

   | Event | When |
   |---|---|
   | Plant | DAP 0 |
   | Water | **Monday, Wednesday and Friday**, from planting until the flood |
   | Treatment | on its DAP |
   | Sacrificial harvest and flood | the day before the first N2O sample |
   | N2O | 24 h and 48 h after the flood |
   | Harvest | after the 48 h sample |

Treatment, flood, N2O and harvest may fall on any day, weekends included.

5. **N2O order**: in each group and cohort, the measured pots are put in loads of 3 (one per N2O column). Each load mixes treatments and genotypes, and every combination is spread over the 3 columns. Pots are flooded in the same order, so every pot has the same time between flood and measurement.
6. **Task list for Asana**: every piece of work as its own dated task. Asana has no repeat option on import, so each repeat (each watering day, each germination check, each N2O day) is its own row. Lab work is put on working days (Mon-Fri), counted from the event it follows:

   | Task | When |
   |---|---|
   | Pre-fill pots with weighed soil | the working day before planting |
   | Label, plant seed, and initial wet-up of pots | planting day |
   | Label tables with treatment and group information per pot | planting day |
   | If using ZeaL genotypes, check for germination | every weekday of the 7 days after planting |
   | Maintenance watering | each watering day (one task per day, for every group and cohort still being watered) |
   | Apply treatments | treatment day |
   | KCl soil collection setup (gathering supplies, printing & labeling tubes) | the working day before a cohort's first harvest |
   | Flood | flood day |
   | N2O measurements | each N2O day (24 h, 48 h) |
   | Plant harvest and soil sample collection | flood day (sacrificial pots) and after the last N2O (final) |
   | KCl extractions | the working day after each harvest |
   | KCl analysis (plate reader) | the working day after each extraction |
   | Plant dry weights measure & record | 3 working days after each harvest |
   | Wash pots for reuse | the working day after the final harvest |

   The offsets are `TASK_LAGS` in `design_core.py`. In the workbook every due date is a formula on the timeline's yellow inputs (`WORKDAY` for the working-day offsets), so editing a planting date, the treatment day, a cohort's first N2O day or the watering days moves the tasks too.

The same random seed always gives the same design.

## The workbook

| Sheet | What it holds |
|---|---|
| `timeline` | A day-by-day calendar, one column per planting group x cohort. **The yellow cells are inputs**: planting date, days between groups, treatment day, watering days (`Mon, Wed, Fri`), first N2O day per cohort, each group's planting date, and capacity. Edit one in Excel and the calendar redraws. *Flood pots* and *N2O pots* count each day, and turn red if N2O is over capacity or a flood lands on an N2O day. |
| `design` | One row per pot: barcode, role, genotype, cohort, treatment, rep, planting group, planting date (linked to the timeline), table, zone, slot, position, tag. |
| `layout` | Each table as a zone x slot grid, in bright colours by planting group (by cohort when there is one group), with a legend. Sacrificial tags are italic. |
| `flood_*` | Per cohort (and planting group): the sacrificial harvest, then the flood in N2O order, with columns to fill in. |
| `measure_*` | One block per N2O time (24 h, 48 h): load, column, seq, pot, and columns for seat mass, times and notes. |
| `harvest_*` | The measured pots after the last N2O time, with shoot and root weight columns. |
| `tags` | Barcode, tag text and position for label printing, by planting group. |
| `groups` | Planting groups and the N2O load per day as designed. |
| `asana_tasks` | One row per task: Name, Due Date, Section/Column, Description, and nothing else. **File > Save As > CSV**, then in Asana: project > Import > CSV. Due dates follow the timeline's inputs. Watering has a few spare rows at the bottom, used when an edit adds watering days; a watering row whose day an edit removes goes blank. Delete blank rows (filter on Name) before saving the CSV. |
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
- A swaps sheet.
- An opening blank in the N2O load order.
