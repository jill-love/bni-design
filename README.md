# BNI experiment setup

Plans a BNI pot experiment and writes its design map workbook: where every pot goes on the bench, when each group is planted, and what happens each day.

## What it does

1. **Plan**: in one window, enter the genotypes, cohorts (time-trial groups), treatments and reps, plus the bench and the N2O capacity. **Spares are per genotype per planting group**, and a planting group is one planting date: a spare can only replace a pot of its genotype planted the same day. Set them as a **% of each group's pots** (10 % of a 39-pot group is 4, of a 20-pot group 2; rounded up, at least 1 per genotype in the group) or as a fixed number per genotype per group. They are added once the groups are known, and never change them. Spare tags carry their group: `B73-G2-SPARE-01`. The **Dosing station** box takes the WFPS target and flood WFPS (in %), the soil's bulk and particle density and air-dry moisture, the air-dry soil per pot and the empty pot weight; any of them may be left blank and filled in the workbook later.
2. **Planting groups**: a planting group is **all the pots planted on one date**. N2O is measured one pot at a time, at most *max N2O measurements / day* (50 by default), so planting is staggered so that no day needs more than that, and **a flood never falls on an N2O day**. The pots are split into the **fewest equally sized rounds** that fit, planted at least 3 days apart by default (one round's 24 h and 48 h N2O days, then the next round's flood), or further apart if needed. The split is random but balanced: every genotype x cohort x treatment cell is spread over the rounds as evenly as it divides. A round is one planting group, or more when a cohort is planted later (below): then that cohort is its own planting group, a few days after the rest of its round.
   **N2O times are set per cohort**, in the cohort table: whole days after the flood, e.g. `24, 48, 72` for one cohort and `24, 48` for another (blank: 24, 48). The first N2O day stays the cohort's first sample; the flood is that many hours before it, and harvest follows the last sample. They are fixed when the workbook is made (shown in each calendar column's header); change them in the window and generate again.
   **The treatment day can be set per cohort**, for a treatment time trial: the cohort table's *Treatment* column, e.g. 14, 16 and 18 days after planting for cohorts A, B and C. A blank cell uses the *Treatment on day* setting; a cohort with its own day is treated even when that setting is off. Each must come before the cohort's flood.
   **Cohorts that would collide are planted later, not respaced.** If one cohort's flood would fall on another's N2O day, the program keeps every cohort's days after planting exactly as entered and plants the later cohort a few days later instead (never earlier than the planting date). The cohort with the earliest first N2O day is planted on the planting date; each next one gets the smallest delay that clears the collision, 0 if none is needed. That cohort is then its own planting group, with its own date on the timeline (which may be typed over), its own spares and its own planting tasks.
3. **Bench layout**: pots are spread over the **tables you say are available**, then the zones of each table. Each cell and each planting group is spread as evenly as it divides, and slots are random. A table holds up to 8 zones x 25 slots. Positions are written `T1-Z2-05`.
4. **Timeline**: per planting group and cohort, all counted in days after planting (DAP):

   | Event | When |
   |---|---|
   | Plant | DAP 0 |
   | Water | **Monday, Wednesday and Friday**, from planting until the flood |
   | Treatment | on its DAP: the cohort's own, else the plan's |
   | Sacrificial harvest and flood | the cohort's first N2O time before its first N2O day (the day before, for 24 h) |
   | N2O | at each of the cohort's own times after the flood, e.g. 24, 48, 72 h |
   | Harvest | after the cohort's last N2O sample |

Treatment, flood, N2O and harvest may fall on any day, weekends included.

5. **N2O order**: in each group and cohort, the measured pots are put in loads of 3 (one per N2O column). Each load mixes treatments and genotypes, and every combination is spread over the 3 columns. Pots are flooded in the same order, so every pot has the same time between flood and measurement.
6. **Task list for Asana**: every piece of work as its own dated task. Asana has no repeat option on import, so each repeat (each watering day, each germination check, each N2O day) is its own row. Lab work is put on working days (Mon-Fri), counted from the event it follows:

   | Task | When |
   |---|---|
   | Pre-fill pots with weighed soil | the working day before planting |
   | Label, plant seed, and initial wet-up of pots | planting day |
   | Label tables with treatment and group information per pot | planting day |
   | If using ZeaL genotypes, check for germination | every weekday of the 7 days after planting |
   | Run swaps: replace failed pots with spares, discard unused spares | a week (5 working days) after each planting day |
   | Maintenance watering | each watering day (one task per day, for every group and cohort still being watered) |
   | Make the treatment solutions | once, the working day before the first treatment (only when the plan has treatments); lists each treatment and its pots |
   | Apply treatments | treatment day (one task per day; cohorts treated on different days are named) |
   | Print all labels | once, the working day before the first planting: 1 per pot (spares included), 2 per harvested plant (shoot and root biomass bags), 4 per soil sample (collection tube, weighed tube, soil water tin, extract tube); copies in `LABEL_COPIES` |
   | KCl soil collection setup (gathering supplies, printing & labeling tubes) | once, for every soil sample of the experiment, with the setup: the working day before the first planting |
   | Flood | flood day |
   | N2O measurements | each of the cohort's N2O days |
   | Plant harvest and soil sample collection | flood day (sacrificial pots) and after the last N2O (final) |
   | KCl extractions | once the **last** harvest is in: all soil samples (sacrificial and final) split into the fewest batches of at most 72, as equal as they divide (117 samples: 59 + 58), one batch per working day starting the working day after the last harvest. Each task lists the harvests in its batch. |
   | Prep KCl extracts for shipping | once, the working day after the last extraction batch |
   | Plant dry weights measure & record | 3 working days after each harvest |
   | Wash pots for reuse | the working day after the final harvest |

   The offsets are `TASK_LAGS` in `design_core.py`. In the workbook every due date is a formula on the timeline's yellow inputs (`WORKDAY` for the working-day offsets), so editing a planting date, a treatment day, a cohort's first N2O day or the watering days moves the tasks too.

The same random seed always gives the same design.

## The workbook

| Sheet | What it holds |
|---|---|
| `timeline` | A day-by-day calendar, one column per planting group x cohort. **The yellow cells are inputs**: planting date, days between groups, treatment day, watering days (`Mon, Wed, Fri`), first N2O day and treatment day per cohort (blank: the treatment day above), each group's planting date, and capacity. Edit one in Excel and the calendar redraws. *Flood pots* and *N2O pots* count each day, and turn red if N2O is over capacity or a flood lands on an N2O day. |
| `design` | One row per pot: barcode, role, genotype, cohort, treatment, rep, planting group, planting date (linked to the timeline), table, zone, slot, position, tag. |
| `layout` | Each table as a zone x slot grid, in bright colours by planting group (by cohort when there is one group), with a legend. Sacrificial tags are italic. |
| `flood_*` | Per cohort (and planting group): the sacrificial harvest, then the flood in N2O order, with columns to fill in. |
| `measure_*` | One block per N2O time of that cohort: load, column, seq, pot, and columns for seat mass, times and notes. |
| `harvest_*` | The measured pots after the last N2O time, with shoot and root weight columns. |
| `tags` | Barcode, tag text and position for label printing, by planting group. |
| `groups` | Planting groups and the N2O load per day as designed. |
| `assignments` | Who does each kind of task (Assignee, Collaborator 1, Collaborator 2) and everyone's Asana email, copied from the team file (below). Edit the yellow cells for this experiment. |
| `asana_tasks` | One row per task: Name, Due Date, Section/Column, Description, Assignee, Collaborators (emails, looked up in `assignments`), and nothing else. **File > Save As > CSV**, then in Asana: project > Import > CSV. Due dates follow the timeline's inputs. Watering has a few spare rows at the bottom, used when an edit adds watering days; a watering row whose day an edit removes goes blank. Delete blank rows (filter on Name) before saving the CSV. |
| `jars` | `barcode` + `tag` (the tag ends with the position). The pots for the dosing station. |
| `admin` | The dosing station's WFPS inputs, as fractions (`wfps_target` 0.40 for 40 %), plus `trial_id`, the planting date and `profile` = `wet_up`. Values only, no formulas. With `jars`, the workbook is a complete trial: **Start New Trial...** on the station takes it as is. Any input left blank in the window is an empty yellow cell marked REQUIRED; fill it in before provisioning. |
| `swaps` | Written by the swap tool: date, failed pot, the spare that replaced it, the role it took over, note. |
| `info` | The plan as entered, the seed and the version. A hidden `plan` tab holds the same as JSON, for the swap tool. |

## The team file: who does what

`asana_team.xlsx`, beside the program, says who does each kind of task, so the Asana import comes with assignees and collaborators already filled in. It is never edited in the setup window.

- Make it once: `python team.py` writes it from the lab's work breakdown (primary worker = Assignee; second and supervisory workers = Collaborators). When the program gains a new kind of task, run `python team.py` again: it adds a row for it (with its default) and leaves everything else as it is.
- **people** tab: everyone's name and **Asana email** (fill these in; a person with no email is left off the import).
- **roles** tab: for each kind of task, the Assignee and up to two Collaborators, by name. Kinds left blank (e.g. Flood) are left unassigned.

It stays on the computer (git ignores `*.xlsx`), so the emails are never pushed to GitHub. Each design map gets a copy as its `assignments` tab: change someone for one experiment there, and the swap tool keeps those edits. Without a team file the tab and the two columns are blank.

## Run it

- **Windows:** double-click `run_design.bat`.
- **Anywhere:** `python design_app.py`.
- **Swaps:** `run_swaps.bat`, or `python swap_app.py` (see below).

The plan can be saved and reopened (**Save plan / Open plan**, a `.json` file).

## Swaps: a spare replaces a failed pot

When a pot fails, a spare of the **same genotype and planting group** takes over its role: cohort, treatment, rep, and its turn in the flood, N2O and harvest order. Neither pot moves and no barcode changes; write "promoted to <tag>" on the spare by hand.

1. Close the design map in Excel.
2. Run `run_swaps.bat` (or `python swap_app.py`) and open the design map.
3. Scan or type the failed pot; pick or scan the spare (only spares of its genotype planted the same day are offered, and any other is refused); **Add swap**. Repeat. A swap added by mistake can be removed until you save.
4. Optional: tick **Discard unused spares planted at least a week before the date above** once all swaps are in. Spares of a later planting day are kept until their own week is up.
5. **Save design map.** The workbook is rebuilt from its own layout (never re-randomised): every tab follows the swaps, failed pots and discarded spares are struck through on the layout and leave `jars`, the spare is bold with "(for Snnn)". Values typed into the flood, measure and harvest sheets (by barcode), edited yellow timeline inputs, admin values and tabs added by hand are kept. The old file is kept as `<name>_before_swaps_<date>.xlsx`.
6. On the greenhouse PC, in the dosing station: pick the trial again in the **Trial** dropdown (it reloads the workbook: the failed and discarded pots are retired, the spares carry their new role), then **Load experiment orders** again. Do this between sessions, not in the middle of a flood.

## Test it

```
pip install -r requirements.txt pytest formulas
python -m pytest -q tests
```

## Not in v0.1 yet

- Moisture levels by table pair (P0: tables 1-2 dry, 3-4 middle, 5-6 wet).
- An opening blank in the N2O load order.
