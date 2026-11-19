#!/usr/bin/env python3
"""
team.py: who does which kind of task, for the Asana import (v0.1, 2026-11-05).

    python team.py            # write asana_team.xlsx from the defaults, if it is missing;
                              # otherwise add rows for task kinds it does not list yet

asana_team.xlsx, beside the program, is the TEAM file: edited in Excel when people change,
never in the setup window. Two tabs:
  people   Name | Asana email
  roles    kind | Task | Assignee | Collaborator 1 | Collaborator 2   (names from `people`)
Every design map gets a copy of both as its `assignments` tab, editable for that
experiment; the asana_tasks tab looks the emails up from there (design_xlsx.py). No team
file: the assignments tab is written blank, and so are the Assignee / Collaborators columns.
"""
from __future__ import annotations

import os
import sys
from dataclasses import dataclass, field

import design_core as dc

TEAM_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "asana_team.xlsx")
COLLABORATORS = 2                # collaborator columns per role

# The lab's work breakdown (Experiment tasks updated, 2026-11): the primary worker is the
# assignee, the others (supervisory included) collaborators. "Undergrad #1 or 2" -> #1.
# Emails are left for the team file; kinds with no default are left blank.
DEFAULT_PEOPLE = ["Jill", "Lina", "Sarah (Undergrad #1)", "Undergrad #2"]
DEFAULT_ROLES = {
    "print_harvest_labels": ("Jill", ()),
    "apply_harvest_labels": ("Lina", ()),
    "prefill": ("Sarah (Undergrad #1)", ("Lina",)),
    "plant": ("Lina", ("Jill", "Sarah (Undergrad #1)")),
    "label_tables": ("Jill", ()),
    "watering": ("Sarah (Undergrad #1)", ("Jill",)),
    "make_treatments": ("Lina", ()),
    "n2o": ("Sarah (Undergrad #1)", ("Jill",)),
    "harvest": ("Undergrad #2", ("Lina",)),
    "dry_weights": ("Sarah (Undergrad #1)", ("Jill",)),
    "kcl_extract": ("Sarah (Undergrad #1)", ("Lina",)),
    "kcl_ship": ("Lina", ()),
    "wash": ("Sarah (Undergrad #1)", ("Lina", "Jill")),
}


@dataclass
class Team:
    people: dict[str, str] = field(default_factory=dict)            # name -> email
    roles: dict[str, tuple[str, tuple[str, ...]]] = field(default_factory=dict)

    def role(self, kind: str) -> tuple[str, tuple[str, ...]]:
        return self.roles.get(kind, ("", ()))


def defaults() -> Team:
    return Team({n: "" for n in DEFAULT_PEOPLE}, dict(DEFAULT_ROLES))


def load(path: str | None = None) -> Team | None:
    """The team file, or None when there is none."""
    import openpyxl
    path = path or TEAM_FILE
    if not os.path.exists(path):
        return None
    wb = openpyxl.load_workbook(path, data_only=True)
    people = {}
    for r in wb["people"].iter_rows(min_row=2, values_only=True):
        if r and r[0] and str(r[0]).strip():
            people[str(r[0]).strip()] = str(r[1] or "").strip()
    roles = {}
    for r in wb["roles"].iter_rows(min_row=2, values_only=True):
        if r and r[0] and str(r[0]).strip() in dc.TASK_KINDS:
            assignee = str(r[2] or "").strip() if len(r) > 2 else ""
            collab = tuple(str(x).strip() for x in r[3:3 + COLLABORATORS] if x and str(x).strip())
            roles[str(r[0]).strip()] = (assignee, collab)
    return Team(people, roles)


def write(team: Team, path: str) -> None:
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    yellow = PatternFill("solid", start_color="FEF9C3", end_color="FEF9C3")
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "people"
    ws.append(["Name", "Asana email"])
    for n, e in team.people.items():
        ws.append([n, e or None])
    for r in range(2, len(team.people) + 12):
        for c in (1, 2):
            ws.cell(r, c).fill = yellow
    ws.column_dimensions["A"].width = 24; ws.column_dimensions["B"].width = 32
    ro = wb.create_sheet("roles")
    ro.append(["kind", "Task", "Assignee"] + [f"Collaborator {i}"
                                              for i in range(1, COLLABORATORS + 1)])
    for kind, what in dc.TASK_KINDS.items():
        a, cs = team.role(kind)
        ro.append([kind, what, a or None] + [cs[i] if i < len(cs) else None
                                             for i in range(COLLABORATORS)])
        for c in range(3, 4 + COLLABORATORS):
            ro.cell(ro.max_row, c).fill = yellow
    for ws_ in (ws, ro):
        for c in ws_[1]:
            c.font = Font(bold=True)
    ro.column_dimensions["A"].width = 13; ro.column_dimensions["B"].width = 58
    for col in "CDE":
        ro.column_dimensions[col].width = 22
    wb.save(path)


def upgrade(path: str) -> list[str]:
    """Add a roles row, with its default, for every task kind the team file does not list
    yet (a kind added to the program after the file was made), in TASK_KINDS order.
    Nothing already in the file changes. Returns the kinds added."""
    import openpyxl
    from openpyxl.styles import PatternFill
    yellow = PatternFill("solid", start_color="FEF9C3", end_color="FEF9C3")
    wb = openpyxl.load_workbook(path)
    ws = wb["roles"]
    have = {str(r[0].value).strip(): r[0].row for r in ws.iter_rows(min_row=2) if r[0].value}
    order = list(dc.TASK_KINDS)
    added = []
    for kind in order:
        if kind in have:
            continue
        later = [have[k] for k in order[order.index(kind) + 1:] if k in have]
        row = min(later) if later else ws.max_row + 1
        ws.insert_rows(row)
        for k, r in have.items():                   # rows below moved down by one
            if r >= row:
                have[k] = r + 1
        a, cs = DEFAULT_ROLES.get(kind, ("", ()))
        for c, v in enumerate([kind, dc.TASK_KINDS[kind], a or None]
                              + [cs[i] if i < len(cs) else None
                                 for i in range(COLLABORATORS)], 1):
            ws.cell(row, c, value=v)
        for c in range(3, 4 + COLLABORATORS):
            ws.cell(row, c).fill = yellow
        have[kind] = row
        added.append(kind)
    if added:
        wb.save(path)
    return added


if __name__ == "__main__":
    if os.path.exists(TEAM_FILE):
        added = upgrade(TEAM_FILE)
        print(f"{TEAM_FILE}: " + (f"added {', '.join(added)}" if added
                                  else "already lists every kind of task"))
        sys.exit(0)
    write(defaults(), TEAM_FILE)
    print(f"wrote {TEAM_FILE}: fill in the Asana emails on the people tab")
