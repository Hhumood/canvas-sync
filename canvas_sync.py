"""
canvas_sync.py - Pulls your assignments from Canvas into Assignment_Tracker.xlsx

How it works (the big picture):
  1. Ask Canvas "which courses am I enrolled in?"          (API call)
  2. For each course, ask "what are the assignments, and
     what's MY submission status on each one?"             (API call)
  3. Turn each assignment into a row: name, type, class, deadline, status
  4. Open the Excel file, add new rows / update existing ones, save.

Run it:        python3 canvas_sync.py
Test run:      python3 canvas_sync.py --dry-run    (shows what WOULD change, saves nothing)
"""

import os
import re
import sys
import shutil
import warnings
from datetime import datetime
from pathlib import Path

import requests                     # talks to the Canvas API over HTTPS
from openpyxl import load_workbook  # reads/writes the .xlsx file

# ============================================================
# 1. SETTINGS  (the only part you should need to edit)
# ============================================================

# Your school's Canvas address: copy it from your browser's address bar.
CANVAS_URL = "https://canvas.aubh.edu.bh"


# The Excel file, expected in the same folder as this script.
EXCEL_FILE = Path(__file__).parent / "Assignment_Tracker.xlsx"

# Canvas course names are long ("CS 350-01 Computer Networks Fall 2026").
# Map a piece of the course name/code -> the short class name you want in Excel.
# Leave it empty at first: the script prints every course it finds, then fill this in.
CLASS_NAMES = {
    "ARHG 106-5":  "History of Bahrain",
    "BIOL 101L-2": "Biology Lab",
    "BIOL 101-1":  "Biology",
    "COSC 410L-1": "CS Certificate",
    "COSC 412-1":  "Database Systems",
    "COSC 485-1":  "Web Engineering",
    "COSC 499A-1": "Design Project",
    "DSAI 422-1":  "Data Mining",
}

# Skip assignments that have no due date (attendance, optional stuff, etc.)
SKIP_UNDATED = True

# Your token is read from an environment variable, NOT written in this file.
# Treat the token like a password: anyone with it can read your Canvas account.
TOKEN = os.environ.get("CANVAS_TOKEN")

# Layout of the tracker (matches Assignment_Tracker.xlsx - don't change)
SHEET = "Assignments"
FIRST_ROW, LAST_ROW = 6, 305
COL = {"name": 1, "type": 2, "class": 3, "deadline": 4,
       "status": 5, "comments": 6, "canvas_id": 7}

# hides harmless openpyxl notices
warnings.filterwarnings("ignore", module="openpyxl")


# ============================================================
# 2. TALKING TO CANVAS
# ============================================================

def canvas_get(path, params=None):
    """GET a Canvas API endpoint and return ALL results as one list.

    Canvas sends results in pages (like a search engine: page 1, page 2...).
    Each response includes a 'next' link in its headers until the last page,
    so we keep following that link and collect everything.
    """
    url = f"{CANVAS_URL.rstrip('/')}/api/v1{path}"
    # this is how Canvas knows it's you
    headers = {"Authorization": f"Bearer {TOKEN}"}
    params = dict(params or {}, per_page=100)
    results = []

    while url:
        resp = requests.get(url, headers=headers, params=params, timeout=30)
        if resp.status_code == 401:
            sys.exit(
                "Canvas said 401 Unauthorized: your token is wrong or expired. Make a new one.")
        resp.raise_for_status()                       # stop on any other HTTP error
        data = resp.json()                            # JSON text -> Python list/dict
        results.extend(data if isinstance(data, list) else [data])
        # None on the last page -> loop ends
        url = resp.links.get("next", {}).get("url")
        params = None                                 # the 'next' URL already has the params
    return results


def get_courses():
    """Courses you're currently enrolled in as a student."""
    courses = canvas_get(
        "/courses", {"enrollment_state": "active", "enrollment_type": "student"})
    # Some old/locked courses come back without a name; skip those.
    return [c for c in courses if c.get("name")]


def get_assignments(course_id):
    """All assignments in a course, each with YOUR submission attached."""
    return canvas_get(f"/courses/{course_id}/assignments", {"include[]": "submission"})


# ============================================================
# 3. TURNING CANVAS DATA INTO TRACKER VALUES
# ============================================================

def class_name(course):
    """Short class name from CLASS_NAMES, or the course code if not mapped."""
    haystack = f"{course.get('name', '')} {course.get('course_code', '')}".lower(
    )
    for piece, short in CLASS_NAMES.items():
        if piece.lower() in haystack:
            return short
    return course.get("course_code") or course["name"]


# Checked top to bottom, first match wins. \b = "start of a word",
# so r"\blab" matches "Lab 3" but NOT "syllabus".
TYPE_RULES = [
    (r"\bmidterm", "Midterm"),
    (r"\bfinal exam|\bfinal\b(?!.*project)", "Final"),
    (r"\bproject", "Project"),
    (r"\bexam\b|\btest\b", "Exam"),
    (r"\bquiz", "Quiz"),
    (r"\blab", "Lab"),
    (r"\bpresentation", "Presentation"),
    (r"\breport", "Report"),
    (r"\breading", "Reading"),
    (r"\bdiscussion", "Discussion"),
    (r"\bhomework|\bhw\s?\d|\bproblem set", "Homework"),
]


def assignment_type(a):
    """Guess the Assignment Type from Canvas info and the assignment's name."""
    kinds = a.get("submission_types") or []
    if a.get("is_quiz_assignment") or "online_quiz" in kinds:
        return "Quiz"
    if "discussion_topic" in kinds:
        return "Discussion"
    name = a["name"].lower()
    for pattern, label in TYPE_RULES:
        if re.search(pattern, name):
            return label
    return "Assignment"


def status(a):
    """'Completed', 'Missing', or None (= pending, leave blank)."""
    sub = a.get("submission") or {}
    if sub.get("excused") or sub.get("submitted_at"):
        return "Completed"
    if sub.get("missing"):
        return "Missing"
    if sub.get("workflow_state") == "graded":   # graded without an upload, e.g. in-class work
        return "Completed"
    return None


def deadline(a):
    """Canvas gives times in UTC ('2026-10-02T20:59:59Z').
    Convert to YOUR Mac's time zone, return (date for Excel, 'Due 11:59 PM' text)."""
    due = a.get("due_at")
    if not due:
        return None, ""
    local = datetime.fromisoformat(due.replace("Z", "+00:00")).astimezone()
    return datetime(local.year, local.month, local.day), "Due " + local.strftime("%-I:%M %p")


def comments(a, due_time):
    parts = [due_time]
    if a.get("points_possible"):
        parts.append(f"{a['points_possible']:g} pts")
    return " · ".join(p for p in parts if p)


# ============================================================
# 4. WRITING TO EXCEL
# ============================================================

def ensure_in_list(lists_ws, column_letter, value):
    """Add value to a dropdown list on the Lists sheet if it isn't there yet
    (the Dashboard only counts classes that are in Lists!A)."""
    existing = [lists_ws[f"{column_letter}{r}"].value for r in range(2, 42)]
    if value in existing:
        return False
    first_blank = existing.index(None) + 2
    lists_ws[f"{column_letter}{first_blank}"] = value
    return True


def main():
    dry_run = "--dry-run" in sys.argv

    if not TOKEN:
        sys.exit("No token found. Run:  echo 'export CANVAS_TOKEN=\"paste-token\"' >> ~/.zshrc  "
                 "then open a NEW terminal window.")
    if "YOUR-SCHOOL" in CANVAS_URL:
        sys.exit("Edit CANVAS_URL at the top of canvas_sync.py first.")
    if not EXCEL_FILE.exists():
        sys.exit(
            f"Can't find {EXCEL_FILE}. Put the script and the Excel file in the same folder.")

    wb = load_workbook(EXCEL_FILE)
    ws, lists_ws = wb[SHEET], wb["Lists"]

    # Clear the EXAMPLE row the template came with
    if str(ws.cell(FIRST_ROW, 1).value or "").startswith("EXAMPLE"):
        for c in range(1, 8):
            ws.cell(FIRST_ROW, c).value = None
            ws.cell(FIRST_ROW, c).hyperlink = None

    # Remember which Canvas assignments are already in the sheet: {canvas_id: row_number}
    existing = {}
    for r in range(FIRST_ROW, LAST_ROW + 1):
        cid = ws.cell(r, COL["canvas_id"]).value
        if cid:
            existing[str(cid)] = r

    def next_empty_row():
        for r in range(FIRST_ROW, LAST_ROW + 1):
            if not ws.cell(r, COL["name"]).value and not ws.cell(r, COL["canvas_id"]).value:
                return r
        sys.exit(
            f"The sheet is full (rows {FIRST_ROW}-{LAST_ROW}). Ask Claude to extend it.")

    added, updated = [], []

    for course in get_courses():
        cls = class_name(course)
        print(f"Course: {course['name']}  ->  '{cls}'")
        if not dry_run and ensure_in_list(lists_ws, "A", cls):
            print(f"   (added '{cls}' to the Class list)")

        for a in get_assignments(course["id"]):
            due_date, due_time = deadline(a)
            if SKIP_UNDATED and not due_date:
                continue
            cid, st = str(a["id"]), status(a)

            if cid in existing:
                # Already in the sheet: only update deadline and status.
                # Your own edits to name/type/comments are left alone.
                r = existing[cid]
                changes = []
                if ws.cell(r, COL["deadline"]).value != due_date:
                    changes.append("deadline")
                    ws.cell(r, COL["deadline"]).value = due_date
                if st and ws.cell(r, COL["status"]).value != st:
                    changes.append(f"status -> {st}")
                    ws.cell(r, COL["status"]).value = st
                if changes:
                    updated.append(
                        f"{a['name']} ({cls}): {', '.join(changes)}")
            else:
                r = next_empty_row()
                typ = assignment_type(a)
                values = {"name": a["name"], "type": typ, "class": cls, "deadline": due_date,
                          "status": st, "comments": comments(a, due_time), "canvas_id": cid}
                for key, v in values.items():
                    ws.cell(r, COL[key]).value = v
                if a.get("html_url"):                       # click the name -> opens it in Canvas
                    ws.cell(r, COL["name"]).hyperlink = a["html_url"]
                if not dry_run:
                    ensure_in_list(lists_ws, "B", typ)
                existing[cid] = r
                added.append(f"{a['name']} ({cls}) due {due_date:%a %d %b}")

    print(f"\nAdded {len(added)}:")
    for line in added:
        print("  +", line)
    print(f"Updated {len(updated)}:")
    for line in updated:
        print("  ~", line)

    if dry_run:
        print("\nDry run: nothing was saved.")
        return

    shutil.copy(EXCEL_FILE, EXCEL_FILE.with_name(
        EXCEL_FILE.stem + "_backup.xlsx"))
    # tell Excel to recalculate the Dashboard on open
    wb.calculation.fullCalcOnLoad = True
    try:
        wb.save(EXCEL_FILE)
    except PermissionError:
        sys.exit("Couldn't save: close the Excel file and run again.")
    print(
        f"\nSaved {EXCEL_FILE.name} (previous version kept as {EXCEL_FILE.stem}_backup.xlsx)")


if __name__ == "__main__":
    main()
