"""
canvas_fetch.py - Pulls your assignments, announcements and calendar exams from Canvas and
saves them as one JSON file for the Work Week dashboard. No Excel involved.

How it works (the big picture):
  1. Load your Canvas token (from the environment, or from the .env file in this folder).
  2. Reuse the Canvas code in canvas_sync.py: same API calls, same type and exam rules.
     Where the dashboard needs something different, this file says so and does it here,
     so canvas_sync.py (the Excel tracker) keeps working exactly as before.
  3. Shape everything into plain data (dates as "2026-10-19", times as "23:59").
  4. Save it as canvas_export.json. The dashboard matches every item by its Canvas ID,
     so running this twice adds nothing twice.

What changed for the dashboard (October 2026):
  - Courses come from the current Canvas term, so a new semester needs no edits.
  - One course that Canvas refuses no longer stops the sync; it is reported instead.
  - A slow or busy Canvas gets one more try before the sync gives up.
  - Assignments with no due date yet are kept when they are real work.
  - "Practice" items are never treated as exams.
  - Exam announcements carry a time and a "sure" flag, and dates are picked more carefully.
  - Exams that only live in the Canvas calendar are read too.
  - Every assignment ID of each course is listed, so the dashboard can tell what was removed.

Run it:   python3 canvas_fetch.py
Output:   canvas_export.json (in this folder)

The token: either the CANVAS_TOKEN environment variable, or a file called .env in this
folder with one line:   CANVAS_TOKEN=your-token
Treat the token like a password. .env is already listed in .gitignore.
"""

import json
import os
import re
import sys
import time
import types
from datetime import datetime, timedelta, timezone
from pathlib import Path

import requests

HERE = Path(__file__).parent
OUTPUT_FILE = HERE / "canvas_export.json"

# Deadlines are written in this time zone, no matter where the script runs.
TIMEZONE = "Asia/Bahrain"

# How much of an announcement's text to keep (the dashboard links to the full post).
ANNOUNCEMENT_CHARS = 3000

# How far ahead to look in the Canvas calendar for exams.
CALENDAR_DAYS_AHEAD = 120

# Seconds to wait before the one retry when Canvas is slow or busy.
RETRY_WAIT = 2

EXPORT_VERSION = 2


# ============================================================
# 1. SETUP: token, time zone, and borrowing canvas_sync.py
# ============================================================

def load_env_file():
    """Read KEY=VALUE lines from .env into the environment.
    A variable that is already set (for example in your terminal) wins."""
    env_file = HERE / ".env"
    if not env_file.exists():
        return
    for line in env_file.read_text().splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def stand_in_for_excel_library():
    """canvas_sync.py imports openpyxl (the Excel library) on its first lines.
    We only use its Canvas functions, so if openpyxl isn't installed we give
    Python an empty stand-in with the names that import expects."""
    names = {
        "openpyxl": ["load_workbook"],
        "openpyxl.styles": ["Font", "PatternFill", "Alignment", "Border", "Side"],
        "openpyxl.formatting": [],
        "openpyxl.formatting.rule": ["FormulaRule"],
        "openpyxl.formatting.formatting": ["ConditionalFormattingList"],
    }
    for module_name, attributes in names.items():
        module = types.ModuleType(module_name)
        for attribute in attributes:
            setattr(module, attribute, None)
        sys.modules[module_name] = module


def import_canvas_sync():
    """Import canvas_sync.py as a module so its functions can be reused."""
    sys.path.insert(0, str(HERE))
    try:
        import canvas_sync
    except ModuleNotFoundError as error:
        if "openpyxl" not in str(error):
            raise
        stand_in_for_excel_library()
        import canvas_sync
    return canvas_sync


def get_time_zone():
    """The TIMEZONE setting as a tzinfo object. Bahrain is UTC+3 all year,
    so that is the fallback if this computer has no time zone database."""
    try:
        from zoneinfo import ZoneInfo
        return ZoneInfo(TIMEZONE)
    except Exception:
        return timezone(timedelta(hours=3))


load_env_file()                 # must happen BEFORE the import: canvas_sync reads the token when it loads
cs = import_canvas_sync()
TZ = get_time_zone()


# ============================================================
# 2. TALKING TO CANVAS, with one retry
# ============================================================

# "Too many requests" and server-side hiccups: worth one more try. Anything else is a real answer.
RETRY_STATUS = {429, 500, 502, 503, 504}


def retry_wait(resp):
    """Canvas may say how long to wait (Retry-After). Use it, but never more than 10 seconds."""
    try:
        return min(10, max(RETRY_WAIT, int(resp.headers.get("Retry-After", RETRY_WAIT))))
    except (TypeError, ValueError):
        return RETRY_WAIT


def canvas_get(path, params=None):
    """Same job as canvas_sync.canvas_get: GET an endpoint and follow every 'next' page.
    The difference: a timeout, a dropped connection or a busy server gets ONE retry
    after a short wait, instead of stopping the whole sync on the first hiccup."""
    url = f"{cs.CANVAS_URL.rstrip('/')}/api/v1{path}"
    headers = {"Authorization": f"Bearer {cs.TOKEN}"}   # this is how Canvas knows it's you
    params = dict(params or {}, per_page=100)
    results = []

    while url:
        resp = None
        for attempt in (1, 2):
            try:
                resp = requests.get(url, headers=headers, params=params, timeout=30)
            except (requests.Timeout, requests.ConnectionError):
                if attempt == 2:
                    raise
                time.sleep(RETRY_WAIT)
                continue
            if resp.status_code in RETRY_STATUS and attempt == 1:
                time.sleep(retry_wait(resp))
                continue
            break
        if resp.status_code == 401:
            sys.exit("Canvas said 401 Unauthorized: your token is wrong or expired. Make a new one.")
        resp.raise_for_status()                       # 403, 404...: the caller decides what to do
        data = resp.json()
        results.extend(data if isinstance(data, list) else [data])
        url = resp.links.get("next", {}).get("url")   # None on the last page -> loop ends
        params = None                                 # the 'next' URL already has the params
    return results


# canvas_sync's helpers (get_assignments, get_announcements) look canvas_get up by name
# when they run, so pointing that name here gives them the retry too.
cs.canvas_get = canvas_get


def describe(error):
    """A short, readable reason for a failed request: '403 Forbidden', 'Timeout'..."""
    resp = getattr(error, "response", None)
    if resp is not None:
        return f"{resp.status_code} {resp.reason or ''}".strip()
    return type(error).__name__


# ============================================================
# 3. WHICH COURSES: the current term
# ============================================================

def parse_utc(iso_text):
    if not iso_text:
        return None
    try:
        return datetime.fromisoformat(iso_text.replace("Z", "+00:00"))
    except ValueError:
        return None


def code_for(course):
    """'[COSC 412-1] Implementation...' -> 'COSC 412': Canvas's course code without the section."""
    raw = course.get("course_code") or ""
    if not raw:
        m = re.search(r"\[([^\]]+)\]", course.get("name", ""))
        raw = m.group(1) if m else ""
    return re.sub(r"\s*-\s*\d+\s*$", "", raw).strip()


def short_name(course):
    """A course that isn't in CLASS_NAMES: its Canvas name without the '[CODE-1]' part."""
    name = re.sub(r"^\s*\[[^\]]*\]\s*", "", course.get("name") or "").strip()
    return name or code_for(course) or "Course"


def current_courses():
    """Your courses for this semester, plus the ones left out and why.

    A course counts when its Canvas term is running now (FALL 2026: 12 Sep to 13 Jan).
    Book Store, Welcome to AUBH and old test courses have no semester dates, so they drop
    out on their own and a new semester needs no list to keep up to date. CLASS_NAMES in
    canvas_sync.py now only gives short names ('Database Systems'); a course without one
    uses its Canvas name. A listed course with no term dates at all is still kept.
    """
    now = datetime.now(timezone.utc)
    keep, left_out = [], []
    courses = canvas_get("/courses", {"enrollment_state": "active", "enrollment_type": "student", "include[]": "term"})
    for course in courses:
        if not course.get("name"):                     # old/locked courses with no name
            continue
        listed = cs.listed_class(course)
        term = course.get("term") or {}
        start, end = parse_utc(term.get("start_at")), parse_utc(term.get("end_at"))
        has_dates = bool(start and end)
        running = has_dates and start <= now <= end
        if running or (listed and not has_dates):
            keep.append((course, listed or short_name(course)))
        elif has_dates:
            left_out.append({"name": course["name"], "reason": f"its term ({term.get('name') or 'unnamed'}) is not running now"})
        else:
            left_out.append({"name": course["name"], "reason": "not part of a semester"})
    return keep, left_out


def course_code(class_name, course):
    """'Data Mining' -> 'DSAI 422', from CLASS_NAMES when it is listed there, else from Canvas."""
    for piece, short in cs.CLASS_NAMES.items():
        if short == class_name:
            return re.sub(r"-\d+$", "", piece)
    return code_for(course)


# ============================================================
# 4. TURNING CANVAS DATA INTO DASHBOARD DATA
# ============================================================

def local(iso_text):
    """Canvas gives times in UTC ('2026-10-19T20:59:59Z').
    Return ('2026-10-19', '23:59') in TIMEZONE, or (None, None) if there is no time."""
    if not iso_text:
        return None, None
    moment = datetime.fromisoformat(iso_text.replace("Z", "+00:00")).astimezone(TZ)
    return moment.strftime("%Y-%m-%d"), moment.strftime("%H:%M")


# Submission types that mean "there is nothing for you to hand in".
NOTHING_TO_HAND_IN = {"none", "external_tool", "not_graded"}
ONLINE = {"online_upload", "online_text_entry", "online_url", "online_quiz", "media_recording",
          "discussion_topic", "student_annotation"}


def submit_kind(assignment):
    """How the work is handed in: 'online', 'paper' or 'none'. The dashboard uses this to know
    whether Canvas will ever show a submission for it."""
    kinds = set(assignment.get("submission_types") or [])
    if kinds & ONLINE:
        return "online"
    if "on_paper" in kinds:
        return "paper"
    return "none"


def is_practice(assignment):
    """'PRACTICE QUESTIONS FOR TEST 1' is study material, not the test itself:
    anything named 'practice', or worth 0 points and only a link to an outside tool."""
    kinds = assignment.get("submission_types") or []
    if re.search(r"\bpractice\b", assignment["name"], re.I):
        return True
    return not assignment.get("points_possible") and kinds == ["external_tool"]


def tracker_type(assignment):
    """The same type canvas_sync.py gives, with two fixes:
    practice items are 'Practice', never an exam, and 'Final Report' or
    'Final Proposal' is a report or an assignment, not a final exam."""
    if is_practice(assignment):
        return "Practice"
    label = cs.assignment_type(assignment)
    name = assignment["name"].lower()
    if label == "Final" and not re.search(r"\bexam|\btest\b", name):
        for pattern, other in cs.TYPE_RULES:
            if other not in ("Final", "Midterm") and re.search(pattern, name):
                return other
        if re.search(r"proposal|paper|essay|draft|submission", name):
            return "Assignment"
    return label


def worth_tracking_undated(assignment):
    """An assignment with no due date yet is kept when it is real work: it is worth points
    and you hand something in (online or on paper). In-class marks with nothing to hand in
    ('Participation 2', 'Quiz1 (A1)') and plain links ('EBOOK ACCESS Link') are left out."""
    kinds = set(assignment.get("submission_types") or [])
    return bool(assignment.get("points_possible")) and bool(kinds - NOTHING_TO_HAND_IN)


def shape_assignment(assignment, class_name, code):
    """One Canvas assignment -> the fields the dashboard needs. None = skip it."""
    due_date, due_time = local(assignment.get("due_at"))
    if not due_date and not worth_tracking_undated(assignment):
        return None
    submission = assignment.get("submission") or {}
    submitted_date, _ = local(submission.get("submitted_at"))
    return {
        "id": str(assignment["id"]),
        "course": class_name,
        "code": code,
        "name": assignment["name"],
        "type": tracker_type(assignment),
        "due": due_date,                             # None = no due date yet
        "due_time": due_time,
        "status": cs.status(assignment),             # "Completed", "Missing" or None
        "submitted": submitted_date,
        "points": assignment.get("points_possible"),
        "submit": submit_kind(assignment),           # "online", "paper" or "none"
        "url": assignment.get("html_url"),
    }


# ---------- exam dates and times in announcement text ----------

EXAM_WORD = r"(?:mid-?terms?|finals?|quiz(?:zes)?|exams?|examination|tests?)"
# Text right before a date that says "this is when the exam is".
POINTS_TO_EXAM_DATE = re.compile(
    rf"(?:\b{EXAM_WORD}\b|\bon|\bheld|\bscheduled|\bwill be|\btakes? place|\bdate|\bis now"
    rf"|\b(?:moved|postponed|rescheduled|changed|shifted)\b.{{0,30}}\bto)\W*(?:\w+\W+){{0,3}}$",
    re.I)
# Text right before a date that says "this is NOT the exam date": the material it covers,
# a deadline, or the old date in "moved from 21 Oct".
NOT_THE_EXAM_DATE = re.compile(
    r"(?:\bup to|\buntil|\btill|\bcover(?:s|ing)?|\bfrom|\bdue|\bdeadline|\bsubmit\w*|\bposted|\bsince)\W*(?:the\s+)?(?:\w+\W+)?$",
    re.I)

TIME_RANGE = re.compile(r"\b(\d{1,2})(?:[:.](\d{2}))?\s*(?:-|–|to)\s*\d{1,2}(?:[:.]\d{2})?\s*([ap])\.?\s*m\b\.?", re.I)
TIME_12H = re.compile(r"\b(\d{1,2})(?:[:.](\d{2}))?\s*([ap])\.?\s*m\b\.?", re.I)
TIME_24H = re.compile(r"\b([01]?\d|2[0-3]):([0-5]\d)\b")


def clock(hour, minute, meridiem):
    """(4, 0, 'p') -> '16:00'. With no am/pm, an hour before 7 is read as afternoon:
    an exam 'at 4:00' is at 16:00, not at 4 in the morning."""
    hour, minute = int(hour), int(minute or 0)
    if hour > 23 or minute > 59:
        return None
    if meridiem:
        meridiem = meridiem.lower()
        if hour > 12:
            return None
        hour = hour % 12 + (12 if meridiem == "p" else 0)
    elif hour < 7:
        hour += 12
    return f"{hour:02d}:{minute:02d}"


def find_time(text):
    """The first time of day written in text, as 'HH:MM', or None."""
    for pattern in (TIME_RANGE, TIME_12H):
        m = pattern.search(text)
        if m:
            return clock(m.group(1), m.group(2), m.group(3))
    m = TIME_24H.search(text)
    return clock(m.group(1), m.group(2), None) if m else None


def find_exam_when(title, text, posted):
    """Best guess at an exam's date and time in an announcement: (date, 'HH:MM' or None, sure).

    Every date in the title and text is a candidate. Dates before the posting day are
    dropped, and so are dates right after 'up to', 'until', 'covers', 'due', 'from'...
    (the material or a deadline, or the old date in 'moved from 21 Oct').
    The rest are scored: +3 when the date is in the title, +2 when it follows words that
    point at the exam ('quiz on', 'held', 'moved to'). The best score wins, the earlier one
    on a tie. 'sure' means the score was at least 2; anything else the dashboard marks
    'check date'.
    """
    combined = f"{title} . {text}"
    in_title = len(title)
    best = None
    for kind, pattern in cs.DATE_PATTERNS:
        for m in pattern.finditer(combined):
            d = date_from_match(kind, m, posted)
            if d is None or d < posted - timedelta(days=1):
                continue
            before = combined[max(0, m.start() - 45):m.start()]
            if NOT_THE_EXAM_DATE.search(before):
                continue
            score = (3 if m.start() < in_title else 0) + (2 if POINTS_TO_EXAM_DATE.search(before) else 0)
            if best is None or score > best[0] or (score == best[0] and m.start() < best[1]):
                best = (score, m.start(), m.end(), d)
    if best is None:
        return None, None, False
    score, start, end, d = best
    near = combined[end:end + 60]                    # "Sunday 18 October, 4:00 PM"
    when = find_time(near) or find_time(combined[max(0, start - 35):start]) or find_time(title)
    return d.strftime("%Y-%m-%d"), when, score >= 2


def date_from_match(kind, m, posted):
    """Turn one date match into a datetime, the same way canvas_sync.find_date does."""
    if kind == "dm":
        day, month, year = int(m.group(1)), cs.MONTH_NUM[m.group(2).lower()[:3]], m.group(3)
    elif kind == "md":
        month, day, year = cs.MONTH_NUM[m.group(1).lower()[:3]], int(m.group(2)), m.group(3)
    else:
        a, b, year = int(m.group(1)), int(m.group(2)), m.group(3)
        day, month = (a, b) if cs.DATE_ORDER == "DMY" else (b, a)
        if month > 12 and day <= 12:                   # "10/15" can only be Oct 15
            day, month = month, day
    year_written = year is not None
    year = int(year) if year_written else posted.year
    if year < 100:
        year += 2000
    try:
        d = datetime(year, month, day)
    except ValueError:                                 # e.g. 31/02, not a real date
        return None
    if not year_written and d < posted - timedelta(days=60):
        try:
            d = d.replace(year=d.year + 1)             # "5 Jan" posted in December = next year
        except ValueError:
            return None
    return d


# canvas_sync.exam_kind only trusts the word "test" in a title, because a CS course body often says
# "unit test" or "test your code". These body phrases are clearly an exam: "lab test", "practical test",
# "the test will be on...", "test 2 is on...". "unit test(s)" never counts.
BODY_TEST = re.compile(
    r"(?<!unit )\b(?:(?:lab|practical|written|class|in-class|online)\s+tests?\b"
    r"|tests?\s*(?:\d+\s*)?(?:will be|is|takes place|is scheduled)\s+(?:on|held|in|at)\b)", re.I)


def exam_kind(title, text):
    """canvas_sync.exam_kind, plus the body phrases above."""
    kind = cs.exam_kind(title, text)
    if kind is None and BODY_TEST.search(text or ""):
        kind = "Exam"
    return kind


def shape_announcement(topic, class_name, code):
    """One Canvas announcement -> the fields the dashboard needs,
    including the exam it talks about (if any) with its date, time and how sure we are."""
    posted_date, _ = local(topic["posted_at"])
    title = topic.get("title") or "(no title)"
    text = cs.html_to_text(topic.get("message") or "")

    exam = None
    kind = exam_kind(title, text)                   # "Quiz", "Exam", "Midterm", "Final" or None
    if kind:
        posted = datetime.strptime(posted_date, "%Y-%m-%d")
        exam_date, exam_time, sure = find_exam_when(title, text, posted)
        exam = {"type": kind, "date": exam_date, "time": exam_time if exam_date else None, "sure": sure}

    return {
        "id": f"ann-{topic['id']}",
        "course": class_name,
        "code": code,
        "title": title,
        "posted": posted_date,
        "text": cs.preview(text, ANNOUNCEMENT_CHARS),
        "url": topic.get("html_url"),
        "exam": exam,
    }


def get_exam_events(courses):
    """Exams the instructor put in the Canvas calendar instead of an assignment or announcement.
    Only events whose TITLE says quiz, test, exam, midterm or final count; the calendar
    gives the real date and time, so they are always 'sure'."""
    start = datetime.now(TZ).date()
    end = start + timedelta(days=CALENDAR_DAYS_AHEAD)
    found = []
    for i in range(0, len(courses), 10):              # Canvas takes at most 10 courses per request
        chunk = courses[i:i + 10]
        by_context = {f"course_{course['id']}": (name, code) for course, name, code in chunk}
        events = canvas_get("/calendar_events", {
            "type": "event",
            "context_codes[]": list(by_context),
            "start_date": start.isoformat(),
            "end_date": end.isoformat(),
        })
        for event in events:
            title = event.get("title") or ""
            kind = cs.exam_kind(title, "")           # title only: "test your code" in a body is not an exam
            if not kind or not event.get("start_at"):
                continue
            name, code = by_context.get(event.get("context_code"), (None, None))
            if not name:
                continue
            day, at = local(event.get("start_at"))
            found.append({
                "id": f"ev-{event['id']}",
                "course": name,
                "code": code,
                "title": title,
                "date": day,
                "url": event.get("html_url"),
                "exam": {"type": kind, "date": day, "time": None if event.get("all_day") else at, "sure": True},
            })
    return found


def build_export():
    """Ask Canvas for everything and return it as one dictionary."""
    assignments, announcements, events, problems = [], [], [], []
    courses, course_status = [], []
    kept, left_out = current_courses()
    with_codes = []

    for course, class_name in kept:
        code = course_code(class_name, course)
        with_codes.append((course, class_name, code))
        courses.append(class_name)
        entry = {"name": class_name, "code": code, "ok": True, "error": None, "assignment_ids": []}
        course_status.append(entry)

        try:
            raw = cs.get_assignments(course["id"])
        except Exception as error:                  # one locked course must not stop the others
            entry["ok"], entry["error"] = False, describe(error)
            problems.append(f"Couldn't read {class_name}: {entry['error']}")
            continue
        entry["assignment_ids"] = [str(a["id"]) for a in raw]
        for assignment in raw:
            shaped = shape_assignment(assignment, class_name, code)
            if shaped:
                assignments.append(shaped)

        if not cs.INCLUDE_ANNOUNCEMENTS:
            continue
        try:
            topics = cs.get_announcements(course["id"])
        except Exception as error:                  # e.g. announcements turned off for a course
            problems.append(f"Couldn't read announcements for {class_name}: {describe(error)}")
            continue
        announcements.extend(shape_announcement(topic, class_name, code) for topic in topics)

    if kept and not any(c["ok"] for c in course_status):
        raise RuntimeError(f"Canvas didn't give any of your courses. {problems[0] if problems else ''}".strip())

    try:
        events = get_exam_events(with_codes)
    except Exception as error:                      # the calendar is a bonus: never fatal
        problems.append(f"Couldn't read the Canvas calendar: {describe(error)}")

    return {
        "version": EXPORT_VERSION,
        "stamp": int(time.time() * 1000),           # milliseconds since 1970: newer export = bigger number
        "fetched_at": datetime.now(TZ).isoformat(timespec="seconds"),
        "canvas": cs.CANVAS_URL,
        "courses": courses,
        "skipped_courses": [c["name"] for c in left_out],
        "left_out": left_out,                       # [{name, reason}]
        "course_status": course_status,             # [{name, code, ok, error, assignment_ids}]
        "problems": problems,
        "assignments": assignments,
        "announcements": announcements,
        "events": events,
    }


# ============================================================
# 5. MAIN
# ============================================================

def main():
    if not cs.TOKEN:
        sys.exit("No Canvas token found. Put it in a file called .env next to this script, "
                 "as one line:  CANVAS_TOKEN=your-token")
    if "YOUR-SCHOOL" in cs.CANVAS_URL:
        sys.exit("Edit CANVAS_URL at the top of canvas_sync.py first.")

    export = build_export()
    OUTPUT_FILE.write_text(json.dumps(export, ensure_ascii=False, indent=1))

    items = export["assignments"]
    completed = sum(1 for a in items if a["status"] == "Completed")
    missing = sum(1 for a in items if a["status"] == "Missing")
    undated = sum(1 for a in items if not a["due"])
    exams = sum(1 for n in export["announcements"] if n["exam"])
    print(f"Courses: {', '.join(export['courses']) or 'none'}")
    if export["left_out"]:
        print(f"Left out {len(export['left_out'])} courses that aren't in the current term")
    print(f"Assignments: {len(items)} ({completed} completed, {missing} missing, {undated} with no due date yet)")
    print(f"Announcements: {len(export['announcements'])} ({exams} mention a quiz or exam)")
    print(f"Calendar exams: {len(export['events'])}")
    for problem in export["problems"]:
        print(f"Problem: {problem}")
    print(f"Saved {OUTPUT_FILE.name}")


if __name__ == "__main__":
    main()
