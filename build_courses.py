#!/usr/bin/env -S uv run --script
# /// script
# dependencies = ["openpyxl"]
# ///
# build_courses.py — build the MATH/STAT master course list from Informer exports.
# Usage:
#   ./build_courses.py                      rebuild the master list in place
#   ./build_courses.py -o OUT.xlsx          write a trial copy somewhere else
#   ./build_courses.py --previous OLD.xlsx -o OUT.xlsx
#   ./build_courses.py --timetable DIR_OR_CSV [--timetable ...] --summaries DIR
#   ./build_courses.py -h | --help
#
#   -o, --output      workbook to write (default: COURSES-math-stat.xlsx in
#                     the curriculum folder). Always written from scratch. A
#                     file already there is first copied beside it as
#                     NAME-bak.xlsx, replacing the last one.
#                     Refuses to run while the workbook is open in Excel.
#                     The real one (not a trial -o) is then copied to the
#                     Advising team as COURSES-math-stat-DO-NOT-EDIT.xlsx
#   -p, --previous    workbook read ONLY for the columns a person typed
#                     (status, notes, ...) and the calendar text (default: the
#                     output workbook itself). Either the old
#                     courses-math-stat-MASTER layout or one this script wrote
#   -t, --timetable   an Informer "UNB Timetable Data" export, or a directory
#                     of them; repeat the flag for several. One row per section
#   -s, --summaries   directory of Informer "Course Summary" exports, one CSV
#                     per term
#       --years N     how many academic years the offering note looks back
#                     (default 6)
#       --sweep DIR   calendar-year folder holding sweep_references.py's
#                     output (default: the newest one that has it)
#
# With no -p/-t/-s the defaults near the top of this file are used; they point
# at the team folders named in lib/paths.py.
#
# WHAT THIS IS FOR
# The master list has two kinds of column. Most can be worked out from the
# registrar's data: when a course runs, in which slot, how many sections, who
# taught it. A few cannot: its status, why it is on hiatus, who else requires
# it. The first kind used to be typed by hand and went stale. This script
# regenerates the first kind and carries the second kind across untouched, so
# the yearly upkeep is "download a new export, run this".
#
# THE TWO SOURCES, AND WHY BOTH
#   timetable exports  every undergraduate section with meeting times, from
#                      Winter 2017. The main source.
#   course summaries   MATH/STAT sections back to Fall 2013, including
#                      graduate ones, but with no meeting times. Used for the
#                      2013-2016 history, and to notice when a cross-listed
#                      course ran under its graduate number only.
#
# THREE MODES OF DELIVERY, KEPT APART
#   in person, fall/winter   what the slot, sections and pattern columns mean
#   in person, summer        its own column
#   online                   its own column, and never counted as a term
#                            offering: until 2025 these were open-entry
#                            courses with no real term, and they have never
#                            been part of the regular teaching load

import argparse
import csv
import datetime
import os
import re
import shutil
import statistics
import sys
from collections import defaultdict
from itertools import zip_longest
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# The helper modules live in lib/, beside this script.
sys.path.insert(0, str(Path(__file__).resolve().parent / "lib"))

import paths

# ----------------------------------------------------------------------------
# Defaults. The data lives outside this repository, on purpose; lib/paths.py
# says where.
# ----------------------------------------------------------------------------
DEFAULT_TIMETABLES = [paths.TIMETABLE_EXPORTS]
DEFAULT_SUMMARIES = paths.COURSE_SUMMARIES
DEFAULT_OUTPUT = paths.MASTER_LIST
# sweep_references.py writes its results into each calendar year's folder;
# the newest year that has them is used for "Required by others".
CALENDAR_DIR = paths.CALENDAR

# Subjects that belong on the list, in the order they are sorted.
SUBJECTS = ["MATH", "STAT", "DATA"]

# A timetable export only "covers" a term if it holds the whole term. The
# 2016-17 file has a single stray Fall 2016 row, which must not be mistaken
# for "nothing but that one section ran".
MIN_ROWS_FOR_COVERAGE = 100

SEASON_ORDER = {"WI": 0, "SM": 1, "FA": 2}
SEASON_LETTER = {"WI": "W", "SM": "S", "FA": "F"}


# ----------------------------------------------------------------------------
# Small helpers
# ----------------------------------------------------------------------------
def term_key(term):
    """Sort key for a term like '2026/FA': by year, then winter-summer-fall."""
    year, season = term.split("/")
    return (int(year), SEASON_ORDER[season])


def term_label(term):
    """'2026/FA' -> '2026 F', the form used in the old workbook."""
    year, season = term.split("/")
    return f"{year} {SEASON_LETTER[season]}"


def course_key(text):
    """'MATH*1003', 'MATH 1003' and 'MATH1003' all become 'MATH1003'."""
    return re.sub(r"[\s*]", "", str(text or "")).upper()


def course_display(key):
    """'MATH1003' -> 'MATH 1003', the way the calendar writes it."""
    return f"{key[:-4]} {key[-4:]}"


def is_course(key):
    return re.fullmatch(r"(%s)\d{4}" % "|".join(SUBJECTS), key) is not None


def sort_key(key):
    return (SUBJECTS.index(key[:-4]), key[-4:])


def compact_time(text):
    """' 8:30AM -  9:20AM' -> '8:30-9:20'."""
    parts = [p.strip()[:-2] for p in text.split("-") if p.strip()]
    return "-".join(parts)


def last_name(full_name):
    """'Dr. Patrick Reynolds' -> 'Reynolds'."""
    words = full_name.strip().split()
    return words[-1] if words else ""


def read_csv(path):
    with open(path, newline="", encoding="utf-8-sig") as handle:
        return list(csv.DictReader(handle))


# ----------------------------------------------------------------------------
# Reading the timetable exports
# ----------------------------------------------------------------------------
def section_kind(section_code, method=""):
    """Sort a section into 'online', 'challenge' or 'inperson'.

    The section code is the reliable signal. The instruction method only says
    WEB from 2021 on; before that it is blank for online sections.
      FRL, FRLN   open-entry online
      FRCH        challenge exam, not a class at all
      WEB         term-based online (new in 2026; these reuse the FRE codes
                  that in-person summer sections also use, hence the check)
    """
    if section_code.startswith("FRCH"):
        return "challenge"
    if section_code.startswith("FRL") or "WEB" in method:
        return "online"
    return "inperson"


def section_meetings(row):
    """Split one timetable row into its lecture slot and tutorial slot.

    A section's meetings are packed into comma-separated columns, for example
    methods 'LEC,TUT', days 'M  W  F, T   ', times '11:30AM - 12:20PM, 8:30AM
    -  9:20AM'. Returns (lecture_text, tutorial_text), e.g.
    ('MWF 11:30-12:20', 'T 8:30-9:20'). Either can be ''.
    """
    methods = row["Instr Method"].split(",")
    days = row["Meeting Days"].split(",")
    times = row["Meeting Times"].split(",")
    lecture, tutorial = [], []
    for method, day, time in zip_longest(methods, days, times, fillvalue=""):
        day = day.replace(" ", "")
        if not day or not time.strip():
            continue  # a part with no set time contributes nothing to the slot
        text = f"{day} {compact_time(time)}"
        (tutorial if method in ("TUT", "LAB") else lecture).append(text)
    return ", ".join(lecture), ", ".join(tutorial)


def export_year(path):
    """'timetable-2026-27-as-of-2026-12-01.csv' -> '2026-27', or None for another name."""
    match = re.match(r"timetable-(\d{4}-\d{2})", Path(path).name)
    return match[1] if match else None


def newest_per_year(files):
    """Of several exports of one academic year, keep only the newest, sorted by year.

    A year still being scheduled is exported more than once, named with its
    date ('timetable-2026-27-as-of-2026-10-12.csv'). Reading two of them would
    keep a section cancelled between the two, so only the latest name counts:
    the dates sort as text, and a bare 'timetable-2026-27.csv' sorts first.
    Sorted by stem, without '.csv': by full name the '.' of a bare name would
    sort after the '-' of a dated one.
    """
    by_year = {}
    for file in sorted(files, key=lambda f: Path(f).stem):
        by_year[export_year(file) or file.name] = file
    return [by_year[year] for year in sorted(by_year)]


def load_timetables(paths, subjects=SUBJECTS):
    """Read every timetable export and return (rows_by_course_term, covered_terms).

    rows_by_course_term maps (course, term) to the list of section rows.
    covered_terms is the set of terms for which an export holds the full term.
    subjects limits which subjects are kept; None keeps every subject.
    """
    files = []
    for path in paths:
        path = Path(path)
        if path.is_dir():
            files += newest_per_year(path.glob("timetable-*.csv"))
        else:
            files.append(path)
    if not files:
        sys.exit("No timetable exports found.")

    rows_in_term = defaultdict(int)
    sections = {}  # (term, section name) -> row; a later file replaces an earlier one
    for file in files:
        for row in read_csv(file):
            rows_in_term[row["Term"]] += 1
            if subjects is None or row["Subject"] in subjects:
                sections[(row["Term"], row["Section Name"])] = row

    covered = {t for t, n in rows_in_term.items() if n >= MIN_ROWS_FOR_COVERAGE}
    by_course_term = defaultdict(list)
    for (term, _), row in sections.items():
        if term in covered:
            by_course_term[(course_key(row["Course Name"]), term)].append(row)
    return by_course_term, covered, files


def load_summaries(directory):
    """Read the course summaries. Only ACTIVE rows are kept.

    The exports list some sections twice, once active and once as a pending
    or cancelled record with no students. Keeping status 'A' alone removes
    those twins; counting every row would overcount some terms by a third.
    """
    by_course_term = defaultdict(list)
    terms = set()
    files = sorted(Path(directory).glob("*.csv"))
    for file in files:
        for row in read_csv(file):
            terms.add(row["Term"])
            if row["Status"] == "A":
                key = course_key(row["Subject"] + row["Course Number"])
                by_course_term[(key, row["Term"])].append(row)
    return by_course_term, terms, files


# ----------------------------------------------------------------------------
# Turning sections into "offerings": one record per course per term
# ----------------------------------------------------------------------------
def offering_from_timetable(rows):
    """Summarise one course's sections in one term (timetable export)."""
    record = new_offering("timetable")
    lectures = {}  # (lecture slot, instructors) -> True; one entry per real lecture
    tutorial_slots = []
    tutorial_sections = 0
    for row in rows:
        kind = section_kind(row["Number"], row["Instr Method"])
        if kind == "challenge":
            continue
        if kind == "online":
            record["online"] = True
            continue
        record["inperson"] = True
        record["start"] = min(record["start"] or row["Section Start Date"], row["Section Start Date"])
        record["enrolment"] += int(row["Registered Student Count"] or 0)
        lecture, tutorial = section_meetings(row)
        names = tuple(last_name(n) for n in row["Faculty Name"].split(",") if n.strip())
        # MATH 1003 has 24 "sections" but four lectures: a section there is a
        # tutorial group. Sections that share a lecture slot AND an instructor
        # are one lecture.
        lectures[(lecture, names)] = True
        for name in names:
            if name not in record["instructors"]:
                record["instructors"].append(name)
        if tutorial:
            tutorial_sections += 1
            if tutorial not in tutorial_slots:
                tutorial_slots.append(tutorial)
    if record["inperson"]:
        record["lectures"] = len(lectures)
        record["tutorials"] = tutorial_sections
        slots = sorted({slot for slot, _ in lectures if slot})
        record["lecture_slots"] = slots
        record["tutorial_slots"] = sorted(tutorial_slots)
    return record


def offering_from_summary(rows):
    """Summarise one course in one term from the course summaries (no times).

    Used only for terms the timetable exports do not reach. There are no
    meeting times, so the number of lectures is estimated as the number of
    distinct instructors, which is right for everything but a course where
    one person taught two lectures.
    """
    record = new_offering("summary")
    instructor_ids = set()
    for row in rows:
        if row["Academic Level"] != "UG":
            continue
        kind = section_kind(row["Section Number"])
        if kind == "challenge":
            continue
        if kind == "online":
            record["online"] = True
            continue
        record["inperson"] = True
        record["enrolment"] += int(row["Registration Count"] or 0)
        if row["Instructor ID"]:
            instructor_ids.add(row["Instructor ID"])
        # Two people sharing a section arrive as 'Tokaryk,Gegenberg'.
        for name in row["Instructor Last Name"].split(","):
            name = name.strip()
            if name and name not in record["instructors"]:
                record["instructors"].append(name)
    if record["inperson"]:
        record["lectures"] = max(1, len(instructor_ids))
    return record


def new_offering(source):
    return {
        "source": source,
        "inperson": False,
        "online": False,
        "grad_only": False,  # ran under the graduate number only
        "lectures": 0,
        "tutorials": 0,
        "lecture_slots": [],
        "tutorial_slots": [],
        "enrolment": 0,
        "instructors": [],
        "start": "",
    }


def grad_enrolment(summaries, grad_course, term):
    """Students registered in a graduate course in a term (0 if it did not run)."""
    rows = summaries.get((grad_course, term), [])
    return sum(int(r["Registration Count"] or 0) for r in rows if r["Academic Level"] == "GR")


def infer_grad_twins(summaries):
    """Find 4000/6000 pairs like MATH 4443 / MATH 6443.

    A pair is accepted when the two numbers differ only in the first digit AND
    the same instructor taught both in the same term at least twice. Matching
    digits alone is not enough: MATH 4903 and MATH 6903 match, but one is an
    independent study and the other a topics course. Irregular pairs (STAT
    3373 / STAT 6473, say) are not guessed at; they belong in the Cross-listed
    column.
    """
    # (term, instructor) -> the courses that person taught that term
    taught = defaultdict(set)
    for (course, term), rows in summaries.items():
        for row in rows:
            if row["Instructor ID"]:
                taught[(term, row["Instructor ID"])].add(course)
    together = defaultdict(int)
    for courses in taught.values():
        for course in courses:
            grad = course[:-4] + "6" + course[-3:]
            if course[-4] == "4" and grad in courses:
                together[(course, grad)] += 1
    return {course: grad for (course, grad), terms in together.items() if terms >= 2}


def build_offerings(timetable, covered, summaries, summary_terms, aliases, twins):
    """Return {course: {term: offering}} for every course and term we know of.

    aliases maps a course to other numbers whose sections count as its own
    (DATA 1001 is filed by the registrar as STAT 1001). twins maps a course to
    the graduate numbers it is cross-listed with.
    """
    # Which list course does each registrar number belong to?
    owner = {}
    for course, others in aliases.items():
        for other in others:
            owner[other] = course

    grouped_tt = defaultdict(list)
    for (course, term), rows in timetable.items():
        grouped_tt[(owner.get(course, course), term)] += rows
    grouped_sum = defaultdict(list)
    for (course, term), rows in summaries.items():
        grouped_sum[(owner.get(course, course), term)] += rows

    offerings = defaultdict(dict)
    for (course, term), rows in grouped_tt.items():
        offerings[course][term] = offering_from_timetable(rows)
    for (course, term), rows in grouped_sum.items():
        if term in covered:
            continue  # the timetable export is the better source for this term
        if not any(r["Academic Level"] == "UG" for r in rows):
            continue
        offerings[course][term] = offering_from_summary(rows)

    # A cross-listed course that ran with graduate students only never shows
    # up under its undergraduate number. Record it, so "last offered" is right.
    for course, grads in twins.items():
        for term in summary_terms:
            students = sum(grad_enrolment(summaries, g, term) for g in grads)
            existing = offerings[course].get(term)
            if students and not (existing and existing["inperson"]):
                record = existing or new_offering("summary")
                record["grad_only"] = True
                offerings[course][term] = record
    return offerings


# ----------------------------------------------------------------------------
# Reading the previous workbook: only what a person typed
# ----------------------------------------------------------------------------
# Each field can be headed differently in the old layout and the new one.
# Matching is exact and case-sensitive on purpose: the old workbook has both
# "MY NOTES" (Patrick's) and "NOTES" (the calendar's).
HEADINGS = {
    "course": ["Course", "NUMBER"],
    "title": ["Title", "TITLE"],
    "ch": ["CH", "CREDIT HOURS"],
    "format": ["Format", "FORMAT"],
    "other": ["Other numbers"],
    "status": ["Status", "STATUS"],
    "cross": ["Cross-listed", "Cross listed"],
    "staffing": ["Instructor note", "Staffing note"],
    "notes": ["Notes", "MY NOTES", "my NOTES"],
    "check": ["Check"],
    # These two exist only on the removed sheet.
    "lastcal": ["Last calendar"],
    "reason": ["Reason"],
    "description": ["Description", "DESCRIPTION"],
    "prereq": ["Prereq", "PREREQ"],
    "calnotes": ["Calendar notes", "NOTES"],
}


def read_sheet_fields(sheet):
    """Return ({course: {field: text}}, set of fields the sheet has columns for)."""
    rows = sheet.iter_rows(values_only=True)
    header = [str(h).strip() if h is not None else "" for h in next(rows, [])]
    column = {}
    for field, names in HEADINGS.items():
        for name in names:
            if name in header:
                column[field] = header.index(name)
                break
    found = {}
    if "course" not in column:
        return found, set()
    for row in rows:
        key = course_key(row[column["course"]] if column["course"] < len(row) else "")
        if not is_course(key):
            continue  # totals, legends and blank rows
        record = {}
        for field, index in column.items():
            value = row[index] if index < len(row) else None
            if value is None:
                continue
            if isinstance(value, float) and value.is_integer():
                value = int(value)
            record[field] = str(value).strip()
        found[key] = record
    return found, set(column)


def read_previous(path):
    """Read the hand-kept fields from the previous workbook.

    Returns (main, removed): {course: {field: text}} for MAIN and removed.
    """
    # data_only=True returns the values Excel last calculated, not formulas.
    workbook = load_workbook(path, data_only=True)
    main, fields = read_sheet_fields(workbook["MAIN"])
    removed = {}
    if "removed" in workbook.sheetnames:
        removed, _ = read_sheet_fields(workbook["removed"])
    # The new layout keeps the long calendar text on its own sheet.
    if "Calendar" in workbook.sheetnames:
        calendar, _ = read_sheet_fields(workbook["Calendar"])
        for key, record in calendar.items():
            target = main.get(key) or removed.get(key)
            if target is not None:
                for field in ("description", "prereq", "calnotes"):
                    if field in record:
                        target[field] = record[field]
    return main, removed


# ----------------------------------------------------------------------------
# Working out the generated columns for one course
# ----------------------------------------------------------------------------
def season_terms(offerings, season, counts):
    """Terms of one season in which the course ran, oldest first."""
    return sorted((t for t, o in offerings.items() if t.endswith(season) and counts(o)), key=term_key)


def ran_in_class(offering):
    """In person, or under the graduate number with students: it was taught."""
    return offering["inperson"] or offering["grad_only"]


def slot_text(offering):
    """'MWF 11:30-12:20 + tut T 8:30-9:20 / Th 8:30-9:20' for one offering."""
    if not offering["lecture_slots"]:
        return "(no set time)"
    text = " | ".join(offering["lecture_slots"])
    if offering["tutorial_slots"]:
        text += " + tut " + " / ".join(offering["tutorial_slots"])
    return text


def sections_text(offering):
    """'1', '2', or '4 lec, 24 tut' when tutorial groups outnumber lectures."""
    if offering["tutorials"] > offering["lectures"]:
        return f"{offering['lectures']} lec, {offering['tutorials']} tut"
    return str(offering["lectures"])


def academic_year(term):
    """'2025/FA' and '2026/WI' both belong to the academic year 2025 (2025-26)."""
    year = int(term[:4])
    return year if term.endswith("FA") else year - 1


def year_label(year):
    """2025 -> '2025-26'."""
    return f"{year}-{(year + 1) % 100:02d}"


def complete_years(terms, count):
    """The last `count` academic years whose fall AND winter are both in `terms`.

    A year whose winter is not in the data yet is left out, or a course that
    runs every winter would look as if it had skipped that year.
    """
    present = set(terms)
    years = sorted({academic_year(t) for t in present if t.endswith("FA")
                    and f"{academic_year(t) + 1}/WI" in present})
    return years[-count:]


def season_letters(letters):
    """{'FA', 'WI'} -> 'F and W'; {'WI'} -> 'W'."""
    return " and ".join(SEASON_LETTER[s] for s in ("FA", "WI") if s in letters)


def offering_note(offerings, window):
    """Say in one cell how regularly, and in which term, a course runs.

    window is the academic years to judge by, oldest first (see
    complete_years). Fall and winter are taken together: a course counts as
    offered in 2025-26 if it was taught in Fall 2025 or Winter 2026.

    Blank means every year, always in the same term (or terms). Otherwise:
      'every year, F or W'          every year, but the term moves about
      'every year; W since 2025-26' every year, and moved term recently
      'every year since 2024-25'    an unbroken recent run
      'alternate years; last 2027 W'
      '3 of last 6 years; last 2025 F'
      'not since 2016 W'            skipped the last two years running
      '-'                           never taught in a fall or winter

    The wording describes the current RUN rather than counting, so that a
    deliberate change shows up after two years rather than six.
    """
    ran = sorted((t for t, o in offerings.items()
                  if not t.endswith("SM") and ran_in_class(o)), key=term_key)
    if not ran:
        return "-"
    last = ran[-1]
    label = term_label(last) + (" (grad only)" if not offerings[last]["inperson"] else "")
    # For each academic year, the seasons the course was taught in.
    seasons = defaultdict(set)
    for term in ran:
        seasons[academic_year(term)].add(term[5:])
    flags = [year in seasons for year in window]

    # How many of the most recent years in a row was it taught?
    run = 0
    for flag in reversed(flags):
        if not flag:
            break
        run += 1

    if all(flags):
        patterns = [frozenset(seasons[year]) for year in window]
        if len(set(patterns)) == 1:
            return ""  # the unremarkable case, left blank on purpose
        # The same term(s) for the last few years running: a move.
        same = 0
        for pattern in reversed(patterns):
            if pattern != patterns[-1]:
                break
            same += 1
        if same >= 2:
            return f"every year; {season_letters(patterns[-1])} since {year_label(window[-same])}"
        return "every year, F or W"
    if not any(flags) or not any(flags[-2:]):
        # Never in the window, or skipped the last two years running.
        return f"not since {label}"
    if run >= 2:
        return f"every year since {year_label(window[-run])}"
    alternating = all(a != b for a, b in zip(flags, flags[1:]))
    if alternating and len(window) >= 4:
        return f"alternate years; last {label}"
    return f"{sum(flags)} of last {len(window)} years; last {label}"


def season_columns(offerings, season, window, today):
    """The three generated cells for fall or for winter: slot, sections, enrolment."""
    ran = season_terms(offerings, season, ran_in_class)
    if not ran:
        return ["-", "", ""]
    in_person = season_terms(offerings, season, lambda o: o["inperson"])
    # The slot comes from the most recent term we have meeting times for.
    timed = [t for t in in_person if offerings[t]["source"] == "timetable"]
    slot = sections = ""
    if timed:
        sections = sections_text(offerings[timed[-1]])
        slot = "(no set time)"
    # Prefer an offering whose time is actually set. A section for next term
    # with no time yet, or a one-off reading course, should not hide the slot
    # the course normally sits in.
    slotted = [t for t in timed if offerings[t]["lecture_slots"]]
    if slotted:
        latest = offerings[slotted[-1]]
        slot = slot_text(latest)
        # Walk back while the lecture slot is unchanged, to find when the
        # course moved into this slot. Only a RECENT move is worth a tag: it
        # tells the reader the slot is not yet one to rely on.
        first_same = slotted[-1]
        for term in reversed(slotted):
            if offerings[term]["lecture_slots"] != latest["lecture_slots"]:
                break
            first_same = term
        moved = first_same != slotted[0]  # there was a different slot before
        if moved and first_same in window[-2:]:
            slot += f" (new {term_label(first_same)})"
    # Typical size: the median of the last three offerings that have started.
    # A term still open for registration would drag the figure down.
    started = [t for t in in_person if not offerings[t]["start"] or offerings[t]["start"] <= today]
    sizes = [offerings[t]["enrolment"] for t in started[-3:]]
    enrolment = round(statistics.median(sizes)) if sizes else ""
    return [slot, sections, enrolment]


def summer_text(offerings, covered):
    """In-person summer offerings, e.g. '8 of 10 summers; last 2026 S'."""
    summers = sorted((t for t in covered if t.endswith("SM")), key=term_key)
    ran = [t for t in summers if t in offerings and offerings[t]["inperson"]]
    if not ran:
        return "-"
    return f"{len(ran)} of {len(summers)} summers; last {term_label(ran[-1])}"


def online_text(offerings, latest_term):
    """'yes' if offered online in the last two years, else when it last was."""
    ran = sorted((t for t, o in offerings.items() if o["online"]), key=term_key)
    if not ran:
        return "-"
    recent = int(latest_term[:4]) - int(ran[-1][:4]) <= 1
    return "yes" if recent else f"last {term_label(ran[-1])}"


def instructors_text(offerings, today, limit=5):
    """Who taught the last three in-person fall/winter offerings, newest first."""
    terms = sorted(
        (t for t, o in offerings.items() if o["inperson"] and not t.endswith("SM")),
        key=term_key,
        reverse=True,
    )
    names = []
    for term in terms[:3]:
        for name in offerings[term]["instructors"]:
            if name not in names:
                names.append(name)
    extra = f" +{len(names) - limit}" if len(names) > limit else ""
    return ", ".join(names[:limit]) + extra


def grad_numbers(cross_text):
    """Graduate course numbers mentioned in a Cross-listed cell."""
    found = re.findall(r"(MATH|STAT|DATA)\s*(6\d{3})", cross_text or "")
    return [subject + number for subject, number in found]


# ----------------------------------------------------------------------------
# Who else names each course ("Required by others")
# ----------------------------------------------------------------------------
# Only the short list of units is kept here. Which program, and whether the
# course is required, one of a list or an option, is in the department's
# reference, curriculum/COURSES-required-by-other-units.md, which a person
# keeps; a script cannot judge that reliably.
#
# A program page is turned into a unit by the first rule below whose text
# appears in the page's path. "" means the page does not count: our own
# programs.
PAGE_UNITS = [
    ("mathematicsstatisticsoption", ""),
    ("bachelorofscienceinengineering", "Engineering"),
    ("softwarenegineering", "Engineering"),  # the calendar's own spelling
    ("bachelorofgeomatics", "Engineering"),
    ("computerscience", "CS"),
    ("concurrent-bscbsc", "CS"),
    ("bachelorofbusinessadministration", "Management"),
    ("biologyoption", "Biology"),
    ("chemistryoption", "Chemistry"),
    ("earthsciencesoption", "Earth Sciences"),
    ("environmentalgeosciencesoption", "Earth Sciences"),
    ("economicsoption", "Economics"),
    ("physicsoption", "Physics"),
    ("psychologyoption", "Psychology"),
    ("medicallaboratoryscience", "Med Lab Science"),
    ("foundationsofbiomedicalsciences", "Biomedical Sciences cert"),
    ("bachelorofscience/index", "BSc first year"),
    ("kinesiology", "Kinesiology"),
    ("recreationandsport", "Recreation and Sport"),
    ("nursing", "Nursing"),
    ("forestry", "Forestry and EM"),
    ("environmentalmanagement", "Forestry and EM"),
    ("bachelorofarts", "Arts"),
]
# The interdepartmental page holds many programs; the second part of the
# section heading names the program, and each program names its units.
INTERDEPARTMENTAL_UNITS = {
    "Biology-Chemistry": ["Biology", "Chemistry"],
    "Biology-Physics": ["Biology", "Physics"],
    "Biology-Psychology": ["Biology", "Psychology"],
    "Chemistry-Physics": ["Chemistry", "Physics"],
    "Earth Sciences-Physics": ["Earth Sciences", "Physics"],
    "Environmental Sciences": ["Environmental Sciences"],
}
# A line of program text that names our courses only to rule them out, as in
# "Credit is not given toward the BCS degree for ... MATH 1843" or "Service
# courses ... are not acceptable for BBA credit: STAT 2263 ...", does not count.
RULED_OUT = re.compile(r"not (?:be )?given|not acceptab|cannot receive credit|will not be approved",
                       re.I)
# Another unit's course counts only when ours is its prerequisite or
# co-requisite, not when the entry merely excludes or recommends ours.
COURSE_ROLES = {"prerequisite", "co-requisite"}
SUBJECT_UNITS = {
    "ADM": "Management",
    "APSC": "Engineering", "CE": "Engineering", "CHE": "Engineering", "ECE": "Engineering",
    "ENGG": "Engineering", "GGE": "Engineering", "ME": "Engineering", "SWE": "Engineering",
    "BIOL": "Biology", "CHEM": "Chemistry", "CS": "CS", "ECON": "Economics",
    "ESCI": "Earth Sciences", "KIN": "Kinesiology", "NURS": "Nursing", "PHYS": "Physics",
    "ENVM": "Forestry and EM", "FOR": "Forestry and EM",
}


def latest_sweep(calendar_dir):
    """The newest calendar-year folder that holds sweep_references.py's output."""
    found = sorted(p.parent for p in Path(calendar_dir).glob("*/sweep-programs.tsv"))
    return found[-1] if found else None


def program_units(page, section):
    """The units a program page (and section heading) belongs to; [] if it does not count."""
    if page.endswith("interdepartmentalprograms"):
        # 'Interdepartmental programs > Biology-Physics > Requirements ...'
        parts = [part.strip() for part in section.split(">")]
        program = parts[1] if len(parts) > 1 else ""
        return INTERDEPARTMENTAL_UNITS.get(program, [])
    for text, unit in PAGE_UNITS:
        if text in page:
            return [unit] if unit else []
    return [None]  # a page no rule knows: reported, so a rule can be added


def required_by_others(sweep_dir):
    """{course key: 'Biology, Physics, ...'} from one year's sweep, and the pages no rule knew."""
    units = defaultdict(set)
    unknown = set()
    with open(sweep_dir / "sweep-programs.tsv", newline="", encoding="utf-8") as handle:
        for row in csv.DictReader(handle, delimiter="\t"):
            if RULED_OUT.search(row["Text"]):
                continue
            for unit in program_units(row["Program page"], row["Section"]):
                if unit is None:
                    unknown.add(row["Program page"])
                else:
                    units[course_key(row["Cited course"])].add(unit)
    with open(sweep_dir / "sweep-courses.tsv", newline="", encoding="utf-8") as handle:
        for row in csv.DictReader(handle, delimiter="\t"):
            if row["Role"] in COURSE_ROLES:
                subject = row["Their course"].split()[0]
                units[course_key(row["Cited course"])].add(SUBJECT_UNITS.get(subject, subject))
    return {key: ", ".join(sorted(names)) for key, names in units.items()}, sorted(unknown)


# ----------------------------------------------------------------------------
# Writing the workbook
# ----------------------------------------------------------------------------
# (heading, width, group). The group decides the header colour, and "hand"
# marks the columns that are typed and carried across, never generated.
MAIN_COLUMNS = [
    ("Course", 11, "id"), ("Title", 34, "id"), ("CH", 5, "id"), ("Format", 11, "id"),
    ("Other numbers", 13, "hand"),
    ("Offering note", 30, "other"),
    ("Fall timeslot", 34, "fall"), ("Fall sections", 12, "fall"), ("Fall enrol", 9, "fall"),
    ("Winter timeslot", 34, "winter"), ("Winter sections", 12, "winter"), ("Winter enrol", 9, "winter"),
    ("Summer (in person)", 20, "other"), ("Online", 11, "other"), ("Recent instructors", 28, "other"),
    ("Status", 10, "hand"), ("Required by others", 48, "other"), ("Cross-listed", 16, "hand"),
    ("Notes", 44, "hand"),
]
# The removed sheet is a record, not a planning tool, so it has none of MAIN's
# timetable columns. Only "Last taught" is generated; the rest is typed.
REMOVED_COLUMNS = [
    ("Course", 11, "id"), ("Title", 34, "id"), ("CH", 5, "id"), ("Format", 11, "id"),
    ("Last calendar", 16, "hand"), ("Last taught", 26, "other"),
    ("Reason", 60, "hand"), ("Notes", 60, "hand"),
]
FILLS = {
    "id": "D9D9D9", "fall": "F8CBAD", "winter": "BDD7EE", "other": "E2EFDA", "hand": "FFF2CC",
}

# Explanations shown as an Excel note (the little red corner) on a header,
# for headings that would be too long if they explained themselves.
HEADER_NOTES = {
    "Fall enrol": "Median of the last 3 offerings.",
    "Winter enrol": "Median of the last 3 offerings.",
    "Required by others": "Units whose calendar text names this course in a program, or as a "
                          "prerequisite or co-requisite. Whether it is required, one of a list or "
                          "an option: see curriculum/COURSES-required-by-other-units.md.",
}

# Each sheet is declared as an Excel table. None means no table style, so the
# sheet keeps the header colours above and nothing else. To try one of Excel's
# own styles (banded rows and so on), put its name here, e.g. "TableStyleLight1".
TABLE_STYLE = None

# Rows whose Status says the course is not currently running are greyed out.
# (text colour for the whole row, fill colour for the Status cell itself)
STATUS_COLOURS = {
    "hiatus": ("9C5700", "FFEB9C"),    # amber: paused, expected back
    "inactive": ("808080", "D9D9D9"),  # grey: not expected back soon
}


def shade_by_status(sheet, columns):
    """Colour the rows on hiatus or inactive, using conditional formatting.

    Conditional formatting (rather than colouring the cells here) means the
    colour follows the Status cell: type 'hiatus' in Excel and the row changes
    at once, without a rebuild. Excel's = ignores case, so HIATUS matches too.
    """
    headings = [heading for heading, _, _ in columns]
    status = get_column_letter(headings.index("Status") + 1)
    last = get_column_letter(len(columns))
    if sheet.max_row < 2:
        return  # no data rows to colour
    for word, (text_colour, fill_colour) in STATUS_COLOURS.items():
        # $Q2: the $ pins the column, so every cell in a row tests that row's Status.
        test = [f'TRIM(${status}2)="{word}"']
        sheet.conditional_formatting.add(
            f"A2:{last}{sheet.max_row}",
            FormulaRule(formula=test, font=Font(color=text_colour)),
        )
        sheet.conditional_formatting.add(
            f"{status}2:{status}{sheet.max_row}",
            FormulaRule(formula=test, fill=PatternFill("solid", bgColor=fill_colour)),
        )


def declare_table(sheet):
    """Tell Excel that the cells just written are a table (Insert > Table).

    Without this the sheet only looks like a table. Declaring it gives the
    Table tab in the ribbon, the table-aware Insert/Delete menu, and a table
    that grows by itself when a row is typed underneath it.
    """
    last = get_column_letter(sheet.max_column)
    area = f"A1:{last}{sheet.max_row}"
    if sheet.max_row < 2:
        # Excel refuses a table that has a header and no data rows, so an
        # empty sheet gets plain filter arrows instead.
        sheet.auto_filter.ref = area
        return
    # A table's name must be unique in the workbook and may not contain
    # spaces, so it is made from the sheet's name: tblMAIN, tblHistory, ...
    name = "tbl" + re.sub(r"\W", "", sheet.title)
    table = Table(displayName=name, ref=area)
    if TABLE_STYLE:
        table.tableStyleInfo = TableStyleInfo(name=TABLE_STYLE, showRowStripes=True)
    # The filter arrows now belong to the table. sheet.auto_filter must NOT be
    # set as well: with both, Excel says the file is damaged and repairs it.
    sheet.add_table(table)


def write_table(sheet, columns, rows, freeze):
    """Write a header row and data rows as an Excel table, with widths and frozen panes."""
    # Excel insists that every heading of a table is filled in and that no two
    # are the same (it ignores case). Stop here with a clear message, rather
    # than write a workbook that Excel would have to repair.
    headings = [str(heading).strip().lower() for heading, _, _ in columns]
    if "" in headings or len(set(headings)) != len(headings):
        sys.exit(f"Sheet {sheet.title}: every column needs its own, non-blank heading.")
    for index, (heading, width, group) in enumerate(columns, start=1):
        cell = sheet.cell(row=1, column=index, value=heading)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", start_color=FILLS[group])
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        sheet.column_dimensions[get_column_letter(index)].width = width
        if heading in HEADER_NOTES:
            cell.comment = Comment(HEADER_NOTES[heading], "build_courses.py")
    for row in rows:
        sheet.append(row)
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    sheet.freeze_panes = freeze
    declare_table(sheet)


def main_row(key, hand, offerings, windows, years, covered, latest_term, today, required):
    """One row of MAIN (or of removed): generated cells around the typed ones."""
    return (
        [course_display(key), hand.get("title", ""), hand.get("ch", ""), hand.get("format", ""),
         hand.get("other", ""), offering_note(offerings, years)]
        + season_columns(offerings, "FA", windows["FA"], today)
        + season_columns(offerings, "WI", windows["WI"], today)
        + [summer_text(offerings, covered), online_text(offerings, latest_term),
           instructors_text(offerings, today)]
        + [hand.get("status", ""), required, hand.get("cross", ""), hand.get("notes", "")]
    )


def removed_row(key, hand, offerings, first_term):
    """One row of the removed sheet."""
    # Every term the course was taught, in any season, oldest first.
    ran = sorted((t for t, o in offerings.items() if ran_in_class(o)), key=term_key)
    if ran:
        names = instructors_cell(offerings[ran[-1]])
        last_taught = term_label(ran[-1]) + (f" ({names})" if names else "")
    else:
        # The registrar's data only reaches back so far; say how far.
        last_taught = f"before {term_label(first_term)}"
    return [course_display(key), hand.get("title", ""), hand.get("ch", ""), hand.get("format", ""),
            hand.get("lastcal", ""), last_taught, hand.get("reason", ""), hand.get("notes", "")]


def history_cell(offering):
    """Lectures taught in person; 'G' if it ran for graduate students only."""
    if offering is None:
        return None
    if offering["inperson"]:
        return offering["lectures"]
    return "G" if offering["grad_only"] else None


def instructors_cell(offering):
    """Who taught it in person that term, e.g. 'Reynolds, Allen, Lissner'."""
    if offering is None:
        return None
    if offering["inperson"]:
        return ", ".join(offering["instructors"]) or "(not recorded)"
    return "(grad only)" if offering["grad_only"] else None


def backup_existing(output):
    """Copy the workbook about to be overwritten, and return where the copy went.

    The workbook holds typed columns that exist nowhere else, so it is never
    overwritten without a copy. There is one copy, NAME-bak.xlsx right beside
    it where it will not be forgotten, replaced at every run; anything older
    is in OneDrive's version history.
    """
    if not output.exists():
        return None
    backup = output.with_name(f"{output.stem}-bak{output.suffix}")
    shutil.copy2(output, backup)  # copy2 keeps the file's modified date
    return backup


def copy_for_another_team(original, copy):
    """Copy a finished file into another team's folder, as its -DO-NOT-EDIT copy.

    Returns the copy's path, or None when that folder is missing (a trial
    run on another machine, say), which is reported rather than fatal.
    """
    if not copy.parent.is_dir():
        print(f"  No folder {copy.parent}, so no copy was made there.")
        return None
    shutil.copy2(original, copy)  # copy2 keeps the file's modified date
    return copy


def refuse_if_open(output):
    """Stop if Excel has the workbook open.

    Excel leaves a '~$name.xlsx' file beside a workbook it has open. Writing
    underneath an open workbook invites Excel to save its old copy back on top.
    """
    lock = output.with_name("~$" + output.name)
    if lock.exists():
        sys.exit(f"{output.name} looks open in Excel ({lock.name} exists). Close it and run again.")


def save_safely(workbook, output):
    """Back up whatever is at output, then save there. Returns the backup's path."""
    backup = backup_existing(output)
    # Save under a temporary name and then swap it in, so a crash part-way
    # through saving cannot leave a half-written workbook.
    temporary = output.with_name(output.stem + ".tmp.xlsx")
    workbook.save(temporary)
    os.replace(temporary, output)
    return backup


def main():
    parser = argparse.ArgumentParser(
        description="Build the MATH/STAT master course list from Informer exports."
    )
    parser.add_argument("-o", "--output", default=DEFAULT_OUTPUT,
                        help="workbook to write (default: the master list itself)")
    parser.add_argument("-p", "--previous",
                        help="workbook to read the typed columns from (default: the output "
                             "workbook, so the master list is rebuilt in place)")
    parser.add_argument("-t", "--timetable", action="append",
                        help="timetable export or directory of them (repeatable)")
    parser.add_argument("-s", "--summaries", default=DEFAULT_SUMMARIES,
                        help="directory of course-summary exports")
    parser.add_argument("--years", type=int, default=6,
                        help="academic years the offering note looks back (default 6)")
    parser.add_argument("--sweep",
                        help="calendar-year folder holding the sweep output (default: the newest)")
    args = parser.parse_args()

    output = Path(args.output).expanduser()
    # With no --previous the workbook is rebuilt in place: its typed columns
    # are read, then the whole file is written afresh around them.
    previous = Path(args.previous).expanduser() if args.previous else output
    if not previous.exists():
        sys.exit(f"Cannot find {previous}.\n"
                 "Give --previous the workbook that holds the typed columns.")
    refuse_if_open(output)

    today = datetime.date.today().isoformat()
    timetable, covered, tt_files = load_timetables(args.timetable or DEFAULT_TIMETABLES)
    summaries, summary_terms, sum_files = load_summaries(args.summaries)
    main_hand, removed_hand = read_previous(previous)

    # "Instructor note" and "Check" were dropped on 2026-10-14. A workbook
    # that still has them keeps their text by folding it into Notes; the
    # automatic "NEW: ..." flags in Check are simply dropped.
    for hand in main_hand.values():
        extra = [hand.pop("staffing", "")]
        check = hand.pop("check", "")
        if not check.startswith("NEW:"):
            extra.append(check)
        for text in extra:
            if text and text not in hand.get("notes", ""):
                hand["notes"] = f"{hand['notes']}; {text}" if hand.get("notes") else text

    sweep_dir = Path(args.sweep).expanduser() if args.sweep else latest_sweep(CALENDAR_DIR)
    if sweep_dir:
        required, unknown_pages = required_by_others(sweep_dir)
    else:
        required, unknown_pages = {}, []

    # Other numbers a course is filed under (typed in "Other numbers").
    aliases = {
        key: [course_key(c) for c in re.findall(r"[A-Z]{4}\s*\d{4}", hand.get("other", ""))]
        for key, hand in {**removed_hand, **main_hand}.items()
    }
    aliased = {other for others in aliases.values() for other in others}

    # Courses the registrar has undergraduate sections for that the list has
    # never held. They are added as rows and flagged, for Patrick to judge.
    seen = {}  # course -> (title, credit hours), the latest we come across
    for (key, term), rows in sorted(summaries.items(), key=lambda kv: term_key(kv[0][1])):
        if any(r["Academic Level"] == "UG" and section_kind(r["Section Number"]) != "challenge"
               for r in rows):
            seen.setdefault(key, ("", ""))
    for (key, term), rows in sorted(timetable.items(), key=lambda kv: term_key(kv[0][1])):
        for row in rows:
            if section_kind(row["Number"], row["Instr Method"]) != "challenge":
                seen[key] = (row["Course Title"], row["Credit Hours"])
    new_courses = []
    for key in sorted(seen):
        known = key in main_hand or key in removed_hand or key in aliased
        # Graduate numbers are not listed in their own right.
        if known or not is_course(key) or int(key[-4]) >= 6:
            continue
        title, hours = seen[key]
        main_hand[key] = {
            "title": title,
            "ch": str(int(float(hours))) if hours else "",
            "notes": f"Added by build_courses.py on {today}: it has sections in Informer "
                     "but was not on the list. Delete this note once checked.",
        }
        new_courses.append(key)

    # The graduate twins each course shares a classroom with: whatever is
    # typed in Cross-listed, plus the safe inference.
    everything = {**removed_hand, **main_hand}
    inferred = infer_grad_twins(summaries)
    twins = {}
    for key, hand in everything.items():
        typed = grad_numbers(hand.get("cross", ""))
        guess = inferred.get(key)
        if guess and guess not in typed:
            # Shown in the cell so Patrick can confirm or correct it.
            note = f"{course_display(guess)} (inferred)"
            hand["cross"] = f"{hand['cross']}; {note}" if hand.get("cross") else note
            typed.append(guess)
        if typed:
            twins[key] = typed

    offerings = build_offerings(timetable, covered, summaries, summary_terms, aliases, twins)

    # The offering note judges by the last N academic years; the timeslot's
    # "(new ...)" tag still looks at the last falls or winters on their own.
    all_terms = sorted({t for per in offerings.values() for t in per}, key=term_key)
    latest_term = all_terms[-1]
    years = complete_years(all_terms, args.years)
    windows = {}
    for season in ("FA", "WI"):
        seasonal = [t for t in all_terms if t.endswith(season)]
        windows[season] = seasonal[-args.years:]

    workbook = Workbook()

    # ---- MAIN ---------------------------------------------------------------
    sheet = workbook.active
    sheet.title = "MAIN"
    main_keys = sorted(main_hand, key=sort_key)

    def required_cell(key):
        """Units naming this course under its own number or any of its other numbers."""
        names = set()
        for number in [key] + aliases.get(key, []):
            if number in required:
                names.update(required[number].split(", "))
        return ", ".join(sorted(names))

    write_table(
        sheet, MAIN_COLUMNS,
        [main_row(k, main_hand[k], offerings.get(k, {}), windows, years, covered, latest_term, today,
                  required_cell(k))
         for k in main_keys],
        freeze="C2",
    )
    shade_by_status(sheet, MAIN_COLUMNS)

    # ---- History------------------------------------------------------------
    sheet = workbook.create_sheet("History")
    newest_first = list(reversed(all_terms))
    columns = [("Course", 11, "id"), ("Title", 34, "id")]
    for term in newest_first:
        group = {"FA": "fall", "WI": "winter", "SM": "other"}[term[5:]]
        columns.append((term_label(term), 7, group))
    rows = []
    for key in main_keys:
        per = offerings.get(key, {})
        rows.append([course_display(key), main_hand[key].get("title", "")]
                    + [history_cell(per.get(term)) for term in newest_first])
    write_table(sheet, columns, rows, freeze="C2")

    # ---- Instructors --------------------------------------------------------
    # The same grid as History, holding names instead of counts. Kept apart so
    # History stays numbers that can be read at a glance, summed and filtered.
    sheet = workbook.create_sheet("Instructors")
    columns = [("Course", 11, "id"), ("Title", 34, "id")]
    for term in newest_first:
        group = {"FA": "fall", "WI": "winter", "SM": "other"}[term[5:]]
        columns.append((term_label(term), 18, group))
    rows = []
    for key in main_keys:
        per = offerings.get(key, {})
        rows.append([course_display(key), main_hand[key].get("title", "")]
                    + [instructors_cell(per.get(term)) for term in newest_first])
    write_table(sheet, columns, rows, freeze="C2")

    # ---- Calendar -----------------------------------------------------------
    sheet = workbook.create_sheet("Calendar")
    columns = [("Course", 11, "id"), ("Title", 34, "id"), ("Description", 80, "hand"),
               ("Prereq", 45, "hand"), ("Calendar notes", 45, "hand")]
    rows = [[course_display(k), main_hand[k].get("title", ""), main_hand[k].get("description", ""),
             main_hand[k].get("prereq", ""), main_hand[k].get("calnotes", "")] for k in main_keys]
    rows += [[course_display(k), h.get("title", ""), h.get("description", ""),
              h.get("prereq", ""), h.get("calnotes", "")]
             for k, h in sorted(removed_hand.items(), key=lambda kv: sort_key(kv[0]))]
    write_table(sheet, columns, rows, freeze="C2")

    # ---- removed ------------------------------------------------------------
    # Courses taken out of the calendar: kept on record, out of MAIN's way.
    sheet = workbook.create_sheet("removed")
    removed_keys = sorted(removed_hand, key=sort_key)
    write_table(
        sheet, REMOVED_COLUMNS,
        [removed_row(k, removed_hand[k], offerings.get(k, {}), all_terms[0]) for k in removed_keys],
        freeze="C2",
    )

    # ---- About--------------------------------------------------------------
    sheet = workbook.create_sheet("About")
    covered_sorted = sorted(covered, key=term_key)
    about = [
        ("Built", f"{today} by build_courses.py"),
        ("Timetable exports", f"{len(tt_files)} files, {term_label(covered_sorted[0])} to "
                              f"{term_label(covered_sorted[-1])} (slots, sections, enrolment, instructors)"),
        ("Course summaries", f"{len(sum_files)} files, {term_label(min(summary_terms, key=term_key))} to "
                             f"{term_label(max(summary_terms, key=term_key))} (earlier history, graduate twins)"),
        ("Typed columns carried from", str(previous)),
        ("", ""),
        ("Yellow headers", "Typed by a person and carried across unchanged on every rebuild. "
                           "Edit these freely."),
        ("Other headers", "Generated from the registrar's data. Anything typed there is lost on "
                          "the next rebuild."),
        ("", ""),
        ("Timeslot", "Lecture slot at the most recent in-person offering in that season. "
                 "'|' separates lectures at different times; '+ tut' lists the tutorial options. "
                 "'-' means never offered in that season. '(new 2026 F)' means the lecture slot "
                 "changed that term, so it is not yet one to rely on."),
        ("Sections", "Lectures at the most recent offering. '4 lec, 24 tut' when students register "
                     "in tutorial groups that outnumber the lectures."),
        ("Offering note", f"BLANK MEANS EVERY YEAR IN THE SAME TERM, judged over the academic years "
                          f"{year_label(years[0])} to {year_label(years[-1])}; fall and winter count "
                          "together. Otherwise: 'every year, F or W' when the term moves about, "
                          "'every year; W since ...' after a move, 'every year since ...' for a recent "
                          "unbroken run, 'not since ...' once two years running are skipped, or a "
                          "pattern with the last term taught. '(grad only)' means it ran under its "
                          "graduate number with no undergraduate section."),
        ("Enrol", "Median total registration over the last three offerings that have started."),
        ("Summer (in person)", "Online sections do not count. Summers are covered from 2017."),
        ("Online", "'yes' if offered online in the last two years. Online is never counted as a "
                   "fall, winter or summer offering."),
        ("Status", "Rows marked hiatus or inactive are coloured automatically; the colour "
                   "follows whatever is typed in the Status cell."),
        ("Required by others", f"From the calendar sweep in {sweep_dir.name if sweep_dir else '(none found)'}: "
                               "the units whose calendar text names the course in a program, or as a "
                               "prerequisite or co-requisite of their own course. Our own programs, and "
                               "lines that say a course is NOT accepted for credit, do not count. For which program "
                               "and in what role, see curriculum/COURSES-required-by-other-units.md."),
        ("removed", "Courses taken out of the calendar. 'Last calendar' is the last calendar year "
                    "with an entry for the course; 'Reason' is why it went, quoted from the "
                    "department's submission where one was found. 'Last taught' is generated: the "
                    "most recent term with a section, and who taught it. To bring a course back, "
                    "type its row on MAIN and delete it here."),
        ("History", "Lectures taught in person each term. G = graduate number only. Blank = not "
                    "taught in person. Before 2017 the count is the number of instructors."),
        ("Instructors", "The same grid as History, with the names of who taught in person."),
        ("Other numbers", "Numbers the registrar files this course under. Their sections are "
                          "counted as this course's."),
        ("Cross-listed", "A graduate number here (e.g. MATH 6443) makes graduate-only offerings "
                         "count. '(inferred)' was guessed from matching digits; delete the word "
                         "once confirmed."),
    ]
    sheet.column_dimensions["A"].width = 26
    sheet.column_dimensions["B"].width = 110
    for label, text in about:
        sheet.append([label, text])
    for row in sheet.iter_rows():
        row[0].font = Font(bold=True)
        row[1].alignment = Alignment(wrap_text=True, vertical="top")

    backup = save_safely(workbook, output)

    print(f"Wrote {output}")
    if backup:
        print(f"  The copy it replaced was kept as {backup}")
    # Advisors read a copy. Only the real master list is copied, never a trial one.
    if output.resolve() == Path(DEFAULT_OUTPUT).resolve():
        if copy := copy_for_another_team(output, paths.MASTER_LIST_COPY):
            print(f"  Copied for the Advising team to {copy}")
    print(f"  MAIN: {len(main_keys)} courses ({len(new_courses)} new), removed: {len(removed_keys)}")
    print(f"  History: {term_label(all_terms[0])} to {term_label(all_terms[-1])}, {len(all_terms)} terms")
    if sweep_dir:
        print(f"  Required by others: from the {sweep_dir.name} calendar sweep")
    else:
        print(f"  Required by others: left blank, no sweep found under {CALENDAR_DIR}")
    if unknown_pages:
        print("  Program pages with no unit in PAGE_UNITS (add a rule for each):")
        print("    " + ", ".join(unknown_pages))
    if new_courses:
        print("  New courses, flagged in the Notes column:")
        print("    " + ", ".join(course_display(k) for k in new_courses))
    guessed = [course_display(k) for k in sorted(inferred, key=sort_key) if k in everything]
    if guessed:
        print("  Graduate twins inferred for: " + ", ".join(guessed))


if __name__ == "__main__":
    main()
