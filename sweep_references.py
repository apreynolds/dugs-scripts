# /// script
# dependencies = []
# ///
# sweep_references.py — find every mention of a MATH/STAT course in a mirrored calendar.
# Usage:
#   uv run sweep_references.py YEARDIR
#       e.g. uv run sweep_references.py "$HOME/OneDrive - University of New Brunswick/TEAM-CURRICULUM/General/academic-calendar/2026-27"
#   uv run sweep_references.py YEARDIR --subjects MATH STAT DATA
#   uv run sweep_references.py -h | --help
#
#   YEARDIR       a folder written by mirror_calendar.py (it reads YEARDIR/text)
#   --subjects    the subjects to look for (default: MATH STAT DATA)
#
# It writes three tab-separated files into YEARDIR, to open in a spreadsheet:
#
#   sweep-by-course.tsv  one row per course of ours: which other units'
#                        programs and courses name it. The quick answer to
#                        "who else depends on MATH 2513?".
#   sweep-programs.tsv   one row per (program page, section, cited course):
#                        where a PROGRAM REGULATION names one of our courses.
#   sweep-courses.tsv    one row per (other unit's course, cited course):
#                        where ANOTHER UNIT'S COURSE names one of ours. Its
#                        "Role" column says how: prerequisite, co-requisite,
#                        exclusion ("credit for only one of"), recommended,
#                        cross-listed, or unclear. The role is read from the
#                        wording on the line, not judged by a person.
#
# Both have an "In calendar" column: "no" means the cited number has no course
# entry in this calendar, i.e. the text refers to a course that does not exist.
#
# It also prints, for the pages given by --check, every cited course of ANY
# subject that has no entry: the check for our own program regulations.

import argparse
import re
import sys
from pathlib import Path

# A course entry's first line in a department page: 'MATH3003 | Applied Analysis | 3 ch (3C)'.
ENTRY = re.compile(r"^([A-Z]{2,6})\s?(\d{4})\s*\|\s*([^|]*)")

# A subject followed by one or more numbers: 'MATH 1003', 'MATH 1003/1013',
# 'STAT 3083 and 3093', 'MATH 2003, 2013'. The shorthand forms matter: a plain
# search for 'MATH 1013' would miss the second course in each of them.
# A second subject can sit between the first and the number, as in
# 'STAT(DATA) 1001' or 'MATH/LING 4903'; the number then counts for both.
SUBJECT = r"[A-Z]{2,6}"
CITED = re.compile(
    r"\b(" + SUBJECT + r")(?:[(/](" + SUBJECT + r")\)?)?"
    r"[\s*]?(\d{4})((?:\s*(?:/|,|,?\s+and|,?\s+or|&)\s*\(?\d{4}\)?(?!\s*ch\b))*)"
)

# Our own program regulations: pages where every cited course is checked.
DEFAULT_CHECK = [
    "frederictonprograms/bachelorofscience/mathematicsstatisticsoption.txt",
    "frederictonprograms/bachelorofscience/interdepartmentalprograms.txt",
    "frederictonprograms/bachelorofscience/economicsoption.txt",
    "frederictonprograms/bachelorofscience/index.txt",
]


def cited_courses(line):
    """Every course a line of text names, as 'MATH 1003' strings, in order."""
    found = []
    for subject, second, first, more in CITED.findall(line):
        for number in [first] + re.findall(r"\d{4}", more):
            for name in [subject, second]:
                code = f"{name} {number}"
                if name and code not in found:
                    found.append(code)
    return found


# Words that say WHY a course entry names another course. The role of a cited
# course is decided by the last of these that appears before it on the line:
# 'Credit can be obtained for only one of CS 2545, STAT 1001. Prerequisite:
# CS 1073' makes STAT 1001 an exclusion and CS 1073 a prerequisite.
ROLE_WORDS = [
    ("prerequisite", re.compile(r"pre-?requisites?", re.I)),
    ("co-requisite", re.compile(r"co-?requisites?", re.I)),
    ("exclusion", re.compile(r"credit\s+(?:can|cannot|may|will|for)\b|only one of|equivalent to", re.I)),
    ("recommended", re.compile(r"recommended", re.I)),
    ("cross-listed", re.compile(r"cross-?\s?listed", re.I)),
]


def cited_roles(line):
    """{course: role} for a line of a course entry; role is 'unclear' if no role word precedes it."""
    roles = {}
    for match in CITED.finditer(line):
        before = line[:match.start()]
        role, where = "unclear", -1
        for name, pattern in ROLE_WORDS:
            for hit in pattern.finditer(before):
                # 'co-requisite' contains no 'prerequisite', but check the
                # longer word wins if both end at the same place.
                if hit.end() > where:
                    role, where = name, hit.end()
        subject, second, first, more = match.groups()
        for number in [first] + re.findall(r"\d{4}", more):
            for name in [subject, second]:
                if name:
                    roles.setdefault(f"{name} {number}", role)
    return roles


# Headings that name a PART of a program rather than a program. The calendar
# sets 'Biology-Mathematics/Statistics' and 'Requirements for the Majors
# Program' at the same heading level, so without this the second would replace
# the first and the program's name would be lost.
GENERIC = re.compile(r"^(Requirements|Honours|Majors?|Minors?|Co-?op|General|Programs?\b|Streams|"
                     r"First|Second|Third|Fourth|Year|Notes?\b|Electives?|Core|Curriculum)", re.I)


def track_heading(headings, line):
    """Update {level: heading} with a '## Heading' line and return it."""
    level = len(line) - len(line.lstrip("#"))
    title = line.lstrip("# ").strip()
    if GENERIC.match(title):
        level += 0.5  # file it beneath the program heading of the same level
    headings = {k: v for k, v in headings.items() if k < level}
    headings[level] = title
    return headings


def read_page(path):
    """The lines of a mirrored page, without the SOURCE line and the site menus."""
    lines = path.read_text(encoding="utf-8").splitlines()
    # The site's navigation uses level-5 headings; the page proper starts at
    # its first level-1 heading that is not the search box.
    start = 0
    for index, line in enumerate(lines):
        if line.startswith("# ") and line != "# Search UNB":
            start = index
            break
    return lines[start:]


def main():
    parser = argparse.ArgumentParser(description="Find every mention of a MATH/STAT course in a mirrored calendar.")
    parser.add_argument("yeardir", help="folder written by mirror_calendar.py")
    parser.add_argument("--subjects", nargs="+", default=["MATH", "STAT", "DATA"])
    parser.add_argument("--check", nargs="*", default=DEFAULT_CHECK,
                        help="pages (relative to text/) where every cited course is checked")
    args = parser.parse_args()

    year = Path(args.yeardir).expanduser()
    text = year / "text"
    if not text.is_dir():
        sys.exit(f"{text} not found. Run mirror_calendar.py first.")
    ours = set(args.subjects)

    # ---- Pass 1: which courses have an entry in this calendar? --------------
    # entries maps 'MATH 3003' to its title; home maps it to its department page.
    entries, home = {}, {}
    course_pages = sorted((text / "frederictoncourses").rglob("*.txt"))
    for path in course_pages:
        for line in read_page(path):
            match = ENTRY.match(line)
            if match:
                code = f"{match.group(1)} {match.group(2)}"
                entries.setdefault(code, match.group(3).strip())
                home.setdefault(code, path.parent.name)

    def in_calendar(code):
        return "yes" if code in entries else "no"

    # ---- Pass 2: other units' courses that name one of ours -----------------
    course_rows = []
    for path in course_pages:
        current = None  # the course entry the lines being read belong to
        for line in read_page(path):
            match = ENTRY.match(line)
            if match:
                current = f"{match.group(1)} {match.group(2)}"
                continue
            if current is None or current.split()[0] in ours:
                continue  # preamble, or one of our own courses
            for code, role in cited_roles(line).items():
                if code.split()[0] in ours:
                    course_rows.append([home[current], current, entries[current], code,
                                        entries.get(code, ""), in_calendar(code), role, line])

    # ---- Pass 3: program regulations that name one of ours ------------------
    program_rows = []
    for path in sorted((text / "frederictonprograms").rglob("*.txt")):
        page = str(path.relative_to(text / "frederictonprograms"))[:-4]
        headings = {}  # level -> the heading currently in force at that level
        for line in read_page(path):
            if line.startswith("#"):
                headings = track_heading(headings, line)
                continue
            for code in cited_courses(line):
                if code.split()[0] in ours:
                    section = " > ".join(headings[k] for k in sorted(headings) if k > 1)
                    program_rows.append([page, section, code, entries.get(code, ""),
                                         in_calendar(code), line])

    def write(name, header, rows):
        with open(year / name, "w", encoding="utf-8") as handle:
            for row in [header] + rows:
                handle.write("\t".join(cell.replace("\t", " ") for cell in row) + "\n")
        print(f"Wrote {year / name}: {len(rows)} rows")

    write("sweep-programs.tsv",
          ["Program page", "Section", "Cited course", "Title", "In calendar", "Text"], program_rows)
    write("sweep-courses.tsv",
          ["Unit", "Their course", "Their title", "Cited course", "Title", "In calendar", "Role", "Text"],
          course_rows)

    # ---- The same facts turned round: one row per course of ours ------------
    # Our own program pages are left out of the "Programs" column, so that it
    # answers "who ELSE names this course?".
    own_pages = {Path(name).stem for name in args.check[:1]}
    own_sections = ("Mathematics/Statistics-Economics", "Mathematics-Physics",
                    "Biology-Mathematics/Statistics")
    programs_citing, courses_citing = {}, {}
    for page, section, code, _, _, _ in program_rows:
        if Path(page).name in own_pages or any(s in section for s in own_sections):
            continue
        # The last part of the section is the most specific; the page says whose it is.
        label = Path(page).name + (f" ({section.split(' > ')[-1]})" if section else "")
        programs_citing.setdefault(code, [])
        if label not in programs_citing[code]:
            programs_citing[code].append(label)
    for _, theirs, _, code, _, _, role, _ in course_rows:
        label = f"{theirs} ({role})"
        courses_citing.setdefault(code, [])
        if label not in courses_citing[code]:
            courses_citing[code].append(label)
    by_course = []
    for code in sorted(set(programs_citing) | set(courses_citing)):
        by_course.append([code, entries.get(code, ""), in_calendar(code),
                          "; ".join(programs_citing.get(code, [])),
                          ", ".join(courses_citing.get(code, []))])
    write("sweep-by-course.tsv",
          ["Our course", "Title", "In calendar", "Other units' programs naming it",
           "Other units' courses naming it"], by_course)

    # ---- Our own regulations: any cited course with no entry ----------------
    print(f"\n{len(entries)} course entries found in this calendar.")
    print("\nCited in our own program pages but with no course entry:")
    for name in args.check:
        path = text / name
        if not path.exists():
            print(f"  (missing page: {name})")
            continue
        headings = {}
        for line in read_page(path):
            if line.startswith("#"):
                headings = track_heading(headings, line)
                continue
            missing = [c for c in cited_courses(line) if c not in entries]
            if missing:
                section = " > ".join(headings[k] for k in sorted(headings) if k > 1)
                print(f"  {Path(name).stem} | {section} | {', '.join(missing)}")

    cited_ours = {row[2] for row in program_rows} | {row[3] for row in course_rows}
    gone = sorted(code for code in cited_ours if code not in entries)
    print("\nOur courses cited anywhere but with no course entry:")
    print("  " + (", ".join(gone) or "none"))


if __name__ == "__main__":
    main()
