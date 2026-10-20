#!/usr/bin/env -S uv run --script
# /// script
# dependencies = []
# ///
# course_history.py — which calendar years had an entry for each MATH/STAT course.
# Usage:
#   ./course_history.py
#   ./course_history.py --root DIR --out FILE.tsv
#   ./course_history.py -h | --help
#
#   --root   the academic-calendar folder (default: CALENDAR in lib/paths.py)
#   --out    the file to write (default: ROOT/course-history.tsv)
#
# It reads, for the Fredericton campus only:
#   2014-15            the department's submission documents (.doc), which hold
#                      the full MATH and STAT course listings of that year
#   2018-19 .. 2025-26 the saved calendar PDFs (undergradcalendar-YYYY-YYYY.pdf)
#   2026-27            the mirrored online calendar (2026-27/text)
# The calendars for 2015-16, 2016-17 and 2017-18 are not saved, so a course
# whose last year reads 2014-15 left some time before 2018-19.
#
# It needs two programs that are already on this Mac: pdftotext (from poppler,
# installed with macports) for the PDFs and textutil (part of macOS) for .doc.
#
# The PDFs are printed in two columns and do not always read back cleanly, so
# an isolated blank between two x's is a reading miss, not a year without the
# course. Before relying on a LAST year, search that year's PDF for the code.
#
# The output has one row per course: its most recent title, credit hours and
# format, the first and last year it has an entry, and an x under every year.

import argparse
import re
import subprocess
import sys
from pathlib import Path

# The helper modules live in lib/, beside this script.
sys.path.insert(0, str(Path(__file__).resolve().parent / "lib"))

import paths

SUBJECTS = ("MATH", "STAT", "DATA")
CODE = r"(MATH|STAT|DATA)\s?(\d{4})"
# '3 ch (3C)', '3 ch (3C 1T) [W]', '6 ch'
HOURS = re.compile(r"^(\d+(?:\.\d)?)\s*ch\b\s*(\([^)]*\))?")


def run(command):
    """Run a program and return what it prints, as text."""
    done = subprocess.run(command, capture_output=True)
    if done.returncode != 0:
        sys.exit(f"{command[0]} failed on {command[-1]}:\n{done.stderr.decode(errors='replace')}")
    return done.stdout.decode("utf-8", errors="replace")


def entries_from_lines(lines, start=0, stop=None, code=CODE, bare_counts=False):
    """Course entries in a run of lines laid out as: code, title, hours.

    The three parts are on separate lines, with blank lines between.
    Returns {('MATH', '3353'): (title, ch, format)}.

    A title is only believed when the hours follow it. The PDFs are printed in
    two columns, and pdftotext sometimes gives two codes in a row and their
    titles afterwards; then the course is recorded with a blank title (if
    bare_counts is True) rather than with the wrong one.
    """
    found = {}
    lines = [line.strip() for line in lines[start:stop]]
    for index, line in enumerate(lines):
        match = re.fullmatch(code, line)
        if not match:
            continue
        # The next two non-blank lines should be the title and the hours.
        following = [text for text in lines[index + 1:index + 8] if text][:2]
        hours = HOURS.match(following[1]) if len(following) > 1 else None
        title = following[0] if hours else ""
        if not hours and not bare_counts:
            continue  # a code on a line of its own inside running text
        found[(match.group(1), match.group(2))] = (
            title, hours.group(1) if hours else "", hours.group(2) or "" if hours else "")
    return found


def fredericton_block(lines, subject):
    """(start, stop) of the Fredericton course listing for one subject.

    The calendar PDF lists the Saint John courses first and the Fredericton
    ones after. A listing shows up as a dense run of lines that are nothing
    but a course code, so the LAST such run for a subject is Fredericton's.
    """
    code_lines = [i for i, line in enumerate(lines)
                  if re.fullmatch(subject + r"\s?\d{4}", line.strip())]
    if not code_lines:
        return None
    start = stop = code_lines[-1]
    for i in reversed(code_lines[:-1]):
        if start - i > 400:  # a gap this long means a different listing
            break
        start = i
    return start, stop + 8


def from_pdf(path):
    """Fredericton MATH/STAT/DATA entries in one calendar PDF."""
    # Without -layout, pdftotext gives the text in reading order, which puts
    # each of code, title and hours on a line of its own.
    lines = run(["pdftotext", str(path), "-"]).splitlines()
    found = {}
    for subject in SUBJECTS:
        block = fredericton_block(lines, subject)
        if block:
            # In a listing the code is printed without a space ('MATH3353');
            # in running text it has one. So a spaceless code on a line of its
            # own is an entry even when its title could not be read.
            found.update(entries_from_lines(lines, *block, code=r"(MATH|STAT|DATA)(\d{4})",
                                            bare_counts=True))
            start, stop = block
            for index in range(start, min(stop, len(lines))):
                line = lines[index].strip()
                # Sometimes the whole heading is on one line:
                # 'MATH3543 Differential Geometry for Geomatics Engineers 3 ch (3L 1T)'
                one_line = re.fullmatch(r"(MATH|STAT|DATA)(\d{4}) (.+?) (\d+(?:\.\d)? ?ch\b.*)", line)
                if one_line:
                    hours = HOURS.match(one_line.group(4))
                    found[(one_line.group(1), one_line.group(2))] = (
                        one_line.group(3), hours.group(1), hours.group(2) or "")
                # Once in a while a heading's code is printed WITH a space. It
                # is believed only when the lines around it are other headings.
                spaced = re.fullmatch(r"(MATH|STAT|DATA) (\d{4})", line)
                nearby = [text.strip() for text in lines[max(index - 3, 0):index + 4]]
                if spaced and any(re.fullmatch(r"(MATH|STAT|DATA)\d{4}", text) for text in nearby):
                    found.setdefault((spaced.group(1), spaced.group(2)), ("", "", ""))
    return found


def from_doc(path):
    """Entries in one of the 2014-15 submission documents."""
    text = run(["textutil", "-convert", "txt", "-stdout", str(path)])
    return entries_from_lines(text.splitlines())


def from_mirror(text_dir):
    """Entries in the mirrored online calendar: 'MATH3003 | Title | 3 ch (3C)'."""
    found = {}
    for path in sorted((text_dir / "frederictoncourses").rglob("*.txt")):
        for line in path.read_text(encoding="utf-8").splitlines():
            parts = [part.strip() for part in line.split("|")]
            match = re.fullmatch(CODE, parts[0])
            if match and len(parts) >= 2:
                hours = HOURS.match(parts[2]) if len(parts) > 2 else None
                found[(match.group(1), match.group(2))] = (
                    parts[1], hours.group(1) if hours else "", hours.group(2) or "" if hours else "")
    return found


def main():
    parser = argparse.ArgumentParser(description="Which calendar years had an entry for each MATH/STAT course.")
    parser.add_argument("--root", default=str(paths.CALENDAR))
    parser.add_argument("--out")
    args = parser.parse_args()
    root = Path(args.root).expanduser()
    out = Path(args.out).expanduser() if args.out else root / "course-history.tsv"

    by_year = {}  # '2018-19' -> {course: (title, ch, format)}
    docs = sorted((root / "2014-15" / "curriculumsubmissionstolevel2").glob("*courses.doc"))
    if docs:
        by_year["2014-15"] = {}
        for path in docs:
            by_year["2014-15"].update(from_doc(path))
    for path in sorted(root.glob("*/undergradcalendar-*.pdf")):
        by_year[path.parent.name] = from_pdf(path)
    for text_dir in sorted(root.glob("*/text")):
        by_year[text_dir.parent.name] = from_mirror(text_dir)

    years = sorted(by_year)
    for year in years:
        print(f"{year}: {len(by_year[year])} entries")

    courses = sorted({course for year in years for course in by_year[year]})
    with open(out, "w", encoding="utf-8") as handle:
        handle.write("\t".join(["Course", "Title", "CH", "Format", "First year", "Last year"] + years) + "\n")
        for course in courses:
            present = [year for year in years if course in by_year[year]]
            # Use the most recent year that gives a title.
            title, ch, fmt = "", "", ""
            for year in reversed(present):
                title, ch, fmt = by_year[year][course]
                if title:
                    break
            marks = ["x" if year in present else "" for year in years]
            handle.write("\t".join([f"{course[0]} {course[1]}", title, ch, fmt,
                                    present[0], present[-1]] + marks) + "\n")
    print(f"Wrote {out}: {len(courses)} courses")


if __name__ == "__main__":
    main()
