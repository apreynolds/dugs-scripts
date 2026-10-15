#!/usr/bin/env -S uv run --script
# /// script
# dependencies = ["openpyxl"]
# ///
# timetable_grid.py — draw one term of the proposed timetable as a week grid (PDF).
# Usage:
#   ./timetable_grid.py                     Fall 2027 -> timetable-grid-fall-2027.tex and .pdf
#   ./timetable_grid.py -s "Winter 2028"    another term sheet
#   ./timetable_grid.py -o OUT.tex          write the LaTeX somewhere else
#   ./timetable_grid.py --no-pdf            write the .tex only
#   ./timetable_grid.py -h | --help
#
#   -w, --workbook   the proposed timetable (default: proposed-timetable-2027-28.xlsx
#                    in timetable-committee). Only read
#   -s, --sheet      the term sheet to draw (default: "Fall 2027")
#   -o, --output     the .tex to write (default: timetable-grid-<sheet>.tex beside
#                    the proposed timetable). Overwritten every run: it is output,
#                    and nothing in it is typed by a person
#       --no-pdf     do not run pr-latexmk afterwards
#       --no-programs  leave out the programs under each course
#
# WHAT THIS IS FOR
# A one-page picture of a term for a department meeting: Monday to Friday
# across the top, each day split into 1000-, 2000-, 3000- and 4000-level
# columns, half-hour rows down the side. Each course sits in the rows its
# suggested time covers, coloured by level, so a crowded slot or a gap stands
# out. The layout is Patrick's, from a sketch he made in the workbook.
# The page is landscape ledger (17in x 11in; it is read on screen, not
# printed), typeset with tabularray in Patrick's default fonts (the
# Libertinus block of the `basicdoc` snippet).
#
# WHAT IT SHOWS, AND WHAT IT LEAVES OUT
#   - Suggested times only; alternative times and tutorials are left out.
#   - A course whose sections meet at different times appears at each.
#   - Courses at the same level whose times overlap share one block, which
#     spans all of them; the times are written out when they differ.
#   - MATH 0863, a pre-university course, goes in the 1000-level column.
#   - Under each course, very small: the degree programs that require it,
#     as check_timetable.py reads the requirements workbooks ("Hons:" for
#     Honours only). A course that is one choice among several is not marked.
#     --no-programs leaves these out.
# The grid is generated: change a time on the term sheet and run this again.
#
# This script borrows its time reading from check_timetable.py and its
# course-number helpers from build_courses.py, which must sit in the same folder.

import argparse
import datetime
import re
import shutil
import subprocess
import sys
from collections import defaultdict
from pathlib import Path

import build_advising as ba
import build_courses as bc
import check_timetable as ct

DAYS = [("M", "Monday"), ("T", "Tuesday"), ("W", "Wednesday"), ("Th", "Thursday"), ("F", "Friday")]
LEVELS = [1, 2, 3, 4]
# One colour per level, light enough to print and to read black text on.
COLOURS = {1: "DDEBF7", 2: "E2EFDA", 3: "FFF2CC", 4: "E4DFEC"}  # blue, green, yellow, lilac
FIRST_ROW_TIME = 8 * 60 + 30  # the grid starts at 8:30
STEP = 30  # minutes per row
# The page: US ledger (11in x 17in, ANSI B), landscape. Nobody prints it, so
# the largest common US size gives the most room.
PAPER = "ansibpaper, landscape, margin=0.5in"
# Height shared out among the half-hour rows: the 10in of the page inside
# half-inch margins, less room for the title, headings and footer. (On legal
# paper 152mm fitted and 162mm spilled onto a second page; here 212mm pushed
# the footer onto a second page.)
GRID_HEIGHT_MM = 200

# How each degree-requirements workbook is named under a course, in this order.
# Only these are shown. The BA Math is left out: the MATH/STAT courses it
# requires are essentially the BSc Math's (Patrick, 2026-10-15).
PROGRAM_NAMES = {
    "bsc-math": "BSc Math", "bsc-biomath": "Bio-Math",
    "bsc-mathecon": "Math-Econ", "bsc-mathphys": "Math-Phys",
}


def time_text(minutes):
    """510 -> '8:30', 780 -> '1:00': the 12-hour clock, as the timetable writes it."""
    hours, mins = divmod(minutes, 60)
    return f"{(hours - 1) % 12 + 1}:{mins:02d}"


def level(key):
    """'MATH 3213' -> 3. MATH 0863 counts as first year."""
    return min(max(int(key[-4]), 1), 4)


def required_by(programs, course):
    """The programs that require a course, as short text for under it.

    'All programs', or a list; programs that require it on the Honours track
    only come after 'Hons:', e.g. 'BSc Math, BA Math; Hons: Math-Phys'.
    "Required" is check_timetable.py's reading of the requirements workbooks:
    a course in a category that needs all of its courses.
    """
    order = list(PROGRAM_NAMES)
    programs = sorted((p for p in programs if p["name"] in PROGRAM_NAMES), key=lambda p: order.index(p["name"]))
    every, honours = [], []
    for program in programs:
        tracks = [track for track, entry in program["tracks"].items()
                  if (role := ct.role_in(entry, course)) and role[0] == "required"]
        if not tracks:
            continue
        name = PROGRAM_NAMES.get(program["name"], program["name"])
        (every if len(tracks) == len(program["tracks"]) else honours).append(name)
    if len(every) == len(programs):
        text = "All programs"
    else:
        text = ", ".join(every)
    if honours:
        text += ("; " if text else "") + "Hons: " + ", ".join(honours)
    return text


def place_courses(courses, programs):
    """Sort a term's courses into the grid.

    programs is check_timetable's reading of the requirements workbooks, or
    None to leave the program names out.
    Returns ({(day, level): [(first row, last row, label, time, programs text)]},
    latest end time, [courses whose time could not be read]).
    """
    placed = defaultdict(list)
    skipped = []
    latest = FIRST_ROW_TIME
    for key, record in courses.items():
        time = record.get("Suggested time", "")
        options = ct.parse_options(time)
        if not options:
            skipped.append(f"{bc.course_display(key)} {time!r}")
            continue
        label = bc.course_display(key)
        # Several sections at one time read 'MATH 1003 (4 sec)'. Sections at
        # different times are placed separately, so need no count.
        sections = record.get("Sections", "")
        if len(options) == 1 and sections.isdigit() and int(sections) > 1:
            label += f" ({sections} sec)"
        needed_by = required_by(programs, key) if programs else ""
        for option in options:
            for day, start, end in option:
                if day not in dict(DAYS):
                    continue
                # A row covers [t, t + 30); the course fills every row it touches.
                first = (start - FIRST_ROW_TIME) // STEP
                last = (end - 1 - FIRST_ROW_TIME) // STEP
                placed[(day, level(key))].append(
                    (first, last, label, f"{time_text(start)}-{time_text(end)}", needed_by))
                latest = max(latest, end)
    return placed, latest, skipped


def blocks_in(items):
    """Group one column's courses into blocks: courses whose rows overlap share a block.

    Returns [(first row, last row, [(course text, programs text)])].
    """
    groups = []
    for item in sorted(items):
        # Sorted by first row, so an overlap can only be with the last group.
        if groups and item[0] <= max(x[1] for x in groups[-1]):
            groups[-1].append(item)
        else:
            groups.append([item])
    blocks = []
    for group in groups:
        same_time = len({x[3] for x in group}) == 1
        lines = []
        for _, _, label, time, needed_by in group:
            line = (label if same_time else f"{label} {time}", needed_by)
            if line not in lines:
                lines.append(line)
        blocks.append((min(x[0] for x in group), max(x[1] for x in group), lines))
    return blocks


def tex_grid(placed, rows):
    """The body of the tblr: one line of LaTeX per table row."""
    # cells[row][column] is the LaTeX for that cell, '' when empty or covered
    # by a block above it. Column 0 is the time.
    columns = len(DAYS) * len(LEVELS)
    cells = [[""] * columns for _ in range(rows)]
    for d, (day, _) in enumerate(DAYS):
        for i, lvl in enumerate(LEVELS):
            col = d * len(LEVELS) + i
            for first, last, lines in blocks_in(placed.get((day, lvl), [])):
                span = last - first + 1
                # Each course is its own paragraph, with the programs that
                # require it in tiny grey type right under it (\progs, defined
                # in the preamble), and a little space before the next course.
                parts = [course + r"\par" + (rf"\progs{{{needed_by}}}" if needed_by else "")
                         for course, needed_by in lines]
                text = r"\vspace{3pt}".join(parts)
                cells[first][col] = rf"\SetCell[r={span}]{{bg=lvl{lvl}}} {text}"
    body = []
    for r in range(rows):
        body.append(" & ".join([time_text(FIRST_ROW_TIME + r * STEP)] + cells[r]) + r" \\")
    return body


def tex_document(title, placed, rows, footer):
    """The whole .tex file, with comments so it can be read on its own."""
    columns = len(DAYS) * len(LEVELS)
    row_height = GRID_HEIGHT_MM / rows
    day_heads = " & ".join(rf"\SetCell[c={len(LEVELS)}]{{c}} {name}" + " &" * (len(LEVELS) - 1)
                           for _, name in DAYS)
    level_heads = " & ".join(rf"\SetCell{{bg=lvl{lvl}}} {lvl}000-level" for _ in DAYS for lvl in LEVELS)
    # A heavier rule at the left of each day: after the time column, and
    # every four columns from there.
    day_rules = ",".join(str(2 + d * len(LEVELS)) for d in range(len(DAYS) + 1))
    colours = "\n".join(rf"\definecolor{{lvl{lvl}}}{{HTML}}{{{c}}}  % {lvl}000-level"
                        for lvl, c in COLOURS.items())
    lines = [
        "% !TeX program = pdflatex",
        "% GENERATED by timetable_grid.py (~/3work/course-scripts). Do not edit:",
        "% change the times on the term sheet of the proposed timetable and rerun.",
        r"\documentclass[10pt]{article}",
        "",
        "% Landscape ledger (ANSI B): 17in x 11in, half-inch margins all round.",
        rf"\usepackage[{PAPER}]{{geometry}}",
        r"\usepackage[dvipsnames]{xcolor}",
        r"\let\CheckCommand\providecommand % stops a microtype warning",
        r"\usepackage[tracking = true]{microtype}",
        r"\usepackage[T1]{fontenc}",
        "",
        "% >>> FONTS (default: libertinus), as in the basicdoc snippet",
        r"\usepackage[semibold, lining, serif, tt=false]{libertinus}",
        "% sflining, not basicdoc's sfoldstyle: course numbers and times read better in lining figures",
        r"\usepackage[sf, sflining, scaled=0.85]{merriweather}",
        r"\usepackage[scale=0.96]{newtxtt}",
        r"\usepackage[libertinus, libaltvw, liby, amsthm, nosymbolsc, cmbraces, smallerops,"
        r" timesmathacc, upint, scaled=1.0]{newtxmath}",
        "% \\useosf is left off: a timetable reads better in lining figures.",
        "% <<< FONTS",
        "",
        r"\usepackage{tabularray}",
        "",
        "% Fill colours for the four levels of course.",
        colours,
        "",
        "% The programs that require a course, set very small and grey under it.",
        "% The \\par inside the braces ends the paragraph while \\tiny is still in",
        "% force, so its lines use \\tiny's own close spacing.",
        r"\newcommand{\progs}[1]{{\tiny\color{black!65}#1\par}}",
        "",
        r"\pagestyle{empty} % no page number",
        r"\setlength{\parindent}{0pt}",
        "",
        r"\begin{document}",
        "",
        rf"{{\sffamily\bfseries\Large {title}}}",
        "",
        r"\vspace{2mm}",
        r"\begin{tblr}{",
        f"  % the time, then {columns} equal columns: four levels for each day",
        f"  colspec = {{Q[r,t,8mm] *{{{columns}}}{{X[c,m]}}}},",
        "  colsep = 2pt,",
        "  rowsep = 1pt,",
        "  % every half-hour row the same height, so time reads straight down the page",
        f"  row{{3-Z}} = {{ht={row_height:.1f}mm}},",
        r"  row{1} = {font=\sffamily\bfseries},",
        r"  row{2} = {font=\sffamily\scriptsize},",
        r"  column{1} = {font=\sffamily\footnotesize},",
        r"  cell{3-Z}{2-Z} = {font=\sffamily\footnotesize},",
        "  % faint lines everywhere, then a heavier one to the left of each day",
        "  hlines = {0.3pt, black!25},",
        "  vlines = {0.3pt, black!25},",
        f"  vline{{{day_rules}}} = {{0.9pt, black}},",
        "}",
        f" & {day_heads} \\\\",
        f" & {level_heads} \\\\",
    ]
    lines += tex_grid(placed, rows)
    lines += [
        r"\end{tblr}",
        "",
        r"\vspace{2mm}",
        rf"{{\footnotesize\itshape {footer}}}",
        "",
        r"\end{document}",
        "",
    ]
    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(description="Draw one term of the proposed timetable as a week grid.")
    parser.add_argument("-w", "--workbook", default=ct.DEFAULT_WORKBOOK,
                        help="the proposed timetable (only read)")
    parser.add_argument("-s", "--sheet", default="Fall 2027", help='term sheet to draw (default "Fall 2027")')
    parser.add_argument("-o", "--output", help="the .tex to write")
    parser.add_argument("--no-pdf", action="store_true", help="do not run pr-latexmk")
    parser.add_argument("--no-programs", action="store_true",
                        help="leave out the programs that require each course")
    args = parser.parse_args()

    source = Path(args.workbook).expanduser()
    slug = re.sub(r"\W+", "-", args.sheet.strip().lower())
    output = Path(args.output).expanduser() if args.output else source.with_name(f"timetable-grid-{slug}.tex")

    terms = {title: courses for title, _, courses in ct.read_proposal(source)}
    if args.sheet not in terms:
        sys.exit(f"No sheet {args.sheet!r} in {source.name}; it has {', '.join(terms)}.")

    programs = None
    if not args.no_programs:
        programs, *_ = ct.load_programs(ba.DEFAULT_REQUIREMENTS, ba.DEFAULT_OUTPUT,
                                        bc.DEFAULT_TIMETABLES, set(terms[args.sheet]))
    placed, latest, skipped = place_courses(terms[args.sheet], programs)
    rows = (latest - 1 - FIRST_ROW_TIME) // STEP + 1
    today = datetime.date.today().isoformat()
    title = f"Mathematics and Statistics, {args.sheet}: proposed timetable"
    footer = (f"Suggested times from {source.name}, drawn {today}. "
              "Alternative times and tutorials are not shown. "
              "A course with sections at different times appears at each; "
              "``(4 sec)'' means four sections at the same time.")
    if programs:
        footer += (" Under each course, in grey: the degree programs that require it, "
                   "read from the degree-requirements workbooks; ``Hons:'' marks programs that "
                   "require it on the Honours track only. "
                   "Courses that are only one choice among several are not marked.")
    output.write_text(tex_document(title, placed, rows, footer))
    print(f"Wrote {output}")
    if skipped:
        print("  Left out, because the time could not be read: " + ", ".join(skipped))

    if args.no_pdf:
        return
    if not shutil.which("pr-latexmk"):
        print("  pr-latexmk not found; compile the .tex yourself.")
        return
    result = subprocess.run(["pr-latexmk", str(output)], capture_output=True, text=True)
    if result.returncode != 0:
        print(result.stdout[-3000:] + result.stderr[-2000:])
        sys.exit("pr-latexmk failed; see above.")
    print(f"Wrote {output.with_suffix('.pdf')}")


if __name__ == "__main__":
    main()
