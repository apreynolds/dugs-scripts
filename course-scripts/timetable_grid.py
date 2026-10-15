#!/usr/bin/env -S uv run --script
# /// script
# dependencies = ["openpyxl"]
# ///
# timetable_grid.py — draw the proposed timetable as week grids (PDF).
# Usage:
#   ./timetable_grid.py                     department grid, Fall 2027 -> timetable-grid-fall-2027.pdf
#   ./timetable_grid.py -s "Winter 2028"    department grid for another term sheet
#   ./timetable_grid.py --programs          one PDF per degree program, a page per term
#   ./timetable_grid.py -o OUT.tex          write the department grid's LaTeX somewhere else
#   ./timetable_grid.py --no-pdf            write the .tex only
#   ./timetable_grid.py -h | --help
#
#   -w, --workbook   the proposed timetable (default: proposed-timetable-2027-28.xlsx
#                    in timetable-committee). Only read
#   -s, --sheet      the term sheet for the department grid (default: "Fall 2027")
#   -o, --output     the department grid's .tex (default: timetable-grid-<sheet>.tex
#                    beside the proposed timetable)
#       --programs   instead, draw each degree-requirements workbook's required
#                    courses, every term sheet, into program-grids/<program>.tex
#                    beside the proposed timetable
#       --no-pdf     do not run pr-latexmk afterwards
#       --no-programs  department grid: leave out the programs under each course
# Every .tex is overwritten each run: it is output, and nothing in it is typed.
#
# WHAT THIS IS FOR
# A one-page picture of a term: Monday to Friday across the top, each day
# split into 1000-, 2000-, 3000- and 4000-level columns, half-hour rows down
# the side. Each course sits in the rows its time covers, coloured by level,
# so a crowded slot or a gap stands out. The layout is Patrick's, from a
# sketch he made in the workbook. The page is landscape ledger (17in x 11in;
# it is read on screen, not printed), typeset with tabularray in Patrick's
# default fonts (the Libertinus block of the `basicdoc` snippet).
#
# THE DEPARTMENT GRID
# Every course on the term sheet, at its suggested time. Under each course,
# very small: our degree programs that require it, as check_timetable.py reads
# the requirements workbooks ("Hons:" for Honours only), then "Also:" and the
# other units with a program that requires it, from the "Required, by course"
# table of curriculum/COURSES-required-by-other-units.md. A course that is one
# choice among several is not marked. --no-programs leaves both out.
#
# THE PROGRAM GRIDS
# For each degree-requirements workbook, the courses it requires, one page per
# term sheet:
#   - our courses at their suggested time; one not on the term sheet is not
#     offered that term, and is left off that page
#   - other units' courses in italic, at the time they had the last time they
#     ran in person that fall (or winter), since their 2027-28 times are not
#     known yet
#   - "Honours only" (or "Major only") under a course one track alone requires
#   - courses with labs and no lecture, and courses with no set time (a
#     thesis), listed under the grid instead
#
# IN BOTH
#   - Alternative times and tutorials are left out.
#   - A course whose sections meet at different times appears at each.
#   - Courses at the same level whose times overlap share one block, which
#     spans all of them; the times are written out when they differ.
#   - MATH 0863, a pre-university course, goes in the 1000-level column.
# The grids are generated: change a time on the term sheet and run this again.
#
# This script borrows its time and requirements reading from check_timetable.py
# and its course-number helpers from build_courses.py, which must sit in the
# same folder.

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
DAY_END = 17 * 60  # and runs to at least 5:00, so every page shows the same day
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

# Other units that require each course: the "Required, by course" table of the
# department's reference document, which is where to change them.
OTHER_UNITS = bc.DUGS / "curriculum/COURSES-required-by-other-units.md"
OTHER_UNITS_HEADING = "# Required, by course"
# Shorter names for the grid, where a cell is narrow.
UNIT_SHORT = {
    "Computer Science": "CS", "Medical Laboratory Science": "Med Lab Sci",
    "Environmental Sciences": "Env Sci", "Earth Sciences": "Earth Sci",
}


def time_text(minutes):
    """510 -> '8:30', 780 -> '1:00': the 12-hour clock, as the timetable writes it."""
    hours, mins = divmod(minutes, 60)
    return f"{(hours - 1) % 12 + 1}:{mins:02d}"


def level(key):
    """'MATH 3213' -> 3. MATH 0863 counts as first year."""
    return min(max(int(key[-4]), 1), 4)


def tex_escape(text):
    """Escape the few characters a course or program name might hold."""
    return text.replace("&", r"\&").replace("%", r"\%")


# ----------------------------------------------------------------------------
# Who requires what
# ----------------------------------------------------------------------------
def read_other_units(path):
    """{course: [units]} from the 'Required, by course' table of the reference document.

    The table's rows look like '| MATH 1843 | Management | BBA (first term); ... |'.
    Returns {} (and says so) if the document or its table cannot be found.
    """
    try:
        lines = Path(path).read_text().splitlines()
    except OSError:
        print(f"  {path} not found; other units are left out.")
        return {}
    if OTHER_UNITS_HEADING not in lines:
        print(f"  No '{OTHER_UNITS_HEADING}' section in {Path(path).name}; other units are left out.")
        return {}
    units = {}
    for line in lines[lines.index(OTHER_UNITS_HEADING) + 1:]:
        if line.startswith("# "):
            break  # the next section
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(cells) >= 2 and ba.any_course(bc.course_key(cells[0])):
            units[bc.course_key(cells[0])] = [u.strip() for u in cells[1].split(",") if u.strip()]
    return units


def required_tracks(program, course):
    """The tracks of one program on which a course is required."""
    return [track for track, entry in program["tracks"].items()
            if (role := ct.role_in(entry, course)) and role[0] == "required"]


def required_by(programs, course):
    """Our programs that require a course, as short text for under it.

    'All our programs', or a list; programs that require it on the Honours
    track only come after 'Hons:', e.g. 'BSc Math, Bio-Math; Hons: Math-Phys'.
    "Required" is check_timetable.py's reading of the requirements workbooks:
    a course in a category that needs all of its courses.
    """
    order = list(PROGRAM_NAMES)
    programs = sorted((p for p in programs if p["name"] in PROGRAM_NAMES), key=lambda p: order.index(p["name"]))
    every, honours = [], []
    for program in programs:
        tracks = required_tracks(program, course)
        if not tracks:
            continue
        name = PROGRAM_NAMES[program["name"]]
        (every if len(tracks) == len(program["tracks"]) else honours).append(name)
    if len(every) == len(programs):
        text = "All our programs"  # "our": the "Also:" line below lists other units
    else:
        text = ", ".join(every)
    if honours:
        text += ("; " if text else "") + "Hons: " + ", ".join(honours)
    return text


# ----------------------------------------------------------------------------
# Laying out a page
# An entry is one course to draw: (course, time as written, [small grey
# lines under it], italic?).
# ----------------------------------------------------------------------------
def place_entries(entries):
    """Sort entries into the grid.

    Returns ({(day, level): [(first row, last row, label, time, small lines)]},
    latest end time, [entries whose time could not be read]).
    """
    placed = defaultdict(list)
    skipped = []
    latest = FIRST_ROW_TIME
    for key, time, small, italic in entries:
        options = ct.parse_options(time)
        if not options:
            skipped.append(f"{bc.course_display(key)} {time!r}")
            continue
        label = bc.course_display(key)
        if italic:
            label = rf"\textit{{{label}}}"
        for option in options:
            for day, start, end in option:
                if day not in dict(DAYS):
                    continue
                # A row covers [t, t + 30); the course fills every row it touches.
                first = (start - FIRST_ROW_TIME) // STEP
                last = (end - 1 - FIRST_ROW_TIME) // STEP
                placed[(day, level(key))].append(
                    (first, last, label, f"{time_text(start)}-{time_text(end)}", small))
                latest = max(latest, end)
    return placed, latest, skipped


def blocks_in(items):
    """Group one column's courses into blocks: courses whose rows overlap share a block.

    Returns [(first row, last row, [(course text, [small lines])])].
    """
    groups = []
    for item in sorted(items, key=lambda x: (x[0], x[1], x[2])):
        # Sorted by first row, so an overlap can only be with the last group.
        if groups and item[0] <= max(x[1] for x in groups[-1]):
            groups[-1].append(item)
        else:
            groups.append([item])
    blocks = []
    for group in groups:
        same_time = len({x[3] for x in group}) == 1
        lines = []
        for _, _, label, time, small in group:
            line = (label if same_time else f"{label} {time}", small)
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
                # Each course is its own paragraph, with its small grey lines
                # right under it (\progs, defined in the preamble), and a
                # little space before the next course.
                parts = [course + r"\par" + "".join(rf"\progs{{{tex_escape(s)}}}" for s in small)
                         for course, small in lines]
                cells[first][col] = rf"\SetCell[r={span}]{{bg=lvl{lvl}}} " + r"\vspace{3pt}".join(parts)
    body = []
    for r in range(rows):
        body.append(" & ".join([time_text(FIRST_ROW_TIME + r * STEP)] + cells[r]) + r" \\")
    return body


def tex_page(title, entries, footer):
    """One page: a title, the grid, and a footer. Returns (LaTeX lines, skipped entries)."""
    placed, latest, skipped = place_entries(entries)
    rows = (max(latest, DAY_END) - 1 - FIRST_ROW_TIME) // STEP + 1
    columns = len(DAYS) * len(LEVELS)
    day_heads = " & ".join(rf"\SetCell[c={len(LEVELS)}]{{c}} {name}" + " &" * (len(LEVELS) - 1)
                           for _, name in DAYS)
    level_heads = " & ".join(rf"\SetCell{{bg=lvl{lvl}}} {lvl}000-level" for _ in DAYS for lvl in LEVELS)
    # A heavier rule at the left of each day: after the time column, and
    # every four columns from there.
    day_rules = ",".join(str(2 + d * len(LEVELS)) for d in range(len(DAYS) + 1))
    lines = [
        rf"{{\sffamily\bfseries\Large {tex_escape(title)}}}",
        "",
        r"\vspace{2mm}",
        r"\begin{tblr}{",
        f"  % the time, then {columns} equal columns: four levels for each day",
        f"  colspec = {{Q[r,t,8mm] *{{{columns}}}{{X[c,m]}}}},",
        "  colsep = 2pt,",
        "  rowsep = 1pt,",
        "  % every half-hour row the same height, so time reads straight down the page",
        f"  row{{3-Z}} = {{ht={GRID_HEIGHT_MM / rows:.1f}mm}},",
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
    ]
    return lines, skipped


def tex_document(pages):
    """The whole .tex file, with comments so it can be read on its own. pages: lists of lines."""
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
        "% The small grey lines under a course (who requires it, where its time came from).",
        "% The \\par inside the braces ends the paragraph while \\tiny is still in",
        "% force, so its lines use \\tiny's own close spacing.",
        r"\newcommand{\progs}[1]{{\tiny\color{black!65}#1\par}}",
        "",
        r"\pagestyle{empty} % no page number",
        r"\setlength{\parindent}{0pt}",
        "",
        r"\begin{document}",
        "",
    ]
    for i, page in enumerate(pages):
        if i:
            lines += [r"\newpage", ""]
        lines += page
    lines += [r"\end{document}", ""]
    return "\n".join(lines)


def compile_pdf(tex, no_pdf):
    """Run pr-latexmk on a .tex, unless told not to. Stops the script if it fails."""
    if no_pdf:
        return
    if not shutil.which("pr-latexmk"):
        print("  pr-latexmk not found; compile the .tex yourself.")
        return
    result = subprocess.run(["pr-latexmk", str(tex)], capture_output=True, text=True)
    if result.returncode != 0:
        print(result.stdout[-3000:] + result.stderr[-2000:])
        sys.exit(f"pr-latexmk failed on {tex.name}; see above.")
    print(f"Wrote {tex.with_suffix('.pdf')}")


# ----------------------------------------------------------------------------
# The department grid
# ----------------------------------------------------------------------------
def department_grid(args, source, terms):
    slug = re.sub(r"\W+", "-", args.sheet.strip().lower())
    output = Path(args.output).expanduser() if args.output else source.with_name(f"timetable-grid-{slug}.tex")
    if args.sheet not in terms:
        sys.exit(f"No sheet {args.sheet!r} in {source.name}; it has {', '.join(terms)}.")

    programs, others = None, {}
    if not args.no_programs:
        programs, *_ = ct.load_programs(ba.DEFAULT_REQUIREMENTS, ba.DEFAULT_OUTPUT,
                                        bc.DEFAULT_TIMETABLES, set(terms[args.sheet]))
        others = read_other_units(OTHER_UNITS)
    entries = []
    for key, record in terms[args.sheet].items():
        # The small grey lines under the course: our programs, then other units.
        small = []
        if programs and (ours := required_by(programs, key)):
            small.append(ours)
        if others.get(key):
            small.append("Also: " + ", ".join(UNIT_SHORT.get(u, u) for u in others[key]))
        entries.append((key, record.get("Suggested time", ""), small, False))

    today = datetime.date.today().isoformat()
    footer = (f"Suggested times from {source.name}, drawn {today}. "
              "Alternative times and tutorials are not shown. "
              "A course with sections at different times appears at each.")
    if programs:
        footer += (" Under each course, in grey: our degree programs that require it, "
                   "read from the degree-requirements workbooks (``Hons:'' for the Honours track only); "
                   "then, after ``Also:'', the other units with a program that requires it, from "
                   f"{OTHER_UNITS.name}. "
                   "Courses that are only one choice among several are not marked.")
    page, skipped = tex_page(f"Mathematics and Statistics, {args.sheet}: proposed timetable", entries, footer)
    output.write_text(tex_document([page]))
    print(f"Wrote {output}")
    if skipped:
        print("  Left out, because the time could not be read: " + ", ".join(skipped))
    compile_pdf(output, args.no_pdf)


# ----------------------------------------------------------------------------
# The program grids
# ----------------------------------------------------------------------------
def program_grids(args, source, proposal):
    terms = {title: courses for title, _, courses in proposal}
    seasons = {title: season for title, season, _ in proposal}
    ours_offered = set().union(*terms.values())
    programs, offerings, covered, _, _ = ct.load_programs(
        ba.DEFAULT_REQUIREMENTS, ba.DEFAULT_OUTPUT, bc.DEFAULT_TIMETABLES, ours_offered)
    folder = source.parent / "program-grids"
    folder.mkdir(exist_ok=True)
    today = datetime.date.today().isoformat()

    for program in programs:
        # Every course this program requires on at least one track.
        required = {}
        for entry in program["tracks"].values():
            for course in entry["exact"]:
                tracks = required_tracks(program, course)
                if tracks:
                    required[course] = tracks
        pages, labs_only, timeless = [], set(), set()
        for title, courses in terms.items():
            season_terms = [t for t in covered if t.endswith(seasons[title])][-ct.LOOKBACK:]
            entries = []
            for course, tracks in sorted(required.items()):
                small = []
                if len(tracks) < len(program["tracks"]):
                    small.append(" and ".join(tracks) + " only")
                if course[:-4] in ct.OURS:
                    if course not in courses:
                        continue  # not offered this term
                    time = courses[course].get("Suggested time", "")
                    if not ct.parse_options(time):
                        timeless.add(course)
                        continue
                    entries.append((course, time, small, False))
                    continue
                # Another unit's course: the time it had the last time it ran
                # in person this season.
                per = offerings.get(course, {})
                ran = [t for t in season_terms if per.get(t) and per[t]["inperson"]]
                if not ran:
                    continue  # not offered in this season lately
                offering = per[ran[-1]]
                if not offering["lecture_slots"]:
                    labs_only.add(course)
                    continue
                small.append(f"{bc.term_label(ran[-1])} time")
                entries.append((course, " | ".join(offering["lecture_slots"]), small, True))
            # Ours with no section in any term sheet, and others with none lately.
            footer = (f"{program['title']}: the courses it requires, at the times proposed in "
                      f"{source.name} for our courses, drawn {today}. "
                      "Other units' courses, in italic, are shown at the time they had the last time "
                      "they ran in person in that term (e.g. ``2026 F time''), since their 2027-28 times "
                      "are not known yet. A course with sections at different times appears at each: "
                      "choose one. Alternative times, tutorials and labs are not shown, nor are "
                      "courses that are one choice among several, or electives.")
            page, skipped = tex_page(f"{program['title']}, {title}: required courses", entries, footer)
            pages.append(page)
            for item in skipped:
                print(f"  {program['name']}, {title}: time not read: {item}")
        # Courses that could not go on any page, listed after the last page.
        not_placed = []
        absent = sorted(c for c in required if c[:-4] in ct.OURS and c not in ours_offered)
        if absent:
            not_placed.append("Required, but not on the 2027-28 course list: "
                              + ", ".join(bc.course_display(c) for c in absent) + ".")
        if labs_only:
            not_placed.append("Labs with no lecture (choose a section): "
                              + ", ".join(bc.course_display(c) for c in sorted(labs_only)) + ".")
        if timeless:
            not_placed.append("No set time: " + ", ".join(bc.course_display(c) for c in sorted(timeless)) + ".")
        if not_placed:
            pages[-1] += [rf"{{\footnotesize\itshape {' '.join(not_placed)}}}", ""]
        output = folder / f"{program['name']}.tex"
        output.write_text(tex_document(pages))
        print(f"Wrote {output}")
        compile_pdf(output, args.no_pdf)


def main():
    parser = argparse.ArgumentParser(description="Draw the proposed timetable as week grids.")
    parser.add_argument("-w", "--workbook", default=ct.DEFAULT_WORKBOOK,
                        help="the proposed timetable (only read)")
    parser.add_argument("-s", "--sheet", default="Fall 2027", help='term sheet to draw (default "Fall 2027")')
    parser.add_argument("-o", "--output", help="the department grid's .tex")
    parser.add_argument("--programs", action="store_true",
                        help="draw one PDF per degree program instead of the department grid")
    parser.add_argument("--no-pdf", action="store_true", help="do not run pr-latexmk")
    parser.add_argument("--no-programs", action="store_true",
                        help="department grid: leave out the programs under each course")
    args = parser.parse_args()

    source = Path(args.workbook).expanduser()
    proposal = ct.read_proposal(source)
    if args.programs:
        program_grids(args, source, proposal)
    else:
        department_grid(args, source, {title: courses for title, _, courses in proposal})


if __name__ == "__main__":
    main()
