#!/usr/bin/env -S uv run --script
# /// script
# dependencies = ["openpyxl"]
# ///
# timetable_grid.py — draw the proposed timetable as week grids (PDF).
# Usage:
#   ./timetable_grid.py                     department grid, Fall 2027 -> timetable-grid-fall-2027.pdf
#   ./timetable_grid.py -s "Winter 2028"    department grid for another term sheet
#   ./timetable_grid.py --by-level          one PDF per degree program, a page per level per term
#   ./timetable_grid.py --alternatives      either of those, with every alternative time taken
#   ./timetable_grid.py -o OUT.tex          write the department grid's LaTeX somewhere else
#   ./timetable_grid.py --no-pdf            write the .tex only
#   ./timetable_grid.py -h | --help
#
#   -w, --workbook   the proposed timetable (default: proposed-timetable-2027-28.xlsx
#                    in timetable-committee). Only read
#   -s, --sheet      the term sheet for the department grid (default: "Fall 2027")
#   -o, --output     the department grid's .tex (default: timetable-grid-<sheet>.tex
#                    beside the proposed timetable)
#       --by-level   instead, draw each degree-requirements workbook's courses,
#                    a page per level per term sheet, into
#                    program-grids-by-level/<program>.tex beside the proposed timetable
#       --alternatives  draw each course that has an Alternative time at that
#                    time instead of its Suggested one, marked "alternative; was
#                    ...". Written beside the usual output, with "-alternatives"
#                    in the name, so the two can be compared
#       --no-pdf     do not run pr-latexmk afterwards
#       --no-programs  department grid: leave out the programs under each course
# Every .tex is overwritten each run: it is output, and nothing in it is typed.
#
# WHAT THIS IS FOR
# A one-page picture of a term: Monday to Friday across the top, each day
# split into 1000-level, 2000-level and 3000/4000-level columns, half-hour rows down
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
# THE PROGRAM GRIDS BY LEVEL
# For each degree-requirements workbook, a page per level (1000, 2000,
# 3000/4000) per term sheet, each day split into Required and Options:
#   - our courses at their suggested time; one running in 2027-28 in the other
#     term is left off; one not running in 2027-28 at all sits where it last
#     sat in this term, marked "not in 2027-28"
#   - a brief note from the master list's Offering note: "alternate years",
#     "occasional", ...
#   - other units' courses in italic, at the time they had the last time they
#     ran in person that fall (or winter), since their 2027-28 times are not
#     known yet; a course with labs and no lecture at its lab times
#   - "Honours only" (or "Major only") under a course one track alone requires
#   - Options: named choices, first-year science (FIRST_YEAR_SCIENCE), and
#     every 3000/4000-level MATH/STAT course that is not inactive
#
# IN BOTH
#   - Tutorials are left out, and alternative times unless --alternatives.
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

from openpyxl import load_workbook

import build_advising as ba
import build_courses as bc
import check_timetable as ct

DAYS = [("M", "Monday"), ("T", "Tuesday"), ("W", "Wednesday"), ("Th", "Thursday"), ("F", "Friday")]
# 1000-level, 2000-level, and the upper years together: 3000- and 4000-level
# courses are taken side by side, so they share a column (Patrick, 2026-10-15).
LEVELS = [1, 2, 3]
LEVEL_NAMES = {1: "1000-level", 2: "2000-level", 3: "3000/4000-level"}
# One colour per level, light enough to print and to read black text on.
COLOURS = {1: "DDEBF7", 2: "E2EFDA", 3: "FFF2CC"}  # blue, green, yellow
OPTION_FILL = "F2F2F2"  # light grey, for the Options column of the by-level pages

# The first-year science lectures and labs a BSc chooses among. The workbooks'
# SciLec and SciLab categories name no courses (an advisor assigns them); this
# list is from the notes in bsc-mathecon-2026 ("6 of 8 from BIOL 1001/1012,
# CHEM 1001/1012, ESCI 1052/1041, PHYS 1051/1052"; labs "BIOL 1006/1017, CHEM
# 1006/1017, ESCI 1057 (or 1026)/1046, PHYS 1091/1092").
FIRST_YEAR_SCIENCE = [
    "BIOL1001", "BIOL1012", "CHEM1001", "CHEM1012", "ESCI1052", "ESCI1041", "PHYS1051", "PHYS1052",
    "BIOL1006", "BIOL1017", "CHEM1006", "CHEM1017", "ESCI1057", "ESCI1046", "PHYS1091", "PHYS1092",
]
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
    """'MATH 2213' -> 2; 'MATH 3213' and 'MATH 4153' -> 3. MATH 0863 counts as first year."""
    return min(max(int(key[-4]), 1), 3)


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

    An entry may carry a fifth item, the sub-column it goes in under each day;
    without one it goes under its level.
    Returns ({(day, sub-column): [(first row, last row, label, time, small lines)]},
    latest end time, [entries whose time could not be read]).
    """
    placed = defaultdict(list)
    skipped = []
    latest = FIRST_ROW_TIME
    for key, time, small, style, *column in entries:
        column = column[0] if column else level(key)
        options = ct.parse_options(time)
        if not options:
            skipped.append(f"{bc.course_display(key)} {time!r}")
            continue
        label = bc.course_display(key)
        # style: True or "italic" (another unit's course), "moved" (at its new
        # time), "left" (its old slot, struck out), or False.
        if style is True or style == "italic":
            label = rf"\textit{{{label}}}"
        elif style == "moved":
            label = rf"\moved{{{label}}}"
        elif style == "left":
            label = rf"\leftslot{{{label}}}"
        for option in options:
            for day, start, end in option:
                if day not in dict(DAYS):
                    continue
                # A row covers [t, t + 30); the course fills every row it touches.
                first = (start - FIRST_ROW_TIME) // STEP
                last = (end - 1 - FIRST_ROW_TIME) // STEP
                placed[(day, column)].append(
                    (first, last, label, f"{time_text(start)}-{time_text(end)}", small, start, end))
                latest = max(latest, end)
    return placed, latest, skipped


def blocks_in(items):
    """Group one column's courses into blocks: courses whose rows overlap share a block.

    Returns [(first row, last row, [(course text, [small lines], start, end)],
    same time?)]; start and end are minutes after midnight.
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
        for _, _, label, time, small, start, end in sorted(group, key=lambda x: (x[5], x[6])):
            line = (label if same_time else f"{label} {time}", small, start, end)
            if line not in lines:
                lines.append(line)
        blocks.append((min(x[0] for x in group), max(x[1] for x in group), lines, same_time))
    return blocks


def tex_brackets(first, span, lines, row_height):
    """A block whose courses run at different times, drawn as a TikZ picture.

    The text stays centred as in any block. Each course also gets a bracket
    spanning exactly its own minutes: the first on the left edge "[", the
    second on the right "]", and so on alternately; a third or later bracket
    on a side already used steps in a little, so two never sit on top of
    each other.
    """
    top = FIRST_ROW_TIME + first * STEP  # the time at the block's top edge
    # A row is taller than its ht by its rowsep above and below (1pt each)
    # and its rule (0.3pt); the cell's content starts 1pt below the rule.
    extra_pt = 2.3

    def depth(minutes):
        """How far below the top of the cell's content a time falls, as LaTeX."""
        rows = max(0.0, min(span, (minutes - top) / STEP))
        return f"{rows * row_height:.2f}mm+{rows * extra_pt - 1:.2f}pt"

    height = f"{span * row_height:.2f}mm+{span * extra_pt - 2:.2f}pt"
    middle = f"{span * row_height / 2:.2f}mm+{(span * extra_pt - 2) / 2:.2f}pt"
    draw = []
    for i, (_, _, start, end) in enumerate(lines):
        inset = 0.5 + (i // 2) * 1.5  # mm in from the cell edge
        tick = 1.2  # mm: the bracket's top and bottom arms
        # The bracket covers the whole slot: a 50-minute class its 60 minutes,
        # an 80-minute one its 90, so it ends on a half-hour line.
        slot_end = FIRST_ROW_TIME + -(-(end - FIRST_ROW_TIME) // STEP) * STEP
        y1, y2 = depth(start), depth(slot_end)
        if i % 2 == 0:  # left: [
            x, arm = f"{inset}mm", f"{inset + tick}mm"
        else:  # right: ]
            x, arm = f"\\cellw-{inset}mm", f"\\cellw-{inset + tick}mm"
        # {-(...)}: the minus must cover both parts, mm and pt.
        draw.append(rf"\draw[bracket] ({arm},{{-({y1})}}) -- ({x},{{-({y1})}}) -- "
                    rf"({x},{{-({y2})}}) -- ({arm},{{-({y2})}});")
    text = r"\vspace{3pt}".join(course + r"\par" + "".join(rf"\progs{{{tex_escape(s)}}}" for s in small)
                                for course, small, _, _ in lines)
    # The text goes in a centred minipage, so it is set exactly as in any
    # other block (same font, same paragraphs), just with room for brackets.
    # \cellw holds the cell's width, taken before the picture starts: inside
    # a TikZ node \linewidth is no longer the cell's.
    # baseline at the middle: sat on its bottom edge, the picture would get a
    # line's depth added below it, and the rows would grow.
    return (r"\setlength{\cellw}{\linewidth}\begin{tikzpicture}[baseline=(current bounding box.center)]"
            rf"\useasboundingbox (0,0) rectangle (\cellw,{{-({height})}});"
            + "".join(draw)
            + rf"\node[inner sep=0pt] at (0.5\cellw,{{-({middle})}}) "
              rf"{{\begin{{minipage}}{{\dimexpr\cellw-9mm\relax}}\centering {text}\end{{minipage}}}};"
            r"\end{tikzpicture}")


# The sub-columns under each day: (key, heading, fill colour, extra cell
# options). The usual one is a column per level; the by-level pages use a
# Required and an Options column instead.
LEVEL_COLUMNS = [(lvl, LEVEL_NAMES[lvl], f"lvl{lvl}", "") for lvl in LEVELS]


def tex_grid(placed, rows, spec, row_height):
    """The body of the tblr: one line of LaTeX per table row. row_height is in mm."""
    # cells[row][column] is the LaTeX for that cell, '' when empty or covered
    # by a block above it. Column 0 is the time.
    columns = len(DAYS) * len(spec)
    cells = [[""] * columns for _ in range(rows)]
    for d, (day, _) in enumerate(DAYS):
        for i, (key, _, fill, extra) in enumerate(spec):
            col = d * len(spec) + i
            for first, last, lines, same_time in blocks_in(placed.get((day, key), [])):
                span = last - first + 1
                if not same_time:
                    # Courses at different times: brackets show each one's span.
                    cells[first][col] = (rf"\SetCell[r={span}]{{bg={fill}{extra}}} "
                                         + tex_brackets(first, span, lines, row_height))
                    continue
                # Each course is its own paragraph, with its small grey lines
                # right under it (\progs, defined in the preamble), and a
                # little space before the next course.
                parts = [course + r"\par" + "".join(rf"\progs{{{tex_escape(s)}}}" for s in small)
                         for course, small, _, _ in lines]
                cells[first][col] = rf"\SetCell[r={span}]{{bg={fill}{extra}}} " + r"\vspace{3pt}".join(parts)
    body = []
    for r in range(rows):
        body.append(" & ".join([time_text(FIRST_ROW_TIME + r * STEP)] + cells[r]) + r" \\")
    return body


def tex_page(title, entries, footer, spec=LEVEL_COLUMNS):
    """One page: a title, the grid, and a footer. Returns (LaTeX lines, skipped entries).

    spec is the sub-columns under each day (see LEVEL_COLUMNS).
    """
    placed, latest, skipped = place_entries(entries)
    rows = (max(latest, DAY_END) - 1 - FIRST_ROW_TIME) // STEP + 1
    columns = len(DAYS) * len(spec)
    day_heads = " & ".join(rf"\SetCell[c={len(spec)}]{{c}} {name}" + " &" * (len(spec) - 1)
                           for _, name in DAYS)
    level_heads = " & ".join(rf"\SetCell{{bg={fill}}} {heading}" for _ in DAYS for _, heading, fill, _ in spec)
    # A heavier rule at the left of each day: after the time column, and
    # at every day's first sub-column from there.
    day_rules = ",".join(str(2 + d * len(spec)) for d in range(len(DAYS) + 1))
    lines = [
        rf"{{\sffamily\bfseries\Large {tex_escape(title)}}}",
        "",
        r"\vspace{2mm}",
        r"\begin{tblr}{",
        f"  % the time, then {columns} equal columns: {len(spec)} for each day",
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
    lines += tex_grid(placed, rows, spec, GRID_HEIGHT_MM / rows)
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
    colours = "\n".join(rf"\definecolor{{lvl{lvl}}}{{HTML}}{{{c}}}  % {LEVEL_NAMES[lvl]}"
                        for lvl, c in COLOURS.items())
    colours += "\n" + rf"\definecolor{{optfill}}{{HTML}}{{{OPTION_FILL}}}  % options, on the by-level pages"
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
        "% A block of courses at different times is a small TikZ picture, with a",
        "% bracket spanning each course's own minutes.",
        r"\usepackage{tikz}",
        r"\tikzset{bracket/.style={line width=1.2pt, black!70}}",
        r"\newlength{\cellw} % such a cell's width, measured before its picture starts",
        "",
        "% Fill colours for the four levels of course.",
        colours,
        "",
        "% The small grey lines under a course (who requires it, where its time came from).",
        "% The \\par inside the braces ends the paragraph while \\tiny is still in",
        "% force, so its lines use \\tiny's own close spacing.",
        r"\newcommand{\progs}[1]{{\tiny\color{black!65}#1\par}}",
        "",
        "% With --alternatives: a moved course in red at its new time, and struck",
        "% out in grey in the slot it left. normalem keeps \\emph as italics.",
        r"\usepackage[normalem]{ulem}",
        r"\newcommand{\moved}[1]{\textcolor{BrickRed}{#1}}",
        r"\newcommand{\leftslot}[1]{\textcolor{black!45}{\sout{#1}}}",
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


def our_time(record, alternatives):
    """One of our courses on a term sheet: (time to draw, note).

    With alternatives, a course that has an Alternative time is drawn there,
    and the note says where it was.
    """
    suggested = record.get("Suggested time", "")
    alternative = record.get("Alternative time", "")
    if alternatives and ct.parse_options(alternative):
        return alternative, f"alternative; was {suggested}"
    return suggested, ""


# ----------------------------------------------------------------------------
# The department grid
# ----------------------------------------------------------------------------
def department_grid(args, source, terms):
    slug = re.sub(r"\W+", "-", args.sheet.strip().lower())
    if args.alternatives:
        slug += "-alternatives"
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
        # The small grey lines under the course: whether it moved, our
        # programs, then other units.
        time, moved = our_time(record, args.alternatives)
        small = []
        if programs and (ours := required_by(programs, key)):
            small.append(ours)
        if others.get(key):
            small.append("Also: " + ", ".join(UNIT_SHORT.get(u, u) for u in others[key]))
        if moved:
            # In red at its new time; struck out, in grey, where it was.
            suggested = record.get("Suggested time", "")
            entries.append((key, time, [moved] + small, "moved"))
            entries.append((key, suggested, [f"moved to {time}"], "left"))
        else:
            entries.append((key, time, small, False))

    today = datetime.date.today().isoformat()
    if args.alternatives:
        footer = (f"Times from {source.name}, drawn {today}, with every alternative time taken: "
                  "a course that moved is in red at its new time (``alternative; was'' its suggested "
                  "time), and struck out in grey in the slot it left. "
                  "Tutorials are not shown. A course with sections at different times appears at each.")
    else:
        footer = (f"Suggested times from {source.name}, drawn {today}. "
                  "Alternative times and tutorials are not shown. "
                  "A course with sections at different times appears at each.")
    if programs:
        footer += (" Under each course, in grey: our degree programs that require it, "
                   "read from the degree-requirements workbooks (``Hons:'' for the Honours track only); "
                   "then, after ``Also:'', the other units with a program that requires it, from "
                   f"{OTHER_UNITS.name}. "
                   "Courses that are only one choice among several are not marked.")
    title = f"Mathematics and Statistics, {args.sheet}: proposed timetable"
    if args.alternatives:
        title += ", with the alternative times"
    page, skipped = tex_page(title, entries, footer)
    output.write_text(tex_document([page]))
    print(f"Wrote {output}")
    if skipped:
        print("  Left out, because the time could not be read: " + ", ".join(skipped))
    compile_pdf(output, args.no_pdf)


# ----------------------------------------------------------------------------
# The program grids by level: a page per level per term, Required and Options
# side by side under each day
# ----------------------------------------------------------------------------
BY_LEVEL_COLUMNS = [
    ("req", "Required", "lvl{level}", ""),
    ("opt", "Options", "optfill", ", fg=black!60"),  # grey text: a choice, not a requirement
]


def recent_time(offerings, course, season_terms):
    """Another unit's course: (time, note) from the last time it ran in person this season, or None.

    Its lecture times; or, for a course with labs and no lecture, its lab times.
    """
    per = offerings.get(course, {})
    ran = [t for t in season_terms if per.get(t) and per[t]["inperson"]]
    if not ran:
        return None
    offering, label = per[ran[-1]], bc.term_label(ran[-1])
    if offering["lecture_slots"]:
        return " | ".join(offering["lecture_slots"]), f"{label} time"
    if offering["tutorial_slots"]:
        return " | ".join(offering["tutorial_slots"]), f"lab, {label} times"
    return None


def master_slot(text):
    """A master-list timeslot cell -> a time this script can draw, or ''.

    'MWF 11:30-12:20 + tut T 8:30-9:20 (new 2026 F)' -> 'MWF 11:30-12:20';
    '|' separates sections and is kept; '-' and '(no set time)' give ''.
    """
    pieces = []
    for piece in str(text or "").split("|"):
        piece = piece.split("+")[0]  # drop the tutorial
        if ct.parse_options(piece):
            pieces.append(piece.strip())
    return " | ".join(pieces)


def offering_word(record):
    """The master list's Offering note and Status, made brief: 'alternate years', 'occasional'.

    Blank means every year in the same term, which needs no note.
    """
    if record.get("Status", "").strip().lower() == "hiatus":
        return "on hiatus"
    note = record.get("Offering note", "").strip()
    if note.startswith("alternate years"):
        return "alternate years"
    match = re.match(r"(\d) of last \d+ years", note)
    if match:
        times = int(match[1])
        return "most years" if times >= 5 else "some years" if times >= 3 else "occasional"
    if note.startswith("not since"):
        return note
    return ""  # blank, '-', or 'every year since ...': runs every year now


def program_level_grids(args, source, proposal):
    terms = {title: courses for title, _, courses in proposal}
    seasons = {title: season for title, season, _ in proposal}
    offered = set().union(*terms.values())  # ours that run in 2027-28, in either term
    programs, offerings, covered, _, _ = ct.load_programs(
        ba.DEFAULT_REQUIREMENTS, ba.DEFAULT_OUTPUT, bc.DEFAULT_TIMETABLES, offered)
    # Every MATH/STAT course on the master list that is not inactive: its
    # offering note, and its usual slot for a course not running in 2027-28.
    master = {k: r for k, r in ba.sheet_records(load_workbook(bc.DEFAULT_OUTPUT, data_only=True)["MAIN"]).items()
              if r.get("Status", "").strip().lower() != "inactive"}
    upper = {c for c in master if level(c) == 3}
    folder = source.parent / ("program-grids-by-level" + ("-alternatives" if args.alternatives else ""))
    folder.mkdir(exist_ok=True)
    today = datetime.date.today().isoformat()

    for program in programs:
        # Every course the program names, sorted into required (on some track)
        # or option. A required course wins over an option.
        required, options = {}, set()
        for entry in program["tracks"].values():
            for course, roles in entry["exact"].items():
                tracks = required_tracks(program, course)
                if tracks:
                    required[course] = tracks
                elif any(role == "option" for role, _, _ in roles):
                    options.add(course)
        if program["title"].startswith("BSc"):
            options |= {c for c in FIRST_YEAR_SCIENCE if c not in required}

        pages = []
        for title, courses in terms.items():
            season_terms = [t for t in covered if t.endswith(seasons[title])][-ct.LOOKBACK:]
            slot_column = "Fall timeslot" if seasons[title] == "FA" else "Winter timeslot"
            # Our courses that only a pattern like MATH 3*** catches are
            # electives, and every upper-level course of ours is shown too:
            # all of them go under Options.
            electives = {c for c in master if any((r := ct.role_in(e, c)) and r[0] == "elective"
                                                  for e in program["tracks"].values())}
            for lvl in LEVELS:
                entries = []
                for course in sorted(set(required) | options | electives | upper):
                    if level(course) != lvl:
                        continue
                    column = "req" if course in required else "opt"
                    small = []
                    if course in required and len(required[course]) < len(program["tracks"]):
                        small.append(" and ".join(required[course]) + " only")
                    if course[:-4] in ct.OURS:
                        # In 2027-28 this term: its proposed time. In 2027-28
                        # but the other term: not on this page. Not in 2027-28
                        # at all: where it last sat in this season.
                        word = offering_word(master.get(course, {}))
                        if course in courses:
                            time, moved = our_time(courses[course], args.alternatives)
                            word = "; ".join(w for w in (moved, word) if w)
                        elif course in offered:
                            continue
                        else:
                            time = master_slot(master.get(course, {}).get(slot_column))
                            word = "; ".join(w for w in (word, "not in 2027-28") if w)
                        if not ct.parse_options(time):
                            continue  # no set time (a thesis), or never in this season
                        entries.append((course, time, small + ([word] if word else []), False, column))
                        continue
                    found = recent_time(offerings, course, season_terms)
                    if found:
                        time, note = found
                        entries.append((course, time, small + [note], True, column))
                if not entries:
                    continue  # nothing at this level this term
                spec = [(k, h, f.format(level=lvl), x) for k, h, f, x in BY_LEVEL_COLUMNS]
                footer = (f"{program['title']}, {LEVEL_NAMES[lvl]} courses: required ones in the first column "
                          "under each day; options in grey in the second (named choices, first-year science, "
                          "and every 3000/4000-level MATH/STAT course). Our courses at the times proposed in "
                          f"{source.name}; one not running in 2027-28 at the slot it last had in this term, "
                          "marked ``not in 2027-28''. Notes such as ``alternate years'' or ``occasional'' "
                          "come from the master course list. Other units' courses in italic at the time they "
                          "had the last time they ran in person in that term, since their 2027-28 times are "
                          "not known yet. A course with sections at different times (or a lab) appears at "
                          f"each: choose one. Drawn {today}.")
                heading = f"{program['title']}, {title}: {LEVEL_NAMES[lvl]} courses"
                if args.alternatives:
                    heading += ", with the alternative times"
                    footer += " Every alternative time is taken: a course that moved says ``alternative; was''."
                page, _ = tex_page(heading, entries, footer, spec)
                pages.append(page)
        output = folder / f"{program['name']}.tex"
        output.write_text(tex_document(pages))
        print(f"Wrote {output} ({len(pages)} pages)")
        compile_pdf(output, args.no_pdf)


def main():
    parser = argparse.ArgumentParser(description="Draw the proposed timetable as week grids.")
    parser.add_argument("-w", "--workbook", default=ct.DEFAULT_WORKBOOK,
                        help="the proposed timetable (only read)")
    parser.add_argument("-s", "--sheet", default="Fall 2027", help='term sheet to draw (default "Fall 2027")')
    parser.add_argument("-o", "--output", help="the department grid's .tex")
    parser.add_argument("--by-level", action="store_true",
                        help="one PDF per degree program, a page per level per term, required and options")
    parser.add_argument("--alternatives", action="store_true",
                        help="draw every course that has an Alternative time at that time")
    parser.add_argument("--no-pdf", action="store_true", help="do not run pr-latexmk")
    parser.add_argument("--no-programs", action="store_true",
                        help="department grid: leave out the programs under each course")
    args = parser.parse_args()

    source = Path(args.workbook).expanduser()
    proposal = ct.read_proposal(source)
    if args.by_level:
        program_level_grids(args, source, proposal)
    else:
        department_grid(args, source, {title: courses for title, _, courses in proposal})


if __name__ == "__main__":
    main()
