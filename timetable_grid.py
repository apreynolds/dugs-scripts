# /// script
# dependencies = ["openpyxl"]
# ///
# timetable_grid.py — draw a year's timetable as week grids (PDF).
# Usage:
#   uv run timetable_grid.py 2026-27                  2026-27 as the registrar's export has it
#   uv run timetable_grid.py 2027-28 --draft          2027-28 as a draft: 2026-27's courses and times
#   uv run timetable_grid.py 2027-28 --proposal WORKBOOK [--alternatives]
#                                                     our times from a proposed timetable workbook
#   uv run timetable_grid.py 2026-27 --only programs  only the program grids (or: --only department)
#   uv run timetable_grid.py 2026-27 --no-pdf         write the .tex only
#   uv run timetable_grid.py -h | --help
#
#   YEAR            the academic year to draw. Everything is written into the
#                   Timetable team's folder for that year, which must exist:
#                     YEAR/departmental-grids/timetable-grid-fall-2026.tex (and winter)
#                     YEAR/program-grids-by-level/<program>.tex
#                   each compiled to a PDF beside it with pr-latexmk. Each
#                   program PDF is also copied to the Advising team, as
#                   RESOURCES/program-grids-by-level/YEAR/<program>-DO-NOT-EDIT.pdf
#                   (not for --draft or --proposal: those are the committee's).
#   --draft         use LAST year's export: every course of ours at the time it
#                   had then. For the timetable committee in the fall, before
#                   anything is decided. Files are named with "-draft"
#   --proposal WB   take our courses and times from the "Course" and "Suggested
#                   time" columns of a workbook's "Fall YYYY" and "Winter YYYY"
#                   sheets (the committee's working timetable). Files are named
#                   with "-proposed"
#   --alternatives  with --proposal: draw each course that has an "Alternative
#                   time" there instead, struck out in grey where it was and in
#                   red where it goes. Files get "-alternative" in the name
#   --only KIND     draw only the department grids or only the program grids
#   --no-pdf        do not run pr-latexmk afterwards
# Every .tex is overwritten each run: it is output, and nothing in it is typed.
#
# WHEN
# Four times a year: in the fall, a --draft of next year for the committee;
# in December, this year again for any winter changes; in March, next year
# once the registrar has released it; in August, next year again. Each time
# download the new export into timetable-exports-by-year first.
#
# WHERE THE TIMES COME FROM
# The newest export of the year (lib/paths.py: TIMETABLE_EXPORTS). Our
# courses: every MATH/STAT/DATA course with an in-person lecture, at its
# lecture times. Other units' courses: their times in the same term. With
# --proposal, whose year has no export yet, other units' courses are at the
# time they had the last time they ran in person in that season, and say so.
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
# Every course of ours in the term, one PDF per term. Under each course,
# very small: our degree programs that require it, as lib/programs.py reads
# the requirements workbooks ("Hons:" for Honours only), then "Also:" and the
# other units with a program that requires it, from the "Required, by course"
# table of COURSES-required-by-other-units.md. A course that is one choice
# among several is not marked.
#
# THE PROGRAM GRIDS BY LEVEL
# For each degree-requirements workbook, a page per level (1000, 2000,
# 3000/4000) per term, each day split by length of meeting: 50-minute
# meetings in one column, 80-minute ones and anything longer (labs) in the
# other. Two courses in the same column that overlap then almost always
# share their time exactly, so no times are written: the rows say it
# (Patrick, 2026-10-21). On each page:
#   - every course on a light grey fill; a required one framed in blue, an
#     option (named choices, first-year science, every 3000/4000-level
#     MATH/STAT course) not framed
#   - our courses at their time; one running this year in the other term is
#     left off; one not running this year at all sits where it last sat in
#     this term (the master list's timeslot), struck out, "not in YEAR";
#     one the master list calls inactive or on hiatus is left off too
#   - other units' courses in italic; a course with labs and no lecture at its
#     lab times
#   - "Honours only" (or "Major only") under a course one track alone requires
#   - nothing else under a course
#
# IN BOTH
#   - Tutorials and labs of our courses are left out.
#   - A course whose sections meet at different times appears at each.
#   - Courses in the same column whose times overlap share one block, which
#     spans all of them; the times are written out (and bracketed) only when
#     they differ.
#   - MATH 0863, a pre-university course, goes in the 1000-level column.
#
# This script borrows its requirements reading from lib/programs.py and its
# course-number and export reading from build_courses.py and build_advising.py,
# which must sit in the same folder.

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
import build_courses as bc  # also puts lib/ on sys.path
import programs as pg  # from lib/
import paths

DAYS = [("M", "Monday"), ("T", "Tuesday"), ("W", "Wednesday"), ("Th", "Thursday"), ("F", "Friday")]
# 1000-level, 2000-level, and the upper years together: 3000- and 4000-level
# courses are taken side by side, so they share a column (Patrick, 2026-10-15).
LEVELS = [1, 2, 3]
LEVEL_NAMES = {1: "1000-level", 2: "2000-level", 3: "3000/4000-level"}
# One colour per level, light enough to print and to read black text on.
COLOURS = {1: "DDEBF7", 2: "E2EFDA", 3: "FFF2CC"}  # blue, green, yellow
COURSE_FILL = "F2F2F2"  # light grey, behind every course on the by-level pages
REQUIRED_FRAME = "2F5597"  # dark blue, the frame round a required course there
# A meeting this long or shorter goes in the by-level pages' 50-minute
# column; a longer one (80 minutes, a lab) in the other.
SHORT_MEETING = 60

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
    "bsc-math": "BSc Math", "bsc-stat": "BSc Stat", "bsc-biomath": "Bio-Math",
    "bsc-mathecon": "Math-Econ", "bsc-mathphys": "Math-Phys",
}

# Other units that require each course: the "Required, by course" table of the
# department's reference document, which is where to change them.
OTHER_UNITS = paths.OTHER_UNITS
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
            if (role := pg.role_in(entry, course)) and role[0] == "required"]


def required_by(programs, course):
    """Our programs that require a course, as short text for under it.

    'All our programs', or a list; programs that require it on the Honours
    track only come after 'Hons:', e.g. 'BSc Math, Bio-Math; Hons: Math-Phys'.
    "Required" is lib/programs.py's reading of the requirements workbooks:
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

    An entry may carry a fifth item, the sub-column it goes in under each day
    ("length": by the length of each meeting, "short" or "long"); without one
    it goes under its level. A sixth, if there is one, boxes the course:
    "framed" (required) or "plain" (an option), as on the by-level pages.
    Returns ({(day, sub-column): [(first row, last row, label, time, small lines,
    start, end, box)]}, latest end time, [entries whose time could not be read]).
    """
    placed = defaultdict(list)
    skipped = []
    latest = FIRST_ROW_TIME
    for key, time, small, style, *extra in entries:
        column = extra[0] if extra else level(key)
        box = extra[1] if len(extra) > 1 else None
        options = pg.parse_options(time)
        if not options:
            skipped.append(f"{bc.course_display(key)} {time!r}")
            continue
        label = bc.course_display(key)
        # style: True or "italic" (another unit's course), "moved" (at its new
        # time), "left" (its old slot, struck out), "notrunning" (not this
        # year, struck out), or False.
        if style is True or style == "italic":
            label = rf"\textit{{{label}}}"
        elif style == "moved":
            label = rf"\moved{{{label}}}"
        elif style == "left":
            label = rf"\leftslot{{{label}}}"
        elif style == "notrunning":
            label = rf"\notrunning{{{label}}}"
        for option in options:
            for day, start, end in option:
                if day not in dict(DAYS):
                    continue
                # A row covers [t, t + 30); the course fills every row it touches.
                first = (start - FIRST_ROW_TIME) // STEP
                last = (end - 1 - FIRST_ROW_TIME) // STEP
                where = column
                if column == "length":
                    where = "short" if end - start <= SHORT_MEETING else "long"
                placed[(day, where)].append(
                    (first, last, label, f"{time_text(start)}-{time_text(end)}", small, start, end, box))
                latest = max(latest, end)
    return placed, latest, skipped


def blocks_in(items):
    """Group one column's courses into blocks: courses whose rows overlap share a block.

    Returns [(first row, last row, [(course text, [small lines], start, end, box)],
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
        # Required courses (framed) before options, then by time.
        for _, _, label, time, small, start, end, box in sorted(
                group, key=lambda x: (x[7] != "framed", x[5], x[6])):
            # Times are written out only on the department grid; a boxed
            # course sits at its own rows instead (tex_lanes).
            line = (label if same_time or box else f"{label} {time}", small, start, end, box)
            if line not in lines:
                lines.append(line)
        blocks.append((min(x[0] for x in group), max(x[1] for x in group), lines, same_time))
    return blocks


def tex_course(course, small, box):
    """One course in a block: its name, its small grey lines right under it
    (\\progs, defined in the preamble), boxed if box is "framed" or "plain"."""
    text = course + r"\par" + "".join(rf"\progs{{{tex_escape(s)}}}" for s in small)
    # A box is set inline, so it needs its own \par, or the next sits beside it.
    if box == "framed":
        return rf"\reqcourse{{{text}}}\par"
    if box == "plain":
        return rf"\optcourse{{{text}}}\par"
    return text


def slot_end(end):
    """A course's end, rounded up to the next half-hour line: a 50-minute class
    covers its 60 minutes, an 80-minute one its 90."""
    return FIRST_ROW_TIME + -(-(end - FIRST_ROW_TIME) // STEP) * STEP


def block_geometry(first, span, row_height):
    """For a block drawn as a TikZ picture: (depth, height, middle), as LaTeX.

    depth(minutes) is how far below the top of the cell's content a time
    falls; height is the whole block's; middle is half of it.
    """
    top = FIRST_ROW_TIME + first * STEP  # the time at the block's top edge
    # A row is taller than its ht by its rowsep above and below (1pt each)
    # and its rule (0.3pt); the cell's content starts 1pt below the rule.
    extra_pt = 2.3

    def depth(minutes):
        rows = max(0.0, min(span, (minutes - top) / STEP))
        return f"{rows * row_height:.2f}mm+{rows * extra_pt - 1:.2f}pt"

    height = f"{span * row_height:.2f}mm+{span * extra_pt - 2:.2f}pt"
    middle = f"{span * row_height / 2:.2f}mm+{(span * extra_pt - 2) / 2:.2f}pt"
    return depth, height, middle


def tex_lanes(first, span, lines, row_height):
    """A block of boxed courses at different times (the by-level pages), as a TikZ picture.

    Courses at the same time share a grey box, which sits at exactly its own
    rows; boxes whose times overlap stand side by side in lanes, the way a
    calendar shows a busy day. No time needs writing: the rows say it.
    """
    depth, height, _ = block_geometry(first, span, row_height)
    times = {}  # (start, end) -> its courses, in the block's order
    for course, small, start, end, box in lines:
        times.setdefault((start, slot_end(end)), []).append(tex_course(course, small, box))
    # Each box takes the first lane that is free by its start, so two
    # boxes share a lane whenever they can, and there are as few as can be.
    lane_ends, lane_of = [], {}
    for start, end in sorted(times):
        lane = next((i for i, e in enumerate(lane_ends) if e <= start), len(lane_ends))
        if lane == len(lane_ends):
            lane_ends.append(end)
        else:
            lane_ends[lane] = end
        lane_of[(start, end)] = lane
    share = 1 / len(lane_ends)  # of the cell's width, per lane
    draw = []
    for (start, end), parts in times.items():
        left = lane_of[(start, end)] * share
        # The box stops 0.4mm short of each edge, so neighbours stay apart.
        # {-(...)}: the minus must cover both parts of a depth, mm and pt.
        draw.append(rf"\fill[coursefill] ({left:.4f}\cellw+0.4mm,{{-({depth(start)})-0.4mm}}) rectangle "
                    rf"({left + share:.4f}\cellw-0.4mm,{{-({depth(end)})+0.4mm}});")
        draw.append(rf"\node[inner sep=0pt] at ({left + share / 2:.4f}\cellw,{{-({depth((start + end) / 2)})}}) "
                    rf"{{\begin{{minipage}}{{\dimexpr{share:.4f}\cellw-1.6mm\relax}}\centering "
                    + r"\vspace{3pt}".join(parts) + r"\end{minipage}};")
    # As in tex_brackets: \cellw is the cell's width, and the baseline sits
    # at the middle so the rows do not grow.
    return (r"\setlength{\cellw}{\linewidth}\begin{tikzpicture}[baseline=(current bounding box.center)]"
            rf"\useasboundingbox (0,0) rectangle (\cellw,{{-({height})}});"
            + "".join(draw) + r"\end{tikzpicture}")


def tex_brackets(first, span, lines, row_height):
    """A block whose courses run at different times, drawn as a TikZ picture.

    The text stays centred as in any block. Each course also gets a bracket
    spanning exactly its own minutes: the first on the left edge "[", the
    second on the right "]", and so on alternately; a third or later bracket
    on a side already used steps in a little, so two never sit on top of
    each other.
    """
    depth, height, middle = block_geometry(first, span, row_height)
    draw = []
    for i, (_, _, start, end, _) in enumerate(lines):
        inset = 0.5 + (i // 2) * 1.5  # mm in from the cell edge
        tick = 1.2  # mm: the bracket's top and bottom arms
        y1, y2 = depth(start), depth(slot_end(end))
        if i % 2 == 0:  # left: [
            x, arm = f"{inset}mm", f"{inset + tick}mm"
        else:  # right: ]
            x, arm = f"\\cellw-{inset}mm", f"\\cellw-{inset + tick}mm"
        # {-(...)}: the minus must cover both parts, mm and pt.
        draw.append(rf"\draw[bracket] ({arm},{{-({y1})}}) -- ({x},{{-({y1})}}) -- "
                    rf"({x},{{-({y2})}}) -- ({arm},{{-({y2})}});")
    text = r"\vspace{3pt}".join(tex_course(course, small, box) for course, small, _, _, box in lines)
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
# 50-minute and a longer column instead (BY_LEVEL_COLUMNS).
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
                if not same_time and lines[0][4]:
                    # Boxed courses at different times: each time its own
                    # box, in lanes; the boxes carry the fill, not the cell.
                    cells[first][col] = (rf"\SetCell[r={span}]{{}} "
                                         + tex_lanes(first, span, lines, row_height))
                    continue
                if not same_time:
                    # Courses at different times: brackets show each one's span.
                    cells[first][col] = (rf"\SetCell[r={span}]{{bg={fill}{extra}}} "
                                         + tex_brackets(first, span, lines, row_height))
                    continue
                # Each course is its own paragraph, with a little space
                # before the next.
                parts = [tex_course(course, small, box) for course, small, _, _, box in lines]
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
    colours += "\n" + rf"\definecolor{{coursefill}}{{HTML}}{{{COURSE_FILL}}}  % behind every course, on the by-level pages"
    colours += "\n" + rf"\definecolor{{reqframe}}{{HTML}}{{{REQUIRED_FRAME}}}  % the frame of a required course there"
    lines = [
        "% !TeX program = pdflatex",
        "% GENERATED by timetable_grid.py (~/3work/dugs-scripts). Do not edit:",
        "% rerun timetable_grid.py instead (see its -h).",
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
        "% The by-level pages: a course of ours not running this year, struck out;",
        "% every course in a box the width of its cell, a required one framed in",
        "% blue, an option framed in the fill colour, so both take the same room.",
        r"\newcommand{\notrunning}[1]{\sout{#1}}",
        r"\newcommand{\coursebox}[2]{{\setlength{\fboxsep}{1.5pt}\setlength{\fboxrule}{0.8pt}%",
        r"  \fcolorbox{#1}{coursefill}{\parbox{\dimexpr\linewidth-2\fboxsep-2\fboxrule\relax}{\centering #2}}}}",
        r"\newcommand{\reqcourse}[1]{\coursebox{reqframe}{#1}}",
        r"\newcommand{\optcourse}[1]{\coursebox{coursefill}{#1}}",
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
    """Run pr-latexmk on a .tex, unless told not to. Returns the PDF, or None if none was made.

    Stops the script if the compile fails.
    """
    if no_pdf:
        return None
    if not shutil.which("pr-latexmk"):
        print("  pr-latexmk not found; compile the .tex yourself.")
        return None
    result = subprocess.run(["pr-latexmk", str(tex)], capture_output=True, text=True)
    if result.returncode != 0:
        print(result.stdout[-3000:] + result.stderr[-2000:])
        sys.exit(f"pr-latexmk failed on {tex.name}; see above.")
    print(f"Wrote {tex.with_suffix('.pdf')}")
    return tex.with_suffix(".pdf")


def our_time(record, alternatives):
    """One of our courses in the plan: (time to draw, note).

    With alternatives, a course that has an Alternative time is drawn there,
    and the note says where it was.
    """
    suggested = record.get("Suggested time", "")
    alternative = record.get("Alternative time", "")
    if alternatives and pg.parse_options(alternative):
        return alternative, f"alternative; was {suggested}"
    return suggested, ""


# ----------------------------------------------------------------------------
# What to draw: a plan, one entry per term of the year
# A plan is a list of terms, each {"title": 'Fall 2026', "season": 'FA',
# "data": the term whose times are used ('2026/FA'), "courses": {course: record}},
# a record holding the "Suggested time" (and, from a proposal, an
# "Alternative time") the way a proposed timetable workbook does.
# ----------------------------------------------------------------------------
def year_terms(year):
    """'2026-27' -> [('Fall 2026', 'FA', '2026/FA'), ('Winter 2027', 'WI', '2027/WI')]."""
    match = re.fullmatch(r"(\d{4})-(\d{2})", year)
    if not match or (int(match[1]) + 1) % 100 != int(match[2]):
        sys.exit(f"{year!r} is not an academic year like 2026-27.")
    first = int(match[1])
    return [(f"Fall {first}", "FA", f"{first}/FA"), (f"Winter {first + 1}", "WI", f"{first + 1}/WI")]


def previous_year(year):
    """'2027-28' -> '2026-27'."""
    first = int(year[:4]) - 1
    return f"{first}-{(first + 1) % 100:02d}"


def plan_from_export(year, data_year, offerings, covered):
    """Our courses in each term of data_year, as the export has them, titled as terms of year.

    Every MATH/STAT/DATA course with an in-person lecture, at its lecture
    times; sections at different times are joined with '|' so the course
    appears at each. Online sections, tutorials and labs are left out.
    """
    plan = []
    for (title, season, _), (_, _, data) in zip(year_terms(year), year_terms(data_year)):
        if data not in covered:
            sys.exit(f"No export in {paths.TIMETABLE_EXPORTS} holds {bc.term_label(data)}; "
                     f"is there a timetable-{data_year}*.csv?")
        courses = {}
        for course, per in offerings.items():
            offering = per.get(data)
            if course[:-4] in pg.OURS and offering and offering["inperson"] and offering["lecture_slots"]:
                courses[course] = {"Suggested time": " | ".join(offering["lecture_slots"])}
        plan.append({"title": title, "season": season, "data": data, "courses": courses})
    return plan


def plan_from_proposal(year, workbook):
    """Our courses and times from a proposed timetable's 'Fall YYYY' and 'Winter YYYY' sheets."""
    sheets = {title: courses for title, _, courses in pg.read_proposal(workbook)}
    plan = []
    for title, season, data in year_terms(year):
        if title not in sheets:
            sys.exit(f"{workbook.name} has no {title!r} sheet; it has {', '.join(sheets)}.")
        plan.append({"title": title, "season": season, "data": data, "courses": sheets[title]})
    return plan


def other_time(offerings, course, data, season_terms):
    """Another unit's course: (time, note), or None if it does not run.

    From the term the plan's times come from, if the exports hold it: its
    lecture times, or for a course with labs and no lecture its lab times.
    If they do not (a proposal for a year not yet scheduled), from the last
    time it ran in person in that season, with a note saying when.
    """
    per = offerings.get(course, {})
    if data in season_terms:
        ran, note = [data], ""
    else:
        ran = [t for t in season_terms if per.get(t) and per[t]["inperson"]][-1:]
        note = f"{bc.term_label(ran[0])} time" if ran else ""
    if not ran or not (per.get(ran[0]) and per[ran[0]]["inperson"]):
        return None
    offering = per[ran[0]]
    if offering["lecture_slots"]:
        return " | ".join(offering["lecture_slots"]), note
    if offering["tutorial_slots"]:
        return " | ".join(offering["tutorial_slots"]), "lab" + (f", {note}s" if note else "")
    return None


def describe_source(mode, year, data_year, source):
    """A sentence for the footer: where the times come from."""
    if mode == "export":
        return f"Times from the registrar's timetable, {source.name}."
    if mode == "draft":
        return (f"A DRAFT for the timetable committee, not the {year} timetable: every course at its "
                f"{data_year} time, from {source.name}. Courses added or dropped for {year} are not shown.")
    return f"Our times as proposed in {source.name}."


# ----------------------------------------------------------------------------
# The department grid: every course of ours, one PDF per term
# ----------------------------------------------------------------------------
def department_grid(term, programs, others, folder, suffix, args, describe):
    slug = re.sub(r"\W+", "-", term["title"].lower())  # 'Fall 2026' -> 'fall-2026'
    output = folder / f"timetable-grid-{slug}{suffix}.tex"
    entries = []
    for key, record in term["courses"].items():
        # The small grey lines under the course: whether it moved, our
        # programs, then other units.
        time, moved = our_time(record, args.alternatives)
        small = []
        if ours := required_by(programs, key):
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

    footer = f"{describe} Drawn {datetime.date.today().isoformat()}. "
    if args.alternatives:
        footer += ("Every alternative time is taken: a course that moved is in red at its new time "
                   "(``alternative; was'' its suggested time), and struck out in grey in the slot it left. ")
    footer += ("Tutorials and labs are not shown. A course with sections at different times appears at each. "
               "Under each course, in grey: our degree programs that require it, "
               "read from the degree-requirements workbooks (``Hons:'' for the Honours track only); "
               "then, after ``Also:'', the other units with a program that requires it, from "
               f"{OTHER_UNITS.name}. "
               "Courses that are only one choice among several are not marked.")
    title = f"Mathematics and Statistics, {term['title']}: {args.label}"
    page, skipped = tex_page(title, entries, footer)
    output.write_text(tex_document([page]))
    print(f"Wrote {output}")
    if skipped:
        print("  Left out, because the time could not be read: " + ", ".join(skipped))
    compile_pdf(output, args.no_pdf)


# ----------------------------------------------------------------------------
# The program grids by level: a page per level per term, 50-minute and longer
# meetings side by side under each day
# ----------------------------------------------------------------------------
BY_LEVEL_COLUMNS = [
    ("short", "50 min", "coursefill", ""),
    ("long", "80 min, labs", "coursefill", ""),
]


def master_slot(text):
    """A master-list timeslot cell -> a time this script can draw, or ''.

    'MWF 11:30-12:20 + tut T 8:30-9:20 (new 2026 F)' -> 'MWF 11:30-12:20';
    '|' separates sections and is kept; '-' and '(no set time)' give ''.
    """
    pieces = []
    for piece in str(text or "").split("|"):
        piece = piece.split("+")[0]  # drop the tutorial
        if pg.parse_options(piece):
            pieces.append(piece.strip())
    return " | ".join(pieces)


def program_level_grids(plan, programs, offerings, covered, folder, suffix, args, describe, data_year):
    offered = set().union(*(term["courses"] for term in plan))  # ours that run, in either term
    # Every MATH/STAT course on the master list that is not inactive or on
    # hiatus, and its usual slot, for a course not running this year. One
    # that is inactive or on hiatus is left off the grids altogether, even
    # where a program names it, unless it is in fact running (Patrick, 2026-10-21).
    records = ba.sheet_records(load_workbook(bc.DEFAULT_OUTPUT, data_only=True)["MAIN"])
    dormant = {k for k, r in records.items() if r.get("Status", "").strip().lower() in ("inactive", "hiatus")}
    master = {k: r for k, r in records.items() if k not in dormant}
    upper = {c for c in master if level(c) == 3}
    not_running = f"not in {data_year}"
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
        # Our courses that only a pattern like MATH 3*** catches are
        # electives, and every upper-level course of ours is shown too:
        # all of them are options.
        electives = {c for c in master if any((r := pg.role_in(e, c)) and r[0] == "elective"
                                              for e in program["tracks"].values())}

        pages = []
        for term in plan:
            season_terms = [t for t in covered if t.endswith(term["season"])
                            and bc.term_key(t) <= bc.term_key(term["data"])][-pg.LOOKBACK:]
            slot_column = "Fall timeslot" if term["season"] == "FA" else "Winter timeslot"
            for lvl in LEVELS:
                entries = []
                for course in sorted(set(required) | options | electives | upper):
                    if level(course) != lvl:
                        continue
                    box = "framed" if course in required else "plain"
                    # Under a course: which track alone requires it, and
                    # whether it is not running this year. Nothing else.
                    small = []
                    if course in required and len(required[course]) < len(program["tracks"]):
                        small.append(" and ".join(required[course]) + " only")
                    if course[:-4] in pg.OURS:
                        # Running this term: its time. Running this year but
                        # the other term: not on this page. Not running this
                        # year at all: where it last sat in this season.
                        style = False
                        if course in term["courses"]:
                            time, moved = our_time(term["courses"][course], args.alternatives)
                            style = "moved" if moved else False
                        elif course in offered or course in dormant:
                            continue
                        else:
                            time = master_slot(master.get(course, {}).get(slot_column))
                            style = "notrunning"
                            small.append(not_running)
                        if not pg.parse_options(time):
                            continue  # no set time (a thesis), or never in this season
                        entries.append((course, time, small, style, "length", box))
                        continue
                    found = other_time(offerings, course, term["data"], season_terms)
                    if found:
                        time, _ = found
                        entries.append((course, time, small, True, "length", box))
                if not entries:
                    continue  # nothing at this level this term
                spec = BY_LEVEL_COLUMNS
                footer = (f"{program['title']}, {LEVEL_NAMES[lvl]} courses: under each day, 50-minute meetings "
                          "in the first column and longer ones (80 minutes, labs) in the second. "
                          "Required courses are framed in blue; the rest are options (named choices, "
                          "first-year science, and every 3000/4000-level MATH/STAT course). "
                          f"{describe} A course of ours not running in {data_year} is struck out, at the slot "
                          "it last had in this term. Other units' courses are in italic"
                          + ("" if term["data"] in season_terms else
                             ", at the time they had the last time they ran in person in this term")
                          + ". A course with sections at different times (or a lab) appears at each: choose one. "
                          f"Drawn {today}.")
                heading = f"{program['title']}, {term['title']}: {LEVEL_NAMES[lvl]} courses"
                if args.label != "timetable":
                    heading += f" ({args.label})"
                if args.alternatives:
                    footer += " Every alternative time is taken: a course that moved is in red."
                page, _ = tex_page(heading, entries, footer, spec)
                pages.append(page)
        output = folder / f"{program['name']}{suffix}.tex"
        output.write_text(tex_document(pages))
        print(f"Wrote {output} ({len(pages)} pages)")
        pdf = compile_pdf(output, args.no_pdf)
        # Advisors read a copy of each program's PDF, in their own folder.
        if pdf and args.copy_to:
            if copy := bc.copy_for_another_team(pdf, args.copy_to / f"{program['name']}-DO-NOT-EDIT.pdf"):
                print(f"  Copied for the Advising team to {copy}")


def main():
    parser = argparse.ArgumentParser(description="Draw a year's timetable as week grids.")
    parser.add_argument("year", help="the academic year to draw, e.g. 2026-27")
    parser.add_argument("--draft", action="store_true",
                        help="draw the year as a draft: last year's courses at last year's times")
    parser.add_argument("--proposal", metavar="WORKBOOK",
                        help="take our courses' times from a proposed timetable workbook instead")
    parser.add_argument("--alternatives", action="store_true",
                        help="with --proposal: draw every course that has an Alternative time at that time")
    parser.add_argument("--only", choices=["department", "programs"], help="draw only one kind of grid")
    parser.add_argument("--no-pdf", action="store_true", help="write the .tex only; do not run pr-latexmk")
    args = parser.parse_args()
    if args.draft and args.proposal:
        sys.exit("--draft and --proposal are two different drafts: choose one.")
    if args.alternatives and not args.proposal:
        sys.exit("--alternatives needs --proposal: only a proposed timetable has alternative times.")
    year_terms(args.year)  # stops on a malformed year

    mode = "proposal" if args.proposal else "draft" if args.draft else "export"
    data_year = previous_year(args.year) if args.draft else args.year
    folder = paths.TIMETABLE / args.year
    if not folder.is_dir():
        sys.exit(f"No folder {folder}; make it first, or check the year.")
    # Anything not from the registrar is marked in every file name, so it cannot
    # be taken for the real timetable.
    suffix = {"export": "", "draft": "-draft", "proposal": "-proposed"}[mode]
    suffix += "-alternative" if args.alternatives else ""
    args.label = {"export": "timetable", "draft": "DRAFT, last year's times",
                  "proposal": "proposed timetable"}[mode]
    if args.alternatives:
        args.label += ", with the alternative times"

    if args.proposal:
        source = Path(args.proposal).expanduser()
        plan = plan_from_proposal(args.year, source)
    offered = set().union(*(term["courses"] for term in plan)) if args.proposal else set()
    programs, offerings, covered, files, _ = pg.load_programs(
        ba.DEFAULT_REQUIREMENTS, ba.DEFAULT_OUTPUT, bc.DEFAULT_TIMETABLES, offered)
    if not args.proposal:
        source = next((f for f in files if bc.export_year(f) == data_year), None)
        if source is None:
            sys.exit(f"No timetable-{data_year}*.csv in {paths.TIMETABLE_EXPORTS}.")
        plan = plan_from_export(args.year, data_year, offerings, covered)
    describe = describe_source(mode, args.year, data_year, source)
    print(f"Drawing {args.year} ({args.label}) from {source.name}")

    if args.only != "programs":
        out = folder / "departmental-grids"
        out.mkdir(exist_ok=True)
        others = read_other_units(OTHER_UNITS)
        for term in plan:
            department_grid(term, programs, others, out, suffix, args, describe)
    if args.only != "department":
        # Only the registrar's timetable goes to the Advising team; drafts and
        # proposals are the committee's.
        args.copy_to = None
        if mode == "export" and paths.PROGRAM_GRIDS_COPY.is_dir():
            args.copy_to = paths.PROGRAM_GRIDS_COPY / args.year
            args.copy_to.mkdir(exist_ok=True)
        out = folder / "program-grids-by-level"
        out.mkdir(exist_ok=True)
        program_level_grids(plan, programs, offerings, covered, out, suffix, args, describe, data_year)


if __name__ == "__main__":
    main()
