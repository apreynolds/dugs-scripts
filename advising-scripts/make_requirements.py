#!/usr/bin/env -S uv run --script
# /// script
# dependencies = ["openpyxl"]
# ///
# make_requirements.py — create a degree-requirements workbook to edit by hand.
# Usage:
#   ./make_requirements.py OUT.xlsx --bsc-math       # BSc Math 2020 seed
#   ./make_requirements.py OUT.xlsx --from-tsv IN.tsv  # rebuild from the repo
#   ./make_requirements.py OUT.tsv  --from-xlsx IN.xlsx # export for committing
#   ./make_requirements.py OUT.xlsx                  # blank starter file
#   ./make_requirements.py -h | --help
#
# Requirements files are meant to be maintained in Excel by hand, so this
# script is a BOOTSTRAP, not part of the regular pipeline: run it once to get
# a correctly-shaped file, then edit that file directly. Re-running it on an
# existing path will overwrite your edits, so it refuses unless --force.
#
# Use --bsc-math to reproduce the BSc Mathematics 2020 requirements that were
# previously hardcoded into TEMPLATE-bsc-math-NEW.xlsx. For any other program,
# start from the blank file and fill it in.

import argparse
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(name="Arial", size=10, bold=True, color="FFFFFF")
BODY_FONT = Font(name="Arial", size=10)
HINT_FONT = Font(name="Arial", size=10, italic=True, color="7F7F7F")


# BSc Mathematics, 2020 calendar. Transcribed from the M:P block of
# TEMPLATE-bsc-math-NEW.xlsx.
#
# The "Courses" entries are the machine-readable satisfying-course lists that
# drive auto-categorization. Blank means "only an advisor can decide" — the
# Breadth categories need a judgement call about disciplinary coherence, and
# the first-year science lecture/lab split does not follow from the course
# number in a way that is safe to encode.
BSC_MATH_2020 = [
    # code,        description,                                   courses,                            major, hons, note
    ("Calc",      "MATH 1003, MATH 1013",                         "MATH 1003; MATH 1013",              6,  6, ""),
    ("Sci1001",   "Sci 1001",                                     "SCI 1001",                          1,  1, ""),
    ("SciLec",    "First Year Science Lectures",                  "",                                 18, 18, ""),
    ("SciLab",    "First Year Science Labs",                      "",                                  8,  8, ""),
    ("2ndMath",   "MATH 2003, 2013, 2203, 2213",                  "MATH 2003; MATH 2013; MATH 2203; MATH 2213", 12, 12, ""),
    ("Stat",      "STAT 3083, STAT 3093",                         "STAT 3083; STAT 3093",              6,  6,
     "STAT 2593 can count for 3083 with a grade of B or better"),
    ("CS",        "CS 1003 or 1073",                              "CS 1003; CS 1073",                  4,  4,
     "CS 1003 = Python, CS 1073 = Java. Recommended to take in 1st or 2nd year."),
    ("3rdMath",   "MATH 3033, 3213, 3243",                        "MATH 3033; MATH 3213; MATH 3243",   9,  9, ""),
    ("Analys",    "MATH 3003 or 3103",                            "MATH 3003; MATH 3103",              3,  3, ""),
    ("Breadth1",  "Breadth (3000/4000-level), same discipline",   "",                                  9,  9,
     "Cannot be STAT, as a max of 6ch can count as Breadth."),
    ("Breadth2",  "Breadth (3000/4000-level)",                    "",                                  6,  6, ""),
    ("MathElec",  "Upper Year MATH Electives",                    "MATH 3***; MATH 4***",             15, 15,
     "For Honours: at least 3 ch must be 4000-level. (See Constraints below.)"),
    ("AnalysH",   "MATH 3113",                                    "MATH 3113",                         0,  3, ""),
    ("Thesis",    "MATH 4100",                                    "MATH 4100",                         0,  6, ""),
    ("FreeElec",  "Free Electives",                               "",                                  0,  0, ""),
]

BSC_MATH_META = [
    ("Program", "BSc Mathematics"),
    ("Edition", "2020"),
    ("Slack category", "FreeElec"),
    ("Total CH: Major", 129),
    ("Total CH: Honours", 132),
]

BSC_MATH_CONSTRAINTS = [
    (1, "Max 6 ch STAT counted toward Breadth "
        "(any combination of Breadth1/Breadth2)", ""),
    (2, "Min 3 ch 4000-level MATH", "Honours"),
]

# A blank starter file: the same shape, one example row, nothing committed to.
BLANK_ROWS = [
    ("Example", "Describe the requirement here", "MATH 1003; MATH 1013",
     6, 6, "Optional tooltip shown in the tracker"),
    ("FreeElec", "Free Electives", "", 0, 0, ""),
]

BLANK_META = [
    ("Program", "<program name>"),
    ("Edition", "<calendar year>"),
    ("Slack category", "FreeElec"),
    ("Total CH: Major", 120),
    ("Total CH: Honours", 120),
]

BLANK_CONSTRAINTS = [
    (1, "<a rule the advisor must confirm by eye>", ""),
]


def style_header(sheet, row, count):
    """Apply the dark-blue header styling used across the tracker workbooks."""
    for column in range(1, count + 1):
        cell = sheet.cell(row=row, column=column)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center")


def set_widths(sheet, widths):
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width


def build_requirements_sheet(sheet, rows, tracks):
    header = ["Category", "Description", "Courses"] + tracks + ["Note"]
    sheet.append(header)
    style_header(sheet, 1, len(header))

    for row in rows:
        sheet.append(list(row))

    for line in sheet.iter_rows(min_row=2):
        for cell in line:
            cell.font = BODY_FONT
            cell.alignment = Alignment(vertical="center")

    set_widths(sheet, [14, 42, 46] + [10] * len(tracks) + [56])
    sheet.freeze_panes = "A2"


def build_meta_sheet(sheet, meta):
    sheet.append(["Field", "Value"])
    style_header(sheet, 1, 2)
    for key, value in meta:
        sheet.append([key, value])
    for line in sheet.iter_rows(min_row=2):
        for cell in line:
            cell.font = BODY_FONT
    set_widths(sheet, [24, 34])


def build_constraints_sheet(sheet, constraints):
    sheet.append(["#", "Constraint", "Tracks"])
    style_header(sheet, 1, 3)
    for number, text, tracks in constraints:
        sheet.append([number, text, tracks])
    for line in sheet.iter_rows(min_row=2):
        for cell in line:
            cell.font = BODY_FONT
    set_widths(sheet, [6, 72, 20])


# Guidance lives on its own sheet rather than below the data. Anything written
# under the Requirements table risks being read back as a category row, and a
# colleague adding their own notes there shouldn't be able to corrupt the file.
HOW_TO_USE = [
    ("Requirements sheet", None),
    ("Category", "Short unique code. This is what appears in the tracker's "
                 "Cat. dropdown, so keep it short."),
    ("Description", "Human-readable text, shown beside the category in the "
                    "tracker."),
    ("Courses", "Semicolon-separated list of course codes that satisfy this "
                "category, e.g. 'MATH 1003; MATH 1013'. A '*' stands for one "
                "digit, so 'MATH 3***' means any 3000-level MATH course."),
    ("", "An exact code beats a wildcard: with 'MATH 3033' listed under "
         "3rdMath and 'MATH 3***' under MathElec, MATH 3033 goes to 3rdMath."),
    ("", "Leave Courses BLANK when only an advisor can decide (Breadth, the "
         "science lecture/lab split). The tracker leaves those to you."),
    ("Track columns", "The COLUMN HEADING is the track name. Add a column to "
                      "add a track (e.g. Co-op), then add a matching "
                      "'Total CH: <track>' row on the Meta sheet."),
    ("Note", "Optional. Becomes a hover tooltip in the generated tracker."),
    ("", None),
    ("Meta sheet", None),
    ("Slack category", "Names the category that absorbs leftover credit "
                       "hours (e.g. FreeElec). Its required hours are "
                       "COMPUTED, so the totals always reconcile — leave its "
                       "numbers at 0 on the Requirements sheet."),
    ("Total CH: <track>", "One row per track, giving the total credit hours "
                          "the degree requires on that track."),
    ("", None),
    ("Constraints sheet", None),
    ("Constraint", "A rule the advisor confirms by eye because it is "
                   "unreliable to compute — e.g. whether 9 ch of Breadth form "
                   "a coherent discipline. Becomes a tick-box in the tracker."),
    ("Tracks", "Leave blank to apply to every track; otherwise list track "
               "names separated by semicolons."),
    ("", None),
    ("Editing this file", None),
    ("", "Keep it plain: values only, no formulas. Every formula belongs in "
         "the generated tracker, not here."),
    ("", "Adding or renaming a category is safe. The tracker is regenerated "
         "from this file each run, so its dropdowns and formula ranges "
         "follow automatically."),
]


def build_how_to_sheet(sheet):
    sheet.append(["Field", "What it means"])
    style_header(sheet, 1, 2)

    for label, text in HOW_TO_USE:
        sheet.append([label, text or ""])
        row = sheet.max_row
        label_cell = sheet.cell(row=row, column=1)
        text_cell = sheet.cell(row=row, column=2)
        if text is None and label:
            # A section heading.
            label_cell.font = Font(name="Arial", size=10, bold=True,
                                   color="1F4E79")
        else:
            label_cell.font = Font(name="Arial", size=10, bold=True)
            text_cell.font = BODY_FONT
        text_cell.alignment = Alignment(wrap_text=True, vertical="top")

    set_widths(sheet, [22, 96])


def build_workbook(rows, meta, constraints, tracks):
    workbook = openpyxl.Workbook()

    requirements = workbook.active
    requirements.title = "Requirements"
    build_requirements_sheet(requirements, rows, tracks)

    build_meta_sheet(workbook.create_sheet("Meta"), meta)
    build_constraints_sheet(workbook.create_sheet("Constraints"), constraints)
    build_how_to_sheet(workbook.create_sheet("How to use"))

    return workbook


# --- TSV round-trip ---------------------------------------------------------
#
# A requirements workbook is data, and data belongs in version control — but
# .xlsx is opaque to git, so a change to a calendar edition would show up as
# "binary files differ". These two functions round-trip the CONTENT (not the
# formatting) through tab-separated text, so the repo can track a .tsv and
# `git diff` shows which requirement actually moved.
#
# The .tsv is the committed source; the .xlsx is rebuilt from it and is what
# gets edited by hand. After editing, export back to .tsv and commit that.

TSV_SECTIONS = ("Meta", "Requirements", "Constraints")

TSV_PREAMBLE = [
    "# Degree-requirements data, tab-separated.",
    "#",
    "# This file is the version-controlled source. Rebuild the spreadsheet:",
    "#     make_requirements.py FILE.xlsx --from-tsv FILE.tsv",
    "# Edit the spreadsheet in Excel, then export back before committing:",
    "#     make_requirements.py FILE.tsv --from-xlsx FILE.xlsx",
    "#",
    "# Section headings in [brackets]; blank lines and #-comments ignored.",
]


def _sheet_rows(sheet):
    """Every used row of a sheet as a list of trimmed strings.

    Trailing empty cells are dropped so the TSV stays narrow, and rows that
    are entirely empty end the block — anything below is a colleague's own
    notes, matching how requirements_io reads the workbook.
    """
    rows = []
    for row in sheet.iter_rows(values_only=True):
        cells = ["" if v is None else str(v).strip() for v in row]
        while cells and cells[-1] == "":
            cells.pop()
        if not cells:
            break
        rows.append(cells)
    return rows


def export_tsv(xlsx_path, tsv_path):
    """Dump a requirements workbook to tab-separated text."""
    workbook = openpyxl.load_workbook(str(xlsx_path), data_only=True)

    lines = list(TSV_PREAMBLE)
    for name in TSV_SECTIONS:
        if name not in workbook.sheetnames:
            continue
        lines.append("")
        lines.append("[" + name + "]")
        for cells in _sheet_rows(workbook[name]):
            # A literal tab inside a description would corrupt the file.
            lines.append("\t".join(c.replace("\t", " ") for c in cells))

    Path(tsv_path).parent.mkdir(parents=True, exist_ok=True)
    Path(tsv_path).write_text("\n".join(lines) + "\n", encoding="utf-8")


def _is_comment(line):
    """Is this line a comment rather than data?

    The Constraints sheet's first column is literally headed "#", so a naive
    "starts with #" test swallows that header row. A comment is a "#" followed
    by whitespace or nothing; "#<tab>Constraint<tab>Tracks" is data.
    """
    if not line.startswith("#"):
        return False
    return len(line) == 1 or line[1] not in "\t"


def read_tsv(tsv_path):
    """Parse a requirements .tsv into {section name: [[cell, ...], ...]}."""
    sections = {}
    current = None

    for number, raw in enumerate(
            Path(tsv_path).read_text(encoding="utf-8").splitlines(), start=1):
        line = raw.rstrip("\n")
        if not line.strip() or _is_comment(line.lstrip()):
            continue
        stripped = line.strip()
        if stripped.startswith("[") and stripped.endswith("]"):
            current = stripped[1:-1]
            sections[current] = []
            continue
        if current is None:
            raise ValueError(
                str(tsv_path) + " line " + str(number) +
                ": data before any [Section] heading"
            )
        sections[current].append(line.split("\t"))

    missing = [s for s in ("Meta", "Requirements") if s not in sections]
    if missing:
        raise ValueError(
            str(tsv_path) + ": missing required section(s) " +
            ", ".join("[" + m + "]" for m in missing)
        )
    return sections


def _without_header(rows, first_heading):
    """Drop the header row if it's there, rather than assuming rows[0] is it.

    Blindly slicing [1:] silently eats the first data row whenever the header
    goes missing for any reason — which is exactly how the "#" comment clash
    above lost a constraint.
    """
    if rows and rows[0] and rows[0][0].strip() == first_heading:
        return rows[1:]
    return rows


def _as_number(text):
    """Convert a cell to int/float where it looks numeric, else leave it."""
    try:
        value = float(text)
    except (TypeError, ValueError):
        return text
    return int(value) if value == int(value) else value


def workbook_from_tsv(tsv_path):
    """Rebuild a requirements workbook from tab-separated text."""
    sections = read_tsv(tsv_path)

    req_rows = sections["Requirements"]
    header = req_rows[0]
    body = req_rows[1:]

    # Track names come from the header, exactly as requirements_io reads them.
    reserved = {"category", "description", "courses", "note", "notes"}
    tracks = [h for h in header[3:] if h and h.lower() not in reserved]

    # Pad short rows so every row has a cell per column, and convert the
    # credit-hour columns back to numbers so Excel stores them as numbers.
    width = len(header)
    rows = []
    for cells in body:
        padded = list(cells) + [""] * (width - len(cells))
        for index, name in enumerate(header):
            if name in tracks:
                padded[index] = _as_number(padded[index]) if padded[index] \
                    else 0
        rows.append(padded[:width])

    meta = []
    for cells in _without_header(sections["Meta"], "Field"):
        key = cells[0] if cells else ""
        value = cells[1] if len(cells) > 1 else ""
        if key:
            meta.append((key, _as_number(value)))

    constraints = []
    for cells in _without_header(sections.get("Constraints", []), "#"):
        number = _as_number(cells[0]) if cells else ""
        text = cells[1] if len(cells) > 1 else ""
        applies = cells[2] if len(cells) > 2 else ""
        if text:
            constraints.append((number, text, applies))

    return build_workbook(rows, meta, constraints, tracks)


def main():
    parser = argparse.ArgumentParser(
        description="Create or convert a degree-requirements file.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="The output format is taken from OUTPUT's extension: .xlsx to "
               "build a spreadsheet, .tsv to export one.",
    )
    parser.add_argument("output", help="path to write (.xlsx or .tsv)")
    source = parser.add_mutually_exclusive_group()
    source.add_argument(
        "--bsc-math", action="store_true",
        help="pre-fill with the BSc Mathematics 2020 requirements "
             "(default: a blank starter file)",
    )
    source.add_argument(
        "--from-tsv", metavar="TSV",
        help="build the .xlsx from a version-controlled .tsv",
    )
    source.add_argument(
        "--from-xlsx", metavar="XLSX",
        help="export an existing .xlsx to .tsv for committing",
    )
    parser.add_argument(
        "--force", action="store_true",
        help="overwrite the output file if it already exists",
    )
    args = parser.parse_args()

    output = Path(args.output)
    if output.exists() and not args.force:
        print("Refusing to overwrite " + str(output) +
              " (requirements files are hand-maintained). "
              "Pass --force if you really mean it.", file=sys.stderr)
        return 1

    if args.from_xlsx:
        if output.suffix.lower() != ".tsv":
            parser.error("--from-xlsx exports to a .tsv; "
                         "OUTPUT ends in " + repr(output.suffix))
        export_tsv(args.from_xlsx, output)
        print("Wrote " + str(output) + " (from " + args.from_xlsx + ")")
        return 0

    if output.suffix.lower() != ".xlsx":
        parser.error("building a requirements workbook writes a .xlsx; "
                     "OUTPUT ends in " + repr(output.suffix))

    if args.from_tsv:
        try:
            workbook = workbook_from_tsv(args.from_tsv)
        except ValueError as error:
            print(str(error), file=sys.stderr)
            return 1
        origin = " (from " + args.from_tsv + ")"
    elif args.bsc_math:
        workbook = build_workbook(
            BSC_MATH_2020, BSC_MATH_META, BSC_MATH_CONSTRAINTS,
            tracks=["Major", "Honours"],
        )
        origin = " (BSc Mathematics 2020)"
    else:
        workbook = build_workbook(
            BLANK_ROWS, BLANK_META, BLANK_CONSTRAINTS,
            tracks=["Major", "Honours"],
        )
        origin = " (blank starter)"

    output.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(str(output))
    print("Wrote " + str(output) + origin)
    return 0


if __name__ == "__main__":
    sys.exit(main())
