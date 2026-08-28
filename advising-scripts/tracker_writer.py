# tracker_writer.py — generate a student's degree-tracker workbook.
# Usage: imported by transcript_to_tracker.py; not run directly.
#
# The tracker is ALWAYS written from scratch. "Updating" a student's tracker
# means reading the old one as data (see tracker_reader.py), merging it with a
# fresh transcript, and generating a brand-new workbook here. Nothing is ever
# read-modify-written in place.
#
# That is a deliberate choice. openpyxl silently drops workbook parts it does
# not understand, and a round-trip would quietly degrade the file a little
# more each year. Generating fresh also means every formula range is COMPUTED
# from the data rather than drawn by hand, so the ranges cannot drift out of
# step with the table the way the original hand-built template's did.
#
# ---------------------------------------------------------------------------
# SHEET LAYOUT
# ---------------------------------------------------------------------------
#
#   row 1   NAME   <name>
#   row 2   ID     <id>                      ... gpa1 / gpa2 labels
#   row 3   Track  <dropdown>  Overall GPA <v> ... gpa1 / gpa2 values
#   row 4                                    DEGREE REQUIREMENTS (<edition>)
#   row 5   Courses table header             Category|Description|Req.CH|Progress
#   row 6+  course rows                      category rows, TOTAL, CONSTRAINTS
#
# Courses table columns:
#   Term | Course | Title | CH | Gr. | Pass | Ovr. | Cat. <edition>... | Notes
#        | gpa1 | gpa2
#
# "Pass" is derived, not entered: it looks up the passing bar for whichever
# category the row's Cat. cell names. Everything else in the table is either
# transcript data or the advisor's own judgement.
#
# One "Cat. <edition>" column and one requirements block per requirements file
# supplied, so a single file gives a one-edition tracker and two files
# reproduce the side-by-side comparison layout.

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import (Alignment, Border, Color, Font, PatternFill,
                             Side)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import FormulaRule

import checkbox_xlsx
from requirements_io import GRADE_RANK

SHEET_NAME = "Tracker"
HIDDEN_SHEET_NAME = "_TrackData"

# The Courses-table column holding each row's passing bar, looked up from the
# category the advisor assigned. Its name appears inside structured formula
# references, so it is written once here.
PASS_COLUMN = "Pass"

# Grade letters ordered worst-to-best, for building the rank expressions.
# GRADE_RANK itself lives in requirements_io because that is where a
# requirements file's Pass column is validated against it — the ladder is part
# of what the DATA may say, so the loader owns it and the generator borrows it.
RANKED_GRADES = sorted(GRADE_RANK, key=GRADE_RANK.get)

# The 4.3 grade scale used for every GPA calculation.
GPA_GRADES = ["A+", "A", "A-", "B+", "B", "B-", "C+", "C", "D"]
GPA_POINTS = [4.3, 4, 3.7, 3.3, 3, 2.7, 2.3, 2, 1]

GRADE_CHOICES = ["A+", "A", "A-", "B+", "B", "B-", "C+", "C", "D", "F",
                 "W", "IP", "CR"]

# The Ovr. (override) vocabulary. These read as what they DO — "incl." forces
# the course to count, "excl." forces it not to. The earlier "yes"/"no"
# wording was ambiguous: it left open what question was being answered, and
# "yes" could just as easily be read as "yes, override the default" (i.e.
# exclude it) as "yes, include it".
OVR_INCLUDE = "incl."
OVR_EXCLUDE = "excl."
OVR_INCLUDE_Q = '"' + OVR_INCLUDE + '"'
OVR_EXCLUDE_Q = '"' + OVR_EXCLUDE + '"'

# --- styling ----------------------------------------------------------------

FONT = "Arial"
SIZE = 10

# Colours are written as 8-digit ARGB with an explicit "FF" (opaque) alpha.
# Passing a 6-digit value makes openpyxl pad it with "00" — fully TRANSPARENT
# — which Excel treats as malformed inside a differential format.
HEADER_FILL = PatternFill("solid", fgColor="FF1F4E79")
HEADER_FONT = Font(name=FONT, size=SIZE, bold=True, color="FFFFFFFF")
BODY_FONT = Font(name=FONT, size=SIZE)

# The grade is the fact; the bar beside it is only a reference. Bolding one
# and greying the other puts them in a visual hierarchy, so a column of mostly
# identical "C"s reads as background rather than as data competing for
# attention.
GRADE_FONT = Font(name=FONT, size=SIZE, bold=True)
PASS_FONT = Font(name=FONT, size=SIZE, color="FF9C9C9C")
LABEL_FONT = Font(name=FONT, size=SIZE, bold=True, color="FF375623")
VALUE_FONT = Font(name=FONT, size=SIZE, bold=True)

# Excel hides its default gridlines underneath any cell fill, so a block that
# gets conditionally filled needs explicit light-grey borders to keep looking
# like a normal gridded table.
GRIDLINE = Side(style="thin", color="FFD4D4D4")
GRID_BORDER = Border(left=GRIDLINE, right=GRIDLINE,
                     top=GRIDLINE, bottom=GRIDLINE)

# --- differential formats, used ONLY by conditional formatting --------------
#
# A dxf is not an ordinary cell style and Excel is strict about it:
#
#   * a solid fill must put its colour in bgColor, NOT fgColor, and keep an
#     "automatic" fgColor (indexed 64). openpyxl's usual
#     PatternFill("solid", fgColor=<colour>) produces a dxf that Excel
#     rejects outright — it repairs the file and drops ALL conditional
#     formatting from the sheet.
#   * a dxf font carries only the DIFFERENCES from the cell's own font, so it
#     must not restate the family and size. The cell is already Arial 10.
#
# The shape below is byte-for-byte what Excel itself writes, verified against
# the dxfs in the hand-built TEMPLATE-bsc-math-NEW.xlsx.
AUTOMATIC = Color(indexed=64)

MET_FILL = PatternFill(patternType="solid", fgColor=AUTOMATIC,
                       bgColor="FFC6EFCE")
MET_FONT = Font(color="FF006100")
OFF_FILL = PatternFill(patternType="solid", fgColor=AUTOMATIC,
                       bgColor="FFE7E6E6")
OFF_FONT = Font(color="FF7F7F7F", strike=True)

# Excel's own "Good" and "Bad" colour pairs, used to make the Ovr. column
# self-reporting. Nothing here judges whether an override was WISE — the point
# is only that a sheet carrying several of them looks different at a glance
# from one carrying none, so the exceptions are countable by eye.
INCL_FILL = PatternFill(patternType="solid", fgColor=AUTOMATIC,
                        bgColor="FFC6EFCE")
INCL_FONT = Font(color="FF006100")
EXCL_FILL = PatternFill(patternType="solid", fgColor=AUTOMATIC,
                        bgColor="FFFFC7CE")
EXCL_FONT = Font(color="FF9C0006")


# --- array-safe grade tests -------------------------------------------------
#
# Every grade test below is written as a sum of explicit equality comparisons
# rather than the more compact MATCH/CHOOSE idiom the original template used.
#
# The reason is that SUMPRODUCT feeds a whole COLUMN into these expressions,
# and neither MATCH nor CHOOSE reliably operates elementwise on an array:
#
#   IFERROR(CHOOSE(MATCH(Courses[Gr.],{...},0), 4.3,4,...), 0)
#
# CHOOSE rejects an array index, returns #VALUE!, and the IFERROR then
# collapses the entire factor to the SCALAR 0 — so the GPA numerator silently
# became 0 while the denominator still computed, and every GPA read 0.00.
#
# ISNUMBER(MATCH(...)) fails more insidiously: it collapses to a single
# TRUE/FALSE for the whole column, which made Progress count every course in a
# category regardless of grade — including failures.
#
# `(range="A+")` IS elementwise in every Excel version, so summing those is
# verbose but correct. A grade can satisfy at most one term, so each sum is 0
# or 1 and can be multiplied like a boolean.


def _grade_rank(column_name):
    """Array-safe ordinal rank of a grade column. Yields 0 for anything
    off the ladder — F, W, IP, CR and blank.

    Takes the column name because the SAME expression is needed for two
    columns: the student's grade (Courses[Gr.]) and the bar it has to clear
    (Courses[Pass]). Comparing two of these keeps both sides elementwise,
    which is what a lookup inlined into the SUMPRODUCT would destroy.
    """
    return "(" + "+".join(
        "(Courses[" + column_name + ']="' + grade + '")*' +
        str(GRADE_RANK[grade])
        for grade in reversed(RANKED_GRADES)
    ) + ")"


def _grade_points():
    """Array-safe 4.3-scale grade points. Yields 0 for any non-GPA grade."""
    return "(" + "+".join(
        '(Courses[Gr.]="' + grade + '")*' + str(points)
        for grade, points in zip(GPA_GRADES, GPA_POINTS)
    ) + ")"


class Block:
    """Where one requirements table sits on the sheet."""

    def __init__(self, requirements, cat_column_name, first_column):
        self.req = requirements
        self.cat_column_name = cat_column_name
        # Four visible columns plus two hidden helpers.
        self.col_category = first_column
        self.col_description = first_column + 1
        self.col_req = first_column + 2
        self.col_progress = first_column + 3
        self.col_earned = first_column + 4
        # Each category's passing bar, as a letter. Hidden: the advisor sees
        # it per COURSE in the Courses table's Pass column, which looks its
        # value up here, so showing it twice would just be noise.
        self.col_pass = first_column + 5

    @property
    def width(self):
        return 6

    def letter(self, column):
        return get_column_letter(column)


def _ordered_categories(requirements):
    """Category rows in display order, with the slack category last.

    The slack row's formula subtracts the sum of every OTHER category's
    required hours. Excel treats a SUMIF whose range includes the formula's
    own cell as a circular reference even when the criterion excludes it, so
    the slack row has to sit below the range it sums.
    """
    fixed = [c for c in requirements.categories
             if c.code != requirements.slack_category]
    slack = [c for c in requirements.categories
             if c.code == requirements.slack_category]
    return fixed + slack


# The Courses table is these, then one Cat. column per edition, then these.
# Kept as named lists because the sheet's geometry is derived from their
# LENGTHS before any block exists — the two used to be separate magic numbers
# that had to be edited in step, which is exactly the kind of arithmetic that
# goes wrong the once.
COURSES_LEADING = ["Term", "Course", "Title", "CH", "Gr.", PASS_COLUMN, "Ovr."]
COURSES_TRAILING = ["Notes", "gpa1", "gpa2"]


def _courses_columns(blocks):
    """Column headings for the Courses table, in order."""
    return (list(COURSES_LEADING)
            + [block.cat_column_name for block in blocks]
            + list(COURSES_TRAILING))


def _truthy(column_name):
    """An Excel test for "this flag cell is switched on".

    The original template used native Excel checkboxes, which are a
    Microsoft 365-only feature: they do not render in Office 2019/2021 or
    LibreOffice, and a generated workbook cannot reliably write them. They are
    replaced here by a TRUE/FALSE dropdown.

    Comparing with `=TRUE` alone would be fragile, because a dropdown can
    deposit either the boolean TRUE or the text "TRUE" depending on how the
    value was entered. Concatenating an empty string coerces both to the text
    "TRUE", so this test accepts either.
    """
    return 'UPPER(Courses[' + column_name + ']&"")="TRUE"'


def _passing_expression():
    """The "this course earns its credit hours" test, per row.

    Two terms that must never both fire for the same row:

      met     the grade clears its category's bar, or is transfer credit,
              and the advisor has not forced it out
      rescue  the advisor has forced IN a real grade that fell short

    The bar is per row, read from the Pass column, which the advisor's Cat.
    choice resolves. So a D earns credit in a category whose bar is D and
    earns nothing in one whose bar is C — which is the whole point.

    The `rank < bar` factor in the rescue term is LOAD-BEARING, not tidiness.
    Without it an "incl." on a course that already clears its bar satisfies
    both terms, and SUMPRODUCT counts its credit hours TWICE. The old rule
    was safe from this only by accident: it rescued "D" alone, and D was not
    in the passing list, so the two terms could not overlap.

    `rank > 0` is what keeps F out of the rescue. F, W, IP and a blank grade
    all rank 0, so "incl." cannot resurrect a failure or switch on a planned
    course that has not been taken yet.
    """
    rank = _grade_rank("Gr.")
    bar = _grade_rank(PASS_COLUMN)
    met = ("((" + rank + ">=" + bar + ')+(Courses[Gr.]="CR"))'
           "*(Courses[Ovr.]<>" + OVR_EXCLUDE_Q + ")")
    rescue = ("(" + rank + ">0)*(" + rank + "<" + bar + ")"
              "*(Courses[Ovr.]=" + OVR_INCLUDE_Q + ")")
    return met + "+" + rescue


def _progress_expression(cat_column_name, category_cell=None):
    """SUMPRODUCT of credit hours earned, optionally filtered to a category.

    Reads as: for courses where (the category matches, if we're filtering)
    AND the course earns its credit hours, sum those credit hours.
    """
    passing = _passing_expression()
    if category_cell is None:
        return "SUMPRODUCT((" + passing + ")*Courses[CH])"
    return (
        "SUMPRODUCT(((Courses[" + cat_column_name + "]=" + category_cell + ")"
        "*(" + passing + "))*Courses[CH])"
    )


def _gpa_formula(flag_column=None):
    """Weighted GPA on the 4.3 scale, optionally filtered by a flag column.

    Blank, W, IP and CR grades are excluded from both the numerator and the
    denominator: they carry no grade points and no attempted hours.
    """
    exclusions = (
        '(Courses[Gr.]<>"")*(Courses[Gr.]<>"W")'
        '*(Courses[Gr.]<>"IP")*(Courses[Gr.]<>"CR")'
    )
    flag = "(" + _truthy(flag_column) + ")*" if flag_column else ""
    points = _grade_points()
    return (
        "=IFERROR("
        "SUMPRODUCT(" + flag + exclusions + "*" + points + "*Courses[CH])"
        "/SUMPRODUCT(" + flag + exclusions + "*Courses[CH])"
        ',"N/A")'
    )


def _write_hidden_track_data(workbook, blocks):
    """Write the per-track required-hours lookup tables on a hidden sheet.

    The original template hardcoded `=IF($B$3,3,0)` in every Honours-only row,
    which only ever supports two tracks. Instead each block gets a small table
    here — one column per track — and the visible Req. CH cells look their
    value up by the track name chosen in B3. Adding a track to a requirements
    file therefore needs no change to any formula.

    Returns a dict mapping each block to its (header_row, first_data_row,
    total_row, first_col, last_col) on the hidden sheet.
    """
    sheet = workbook.create_sheet(HIDDEN_SHEET_NAME)
    sheet.sheet_state = "hidden"
    sheet["A1"] = ("Lookup tables for the Track selector on the Tracker "
                   "sheet. Generated - do not edit by hand.")

    positions = {}
    row = 3

    for block in blocks:
        tracks = block.req.tracks
        header_row = row
        sheet.cell(row=row, column=1, value="Category (" + block.req.label + ")")
        for offset, track in enumerate(tracks):
            sheet.cell(row=row, column=2 + offset, value=track)
        row += 1

        first_data_row = row
        for category in _ordered_categories(block.req):
            sheet.cell(row=row, column=1, value=category.code)
            for offset, track in enumerate(tracks):
                sheet.cell(row=row, column=2 + offset,
                           value=category.req_ch.get(track, 0))
            row += 1

        total_row = row
        sheet.cell(row=row, column=1, value="TOTAL")
        for offset, track in enumerate(tracks):
            sheet.cell(row=row, column=2 + offset,
                       value=block.req.totals.get(track, 0))
        row += 2

        positions[id(block)] = {
            "header_row": header_row,
            "first_data_row": first_data_row,
            "total_row": total_row,
            "first_col": 2,
            "last_col": 1 + len(tracks),
        }

    return positions


def common_tracks(requirement_sets):
    """Tracks every supplied edition defines, in the first edition's order.

    B3 drives every requirements block at once, so a track that only some
    editions know about could not be resolved by the others. Offering the
    intersection means every selectable track works in every block. The caller
    warns about anything dropped, so a missing "Co-op" is explained rather
    than mysterious.
    """
    shared = set(requirement_sets[0].tracks)
    for requirements in requirement_sets[1:]:
        shared &= set(requirements.tracks)
    ordered = [t for t in requirement_sets[0].tracks if t in shared]
    # Falling back to the first edition's tracks keeps a pathological file
    # (no overlap at all) producing a usable workbook rather than crashing.
    return ordered or list(requirement_sets[0].tracks)


def dropped_tracks(requirement_sets):
    """Tracks that some edition defines but which are not offered."""
    offered = set(common_tracks(requirement_sets))
    dropped = {}
    for requirements in requirement_sets:
        missing = [t for t in requirements.tracks if t not in offered]
        if missing:
            dropped[requirements.label] = missing
    return dropped


def _track_lookup(position, data_row):
    """An INDEX/MATCH that resolves one hidden-sheet row for the chosen track.

    Wrapped in IFERROR so that a track typed by hand into B3 that an edition
    doesn't define reads as 0 — which greys the row out — rather than
    spraying #N/A across the block.
    """
    first = get_column_letter(position["first_col"])
    last = get_column_letter(position["last_col"])
    header = position["header_row"]
    return (
        "IFERROR(INDEX(" + HIDDEN_SHEET_NAME + "!$" + first + "$" +
        str(data_row) + ":$" + last + "$" + str(data_row) + ","
        "MATCH($B$3," + HIDDEN_SHEET_NAME + "!$" + first + "$" + str(header) +
        ":$" + last + "$" + str(header) + ",0)),0)"
    )


def build_tracker(path, requirement_sets, rows, student,
                  track=None, spare_rows=15, checkboxes=True):
    """Generate the tracker workbook.

    requirement_sets — list of Requirements objects, one per edition
    rows            — list of course dicts (term/course/title/grade/hours,
                      plus optional cat_<edition>/ovr/notes/gpa1/gpa2)
    student         — dict with 'name', 'id'
    track           — which track to preselect in B3
    spare_rows      — blank rows left at the foot of the Courses table so an
                      advisor can add a course by hand
    checkboxes      — render the gpa1/gpa2 flags as native Excel checkboxes.
                      Needs Excel 365; False falls back to the TRUE/FALSE
                      dropdowns, which work everywhere. Either way the cells
                      hold booleans and read back identically.
    """
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = SHEET_NAME

    # --- work out the geometry before writing anything ----------------------

    blocks = []
    used_labels = {}
    # Courses table is COURSES_LEADING + one Cat. column per edition +
    # COURSES_TRAILING, then one blank spacer column before the first
    # requirements block.
    courses_width = (len(COURSES_LEADING) + len(requirement_sets)
                     + len(COURSES_TRAILING))
    first_block_column = courses_width + 2

    column = first_block_column
    for requirements in requirement_sets:
        label = requirements.label
        # Two editions could carry the same label; keep the column names unique.
        if label in used_labels:
            used_labels[label] += 1
            label = label + " (" + str(used_labels[label]) + ")"
        else:
            used_labels[label] = 1
        block = Block(requirements, "Cat. " + label, column)
        blocks.append(block)
        column += block.width + 1  # one blank spacer column between blocks

    headings = _courses_columns(blocks)
    table_first_row = 5
    data_first_row = 6
    data_last_row = data_first_row + max(len(rows) + spare_rows, 20) - 1

    tracks = common_tracks(requirement_sets)
    chosen_track = track if track in tracks else tracks[0]

    positions = _write_hidden_track_data(workbook, blocks)

    # --- top section --------------------------------------------------------

    sheet["A1"] = "NAME"
    sheet["B1"] = student.get("name", "")
    sheet["A2"] = "ID"
    sheet["B2"] = student.get("id", "")
    sheet["A3"] = "Track"
    sheet["B3"] = chosen_track

    for address in ("A1", "A2", "A3"):
        sheet[address].font = LABEL_FONT
    for address in ("B1", "B2", "B3"):
        sheet[address].font = VALUE_FONT
    sheet["B1"].alignment = Alignment(horizontal="left")
    sheet["B2"].alignment = Alignment(horizontal="center")
    sheet["B3"].alignment = Alignment(horizontal="center")

    sheet["C3"] = "Overall GPA"
    sheet["C3"].font = LABEL_FONT
    sheet["C3"].alignment = Alignment(horizontal="right")
    sheet["D3"] = _gpa_formula()
    sheet["D3"].font = VALUE_FONT
    sheet["D3"].alignment = Alignment(horizontal="center")
    sheet["D3"].number_format = "0.00"

    # The two sub-GPA readouts sit directly above the flag columns that drive
    # them, so it is obvious which column feeds which number.
    # Resolve every column we touch by NAME. Inserting the Pass column shifted
    # Ovr. from F to G, and the literals that used to be dotted about here
    # would have gone on quietly addressing the wrong column.
    grade_col = headings.index("Gr.") + 1
    pass_col = headings.index(PASS_COLUMN) + 1
    ovr_col = headings.index("Ovr.") + 1
    gpa1_col = headings.index("gpa1") + 1
    gpa2_col = headings.index("gpa2") + 1
    for column_index, name in ((gpa1_col, "gpa1"), (gpa2_col, "gpa2")):
        letter = get_column_letter(column_index)
        sheet[letter + "2"] = name
        sheet[letter + "2"].font = LABEL_FONT
        sheet[letter + "2"].alignment = Alignment(horizontal="center")
        sheet[letter + "3"] = _gpa_formula(name)
        sheet[letter + "3"].font = VALUE_FONT
        sheet[letter + "3"].alignment = Alignment(horizontal="center")
        sheet[letter + "3"].number_format = "0.00"

    # --- courses table ------------------------------------------------------

    for index, heading in enumerate(headings, start=1):
        cell = sheet.cell(row=table_first_row, column=index, value=heading)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")

    cat_columns = {}
    for block in blocks:
        cat_columns[block.cat_column_name] = \
            headings.index(block.cat_column_name) + 1

    for offset, row in enumerate(rows):
        excel_row = data_first_row + offset
        sheet.cell(row=excel_row, column=1, value=row.get("term", ""))
        sheet.cell(row=excel_row, column=2, value=row.get("course", ""))
        sheet.cell(row=excel_row, column=3, value=row.get("title", ""))
        hours = row.get("hours", "")
        sheet.cell(row=excel_row, column=4,
                   value=hours if hours != "" else None)
        sheet.cell(row=excel_row, column=grade_col, value=row.get("grade", ""))
        sheet.cell(row=excel_row, column=ovr_col, value=row.get("ovr", ""))
        for block in blocks:
            key = "cat_" + block.req.label
            sheet.cell(row=excel_row,
                       column=cat_columns[block.cat_column_name],
                       value=row.get(key, ""))
        sheet.cell(row=excel_row, column=headings.index("Notes") + 1,
                   value=row.get("notes", ""))
        sheet.cell(row=excel_row, column=gpa1_col,
                   value=bool(row.get("gpa1", False)))
        sheet.cell(row=excel_row, column=gpa2_col,
                   value=bool(row.get("gpa2", False)))

    # Spare rows still need their flag cells populated, or the GPA formulas
    # would see empty cells where they expect a boolean.
    for excel_row in range(data_first_row + len(rows), data_last_row + 1):
        sheet.cell(row=excel_row, column=gpa1_col, value=False)
        sheet.cell(row=excel_row, column=gpa2_col, value=False)

    # --- the Pass column ----------------------------------------------------
    #
    # A course's bar is the bar of the category the advisor filed it under, so
    # each cell is an INDEX/MATCH into the first requirements block.
    #
    # This must be a REAL CELL and not the same lookup inlined into the
    # Progress SUMPRODUCT. MATCH does not broadcast: handed a whole column as
    # its lookup value it collapses to one scalar, and every row would then be
    # measured against row 6's threshold — silently. That is trap 2 in
    # CLAUDE.md wearing a different hat.
    #
    # With two editions the first one wins. Two editions normally means two
    # calendar years of one degree, whose bars agree. A BA block beside a BSc
    # block would need two different verdicts on the same row, and is not a
    # supported layout.
    primary = blocks[0]
    block_first_row = 6
    primary_last_cat = (block_first_row
                        + len(_ordered_categories(primary.req)) - 1)
    pass_range = ("$" + primary.letter(primary.col_pass) + "$" +
                  str(block_first_row) + ":$" +
                  primary.letter(primary.col_pass) + "$" +
                  str(primary_last_cat))
    code_range = ("$" + primary.letter(primary.col_category) + "$" +
                  str(block_first_row) + ":$" +
                  primary.letter(primary.col_category) + "$" +
                  str(primary_last_cat))
    courses_cat_letter = get_column_letter(
        cat_columns[primary.cat_column_name])

    for excel_row in range(data_first_row, data_last_row + 1):
        # An uncategorised row still needs a bar to clear: the block TOTAL
        # counts those rows, and a blank cell here would rank 0 and wave
        # everything through. Hence the default as the IFERROR fallback.
        sheet.cell(row=excel_row, column=pass_col, value=(
            "=IFERROR(INDEX(" + pass_range + ",MATCH($" +
            courses_cat_letter + str(excel_row) + "," + code_range +
            ',0)),"' + primary.req.default_pass_grade + '")'))

    for excel_row in range(data_first_row, data_last_row + 1):
        for column_index in range(1, len(headings) + 1):
            cell = sheet.cell(row=excel_row, column=column_index)
            if column_index == grade_col:
                cell.font = GRADE_FONT
            elif column_index == pass_col:
                cell.font = PASS_FONT
            else:
                cell.font = BODY_FONT
            if column_index in (1, 2, 4, grade_col, pass_col, ovr_col,
                                gpa1_col, gpa2_col):
                cell.alignment = Alignment(horizontal="center")
            elif column_index in cat_columns.values():
                cell.alignment = Alignment(horizontal="center")

    table = Table(
        displayName="Courses",
        ref="A" + str(table_first_row) + ":" +
            get_column_letter(len(headings)) + str(data_last_row),
    )
    # No banded rows. The stripe colour and the "not on this track" grey in
    # the requirements block were near enough identical that the second one
    # stopped registering as meaningful. Grey now means exactly one thing in
    # this workbook — struck out, not on your track — and nothing else in it
    # is grey. The header keeps its own explicit fill, which beats the table
    # style, so only the banding goes.
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleLight1", showRowStripes=False)
    sheet.add_table(table)

    # Colour the overrides. Like every other rule here the formula is written
    # for the TOP-LEFT cell of the range and re-evaluated relatively, so the
    # reference is column-absolute but ROW-relative ($G6, not $G$6) — with
    # $G$6 every row would test row 6 and the column would light up as one.
    ovr_letter = get_column_letter(ovr_col)
    ovr_range = ("$" + ovr_letter + "$" + str(data_first_row) +
                 ":$" + ovr_letter + "$" + str(data_last_row))
    for value, fill, font in ((OVR_INCLUDE, INCL_FILL, INCL_FONT),
                              (OVR_EXCLUDE, EXCL_FILL, EXCL_FONT)):
        sheet.conditional_formatting.add(
            ovr_range,
            FormulaRule(
                formula=['$' + ovr_letter + str(data_first_row) + '="' +
                         value + '"'],
                fill=fill, font=font, stopIfTrue=False),
        )

    # --- requirements blocks ------------------------------------------------

    for block in blocks:
        position = positions[id(block)]
        categories = _ordered_categories(block.req)

        cat_letter = block.letter(block.col_category)
        req_letter = block.letter(block.col_req)
        prog_letter = block.letter(block.col_progress)
        earned_letter = block.letter(block.col_earned)

        heading_cell = sheet.cell(
            row=4, column=block.col_category,
            value="DEGREE REQUIREMENTS (" + block.req.label + ")")
        heading_cell.font = Font(name=FONT, size=SIZE, bold=True,
                                 color="1F4E79")

        for offset, name in enumerate(
                ["Category", "Description", "Req. CH", "Progress"]):
            cell = sheet.cell(row=5, column=block.col_category + offset,
                              value=name)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = Alignment(horizontal="center")

        first_row = 6
        for offset, category in enumerate(categories):
            excel_row = first_row + offset
            hidden_row = position["first_data_row"] + offset

            sheet.cell(row=excel_row, column=block.col_category,
                       value=category.code)
            # pass_grade_for() resolves the file's default, so this is always
            # a real letter. A blank here would rank 0 and every grade would
            # clear the bar, failures included.
            sheet.cell(row=excel_row, column=block.col_pass,
                       value=block.req.pass_grade_for(category.code))
            description_cell = sheet.cell(
                row=excel_row, column=block.col_description,
                value=category.description)
            if category.note:
                comment = Comment(category.note, "Degree requirements")
                comment.width = 320
                comment.height = 90
                description_cell.comment = comment

            if category.code == block.req.slack_category:
                # The slack category absorbs whatever the other categories do
                # not account for, so the totals always reconcile.
                last_fixed = excel_row - 1
                sheet.cell(
                    row=excel_row, column=block.col_req,
                    value="=" + _track_lookup(position, position["total_row"])
                    + "-SUM($" + req_letter + "$" + str(first_row) +
                    ":$" + req_letter + "$" + str(last_fixed) + ")")
            else:
                sheet.cell(row=excel_row, column=block.col_req,
                           value="=" + _track_lookup(position, hidden_row))

            sheet.cell(
                row=excel_row, column=block.col_earned,
                value="=" + _progress_expression(
                    block.cat_column_name,
                    "$" + cat_letter + "$" + str(excel_row)))
            sheet.cell(
                row=excel_row, column=block.col_progress,
                value='=TEXT($' + earned_letter + "$" + str(excel_row) +
                      ',"0")&" / "&TEXT($' + req_letter + "$" +
                      str(excel_row) + ',"0")&" ch"')

        total_row = first_row + len(categories)
        last_category_row = total_row - 1

        sheet.cell(row=total_row, column=block.col_category, value="TOTAL")
        sheet.cell(row=total_row, column=block.col_req,
                   value="=" + _track_lookup(position, position["total_row"]))
        sheet.cell(row=total_row, column=block.col_earned,
                   value="=" + _progress_expression(block.cat_column_name))
        sheet.cell(
            row=total_row, column=block.col_progress,
            value='=TEXT($' + earned_letter + "$" + str(total_row) +
                  ',"0")&" / "&TEXT($' + req_letter + "$" + str(total_row) +
                  ',"0")&" ch"')

        for excel_row in range(5, total_row + 1):
            for column_index in range(block.col_category,
                                      block.col_progress + 1):
                cell = sheet.cell(row=excel_row, column=column_index)
                cell.border = GRID_BORDER
                if excel_row > 5:
                    cell.font = BODY_FONT
                if column_index in (block.col_req, block.col_progress):
                    cell.alignment = Alignment(horizontal="center")
        for column_index in range(block.col_category, block.col_progress + 1):
            sheet.cell(row=total_row, column=column_index).font = Font(
                name=FONT, size=SIZE, bold=True)

        # Hide the numeric helper column. It exists so the conditional
        # formatting can compare two NUMBERS; the original template compared
        # by pulling the numbers back out of the "12 / 15 ch" display string
        # with FIND and MID, which is brittle.
        sheet.column_dimensions[earned_letter].hidden = True
        sheet.column_dimensions[block.letter(block.col_pass)].hidden = True

        # --- conditional formatting on this block ---------------------------

        # A conditional-formatting formula is written for the TOP-LEFT cell of
        # its range and then re-evaluated relative to every other cell. So the
        # references here must be column-absolute but ROW-RELATIVE ($N6, not
        # $N$6) — otherwise every row would test row 6's condition and the
        # whole column would light up together.
        progress_range = ("$" + prog_letter + "$" + str(first_row) +
                          ":$" + prog_letter + "$" + str(total_row))
        sheet.conditional_formatting.add(
            progress_range,
            FormulaRule(
                formula=["AND($" + req_letter + str(first_row) + ">0,$" +
                         earned_letter + str(first_row) + ">=$" +
                         req_letter + str(first_row) + ")"],
                fill=MET_FILL, font=MET_FONT, stopIfTrue=False),
        )

        # A category that requires 0 ch on the selected track is not part of
        # that track — grey it out and strike it through. Keying off the
        # resolved requirement rather than an "is Honours?" flag means this
        # works for any track a requirements file cares to define.
        row_range = ("$" + cat_letter + "$" + str(first_row) +
                     ":$" + prog_letter + "$" + str(last_category_row))
        sheet.conditional_formatting.add(
            row_range,
            FormulaRule(
                formula=["$" + req_letter + str(first_row) + "=0"],
                fill=OFF_FILL, font=OFF_FONT, stopIfTrue=True),
        )

        # --- constraints ----------------------------------------------------

        constraints_row = total_row + 2
        heading = sheet.cell(row=constraints_row, column=block.col_category,
                             value="CONSTRAINTS")
        heading.font = Font(name=FONT, size=SIZE, bold=True, color="1F4E79")

        for offset, constraint in enumerate(block.req.constraints):
            excel_row = constraints_row + 1 + offset
            sheet.cell(row=excel_row, column=block.col_category,
                       value=constraint.number)
            text_cell = sheet.cell(row=excel_row,
                                   column=block.col_description,
                                   value=constraint.text)
            text_cell.font = BODY_FONT
            # Deliberately NOT wrapped. Nothing in this workbook wraps: a
            # constraint too long to fit is a constraint that wants shortening
            # with the rest of it moved to its Note, which becomes a hover
            # tooltip. Growing the row instead just hides the problem.
            text_cell.alignment = Alignment(vertical="center")
            if constraint.note:
                # Same construction as the category notes above, so the two
                # kinds of tooltip look and behave identically.
                comment = Comment(constraint.note, "Degree requirements")
                comment.width = 320
                comment.height = 90
                text_cell.comment = comment
            tick = sheet.cell(row=excel_row, column=block.col_progress,
                              value=False)
            tick.alignment = Alignment(horizontal="center")
            tick.font = BODY_FONT

            if constraint.tracks:
                # Grey out a constraint that does not apply to the chosen
                # track. Written as a chain of <> tests rather than
                # MATCH against an array constant: conditional-formatting
                # formulas do not accept array constants, and Excel discards
                # the whole sheet's formatting if one appears.
                tests = ['$B$3<>"' + name + '"' for name in constraint.tracks]
                condition = (tests[0] if len(tests) == 1
                             else "AND(" + ",".join(tests) + ")")
                sheet.conditional_formatting.add(
                    "$" + cat_letter + "$" + str(excel_row) + ":$" +
                    prog_letter + "$" + str(excel_row),
                    FormulaRule(formula=[condition], fill=OFF_FILL,
                                font=OFF_FONT, stopIfTrue=True),
                )

        if block.req.constraints:
            # A ticked constraint goes green, exactly as a met category does
            # in the Progress column just above it — same MET_FILL/MET_FONT,
            # so "green means satisfied" reads the same way down the whole
            # column.
            #
            # Added AFTER the per-row grey rules above, so those keep the
            # higher priority: a constraint that does not apply to the
            # selected track stays greyed and struck through even if it was
            # left ticked from an earlier track. (Their stopIfTrue=True is
            # what makes that stick.)
            tick_range = ("$" + prog_letter + "$" + str(constraints_row + 1) +
                          ":$" + prog_letter + "$" +
                          str(constraints_row + len(block.req.constraints)))
            # Tested as text rather than as `=TRUE`, because the cell can hold
            # either. The tick is WRITTEN as a boolean, and Excel coerces a
            # dropdown pick back to one — but a hand-typed or pasted "TRUE" is
            # a string, and in Excel a string never equals the boolean TRUE.
            # `&""` renders both as "TRUE", which is the same tolerance
            # tracker_reader._flag() already applies when reading flags back.
            sheet.conditional_formatting.add(
                tick_range,
                FormulaRule(
                    formula=['UPPER($' + prog_letter +
                             str(constraints_row + 1) + '&"")="TRUE"'],
                    fill=MET_FILL, font=MET_FONT, stopIfTrue=False),
            )

        _add_block_validation(sheet, block, categories,
                              data_first_row, data_last_row,
                              cat_columns, last_category_row, first_row,
                              constraints_row)

    _add_courses_validation(sheet, headings, data_first_row, data_last_row,
                            gpa1_col, gpa2_col, tracks, checkboxes)
    _set_widths(sheet, headings, blocks)

    sheet.freeze_panes = "A6"
    sheet.sheet_view.showGridLines = True

    workbook.save(str(path))

    # Checkboxes are a post-processing step: openpyxl cannot express them, so
    # they are injected into the saved archive. See checkbox_xlsx.py for why
    # a template cannot work instead. This runs last, on a finished file.
    if checkboxes:
        flag_cells = [
            get_column_letter(column) + str(row)
            for column in (gpa1_col, gpa2_col)
            for row in range(data_first_row, data_last_row + 1)
        ]
        checkbox_xlsx.add_checkboxes(path, SHEET_NAME, flag_cells)

    return path


def _add_block_validation(sheet, block, categories, data_first_row,
                          data_last_row, cat_columns, last_category_row,
                          first_row, constraints_row):
    """Dropdowns belonging to one requirements block."""
    cat_letter = block.letter(block.col_category)
    prog_letter = block.letter(block.col_progress)

    # The Cat. dropdown lists exactly the category codes — and stops short of
    # the TOTAL row, which is not a category a course can be assigned to.
    source = ("=$" + cat_letter + "$" + str(first_row) +
              ":$" + cat_letter + "$" + str(last_category_row))
    validation = DataValidation(
        type="list", formula1=source, allowBlank=True,
        showErrorMessage=True, errorTitle="Invalid category",
        error="Pick one of the category codes from the "
              + block.req.label + " requirements table.",
    )
    sheet.add_data_validation(validation)
    column_letter = get_column_letter(cat_columns[block.cat_column_name])
    validation.add(column_letter + str(data_first_row) + ":" +
                   column_letter + str(data_last_row))

    # Tick-boxes for the advisor-attested constraints.
    if block.req.constraints:
        ticks = DataValidation(type="list", formula1='"TRUE,FALSE"',
                               allowBlank=True)
        sheet.add_data_validation(ticks)
        ticks.add(prog_letter + str(constraints_row + 1) + ":" +
                  prog_letter + str(constraints_row + len(block.req.constraints)))


def _add_courses_validation(sheet, headings, data_first_row, data_last_row,
                            gpa1_col, gpa2_col, tracks, checkboxes=False):
    """Dropdowns on the Courses table and the track selector."""
    grade = DataValidation(
        type="list", formula1='"' + ",".join(GRADE_CHOICES) + '"',
        allowBlank=True, showErrorMessage=True, errorTitle="Invalid grade",
        error="Grade must be one of: " + ", ".join(GRADE_CHOICES)
              + " — or blank for in-progress.",
    )
    sheet.add_data_validation(grade)
    # Resolved by name: the Pass column sits between Gr. and Ovr., so these
    # are no longer E and F, and a stale letter here would attach a dropdown
    # to the wrong column without any complaint from Excel.
    grade_letter = get_column_letter(headings.index("Gr.") + 1)
    ovr_letter = get_column_letter(headings.index("Ovr.") + 1)
    grade.add(grade_letter + str(data_first_row) + ":" +
              grade_letter + str(data_last_row))

    override = DataValidation(
        type="list",
        formula1='"' + OVR_INCLUDE + "," + OVR_EXCLUDE + '"', allowBlank=True,
        showErrorMessage=True, errorTitle="Invalid override",
        error='Use "' + OVR_INCLUDE + '" to make this course count toward '
              'progress, "' + OVR_EXCLUDE + '" to stop it counting, or leave '
              'blank for the default rule.',
    )
    sheet.add_data_validation(override)
    override.add(ovr_letter + str(data_first_row) + ":" +
                 ovr_letter + str(data_last_row))

    override_note = Comment(
        "Override — adjusts PROGRESS counting only, never GPA.\n\n"
        "blank : default rule. The grade must reach the bar in the Pass "
        "column, which comes from the category in Cat. Transfer credit (CR) "
        "always counts.\n"
        + OVR_INCLUDE + " : count it anyway, e.g. a D you have approved in a "
        "category whose bar is C. An F is never counted, and neither is a "
        "blank, W or IP.\n"
        + OVR_EXCLUDE + " : don't count it, e.g. the first attempt at a "
        "repeated course.",
        "Degree requirements")
    override_note.width = 360
    override_note.height = 190
    sheet[ovr_letter + "5"].comment = override_note

    pass_letter = get_column_letter(headings.index(PASS_COLUMN) + 1)
    pass_note = Comment(
        "Lowest grade that earns credit in this course's category.\n\n"
        "Filled in automatically from the Cat. column, so it changes when you "
        "recategorise a course. Set per category in the requirements file — "
        "that is how a BA can accept a D outside the major while a BSc "
        "accepts none.\n\n"
        "Not for typing in: overwriting the formula fixes the bar for that "
        "row. Use Ovr. to make an exception.",
        "Degree requirements")
    pass_note.width = 360
    pass_note.height = 170
    sheet[pass_letter + "5"].comment = pass_note

    # A checkbox already constrains the cell to TRUE/FALSE, and a dropdown
    # arrow sitting beside one looks like a second, different control. So the
    # gpa flags get the dropdown only when they are NOT checkboxes. The
    # constraint ticks keep theirs either way — those are attestations rather
    # than scratch flags, and looking different is the point.
    if not checkboxes:
        flags = DataValidation(type="list", formula1='"TRUE,FALSE"',
                               allowBlank=True)
        sheet.add_data_validation(flags)
        for column_index in (gpa1_col, gpa2_col):
            letter = get_column_letter(column_index)
            flags.add(letter + str(data_first_row) + ":" +
                      letter + str(data_last_row))

    track = DataValidation(
        type="list", formula1='"' + ",".join(tracks) + '"',
        allowBlank=False, showErrorMessage=True, errorTitle="Unknown track",
        error="Track must be one of: " + ", ".join(tracks),
    )
    sheet.add_data_validation(track)
    track.add("B3")


def _set_widths(sheet, headings, blocks):
    widths = {
        "Term": 12, "Course": 12, "Title": 34, "CH": 5, "Gr.": 6,
        PASS_COLUMN: 6, "Ovr.": 6, "Notes": 26, "gpa1": 7, "gpa2": 7,
    }
    for index, heading in enumerate(headings, start=1):
        letter = get_column_letter(index)
        sheet.column_dimensions[letter].width = widths.get(heading, 11)

    for block in blocks:
        sheet.column_dimensions[block.letter(block.col_category)].width = 12
        sheet.column_dimensions[block.letter(block.col_description)].width = 44
        sheet.column_dimensions[block.letter(block.col_req)].width = 9
        sheet.column_dimensions[block.letter(block.col_progress)].width = 13
