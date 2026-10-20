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
#   row 5   Courses table header             Category|Description|Completed|Planned
#   row 6+  course rows                      category rows, TOTAL, CONSTRAINTS
#
# The requirements block also carries four HIDDEN columns: Req. CH (the
# track lookup every conditional-formatting rule tests), the two numeric
# helpers behind Completed and Planned, and each category's passing bar.
#
# "Completed" is hours banked. "Planned" is the projected figure — banked
# plus every course filed here that has no grade yet — over the same
# requirement, so "will this be met?" needs no arithmetic. Completed
# deliberately excludes planned hours: the strike-through rule on the Courses
# table is written as the exact negation of the Completed test so the two can
# never disagree, and folding planned hours in would break that.
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
from requirements_io import DEFAULT_PASS_GRADE, GRADE_RANK

SHEET_NAME = "Tracker"
HIDDEN_SHEET_NAME = "_TrackData"

# The Courses-table column holding each row's passing bar, looked up from the
# category the advisor assigned. Its name appears inside structured formula
# references, so it is written once here.
PASS_COLUMN = "Pass"

# What a category with no target shows in its Req. CH cell. It is written into
# the hidden lookup table as TEXT, and every conditional-formatting rule that
# cares tests ISNUMBER/ISTEXT rather than the value — so "has a target" is a
# question about the cell's TYPE, not about which row it sits on. That is what
# lets a catch-all category appear anywhere in a requirements file.
#
# It must not be blank: Excel reads a blank lookup cell as 0, and 0 already
# means "not part of this track" (greyed and struck out).
NO_TARGET_TEXT = "—"

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

# A course that earns no credit — failed, withdrawn, below its category's bar,
# or excluded by hand. Deliberately the SAME grey strike-through as OFF_FONT,
# and deliberately with NO fill.
#
# The two say the same thing in the two halves of the sheet: struck-out grey
# TEXT means "this does not count towards the degree". What stays unique to
# the requirements block is the grey FILL, which still means "not on this
# track" and nothing else — so the rule in docs/degree-tracker.md is about OFF_FILL, and
# this rule adds no new colour to the workbook.
NO_CREDIT_FONT = Font(color="FF7F7F7F", strike=True)

# A category with no target — a catch-all like FreeElec. Informational, not a
# verdict, so it needs a THIRD colour of its own. Grey is spoken for ("not on
# this track" and nothing else), and simply leaving the cell unfilled is not
# an option either: among a column of green rows, the absence of green reads
# as "unmet", which is precisely the false alarm this row exists to stop
# raising. The blue is the one already used for the block headings.
NO_TARGET_FILL = PatternFill(patternType="solid", fgColor=AUTOMATIC,
                             bgColor="FFDDEBF7")
NO_TARGET_FONT = Font(color="FF1F4E79")

# A category that is NOT met yet but WOULD be if the planned courses land.
# Excel's own "Neutral" pair, which is the missing third beside the "Good" and
# "Bad" pairs already used below for the Ovr. column.
#
# Deliberately NOT green. In this workbook green means "banked", and the
# Planned cell sits directly beside the Completed cell that earns its green
# honestly — a green here would be the first one in the sheet that does not
# mean the hours exist, one column away from ones that do. Amber says "not
# yet, but on plan" without colliding with anything.
PLANNED_FILL = PatternFill(patternType="solid", fgColor=AUTOMATIC,
                           bgColor="FFFFEB9C")
PLANNED_FONT = Font(color="FF9C6500")

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


def _grade_ladder(reference):
    """Ordinal rank of whatever `reference` names, as a sum of comparisons.

    The one grade ladder in the workbook, used from two quite different
    places:

      _grade_rank("Gr.")   a whole COLUMN, inside SUMPRODUCT
      _grade_rank_cell()   a single CELL, inside a conditional format

    They are written the same way for different reasons. In SUMPRODUCT it is
    the array-safety rule above. In a conditional format it is because CF
    formulas do not accept array constants at all — `MATCH($E6,{"D","C",...})`
    makes Excel repair the file and discard every rule on the sheet (trap 1).
    Neither place can use the compact idiom, so neither does.
    """
    return "(" + "+".join(
        "(" + reference + '="' + grade + '")*' + str(GRADE_RANK[grade])
        for grade in reversed(RANKED_GRADES)
    ) + ")"


def _grade_rank(column_name):
    """Array-safe ordinal rank of a grade column. Yields 0 for anything
    off the ladder — F, W, IP, CR and blank.

    Takes the column name because the SAME expression is needed for two
    columns: the student's grade (Courses[Gr.]) and the bar it has to clear
    (Courses[Pass]). Comparing two of these keeps both sides elementwise,
    which is what a lookup inlined into the SUMPRODUCT would destroy.
    """
    return _grade_ladder("Courses[" + column_name + "]")


def _grade_rank_cell(reference):
    """Ordinal rank of ONE grade cell, e.g. "$E6".

    The scalar twin of _grade_rank, for conditional-formatting formulas. Pass
    a column-absolute, ROW-RELATIVE reference: a CF formula is written for the
    top-left cell of its range and re-evaluated relatively, so "$E$6" would
    make every row in the range test row 6's grade.
    """
    return _grade_ladder(reference)


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
        # Four visible columns plus four hidden helpers.
        self.col_category = first_column
        self.col_description = first_column + 1
        # Req. CH is HIDDEN, despite having a header. It is the INDEX/MATCH
        # that resolves the B3 track dropdown, and all four of the rules on
        # the block's category rows test it — but its value is already
        # printed as the denominator of the two columns beside it, so showing
        # it as well says the same number twice per row. Kept as a column
        # rather than inlined for three reasons: the CF rules need a cell to
        # point at; ISNUMBER/ISTEXT on it is how "has a target" is asked (see
        # NO_TARGET_TEXT); and removing it would shift col_progress, which
        # tracker_reader finds by a hardcoded +3 offset and would then lose
        # every constraint tick SILENTLY.
        self.col_req = first_column + 2
        # "Completed" — hours banked. The offset MUST stay +3: tracker_reader
        # reaches the constraint ticks at col_category + 3.
        self.col_progress = first_column + 3
        # "Planned" — the projected figure, banked plus planned.
        self.col_planned = first_column + 4
        self.col_earned = first_column + 5
        self.col_planned_ch = first_column + 6
        # Each category's passing bar, as a letter. Hidden: the advisor sees
        # it per COURSE in the Courses table's Pass column, which looks its
        # value up here, so showing it twice would just be noise.
        self.col_pass = first_column + 7

    @property
    def width(self):
        return 8

    def letter(self, column):
        return get_column_letter(column)


# The Courses table is these, then one Cat. column per edition, then these.
# Kept as named lists because the sheet's geometry is derived from their
# LENGTHS before any block exists — the two used to be separate magic numbers
# that had to be edited in step, which is exactly the kind of arithmetic that
# goes wrong the once.
COURSES_LEADING = ["Term", "Course", "Title", "CH", "Gr.", PASS_COLUMN, "Ovr."]
COURSES_TRAILING = ["Notes", "gpa1", "gpa2"]

# The Cat. column of an ad hoc tracker — one built with no requirements file at
# all, to look at what a student has taken before there is a programme to
# measure it against.
#
# Named WITHOUT an edition label, and that is the whole distinction: every
# other Cat. column is "Cat. <edition>", so the bare name says "these
# categories are yours, not a requirements file's". tracker_reader keys on the
# same difference, reading this one back as "cat" rather than "cat_<label>".
AD_HOC_CAT_COLUMN = "Cat."


def _courses_columns(blocks):
    """Column headings for the Courses table, in order."""
    cat_columns = ([block.cat_column_name for block in blocks]
                   or [AD_HOC_CAT_COLUMN])
    return list(COURSES_LEADING) + cat_columns + list(COURSES_TRAILING)


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


def _planned_test():
    """The "this course is planned or in progress" test, per row.

    A blank grade means registered-or-intended; "IP" means in progress this
    term. The workbook already treats the two identically everywhere else —
    the strike-through rule excludes both up front, and the GPA excludes both
    from numerator and denominator — so they are one concept here too, and no
    new grade token is introduced to say it.

    "excl." takes a row out, exactly as it does for Completed: it is how an
    advisor says "this one does not count", and a planned course is no
    different. "incl." is deliberately NOT honoured — it rescues a real grade
    that fell short of its bar, and a blank is not a grade. That is the same
    `rank > 0` reasoning as in _passing_expression, arrived at from the other
    side.

    Written as sums of explicit comparisons, per trap 2: SUMPRODUCT feeds a
    whole column in, and anything that is not elementwise collapses to a
    scalar and fails silently.
    """
    return ('((Courses[Gr.]="")+(Courses[Gr.]="IP"))'
            "*(Courses[Ovr.]<>" + OVR_EXCLUDE_Q + ")")


def _planned_expression(cat_column_name, category_cell=None):
    """SUMPRODUCT of credit hours PLANNED, optionally filtered to a category.

    The same shape as _progress_expression, with the passing test swapped for
    the planned one, so the two columns can never disagree about which rows
    they are looking at.

    The spare rows at the foot of the Courses table all have a blank Gr. and
    so satisfy the planned test. That is harmless rather than a bug: their CH
    cells are empty, Excel reads an empty cell as 0 in an arithmetic context,
    and a row contributes CH hours or nothing. Worth stating, because the
    alternative reading — "every blank row counts as a planned course" — is
    the obvious one and is wrong.
    """
    planned = _planned_test()
    if category_cell is None:
        return "SUMPRODUCT((" + planned + ")*Courses[CH])"
    return (
        "SUMPRODUCT(((Courses[" + cat_column_name + "]=" + category_cell + ")"
        "*(" + planned + "))*Courses[CH])"
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
        for category in block.req.categories:
            sheet.cell(row=row, column=1, value=category.code)
            for offset, track in enumerate(tracks):
                required = category.req_ch.get(track, 0)
                sheet.cell(
                    row=row, column=2 + offset,
                    value=NO_TARGET_TEXT if required is None else required)
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
    if not requirement_sets:
        return []   # an ad hoc tracker has no programme, so no tracks
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

    requirement_sets — list of Requirements objects, one per edition. EMPTY
                      builds an ad hoc tracker: the Courses table alone, with
                      a free-text Cat. column, no requirements blocks, no
                      track selector, and a hand-editable passing bar.
    rows            — list of course dicts (term/course/title/grade/hours,
                      plus optional cat_<edition> — or "cat" for an ad hoc
                      tracker — and ovr/notes/gpa1/gpa2)
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
    # One Cat. column per edition — or, with no editions at all, the single
    # ad hoc one, so the table keeps its shape either way.
    courses_width = (len(COURSES_LEADING) + max(len(requirement_sets), 1)
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
    chosen_track = (track if track in tracks
                    else (tracks[0] if tracks else ""))

    # Skipped entirely for an ad hoc tracker: with no blocks the lookup tables
    # would be empty, and an empty hidden sheet is just a thing to wonder
    # about later.
    positions = _write_hidden_track_data(workbook, blocks) if blocks else {}

    # --- top section --------------------------------------------------------

    sheet["A1"] = "NAME"
    sheet["B1"] = student.get("name", "")
    sheet["A2"] = "ID"
    sheet["B2"] = student.get("id", "")

    identity_cells = ["A1", "A2"]
    value_cells = ["B1", "B2"]

    # No requirements means no tracks, so row 3 gets no selector. Leaving a
    # "Track" label above an empty unvalidated cell would advertise a control
    # that does nothing — every formula that reads B3 lives in a requirements
    # block, and there are none.
    if blocks:
        sheet["A3"] = "Track"
        sheet["B3"] = chosen_track
        identity_cells.append("A3")
        value_cells.append("B3")
        sheet["B3"].alignment = Alignment(horizontal="center")

    for address in identity_cells:
        sheet[address].font = LABEL_FONT
    for address in value_cells:
        sheet[address].font = VALUE_FONT
    sheet["B1"].alignment = Alignment(horizontal="left")
    sheet["B2"].alignment = Alignment(horizontal="center")

    sheet["C3"] = "Overall GPA"
    sheet["C3"].font = LABEL_FONT
    sheet["C3"].alignment = Alignment(horizontal="right")
    sheet["D3"] = _gpa_formula()
    sheet["D3"].font = VALUE_FONT
    sheet["D3"].alignment = Alignment(horizontal="center")
    sheet["D3"].number_format = "0.00"

    if not blocks:
        # The one number an ad hoc tracker can still honestly report. Every
        # other total in this workbook is "x of the y your programme wants",
        # and there is no y — but "hours banked so far" needs no programme,
        # only the Pass column beside each grade.
        #
        # NOT in B3, which tracker_reader reads as the selected track: a
        # number there would come back as a track name on a later merge.
        sheet["C2"] = "CH earned"
        sheet["C2"].font = LABEL_FONT
        sheet["C2"].alignment = Alignment(horizontal="right")
        sheet["D2"] = "=" + _progress_expression(AD_HOC_CAT_COLUMN)
        sheet["D2"].font = VALUE_FONT
        sheet["D2"].alignment = Alignment(horizontal="center")

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
    if not blocks:
        # The ad hoc column is a Cat. column for every purpose except having a
        # requirements block behind it, so it goes in the same dict and picks
        # up the same centring and width without a second code path.
        cat_columns[AD_HOC_CAT_COLUMN] = \
            headings.index(AD_HOC_CAT_COLUMN) + 1

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
        if not blocks:
            # Read back under the same key by tracker_reader, so an ad hoc
            # tracker survives its own -u re-run.
            sheet.cell(row=excel_row,
                       column=cat_columns[AD_HOC_CAT_COLUMN],
                       value=row.get("cat", ""))
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
    # docs/degree-tracker.md wearing a different hat.
    #
    # With two editions the first one wins. Two editions normally means two
    # calendar years of one degree, whose bars agree. A BA block beside a BSc
    # block would need two different verdicts on the same row, and is not a
    # supported layout.
    block_first_row = 6

    if not blocks:
        # An ad hoc tracker has no category table to look the bar up IN, so
        # these cells hold a plain letter instead of a formula — the default,
        # editable, with a dropdown to keep it on the ladder. It is the one
        # column that changes meaning between the two kinds of tracker: here
        # it is entered rather than derived, and its header note says so.
        for excel_row in range(data_first_row, data_last_row + 1):
            sheet.cell(row=excel_row, column=pass_col,
                       value=DEFAULT_PASS_GRADE)
    else:
        primary = blocks[0]
        primary_last_cat = (block_first_row
                            + len(primary.req.categories) - 1)
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

    # --- strike out the courses that earn nothing ---------------------------
    #
    # Grey and struck through, on the Course and Title cells only: the Term
    # stays legible so the row still reads chronologically, and the grade
    # stays bold because the grade is the REASON, not the casualty.
    #
    # The condition is not a list of bad grades. It is the exact negation of
    # the rule the Progress column already uses — "this course did not earn
    # its credit hours" — so the two can never disagree, and it stays right
    # for a category whose bar is D (a D is not struck out) or B (a C is).
    # That is what makes "F and W, plus D where the bar is C" fall out of it
    # rather than being enumerated.
    #
    # A blank or IP grade is excluded up front. Those are courses in progress
    # or planned, not courses that failed to count, and striking them through
    # would read as a verdict on work that hasn't been marked yet.
    grade_ref = "$" + get_column_letter(grade_col) + str(data_first_row)
    pass_ref = "$" + get_column_letter(pass_col) + str(data_first_row)
    ovr_ref = "$" + ovr_letter + str(data_first_row)
    rank = _grade_rank_cell(grade_ref)
    bar = _grade_rank_cell(pass_ref)

    # The same two terms as _passing_expression, written for one row.
    met = ("AND(OR(" + rank + ">=" + bar + "," + grade_ref + '="CR"),'
           + ovr_ref + "<>" + OVR_EXCLUDE_Q + ")")
    rescue = ("AND(" + rank + ">0," + rank + "<" + bar + ","
              + ovr_ref + "=" + OVR_INCLUDE_Q + ")")
    no_credit = ("AND(" + grade_ref + '<>"",' + grade_ref + '<>"IP",'
                 "NOT(OR(" + met + "," + rescue + ")))")

    course_letter = get_column_letter(headings.index("Course") + 1)
    title_letter = get_column_letter(headings.index("Title") + 1)
    sheet.conditional_formatting.add(
        "$" + course_letter + "$" + str(data_first_row) +
        ":$" + title_letter + "$" + str(data_last_row),
        FormulaRule(formula=[no_credit], font=NO_CREDIT_FONT,
                    stopIfTrue=False),
    )

    # --- requirements blocks ------------------------------------------------

    for block in blocks:
        position = positions[id(block)]
        categories = block.req.categories
        no_target = set(block.req.no_target_codes())

        cat_letter = block.letter(block.col_category)
        req_letter = block.letter(block.col_req)
        prog_letter = block.letter(block.col_progress)
        planned_letter = block.letter(block.col_planned)
        earned_letter = block.letter(block.col_earned)
        planned_ch_letter = block.letter(block.col_planned_ch)

        heading_cell = sheet.cell(
            row=4, column=block.col_category,
            value="DEGREE REQUIREMENTS (" + block.req.label + ")")
        heading_cell.font = Font(name=FONT, size=SIZE, bold=True,
                                 color="1F4E79")

        # "Req. CH" still gets a header even though its column is hidden, so
        # that unhiding it gives a labelled column rather than a mystery one.
        for offset, name in enumerate(
                ["Category", "Description", "Req. CH", "Completed",
                 "Planned"]):
            cell = sheet.cell(row=5, column=block.col_category + offset,
                              value=name)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = Alignment(horizontal="center")

        planned_header = sheet.cell(row=5, column=block.col_planned)
        planned_note = Comment(
            "What this category would hold once the planned courses land: "
            "the Completed figure plus every course filed here whose Gr. is "
            "blank or \"IP\" and whose Ovr. is not \"excl.\".\n\n"
            "Amber means this category is NOT met yet but WOULD be if the "
            "plan comes off. No amber and no green next door means the plan "
            "does not close the gap.\n\n"
            "Green lives in Completed only. Hours are not banked until they "
            "are graded.",
            "Degree requirements")
        planned_note.width = 320
        planned_note.height = 140
        planned_header.comment = planned_note

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

            sheet.cell(row=excel_row, column=block.col_req,
                       value="=" + _track_lookup(position, hidden_row))

            sheet.cell(
                row=excel_row, column=block.col_earned,
                value="=" + _progress_expression(
                    block.cat_column_name,
                    "$" + cat_letter + "$" + str(excel_row)))
            sheet.cell(
                row=excel_row, column=block.col_planned_ch,
                value="=" + _planned_expression(
                    block.cat_column_name,
                    "$" + cat_letter + "$" + str(excel_row)))

            # The Planned cell shows the PROJECTED figure — banked plus
            # planned — over the same requirement, so "will this be met?"
            # needs no mental arithmetic. It reads as a second, hypothetical
            # Completed cell, which is why it takes the same two shapes.
            earned_ref = "$" + earned_letter + "$" + str(excel_row)
            projected = ("$" + earned_letter + "$" + str(excel_row) + "+$" +
                         planned_ch_letter + "$" + str(excel_row))
            if category.code in no_target:
                # No denominator, because there is no target to fall short of.
                # The right formula is emitted per row rather than wrapping
                # every row in an IF: the generator already knows which
                # categories these are, so the sheet need not re-derive it.
                progress = '=TEXT(' + earned_ref + ',"0")&" ch"'
                planned = '=TEXT(' + projected + ',"0")&" ch"'
            else:
                req_ref = "$" + req_letter + "$" + str(excel_row)
                progress = ('=TEXT(' + earned_ref + ',"0")&" / "&TEXT(' +
                            req_ref + ',"0")&" ch"')
                planned = ('=TEXT(' + projected + ',"0")&" / "&TEXT(' +
                           req_ref + ',"0")&" ch"')
            sheet.cell(row=excel_row, column=block.col_progress,
                       value=progress)
            sheet.cell(row=excel_row, column=block.col_planned,
                       value=planned)

        total_row = first_row + len(categories)
        last_category_row = total_row - 1

        sheet.cell(row=total_row, column=block.col_category, value="TOTAL")
        sheet.cell(row=total_row, column=block.col_req,
                   value="=" + _track_lookup(position, position["total_row"]))
        sheet.cell(row=total_row, column=block.col_earned,
                   value="=" + _progress_expression(block.cat_column_name))
        sheet.cell(row=total_row, column=block.col_planned_ch,
                   value="=" + _planned_expression(block.cat_column_name))
        sheet.cell(
            row=total_row, column=block.col_progress,
            value='=TEXT($' + earned_letter + "$" + str(total_row) +
                  ',"0")&" / "&TEXT($' + req_letter + "$" + str(total_row) +
                  ',"0")&" ch"')
        # The "will they graduate on plan?" cell. Like the TOTAL Completed
        # cell it always keeps its denominator: the declared total is a real
        # number on every track, never a dash.
        sheet.cell(
            row=total_row, column=block.col_planned,
            value='=TEXT($' + earned_letter + "$" + str(total_row) + '+$' +
                  planned_ch_letter + "$" + str(total_row) +
                  ',"0")&" / "&TEXT($' + req_letter + "$" + str(total_row) +
                  ',"0")&" ch"')

        for excel_row in range(5, total_row + 1):
            for column_index in range(block.col_category,
                                      block.col_planned + 1):
                cell = sheet.cell(row=excel_row, column=column_index)
                cell.border = GRID_BORDER
                if excel_row > 5:
                    cell.font = BODY_FONT
                if column_index in (block.col_req, block.col_progress,
                                    block.col_planned):
                    cell.alignment = Alignment(horizontal="center")
        for column_index in range(block.col_category, block.col_planned + 1):
            sheet.cell(row=total_row, column=column_index).font = Font(
                name=FONT, size=SIZE, bold=True)

        # Hide the numeric helper columns. They exist so the conditional
        # formatting can compare two NUMBERS; the original template compared
        # by pulling the numbers back out of the "12 / 15 ch" display string
        # with FIND and MID, which is brittle.
        sheet.column_dimensions[earned_letter].hidden = True
        sheet.column_dimensions[planned_ch_letter].hidden = True
        sheet.column_dimensions[block.letter(block.col_pass)].hidden = True
        # Req. CH is hidden too — see the note on Block.col_req. Its value is
        # already the denominator of both visible figures beside it, so
        # showing the column as well prints the same number twice a row (and
        # on a no-target row, prints "—" beside a figure that already has no
        # denominator). It stays a real column because the CF rules need a
        # cell to test and because moving col_progress would break the
        # constraint-tick read-back silently.
        sheet.column_dimensions[req_letter].hidden = True

        # --- conditional formatting on this block ---------------------------

        # A conditional-formatting formula is written for the TOP-LEFT cell of
        # its range and then re-evaluated relative to every other cell. So the
        # references here must be column-absolute but ROW-RELATIVE ($N6, not
        # $N$6) — otherwise every row would test row 6's condition and the
        # whole column would light up together.
        # The ISNUMBER guard on the next two rules is what keeps a no-target
        # row out of them. Its Req. CH cell holds TEXT ("—"), and Excel sorts
        # text above every number, so a bare `>0` would be TRUE there. The
        # second half of each test happens to rescue it, but only by accident
        # of Excel's type ordering — which is exactly the kind of thing that
        # holds until it doesn't. Say what is meant instead.
        has_target = "ISNUMBER($" + req_letter + str(first_row) + ")"

        # Green stays on the Completed column ALONE. The Planned column can
        # never go green: its figure includes hours that have not been earned,
        # and green in this workbook means banked.
        progress_range = ("$" + prog_letter + "$" + str(first_row) +
                          ":$" + prog_letter + "$" + str(total_row))
        sheet.conditional_formatting.add(
            progress_range,
            FormulaRule(
                formula=["AND(" + has_target + ",$" + req_letter +
                         str(first_row) + ">0,$" +
                         earned_letter + str(first_row) + ">=$" +
                         req_letter + str(first_row) + ")"],
                fill=MET_FILL, font=MET_FONT, stopIfTrue=False),
        )

        # A category that requires 0 ch on the selected track is not part of
        # that track — grey it out and strike it through. Keying off the
        # resolved requirement rather than an "is Honours?" flag means this
        # works for any track a requirements file cares to define.
        row_range = ("$" + cat_letter + "$" + str(first_row) +
                     ":$" + planned_letter + "$" + str(last_category_row))
        sheet.conditional_formatting.add(
            row_range,
            FormulaRule(
                formula=["AND(" + has_target + ",$" + req_letter +
                         str(first_row) + "=0)"],
                fill=OFF_FILL, font=OFF_FONT, stopIfTrue=True),
        )

        # A catch-all category — no target, so neither "met" nor "unmet". It
        # gets its own colour rather than being left unfilled, because in a
        # column of green rows an unfilled one reads as a failure.
        #
        # Discriminating on the cell's TYPE rather than on which row it sits
        # on is what lets the range stay contiguous and lets a requirements
        # file put its catch-all anywhere — or have several. The TOTAL row is
        # excluded because it always resolves to a number.
        #
        # Both figure columns, so a catch-all's projection is blue like the
        # Completed cell beside it. Neither is a verdict, and they should not
        # look like two different kinds of non-verdict.
        if no_target:
            sheet.conditional_formatting.add(
                "$" + prog_letter + "$" + str(first_row) +
                ":$" + planned_letter + "$" + str(last_category_row),
                FormulaRule(
                    formula=["ISTEXT($" + req_letter + str(first_row) + ")"],
                    fill=NO_TARGET_FILL, font=NO_TARGET_FONT,
                    stopIfTrue=False),
            )

        # A category the plan would close: short today, met once the planned
        # courses land. Amber rather than green — see PLANNED_FILL.
        #
        # Four terms, and each is load-bearing:
        #
        #   ISNUMBER   trap 4. Without it the rule fires on a no-target row,
        #              because "—" > 0 is TRUE in Excel — text sorts above
        #              every number. The row would then claim its plan closes
        #              a gap that does not exist.
        #   > 0        not on this track. The grey rule above stops first, so
        #              this is belt and braces, but the rule should say what
        #              it means rather than lean on rule ordering.
        #   earned <   NOT met yet. This is what keeps amber off the rows that
        #              are already green: on a met row the projection trivially
        #              also clears the bar, and colouring it would put a second
        #              status beside a green that already answered the
        #              question. Same shape, and the same reason, as the
        #              `rank < bar` bound on the "incl." rescue.
        #   >=         and the projection does clear it.
        #
        # Added after the grey rule so an off-track row stays grey (its
        # stopIfTrue=True is what makes that stick).
        sheet.conditional_formatting.add(
            "$" + planned_letter + "$" + str(first_row) +
            ":$" + planned_letter + "$" + str(total_row),
            FormulaRule(
                formula=["AND(" + has_target +
                         ",$" + req_letter + str(first_row) + ">0" +
                         ",$" + earned_letter + str(first_row) +
                         "<$" + req_letter + str(first_row) +
                         ",$" + earned_letter + str(first_row) +
                         "+$" + planned_ch_letter + str(first_row) +
                         ">=$" + req_letter + str(first_row) + ")"],
                fill=PLANNED_FILL, font=PLANNED_FONT, stopIfTrue=False),
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
            # The tick lives in the Completed column, not Planned: it is an
            # attestation that something IS so, and it goes green there
            # exactly as a met category does. tracker_reader also finds it at
            # col_category + 3, which is why that offset must not move.
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
                    planned_letter + "$" + str(excel_row),
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
                            gpa1_col, gpa2_col, tracks, checkboxes,
                            ad_hoc=not blocks)
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
                            gpa1_col, gpa2_col, tracks, checkboxes=False,
                            ad_hoc=False):
    """Dropdowns on the Courses table and the track selector.

    ad_hoc — no requirements file behind this tracker, so the Cat. column gets
    no dropdown (there is no list of categories to offer) and the Pass column
    gets one (it is typed rather than derived).
    """
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
    if ad_hoc:
        # In an ad hoc tracker this column is the OPPOSITE of what it is
        # everywhere else: there is no requirements file to derive it from, so
        # it is typed. Saying "not for typing in" here would be exactly wrong,
        # which is why the note is swapped rather than shared.
        pass_note = Comment(
            "Lowest grade that earns credit in this course.\n\n"
            "Yours to set. Every row starts at " + DEFAULT_PASS_GRADE +
            "; change any of them — a\n"
            "programme that accepts a D outside the major, say — and the "
            "strike-through and the CH earned total follow.\n\n"
            "Use Ovr. for a one-off exception you would rather not turn into "
            "a rule.",
            "Ad hoc tracker")
        pass_note.height = 170
        grades = DataValidation(
            type="list",
            formula1='"' + ",".join(reversed(RANKED_GRADES)) + '"',
            allowBlank=True, showErrorMessage=True,
            errorTitle="Invalid passing grade",
            error="The bar must be one of: " + ", ".join(reversed(RANKED_GRADES))
                  + ". F is not offered — it ranks below every grade, so a bar "
                    "of F would wave through blanks and failures alike.",
        )
        sheet.add_data_validation(grades)
        grades.add(pass_letter + str(data_first_row) + ":" +
                   pass_letter + str(data_last_row))
    else:
        pass_note = Comment(
            "Lowest grade that earns credit in this course's category.\n\n"
            "Filled in automatically from the Cat. column, so it changes when "
            "you recategorise a course. Set per category in the requirements "
            "file — that is how a BA can accept a D outside the major while a "
            "BSc accepts none.\n\n"
            "Not for typing in: overwriting the formula fixes the bar for "
            "that row. Use Ovr. to make an exception.",
            "Degree requirements")
        pass_note.height = 170
    pass_note.width = 360
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

    # No requirements means no tracks and no B3 selector to validate.
    if tracks:
        track = DataValidation(
            type="list", formula1='"' + ",".join(tracks) + '"',
            allowBlank=False, showErrorMessage=True,
            errorTitle="Unknown track",
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
        sheet.column_dimensions[block.letter(block.col_progress)].width = 13
        sheet.column_dimensions[block.letter(block.col_planned)].width = 13
        # No width for col_req: it is hidden, and a width on a hidden column
        # is a value nobody will ever see being wrong.
