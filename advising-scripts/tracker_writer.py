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
#   Term | Course | Title | CH | Gr. | Ovr. | Cat. <edition>... | Notes
#        | gpa1 | gpa2
#
# One "Cat. <edition>" column and one requirements block per requirements file
# supplied, so a single file gives a one-edition tracker and two files
# reproduce the side-by-side comparison layout.

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import FormulaRule

SHEET_NAME = "Tracker"
HIDDEN_SHEET_NAME = "_TrackData"

# Grades that count toward degree progress. "CR" is transfer credit: it counts
# toward progress and total credit hours but carries no GPA-bearing grade.
PASSING_GRADES = ["A+", "A", "A-", "B+", "B", "B-", "C+", "C", "CR"]

# The 4.3 grade scale used for every GPA calculation.
GPA_GRADES = ["A+", "A", "A-", "B+", "B", "B-", "C+", "C", "D"]
GPA_POINTS = [4.3, 4, 3.7, 3.3, 3, 2.7, 2.3, 2, 1]

GRADE_CHOICES = ["A+", "A", "A-", "B+", "B", "B-", "C+", "C", "D", "F",
                 "W", "IP", "CR"]

# --- styling ----------------------------------------------------------------

FONT = "Arial"
SIZE = 10

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(name=FONT, size=SIZE, bold=True, color="FFFFFF")
BODY_FONT = Font(name=FONT, size=SIZE)
LABEL_FONT = Font(name=FONT, size=SIZE, bold=True, color="375623")
VALUE_FONT = Font(name=FONT, size=SIZE, bold=True)

# Excel hides its default gridlines underneath any cell fill, so a block that
# gets conditionally filled needs explicit light-grey borders to keep looking
# like a normal gridded table.
GRIDLINE = Side(style="thin", color="D4D4D4")
GRID_BORDER = Border(left=GRIDLINE, right=GRIDLINE,
                     top=GRIDLINE, bottom=GRIDLINE)

MET_FILL = PatternFill("solid", fgColor="C6EFCE")
MET_FONT = Font(name=FONT, size=SIZE, color="006100")
OFF_FILL = PatternFill("solid", fgColor="E7E6E6")
OFF_FONT = Font(name=FONT, size=SIZE, color="7F7F7F", strike=True)


def _array_literal(values):
    """Render a Python list as an Excel array constant: {"A+","A",...}."""
    parts = []
    for value in values:
        if isinstance(value, str):
            parts.append('"' + value + '"')
        else:
            parts.append(str(value))
    return "{" + ",".join(parts) + "}"


PASSING_ARRAY = _array_literal(PASSING_GRADES)
GPA_GRADE_ARRAY = _array_literal(GPA_GRADES)
GPA_POINT_ARRAY = _array_literal(GPA_POINTS)


class Block:
    """Where one requirements table sits on the sheet."""

    def __init__(self, requirements, cat_column_name, first_column):
        self.req = requirements
        self.cat_column_name = cat_column_name
        # Four visible columns plus a hidden numeric helper.
        self.col_category = first_column
        self.col_description = first_column + 1
        self.col_req = first_column + 2
        self.col_progress = first_column + 3
        self.col_earned = first_column + 4

    @property
    def width(self):
        return 5

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


def _courses_columns(blocks):
    """Column headings for the Courses table, in order."""
    headings = ["Term", "Course", "Title", "CH", "Gr.", "Ovr."]
    headings += [block.cat_column_name for block in blocks]
    headings += ["Notes", "gpa1", "gpa2"]
    return headings


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


def _progress_expression(cat_column_name, category_cell=None):
    """SUMPRODUCT of credit hours earned, optionally filtered to a category.

    Reads as: for courses where (the category matches, if we're filtering)
    AND ((the grade is C-or-better or CR, and the row is not force-excluded)
    OR (the grade is D and the row is force-included)), sum the credit hours.
    """
    passing = (
        "(ISNUMBER(MATCH(Courses[Gr.]," + PASSING_ARRAY + ",0))"
        '*(Courses[Ovr.]<>"no"))'
        '+((Courses[Gr.]="D")*(Courses[Ovr.]="yes"))'
    )
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
    points = (
        "IFERROR(CHOOSE(MATCH(Courses[Gr.]," + GPA_GRADE_ARRAY + ",0),"
        + ",".join(str(p) for p in GPA_POINTS) + "),0)"
    )
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


def _track_lookup(position, data_row):
    """An INDEX/MATCH that resolves one hidden-sheet row for the chosen track."""
    first = get_column_letter(position["first_col"])
    last = get_column_letter(position["last_col"])
    header = position["header_row"]
    return (
        "INDEX(" + HIDDEN_SHEET_NAME + "!$" + first + "$" + str(data_row) +
        ":$" + last + "$" + str(data_row) + ","
        "MATCH($B$3," + HIDDEN_SHEET_NAME + "!$" + first + "$" + str(header) +
        ":$" + last + "$" + str(header) + ",0))"
    )


def build_tracker(path, requirement_sets, rows, student,
                  track=None, spare_rows=15):
    """Generate the tracker workbook.

    requirement_sets — list of Requirements objects, one per edition
    rows            — list of course dicts (term/course/title/grade/hours,
                      plus optional cat_<edition>/ovr/notes/gpa1/gpa2)
    student         — dict with 'name', 'id'
    track           — which track to preselect in B3
    spare_rows      — blank rows left at the foot of the Courses table so an
                      advisor can add a course by hand
    """
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = SHEET_NAME

    # --- work out the geometry before writing anything ----------------------

    blocks = []
    used_labels = {}
    first_block_column = 6 + len(requirement_sets) + 3 + 2  # see below
    # Courses table is: Term, Course, Title, CH, Gr., Ovr. (6)
    #                 + one Cat. column per edition
    #                 + Notes, gpa1, gpa2 (3)
    # then one blank spacer column before the first requirements block.
    courses_width = 6 + len(requirement_sets) + 3
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

    tracks = requirement_sets[0].tracks
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
        sheet.cell(row=excel_row, column=5, value=row.get("grade", ""))
        sheet.cell(row=excel_row, column=6, value=row.get("ovr", ""))
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

    for excel_row in range(data_first_row, data_last_row + 1):
        for column_index in range(1, len(headings) + 1):
            cell = sheet.cell(row=excel_row, column=column_index)
            cell.font = BODY_FONT
            if column_index in (1, 2, 4, 5, 6, gpa1_col, gpa2_col):
                cell.alignment = Alignment(horizontal="center")
            elif column_index in cat_columns.values():
                cell.alignment = Alignment(horizontal="center")

    table = Table(
        displayName="Courses",
        ref="A" + str(table_first_row) + ":" +
            get_column_letter(len(headings)) + str(data_last_row),
    )
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleLight1", showRowStripes=True)
    sheet.add_table(table)

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
            text_cell.alignment = Alignment(wrap_text=True, vertical="center")
            tick = sheet.cell(row=excel_row, column=block.col_progress,
                              value=False)
            tick.alignment = Alignment(horizontal="center")
            tick.font = BODY_FONT

            if constraint.tracks:
                # Grey out a constraint that does not apply to the chosen track.
                applicable = _array_literal(constraint.tracks)
                sheet.conditional_formatting.add(
                    "$" + cat_letter + "$" + str(excel_row) + ":$" +
                    prog_letter + "$" + str(excel_row),
                    FormulaRule(
                        formula=["ISNA(MATCH($B$3," + applicable + ",0))"],
                        fill=OFF_FILL, font=OFF_FONT, stopIfTrue=True),
                )

        _add_block_validation(sheet, block, categories,
                              data_first_row, data_last_row,
                              cat_columns, last_category_row, first_row,
                              constraints_row)

    _add_courses_validation(sheet, headings, data_first_row, data_last_row,
                            gpa1_col, gpa2_col, tracks)
    _set_widths(sheet, headings, blocks)

    sheet.freeze_panes = "A6"
    sheet.sheet_view.showGridLines = True

    workbook.save(str(path))
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
                            gpa1_col, gpa2_col, tracks):
    """Dropdowns on the Courses table and the track selector."""
    grade = DataValidation(
        type="list", formula1='"' + ",".join(GRADE_CHOICES) + '"',
        allowBlank=True, showErrorMessage=True, errorTitle="Invalid grade",
        error="Grade must be one of: " + ", ".join(GRADE_CHOICES)
              + " — or blank for in-progress.",
    )
    sheet.add_data_validation(grade)
    grade.add("E" + str(data_first_row) + ":E" + str(data_last_row))

    override = DataValidation(
        type="list", formula1='"yes,no"', allowBlank=True,
        showErrorMessage=True, errorTitle="Invalid override",
        error='Use "yes" to force-include a course in progress, "no" to '
              'force-exclude it, or leave blank for the default rule.',
    )
    sheet.add_data_validation(override)
    override.add("F" + str(data_first_row) + ":F" + str(data_last_row))

    override_note = Comment(
        "Override — adjusts PROGRESS counting only, never GPA.\n\n"
        "blank : default rule. C or better (and CR) counts; D, F, W and IP "
        "do not.\n"
        "yes   : force-include, e.g. a D you have approved to count.\n"
        "no    : force-exclude, e.g. the first attempt at a repeated course.",
        "Degree requirements")
    override_note.width = 360
    override_note.height = 130
    sheet["F5"].comment = override_note

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
        "Ovr.": 6, "Notes": 26, "gpa1": 7, "gpa2": 7,
    }
    for index, heading in enumerate(headings, start=1):
        letter = get_column_letter(index)
        sheet.column_dimensions[letter].width = widths.get(heading, 11)

    for block in blocks:
        sheet.column_dimensions[block.letter(block.col_category)].width = 12
        sheet.column_dimensions[block.letter(block.col_description)].width = 44
        sheet.column_dimensions[block.letter(block.col_req)].width = 9
        sheet.column_dimensions[block.letter(block.col_progress)].width = 13
