# /// script
# dependencies = ["openpyxl"]
# ///
# check_blank.py — test the empty tracker blank_tracker.py produces.
# Usage: uv run check_blank.py
#
# A regression check, not part of the shipped pipeline.
#
# A blank tracker is only useful if its empty rows are already *live*: the
# dropdowns, the formatting and the GPA flags have to reach the bottom of the
# table, and the requirements block has to compute against a table with no
# data in it. So this checks three things:
#
#   1. the table really is empty — no stray course rows, and every row carries
#      its FALSE gpa flags (an empty flag cell would break the GPA formulas);
#   2. the dropdowns reach the last row, so a colleague typing on row 40 gets
#      the same Gr./Ovr./Cat. choices as one typing on row 6;
#   3. rows typed in by hand read back through tracker_reader, which is what
#      makes a hand-filled tracker mergeable with a real transcript later.
#
# Note this check needs no transcript at all — it is the one path through the
# generator that does not involve a PDF, which is half of why it exists.

import sys
from pathlib import Path

import openpyxl

# The scripts live one folder up, and their helper modules in lib/.
REPO = Path(__file__).resolve().parent.parent
sys.path[:0] = [str(REPO), str(REPO / "lib")]

import check_support
import tracker_reader

ROWS = 30  # small enough to read in a dump, big enough to be a real table


def data_validation_ranges(sheet, column_letter):
    """Every validation range on the sheet that covers this column."""
    found = []
    for validation in sheet.data_validations.dataValidation:
        for cell_range in validation.sqref.ranges:
            if cell_range.min_col <= openpyxl.utils.column_index_from_string(
                    column_letter) <= cell_range.max_col:
                found.append((cell_range.min_row, cell_range.max_row))
    return found


def main():
    work = check_support.fresh_dir("_blank_test")
    req = check_support.sample_requirements(work)
    tracker = work / "blank.xlsx"

    print(check_support.run_script(
        "blank_tracker.py", "-r", req, "-o", tracker,
        "-t", "Honours", "--rows", ROWS).strip())
    print()

    failures = []

    workbook = openpyxl.load_workbook(str(tracker))
    sheet = workbook["Tracker"]

    # --- 1. the table is empty, but not inert -------------------------------

    headings = {}
    for column in range(1, sheet.max_column + 1):
        name = sheet.cell(row=5, column=column).value
        if name:
            headings[str(name)] = column

    first_row, last_row = 6, 5 + ROWS
    print("1. an empty Courses table")
    # Row 5 also carries the requirements block's own headings further
    # right, so the Courses table is A .. gpa2, not A .. last used column.
    courses_last_column = headings.get("gpa2", len(headings))
    print("   Courses columns:", [name for name, column in headings.items()
                                  if column <= courses_last_column])
    print("   data rows   :", first_row, "-", last_row)

    if "Courses" not in sheet.tables:
        failures.append("no Courses table defined")
    else:
        expected_ref = "A5:" + openpyxl.utils.get_column_letter(
            courses_last_column) + str(last_row)
        if sheet.tables["Courses"].ref != expected_ref:
            failures.append("Courses table spans " +
                            sheet.tables["Courses"].ref + ", expected " +
                            expected_ref)

    typed = [name for name in ("Term", "Course", "Title", "CH", "Gr.", "Ovr.",
                               "Notes") if name in headings]
    stray = []
    for row in range(first_row, last_row + 1):
        for name in typed:
            value = sheet.cell(row=row, column=headings[name]).value
            if value not in (None, ""):
                stray.append(name + str(row) + "=" + repr(value))
    if stray:
        failures.append("cells that should be blank hold data: " +
                        ", ".join(stray[:5]))
    print("   " + str(len(typed) * ROWS) + " course cells checked, " +
          str(len(stray)) + " non-blank")

    # The GPA flags are the exception: they must be FALSE, never empty, or
    # SUMPRODUCT sees a blank where it expects a boolean.
    bad_flags = []
    for name in ("gpa1", "gpa2"):
        if name not in headings:
            failures.append("no " + name + " column")
            continue
        for row in range(first_row, last_row + 1):
            if sheet.cell(row=row, column=headings[name]).value is not False:
                bad_flags.append(name + str(row))
    if bad_flags:
        failures.append("gpa flag cells not set to FALSE: " +
                        ", ".join(bad_flags[:5]))
    print("   gpa1/gpa2 flags all FALSE:", not bad_flags)

    # --- 2. the dropdowns reach the bottom of the table ---------------------

    print()
    print("2. dropdowns on every blank row")
    for name in ("Gr.", "Ovr.", "Cat. 2020"):
        if name not in headings:
            failures.append("no " + name + " column")
            continue
        letter = openpyxl.utils.get_column_letter(headings[name])
        ranges = data_validation_ranges(sheet, letter)
        covers = any(low <= first_row and high >= last_row
                     for low, high in ranges)
        print("   {:10s} column {:2s}  validation rows {}  covers 6-{}: {}"
              .format(name, letter, ranges or "none", last_row, covers))
        if not covers:
            failures.append(name + " dropdown does not reach row " +
                            str(last_row))

    # --- 3. hand-typed rows read back ---------------------------------------
    #
    # Simulating the colleague this script is for: they open the blank file,
    # type two courses, and later hand it back to be merged. Writing to a
    # generated tracker with openpyxl is fine HERE because the file is
    # scratch — the pipeline itself never read-modify-writes a tracker.

    print()
    print("3. hand-entered rows read back")
    entered = [
        {"Term": "2024-3F", "Course": "MATH 1003", "Title": "Introduction to "
         "Calculus I", "CH": 3, "Gr.": "A", "Cat. 2020": "Calc",
         "Notes": "typed by hand"},
        {"Term": "2025-1W", "Course": "MATH 1013", "Title": "Introduction to "
         "Calculus II", "CH": 3, "Gr.": "B+", "Cat. 2020": "Calc",
         "Ovr.": "incl."},
    ]
    for offset, values in enumerate(entered):
        for name, value in values.items():
            sheet.cell(row=first_row + offset, column=headings[name],
                       value=value)
    sheet.cell(row=first_row, column=headings["gpa1"], value=True)
    workbook.save(str(tracker))

    read_back = tracker_reader.read_tracker(tracker)
    print("   student:", read_back["student"], " track:", read_back["track"])
    print("   rows read:", len(read_back["rows"]))
    for row in read_back["rows"]:
        print("     {:9s} {:10s} {:3s} {:5s} cat={:6s} ovr={:5s} gpa1={}"
              .format(row["term"], row["course"], str(row["hours"]),
                      row["grade"], row.get("cat_2020", ""), row["ovr"],
                      row["gpa1"]))

    if len(read_back["rows"]) != len(entered):
        failures.append("read back " + str(len(read_back["rows"])) +
                        " rows, expected " + str(len(entered)) +
                        " (are the blank rows being counted?)")
    else:
        first, second = read_back["rows"]
        if first["grade"] != "A" or first.get("cat_2020") != "Calc":
            failures.append("first hand-typed row came back as " + repr(first))
        if not first["gpa1"]:
            failures.append("gpa1 tick did not survive the read")
        if second["ovr"] != "incl.":
            failures.append("Ovr. did not survive the read: " +
                            repr(second["ovr"]))
        if read_back["track"] != "Honours":
            failures.append("track came back as " +
                            repr(read_back["track"]) + ", expected 'Honours'")

    print()
    if failures:
        print("FAILED:")
        for failure in failures:
            print("  -", failure)
        return 1
    print("PASSED — blank tracker is empty, live to the last row, and reads "
          "back what a human types into it.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
