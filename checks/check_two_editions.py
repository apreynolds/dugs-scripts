# /// script
# dependencies = ["openpyxl", "pdfplumber"]
# ///
# check_two_editions.py — test side-by-side calendar editions in one tracker.
# Usage: uv run check_two_editions.py
#
# A regression check, not part of the shipped pipeline. Builds a hypothetical 2027 edition from the
# 2020 one (a couple of requirements changed, plus a third "Co-op" track that
# does not exist in 2020), then generates a tracker carrying BOTH editions and
# checks the layout came out right.
#
# This is the case the old hand-built template supported by bolting a second
# table onto the sheet. Here it falls out of passing --requirements twice.

import shutil
import subprocess
import sys
from pathlib import Path

import openpyxl

# The scripts live one folder up, and their helper modules in lib/.
REPO = Path(__file__).resolve().parent.parent
sys.path[:0] = [str(REPO), str(REPO / "lib")]

import check_support
import requirements_io

HERE = Path(__file__).resolve().parent


def make_2027_edition(source, destination):
    """Derive a plausible 2027 edition: a few numbers move, a track is added."""
    shutil.copy(source, destination)
    workbook = openpyxl.load_workbook(str(destination))

    sheet = workbook["Requirements"]
    # Add a third track column, "Co-op", after the existing Honours column.
    header = [c.value for c in sheet[1]]
    honours_col = header.index("Honours") + 1
    note_col = header.index("Note") + 1
    sheet.insert_cols(note_col)
    sheet.cell(row=1, column=note_col, value="Co-op")
    for row in range(2, sheet.max_row + 1):
        sheet.cell(row=row, column=note_col,
                   value=sheet.cell(row=row, column=honours_col).value)

    # 2027 changes: Stat rises to 9 ch, SciLec falls to 15 ch.
    for row in range(2, sheet.max_row + 1):
        code = sheet.cell(row=row, column=1).value
        if code == "Stat":
            for column in (honours_col - 1, honours_col, note_col):
                sheet.cell(row=row, column=column, value=9)
        if code == "SciLec":
            for column in (honours_col - 1, honours_col, note_col):
                sheet.cell(row=row, column=column, value=15)

    meta = workbook["Meta"]
    for row in range(1, meta.max_row + 1):
        if meta.cell(row=row, column=1).value == "Edition":
            meta.cell(row=row, column=2, value="2027")
    meta.append(["Total CH: Co-op", 132])

    workbook.save(str(destination))


def main():
    WORK, PDF, REQ2020 = check_support.prepare("_twoedition_test")

    req2027 = WORK / "bsc-math-2027.xlsx"
    make_2027_edition(REQ2020, req2027)

    loaded = requirements_io.load_requirements(req2027)
    print("2027 edition loaded:")
    print("   tracks:", loaded.tracks)
    print("   totals:", loaded.totals)
    stat = [c for c in loaded.categories if c.code == "Stat"][0]
    print("   Stat  :", stat.req_ch)
    print()

    tracker = WORK / "two-edition.xlsx"
    result = subprocess.run(
        [sys.executable, str(REPO / "transcript_to_tracker.py"), str(PDF),
         "-r", str(REQ2020), "-r", str(req2027),
         "-o", str(tracker), "-t", "Honours"],
        capture_output=True, text=True,
    )
    print(result.stdout.strip())
    if result.returncode != 0:
        print(result.stderr, file=sys.stderr)
        return 1

    print()
    workbook = openpyxl.load_workbook(str(tracker))
    sheet = workbook["Tracker"]

    headings = [sheet.cell(row=5, column=c).value
                for c in range(1, sheet.max_column + 1)]
    print("Courses table columns:")
    print("  ", [h for h in headings[:12] if h])

    print()
    print("Requirements block headings on row 4:")
    for column in range(1, sheet.max_column + 1):
        value = sheet.cell(row=4, column=column).value
        if value:
            print("   column", openpyxl.utils.get_column_letter(column),
                  "->", value)

    failures = []
    if "Cat. 2020" not in headings:
        failures.append("missing Cat. 2020 column")
    if "Cat. 2027" not in headings:
        failures.append("missing Cat. 2027 column")

    row4 = [sheet.cell(row=4, column=c).value
            for c in range(1, sheet.max_column + 1)]
    if "DEGREE REQUIREMENTS (2020)" not in row4:
        failures.append("missing 2020 requirements block")
    if "DEGREE REQUIREMENTS (2027)" not in row4:
        failures.append("missing 2027 requirements block")

    # The track dropdown must offer only tracks the FIRST edition defines,
    # since B3 drives both blocks.
    print()
    if failures:
        print("FAILED:")
        for failure in failures:
            print("  -", failure)
        return 1
    print("PASSED — two independent requirement blocks and two Cat. columns.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
