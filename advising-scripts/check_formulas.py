# /// script
# dependencies = ["openpyxl", "pdfplumber", "reportlab"]
# ///
# check_formulas.py — guard against array-unsafe formulas, and work an example.
# Usage: uv run check_formulas.py
#
# A regression check, not part of the shipped pipeline.
#
# SUMPRODUCT feeds a whole column into every grade test, so those tests must
# operate elementwise. Two idioms silently do NOT:
#
#   IFERROR(CHOOSE(MATCH(Courses[Gr.],{...},0),4.3,...),0)
#       CHOOSE rejects an array index -> #VALUE! -> IFERROR collapses the
#       whole factor to the SCALAR 0, so every GPA reads 0.00.
#
#   ISNUMBER(MATCH(Courses[Gr.],{...},0))
#       collapses to one TRUE/FALSE for the entire column, so Progress counts
#       every course in a category regardless of grade — failures included.
#
# Both were present in the first version of the generator and both are now
# written as sums of explicit (range="X") comparisons. This check fails if
# either idiom comes back, and then works a sub-GPA example end to end so the
# number can be compared against what Excel displays.

import re
import subprocess
import sys
import zipfile
from pathlib import Path

import openpyxl

import check_expected_values as expected
import check_support
import tracker_reader

HERE = Path(__file__).resolve().parent

FORBIDDEN = {
    "CHOOSE(": "CHOOSE cannot take an array index; it collapses the GPA "
               "numerator to a scalar 0",
    "ISNUMBER(MATCH(": "MATCH does not broadcast here; Progress would count "
                       "failed courses",
}


def sheet_formulas(xlsx_path):
    """Pull every formula out of the workbook's XML, unescaped."""
    with zipfile.ZipFile(xlsx_path) as archive:
        names = [n for n in archive.namelist()
                 if n.startswith("xl/worksheets/sheet")]
        text = "".join(archive.read(n).decode("utf-8") for n in names)
    formulas = re.findall(r"<f>(.*?)</f>", text, re.S)
    unescape = {"&quot;": '"', "&gt;": ">", "&lt;": "<", "&amp;": "&"}
    out = []
    for formula in formulas:
        for old, new in unescape.items():
            formula = formula.replace(old, new)
        out.append(formula)
    return out


def main():
    WORK, PDF, REQ = check_support.prepare("_formula_test")
    tracker = WORK / "tracker.xlsx"

    result = subprocess.run(
        [sys.executable, str(HERE / "transcript_to_tracker.py"), str(PDF),
         "-r", str(REQ), "-o", str(tracker), "-t", "Honours"],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        print(result.stderr, file=sys.stderr)
        return 1

    print("1. array-safety guard")
    formulas = sheet_formulas(tracker)
    print("   " + str(len(formulas)) + " formulas in the workbook")
    failures = []
    for idiom, why in FORBIDDEN.items():
        hits = [f for f in formulas if idiom in f]
        if hits:
            failures.append(idiom + " is back (" + str(len(hits)) +
                            " formulas) — " + why)
        else:
            print("   OK  no " + idiom)

    if failures:
        print()
        for failure in failures:
            print("  FAIL", failure)
        return 1

    # --- 2. work a sub-GPA example ------------------------------------------

    print()
    print("2. worked sub-GPA example")
    workbook = openpyxl.load_workbook(str(tracker))
    sheet = workbook["Tracker"]

    headings = {sheet.cell(row=5, column=c).value: c
                for c in range(1, sheet.max_column + 1)}
    gpa1_col = headings["gpa1"]

    flagged = []
    for row in range(6, 70):
        course = sheet.cell(row=row, column=headings["Course"]).value
        grade = sheet.cell(row=row, column=headings["Gr."]).value
        if course and str(course).startswith("MATH") and grade:
            sheet.cell(row=row, column=gpa1_col, value=True)
            flagged.append((course, grade,
                            sheet.cell(row=row,
                                       column=headings["CH"]).value))
    workbook.save(str(tracker))

    print("   flagged " + str(len(flagged)) + " graded MATH courses as gpa1:")
    for course, grade, hours in flagged:
        print("     {:12s} {:3s} {} ch".format(course, str(grade), hours))

    data = tracker_reader.read_tracker(tracker)
    rows = data["rows"]

    print()
    print("   Expected values — compare these against Excel:")
    print("     D3  Overall GPA :", expected.gpa(rows))
    print("     I3  gpa1        :", expected.gpa(rows, "gpa1"),
          "  (MATH courses only)")
    print("     J3  gpa2        :", expected.gpa(rows, "gpa2"),
          "  (nothing flagged)")
    print()
    print("   Tracker saved at:", tracker)
    print()
    if expected.gpa(rows) == 0 or expected.gpa(rows, "gpa1") == 0:
        print("  FAIL a GPA computed as 0 — the points factor collapsed")
        return 1
    print("PASSED — no array-unsafe idioms, and both GPAs are non-zero.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
