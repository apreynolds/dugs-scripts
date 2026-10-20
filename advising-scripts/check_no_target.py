# /// script
# dependencies = ["openpyxl", "pdfplumber"]
# ///
# check_no_target.py — test the "no target" category (the catch-all).
# Usage: uv run check_no_target.py
#
# A regression check, not part of the shipped pipeline.
#
# The rule this covers: a category whose track cells hold "—" has NO target.
# It reports what has been banked and is never met or unmet. FreeElec is the
# case that motivated it — free electives have no fixed count, only "enough to
# reach the degree total", which is the TOTAL row's job.
#
# Three things have to hold together, and none of them is visible in the
# generated XML by inspection:
#
#   1. "—" and 0 must stay DIFFERENT. 0 means "not part of this track" and
#      greys the row out; a blank cell still means 0, so only a deliberately
#      typed dash opts out. A blank that meant "no target" would fail open.
#   2. The Req. CH cell holds TEXT, and Excel sorts text above every number.
#      So `$O6>0` is TRUE there, and the "met" rule would fire on a row that
#      cannot be met. Both existing rules carry an ISNUMBER guard for this;
#      without it they pass every offline check and are wrong in Excel.
#   3. The no-target row needs its own colour. Grey is spoken for, and leaving
#      it unfilled makes it read as "unmet" among green rows — the exact false
#      alarm the change exists to remove.
#
# What this canNOT check is whether Excel agrees. Nothing offline evaluates
# conditional formatting. See "Verifying a change" in CLAUDE.md.

import re
import subprocess
import sys
import zipfile
from pathlib import Path

import openpyxl

import check_support
import requirements_io
import tracker_writer

HERE = Path(__file__).resolve().parent

# The fills that must all stay distinguishable. A no-target row that happened
# to share the grey would silently read as "not on this track".
DISTINCT_FILLS = {
    "met (green)": "FFC6EFCE",
    "off-track (grey)": "FFE7E6E6",
    "excluded (red)": "FFFFC7CE",
    "no target (blue)": "FFDDEBF7",
    "on plan (amber)": "FFFFEB9C",
}


def failing(message, failures):
    failures.append(message)
    print("  FAIL", message)


def check_parsing(work, failures):
    """"—" parses to None, and 0 and blank still parse to 0."""
    print("1. parsing the sentinel")

    req = requirements_io.load_requirements(
        check_support.sample_requirements(work))

    codes = req.no_target_codes()
    if codes != ["FreeElec"]:
        failing("expected FreeElec to be the only no-target category, got " +
                repr(codes), failures)
    else:
        print("   OK  FreeElec is the file's only no-target category")

    free = [c for c in req.categories if c.code == "FreeElec"][0]
    if any(v is not None for v in free.req_ch.values()):
        failing("FreeElec should be None on every track, got " +
                repr(free.req_ch), failures)
    else:
        print("   OK  FreeElec is None on every track")

    # The distinction the whole design rests on. Thesis is 0 on Major — "not
    # part of this track" — and that must NOT collapse into "no target".
    thesis = [c for c in req.categories if c.code == "Thesis"][0]
    if thesis.req_ch["Major"] != 0:
        failing("Thesis should be 0 on Major (not on that track), got " +
                repr(thesis.req_ch["Major"]), failures)
    else:
        print("   OK  0 is still 0 — 'not on this track' did not become "
              "'no target'")

    # Every dash a person might actually type, since Excel's autocorrect
    # swaps one for another depending on settings.
    missed = [d for d in sorted(requirements_io.NO_TARGET_LITERALS)
              if requirements_io._cell_req_ch(d, "test", "test") is not None]
    if missed:
        failing("not read as no-target: " + " ".join(missed), failures)
    else:
        print("   OK  all of " +
              " ".join(sorted(requirements_io.NO_TARGET_LITERALS)) +
              " read as no-target")

    if requirements_io._cell_req_ch("", "test", "test") != 0:
        failing("a blank cell should still be 0, not no-target", failures)
    else:
        print("   OK  a blank cell is still 0 — a forgotten cell fails safe")


def check_validation(work, failures):
    """Half-marked categories and unexplained gaps are caught."""
    print()
    print("2. validation")

    # A category marked "—" on one track and numbered on another describes
    # two different things under one code, and would make the tracker's row
    # change kind as B3 changes.
    path = work / "half-marked.xlsx"
    workbook = check_support.workbook_from_tsv(
        check_support.REQUIREMENTS_TSV)
    sheet = workbook["Requirements"]
    for row in range(2, sheet.max_row + 1):
        if sheet.cell(row=row, column=1).value == "FreeElec":
            sheet.cell(row=row, column=4, value=6)   # Major gets a number
            break
    workbook.save(str(path))
    try:
        requirements_io.load_requirements(path)
    except requirements_io.RequirementsError as error:
        if "FreeElec" not in str(error):
            failing("half-marked category rejected without naming it: " +
                    str(error), failures)
        else:
            print("   OK  a category marked on some tracks but not others "
                  "is refused")
    else:
        failing("a category marked '—' on Honours but 6 on Major was "
                "accepted", failures)

    # Deleting the catch-all leaves credit hours unaccounted for. That is a
    # warning, not an error — a programme may legitimately have a gap — but
    # it has to be said, because it is what the old slack mechanism used to
    # guarantee could not happen.
    path = work / "no-catch-all.xlsx"
    workbook = check_support.workbook_from_tsv(
        check_support.REQUIREMENTS_TSV)
    sheet = workbook["Requirements"]
    for row in range(2, sheet.max_row + 1):
        if sheet.cell(row=row, column=1).value == "FreeElec":
            sheet.delete_rows(row)
            break
    workbook.save(str(path))

    captured = _load_capturing_stderr(path)
    if "unaccounted" not in captured:
        failing("removing the catch-all produced no warning about the "
                "unaccounted hours; got " + repr(captured), failures)
    elif "26" not in captured:
        failing("the gap warning does not say how many ch are missing: " +
                repr(captured), failures)
    else:
        print("   OK  removing the catch-all warns about the 26 unaccounted "
              "ch")

    # A retired Meta key must be called out rather than ignored in silence.
    path = work / "stale-slack.xlsx"
    workbook = check_support.workbook_from_tsv(
        check_support.REQUIREMENTS_TSV)
    workbook["Meta"].append(["Slack category", "FreeElec"])
    workbook.save(str(path))
    captured = _load_capturing_stderr(path)
    if "Slack category" not in captured:
        failing("a leftover 'Slack category' Meta row was ignored in "
                "silence; got " + repr(captured), failures)
    else:
        print("   OK  a leftover 'Slack category' row is called out")


def _load_capturing_stderr(path):
    """Load a requirements file and return whatever it wrote to stderr.

    Run in a subprocess rather than redirecting sys.stderr in-process: the
    loader writes through the module-level `sys.stderr` it imported, so
    swapping the attribute here would be one refactor away from silently
    capturing nothing and passing.
    """
    result = subprocess.run(
        [sys.executable, "-c",
         "import sys; sys.path.insert(0, " + repr(str(HERE)) + "); "
         "import requirements_io; "
         "requirements_io.load_requirements(" + repr(str(path)) + ")"],
        capture_output=True, text=True,
    )
    return result.stderr


def check_tracker(work, failures):
    """The generated sheet renders the row, and the CF rules exclude it."""
    print()
    print("3. the generated tracker")

    tracker = work / "tracker.xlsx"
    result = subprocess.run(
        [sys.executable, str(HERE / "transcript_to_tracker.py"),
         str(check_support.sample_transcript()),
         "-r", str(work / "bsc-math-2020.xlsx"),
         "-o", str(tracker), "-t", "Honours"],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        sys.stderr.write(result.stdout + result.stderr)
        failing("could not generate a tracker", failures)
        return

    workbook = openpyxl.load_workbook(str(tracker))
    sheet = workbook["Tracker"]

    block_col = None
    for column in range(1, sheet.max_column + 1):
        if str(sheet.cell(row=4, column=column).value or "").startswith(
                "DEGREE REQUIREMENTS"):
            block_col = column
            break
    if block_col is None:
        failing("no requirements block found on row 4", failures)
        return

    free_row = None
    for row in range(6, sheet.max_row + 1):
        if sheet.cell(row=row, column=block_col).value == "FreeElec":
            free_row = row
            break
    if free_row is None:
        failing("FreeElec has no row in the generated block — the catch-all "
                "must stay a category, or a course judged to be filler has "
                "nowhere to go but a blank Cat. cell", failures)
        return

    # The lookup table is where the dash actually lives; the visible Req. CH
    # cell only resolves it. A blank here would be read by Excel as 0.
    hidden = workbook["_TrackData"]
    dashes = [row for row in hidden.iter_rows(values_only=True)
              if row and row[0] == "FreeElec"]
    if not dashes or not all(v == tracker_writer.NO_TARGET_TEXT
                             for v in dashes[0][1:3]):
        failing("_TrackData should hold text " +
                repr(tracker_writer.NO_TARGET_TEXT) +
                " for FreeElec on every track, got " + repr(dashes), failures)
    else:
        print("   OK  _TrackData holds the dash as TEXT, not a blank")

    progress = sheet.cell(row=free_row, column=block_col + 3).value or ""
    if " / " in progress:
        failing("the FreeElec Progress cell still has a denominator: " +
                progress, failures)
    else:
        print("   OK  the FreeElec Progress cell has no denominator")

    total_row = free_row + 1
    total_progress = sheet.cell(row=total_row, column=block_col + 3).value \
        or ""
    if sheet.cell(row=total_row, column=block_col).value != "TOTAL":
        failing("expected the TOTAL row directly below FreeElec", failures)
    elif " / " not in total_progress:
        failing("the TOTAL row lost its denominator: " + total_progress,
                failures)
    else:
        print("   OK  the TOTAL row still carries one — it is the verdict now")

    _check_rules(tracker, failures)


def _check_rules(tracker, failures):
    """The three conditional-formatting rules, read out of the saved XML."""
    with zipfile.ZipFile(tracker) as archive:
        sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
        styles_xml = archive.read("xl/styles.xml").decode("utf-8")

    formulas = re.findall(r"<formula>(.*?)</formula>", sheet_xml, re.S)

    # The guard that cannot be checked any other way. Without it the "met"
    # rule fires on a text Req. CH cell, because Excel sorts text above every
    # number — and the generated XML looks perfectly well-formed either way.
    #
    # This now matches TWO rules: "met" on the Completed column and the amber
    # "the plan closes this" rule on Planned. Both compare against Req. CH, so
    # both need the guard, and requiring it of every match is the point rather
    # than a looseness in the filter.
    met = [f for f in formulas if "&gt;=" in f and "$O" in f]
    if len(met) < 2:
        failing("expected both the 'met' and the 'on plan' rules to compare "
                "against Req. CH; found " + str(len(met)) + ": " + repr(met),
                failures)
    elif not all("ISNUMBER" in f for f in met):
        failing("a rule comparing against Req. CH has no ISNUMBER guard, so "
                "it will fire on a no-target row: " + repr(met), failures)
    else:
        print("   OK  both rules comparing against Req. CH are guarded by "
              "ISNUMBER")

    off = [f for f in formulas if "=0)" in f or f.endswith("=0")]
    if not off:
        failing("could not find the 'not on this track' rule", failures)
    elif not all("ISNUMBER" in f for f in off):
        failing("the off-track rule has no ISNUMBER guard: " + repr(off),
                failures)
    else:
        print("   OK  the off-track rule is guarded by ISNUMBER")

    if not any(f.startswith("ISTEXT(") for f in formulas):
        failing("no ISTEXT rule — the no-target row has no colour of its own, "
                "so it reads as unmet among green rows", failures)
    else:
        print("   OK  a no-target rule keys off the cell's TYPE, not its row")

    # Every fill distinct, and every one shaped the way Excel demands of a
    # dxf: colour in bgColor, fgColor left automatic. openpyxl's usual
    # PatternFill puts it in fgColor, which makes Excel repair the file and
    # discard EVERY conditional-formatting rule on the sheet.
    dxfs = re.search(r"<dxfs.*?</dxfs>", styles_xml, re.S)
    dxf_text = dxfs.group(0) if dxfs else ""
    for name, colour in DISTINCT_FILLS.items():
        if 'bgColor rgb="' + colour + '"' not in dxf_text:
            failing("the " + name + " fill " + colour + " is missing from "
                    "the dxfs, or is not in bgColor", failures)
        else:
            print("   OK  " + name + " present, colour in bgColor")

    if len(set(DISTINCT_FILLS.values())) != len(DISTINCT_FILLS):
        failing("two of the fills share a colour, so two different states "
                "look identical", failures)


def main():
    failures = []
    work = check_support.fresh_dir("_no_target_test")
    check_parsing(work, failures)
    check_validation(work, failures)
    check_tracker(work, failures)

    print()
    if failures:
        print("FAILED —", len(failures), "problem(s) above")
        return 1
    print("PASSED — the catch-all reports a count, is never met or unmet, "
          "and 0 still means 'not on this track'.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
