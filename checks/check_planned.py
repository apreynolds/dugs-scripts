# /// script
# dependencies = ["openpyxl", "pdfplumber"]
# ///
# check_planned.py — test the Planned column.
# Usage: uv run check_planned.py
#
# A regression check, not part of the shipped pipeline.
#
# The rule this covers: a course typed into the tracker with no grade yet — a
# blank Gr., or "IP" for one in progress this term — earns nothing, but IS
# hours the student is on course to hold. The Planned column reports the
# projected figure, banked plus planned, over the same requirement.
#
# Five things have to hold together, and none of them is visible in the
# generated XML by inspection:
#
#   1. Completed must NOT include planned hours. The strike-through rule on
#      the Courses table is written as the exact negation of the Completed
#      test so the two can never disagree; folding planned hours in would
#      leave a row counted but not struck, and green would stop meaning
#      "banked".
#   2. The planned SUMPRODUCT must be ELEMENTWISE (trap 2). SUMPRODUCT feeds a
#      whole column in, and MATCH/CHOOSE collapse to a scalar — silently, with
#      a well-formed formula and a plausible-looking number.
#   3. The amber rule needs an ISNUMBER guard (trap 4), or it fires on a
#      no-target row, where "—" > 0 is TRUE because Excel sorts text above
#      every number.
#   4. The amber rule needs its `earned < required` bound, or it also fires on
#      rows that are already green — putting a second status beside a green
#      that has already answered the question.
#   5. col_progress must stay at col_category + 3. tracker_reader reaches the
#      constraint ticks by that hardcoded offset, inside a bare `except`, so
#      an insertion ahead of it loses every tick with no error at all. Adding
#      Planned AFTER it, and hiding Req. CH rather than deleting it, is what
#      keeps that true — and this is where a future change is caught.
#
# What this canNOT check is whether Excel agrees. Nothing offline evaluates
# conditional formatting or structured table references. See "Verifying a
# change" in docs/degree-tracker.md.

import re
import subprocess
import sys
import zipfile
from pathlib import Path

import openpyxl

# The scripts live one folder up, and their helper modules in lib/.
REPO = Path(__file__).resolve().parent.parent
sys.path[:0] = [str(REPO), str(REPO / "lib")]

import check_expected_values as expected
import check_support

HERE = Path(__file__).resolve().parent

# (grade, override, counts as planned, why). The pairs that matter are the two
# IP rows — the SAME grade is planned on its own and is nothing once the
# advisor forces it out — and the blank/W pair, which separates "not marked
# yet" from "marked, and it was a withdrawal".
SEMANTICS = [
    ("", "", True, "no grade yet: registered or intended"),
    ("IP", "", True, "in progress this term"),
    ("", "incl.", True, '"incl." rescues a grade; a blank is not a grade'),
    ("IP", "excl.", False, "forced out, exactly as for Completed"),
    ("", "excl.", False, "likewise"),
    ("W", "", False, "withdrawn: marked, and it earned nothing"),
    ("F", "", False, "failed: marked"),
    ("B", "", False, "graded, so it is Completed, not Planned"),
    ("CR", "", False, "transfer credit is banked, not planned"),
]

# Planted into the tracker before the numbers are compared. Course codes that
# cannot collide with the sample transcript's own rows.
PLANTED = [
    ("2028-1W", "MATH 4213", "PLANNED TOPOLOGY", 3, "", "", "3rdMath"),
    ("2028-1W", "MATH 4223", "IN PROGRESS ALGEBRA", 3, "IP", "", "3rdMath"),
    ("2028-1W", "MATH 4233", "PLANNED BUT EXCLUDED", 3, "IP", "excl.",
     "3rdMath"),
    ("2028-1W", "MATH 4243", "WITHDRAWN", 3, "W", "", "3rdMath"),
]


def failing(message, failures):
    failures.append(message)
    print("  FAIL", message)


def build_tracker(work, failures):
    """Generate a tracker, plant the four edge-case rows, save it back."""
    requirements = check_support.sample_requirements(work)
    tracker = work / "tracker.xlsx"
    result = subprocess.run(
        [sys.executable, str(REPO / "transcript_to_tracker.py"),
         str(check_support.sample_transcript()),
         "-r", str(requirements), "-o", str(tracker), "-t", "Honours"],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        sys.stderr.write(result.stdout + result.stderr)
        failing("could not generate a tracker", failures)
        return None, None
    return tracker, requirements


def plant_rows(tracker, failures):
    """Type the four edge-case rows into the spare rows at the foot.

    Deliberately written the way an ADVISOR would add a planned course —
    straight into the blank rows the generator leaves — rather than through
    the merge, because that is the workflow the column exists to serve.
    """
    workbook = openpyxl.load_workbook(str(tracker))
    sheet = workbook["Tracker"]

    headings = {sheet.cell(row=5, column=c).value: c
                for c in range(1, sheet.max_column + 1)
                if sheet.cell(row=5, column=c).value}

    # The first wholly empty row of the Courses table.
    row = 6
    while sheet.cell(row=row, column=1).value:
        row += 1

    for term, course, title, hours, grade, ovr, cat in PLANTED:
        sheet.cell(row=row, column=headings["Term"], value=term)
        sheet.cell(row=row, column=headings["Course"], value=course)
        sheet.cell(row=row, column=headings["Title"], value=title)
        sheet.cell(row=row, column=headings["CH"], value=hours)
        sheet.cell(row=row, column=headings["Gr."], value=grade)
        sheet.cell(row=row, column=headings["Ovr."], value=ovr)
        sheet.cell(row=row, column=headings["Cat. 2020"], value=cat)
        row += 1

    workbook.save(str(tracker))
    return headings


def check_semantics(failures):
    """The rule itself, through the model the manual comparison relies on."""
    print("1. counting semantics")
    for grade, override, should, why in SEMANTICS:
        row = {"grade": grade, "ovr": override, "hours": 3}
        got = expected.counts_as_planned(row)
        described = "{:2s} {:6s}".format(grade or "—", override or "—")
        if got != should:
            failing(described + " -> " + str(got) + ", wanted " + str(should) +
                    " (" + why + ")", failures)
        else:
            print("   OK  " + described + " -> " +
                  ("planned" if should else "not planned") + "  (" + why + ")")

    # The two columns must never both claim the same row. Nothing enforces
    # this structurally — it falls out of "graded" and "not graded" being
    # complementary — so it is worth asserting rather than assuming.
    for grade, override, _, _ in SEMANTICS:
        row = {"grade": grade, "ovr": override, "hours": 3}
        if expected.counts_as_planned(row) \
                and expected.counts_toward_progress(row, "C"):
            failing("grade " + repr(grade) + " with override " +
                    repr(override) + " counts as BOTH completed and planned, "
                    "so its hours would be claimed twice", failures)
            return
    print("   OK  no row can be both completed and planned")


def check_layout(tracker, failures):
    """The columns, and the offset that fails silently if it moves."""
    print()
    print("2. layout")

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
        return None

    def letter(offset):
        return openpyxl.utils.get_column_letter(block_col + offset)

    def header(offset):
        return sheet.cell(row=5, column=block_col + offset).value

    if header(3) != "Completed":
        failing("col_category + 3 is headed " + repr(header(3)) +
                ", expected 'Completed' — tracker_reader reads the constraint "
                "ticks at that offset", failures)
    else:
        print("   OK  Completed sits at col_category + 3")

    if header(4) != "Planned":
        failing("col_category + 4 is headed " + repr(header(4)) +
                ", expected 'Planned'", failures)
    elif sheet.column_dimensions[letter(4)].hidden:
        failing("the Planned column is hidden", failures)
    else:
        print("   OK  Planned sits beside it, and is visible")

    # Req. CH keeps its header, so unhiding it gives a labelled column rather
    # than a mystery one — but it must not be on screen, or every row prints
    # its requirement twice.
    if header(2) != "Req. CH":
        failing("col_category + 2 is headed " + repr(header(2)) +
                ", expected 'Req. CH'", failures)
    elif not sheet.column_dimensions[letter(2)].hidden:
        failing("Req. CH is visible; its value is already the denominator of "
                "both columns beside it", failures)
    else:
        print("   OK  Req. CH is present but hidden")

    for offset, name in ((5, "earned"), (6, "planned hours"), (7, "bar")):
        if not sheet.column_dimensions[letter(offset)].hidden:
            failing("the hidden " + name + " helper at +" + str(offset) +
                    " is visible", failures)
    print("   OK  the three numeric helpers are hidden")

    return block_col


def check_ticks_survive(tracker, failures):
    """Pin the offset tracker_reader depends on, since it fails silently."""
    print()
    print("3. the constraint ticks still read back")

    import tracker_reader

    workbook = openpyxl.load_workbook(str(tracker))
    sheet = workbook["Tracker"]

    # Find the CONSTRAINTS heading and tick the first box in the column
    # tracker_reader will look in — which is the point: if the writer and the
    # reader disagree about where that is, this comes back empty.
    ticked = None
    for row in range(1, sheet.max_row + 1):
        for column in range(1, sheet.max_column + 1):
            if str(sheet.cell(row=row, column=column).value or "") \
                    == "CONSTRAINTS":
                ticked = (row + 1, column + 3)
                break
        if ticked:
            break
    if ticked is None:
        failing("no CONSTRAINTS heading in the generated tracker", failures)
        return

    sheet.cell(row=ticked[0], column=ticked[1], value=True)
    workbook.save(str(tracker))

    data = tracker_reader.read_tracker(tracker)
    ticks = data["constraints"].get("2020")
    if not ticks:
        failing("tracker_reader read no constraint ticks back — col_progress "
                "has moved away from col_category + 3, and every advisor's "
                "attestations would be lost on the next -u, silently",
                failures)
    elif ticks[0] is not True:
        failing("the ticked constraint read back as " + repr(ticks[0]),
                failures)
    else:
        print("   OK  a ticked constraint survives the round trip (" +
              str(len(ticks)) + " constraints read)")


def check_formula_shape(tracker, failures):
    """The planned SUMPRODUCT, and the amber rule's four terms."""
    print()
    print("4. formula shape")

    workbook = openpyxl.load_workbook(str(tracker))
    sheet = workbook["Tracker"]

    planned = [cell.value for row in sheet.iter_rows() for cell in row
               if isinstance(cell.value, str)
               and 'Courses[Gr.]="IP"' in cell.value
               and cell.value.startswith("=SUMPRODUCT")]
    if not planned:
        failing("no planned SUMPRODUCT found in the sheet", failures)
        return

    formula = planned[0]

    # Trap 2. Either idiom collapses the whole factor to a scalar, silently.
    for idiom in ("CHOOSE(", "ISNUMBER(MATCH("):
        if idiom in formula:
            failing("the planned sum uses " + idiom + ", which is not "
                    "elementwise inside SUMPRODUCT: " + formula, failures)
            return
    print("   OK  the planned sum is elementwise (no CHOOSE, no "
          "ISNUMBER(MATCH))")

    if '(Courses[Ovr.]<>"excl.")' not in formula:
        failing('the planned sum does not exclude "excl." rows, so an '
                "override the advisor set would be ignored: " + formula,
                failures)
    else:
        print('   OK  the planned sum honours "excl."')

    if '(Courses[Ovr.]="incl.")' in formula:
        failing('the planned sum honours "incl.", but "incl." rescues a real '
                "grade that fell short and a blank is not a grade: " + formula,
                failures)
    else:
        print('   OK  the planned sum ignores "incl."')

    if '(Courses[Gr.]="")' not in formula:
        failing("the planned sum does not count a blank grade, which is the "
                "main case: " + formula, failures)
    else:
        print("   OK  a blank grade and IP are both counted")


def check_amber_rule(tracker, failures):
    """The amber rule, read out of the saved XML, with its guards."""
    print()
    print("5. the 'on plan' rule")

    with zipfile.ZipFile(tracker) as archive:
        sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode("utf-8")
        styles_xml = archive.read("xl/styles.xml").decode("utf-8")

    formulas = re.findall(r"<formula>(.*?)</formula>", sheet_xml, re.S)

    # The amber rule is the only one that adds two cells together.
    amber = [f for f in formulas if "+$" in f and "&gt;=" in f]
    if len(amber) != 1:
        failing("expected exactly one 'on plan' rule, found " +
                str(len(amber)) + ": " + repr(amber), failures)
        return
    rule = amber[0]

    # Trap 4. "—" > 0 is TRUE in Excel, so without this the rule fires on a
    # no-target row and claims the plan closes a gap that does not exist.
    if "ISNUMBER" not in rule:
        failing("the 'on plan' rule has no ISNUMBER guard, so it fires on a "
                "no-target row: " + rule, failures)
    else:
        print("   OK  guarded by ISNUMBER")

    # The bound that keeps amber off rows that are already green. Same shape,
    # and the same reason, as the `rank < bar` bound on the "incl." rescue.
    if "&lt;$" not in rule:
        failing("the 'on plan' rule has no 'earned < required' bound, so it "
                "also fires on categories that are already met: " + rule,
                failures)
    else:
        print("   OK  bounded below the requirement")

    # CF formulas are written for the top-left cell and re-evaluated
    # relatively, so an absolute row would test row 6 for every row.
    if re.search(r"\$[A-Z]+\$\d", rule):
        failing("the 'on plan' rule has an absolute ROW reference, so every "
                "row would test row 6's condition: " + rule, failures)
    else:
        print("   OK  column-absolute, row-relative")

    if "{" in rule:
        failing("the 'on plan' rule contains an array constant, which Excel "
                "refuses in conditional formatting — it repairs the file and "
                "discards every rule on the sheet: " + rule, failures)

    # The amber must be a real dxf, shaped the way Excel demands.
    if 'bgColor rgb="FFFFEB9C"' not in styles_xml:
        failing("the amber fill is missing from the dxfs, or its colour is "
                "in fgColor — which makes Excel discard every conditional "
                "format on the sheet", failures)
    else:
        print("   OK  the amber dxf puts its colour in bgColor")

    return rule


def _mutate_workbook(tracker, mutant, rewrite):
    """Copy a tracker, rewriting the amber rule's formula in sheet1.xml.

    Mutating the SAVED ARTEFACT rather than the generator is what makes this
    honest: check_amber_rule is run against a real workbook that differs from
    the real one only in the guard being tested.
    """
    with zipfile.ZipFile(tracker) as archive:
        entries = [(item, archive.read(item.filename))
                   for item in archive.infolist()]

    changed = False
    with zipfile.ZipFile(mutant, "w", zipfile.ZIP_DEFLATED) as out:
        for item, data in entries:
            if item.filename == "xl/worksheets/sheet1.xml":
                text = data.decode("utf-8")
                for formula in re.findall(r"<formula>(.*?)</formula>", text,
                                          re.S):
                    # The same predicate check_amber_rule uses to find it.
                    if "+$" in formula and "&gt;=" in formula:
                        replacement = rewrite(formula)
                        if replacement != formula:
                            text = text.replace(formula, replacement)
                            changed = True
                data = text.encode("utf-8")
            out.writestr(item, data)
    return changed


def check_mutations(work, tracker, failures):
    """Delete each guard from a real workbook, confirm the check notices.

    The assertions above are only worth having if they can fail. Both Excel
    traps in docs/degree-tracker.md shipped in code that passed every structural check and
    produced well-formed XML, so a guard nobody has watched fail is a guard
    nobody knows works.

    Each mutant must produce EXACTLY ONE failure — the one being tested. More
    than that means the mutation broke something else too, and the check would
    have gone red for the wrong reason.
    """
    print()
    print("6. mutation test — each guard, removed from a real workbook")

    for name, rewrite in (
            ("ISNUMBER guard",
             lambda f: f.replace("ISNUMBER($O6),", "")),
            ("earned < required bound",
             lambda f: re.sub(r",\$[A-Z]+\d+&lt;\$[A-Z]+\d+", "", f, count=1)),
    ):
        mutant = work / ("mutant-" + name.split()[0] + ".xlsx")
        if not _mutate_workbook(tracker, mutant, rewrite):
            failing("the " + name + " mutation changed nothing, so the "
                    "assertion for it proves nothing", failures)
            continue

        caught = []
        # Silence the mutant's own OK lines; only the verdict matters here.
        stdout = sys.stdout
        sys.stdout = open("/dev/null", "w")
        try:
            check_amber_rule(mutant, caught)
        finally:
            sys.stdout.close()
            sys.stdout = stdout

        if len(caught) != 1:
            failing("removing the " + name + " produced " + str(len(caught)) +
                    " failures, expected exactly 1: " + repr(caught), failures)
        else:
            print("   OK  removing the " + name + " is caught  (" +
                  caught[0].split(":")[0] + ")")


def check_numbers(tracker, requirements, failures):
    """The planted rows land in the right column, with the right hours."""
    print()
    print("7. the planted rows")

    import requirements_io
    import tracker_reader

    req = requirements_io.load_requirements(requirements)
    data = tracker_reader.read_tracker(tracker)
    rows = data["rows"]

    planted = {r["course"]: r for r in rows
               if r["course"] in {p[1] for p in PLANTED}}
    if len(planted) != len(PLANTED):
        failing("only " + str(len(planted)) + " of " + str(len(PLANTED)) +
                " planted rows read back", failures)
        return

    bar = req.pass_grade_for("3rdMath")
    in_cat = [r for r in rows if r.get("cat_2020") == "3rdMath"]
    earned = sum(r["hours"] or 0 for r in in_cat
                 if expected.counts_toward_progress(r, bar))
    planned_ch = sum(r["hours"] or 0 for r in in_cat
                     if expected.counts_as_planned(r))

    # Measured as a DELTA, not against a magic total. The sample transcript
    # already carries genuinely in-progress rows of its own — its last term
    # has no grades yet — so 3rdMath has planned hours before anything is
    # planted. That is a good thing (the fixture exercises the column without
    # help) and a bad number to hardcode.
    planted_codes = {p[1] for p in PLANTED}
    baseline = sum(r["hours"] or 0 for r in in_cat
                   if r["course"] not in planted_codes
                   and expected.counts_as_planned(r))
    if planned_ch - baseline != 6:
        failing("the four planted rows should add 6 planned ch to 3rdMath "
                "(blank + IP count; IP with excl. and W do not); they added " +
                str(planned_ch - baseline), failures)
    else:
        print("   OK  the planted rows add 6 planned ch — blank and IP count, "
              "the excl. and W rows do not")
        print("   ..  3rdMath already had " + str(int(baseline)) +
              " planned ch from the transcript's ungraded final term")

    # And none of them reached Completed. That is the invariant that keeps the
    # strike-through rule and the Completed column in agreement.
    for course, row in planted.items():
        if expected.counts_toward_progress(row, bar):
            failing(course + " (grade " + repr(row["grade"]) + ") counts "
                    "toward Completed; no ungraded course should", failures)
    print("   OK  no planted row reached Completed")

    required = None
    for category in req.categories:
        if category.code == "3rdMath":
            required = category.req_ch.get("Honours", 0)
    print("   ..  3rdMath on Honours: " + str(earned) + " / " +
          str(required) + " completed, " + str(earned + planned_ch) + " / " +
          str(required) + " projected")


def main():
    failures = []
    work = check_support.fresh_dir("_planned_test")

    check_semantics(failures)

    tracker, requirements = build_tracker(work, failures)
    if tracker is None:
        print()
        print("FAILED — could not generate a tracker")
        return 1

    plant_rows(tracker, failures)
    check_layout(tracker, failures)
    check_ticks_survive(tracker, failures)
    check_formula_shape(tracker, failures)
    check_amber_rule(tracker, failures)
    check_mutations(work, tracker, failures)
    check_numbers(tracker, requirements, failures)

    print()
    if failures:
        print("FAILED —", len(failures), "problem(s) above")
        return 1
    print("PASSED — planned hours are counted, kept out of Completed, and the "
          "constraint ticks still read back.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
