# /// script
# dependencies = ["openpyxl"]
# ///
# check_expected_values.py — compute what the tracker's formulas SHOULD show.
# Usage: uv run check_expected_values.py out.xlsx requirements/bsc-math-2020.xlsx [Track]
#
# A regression check, not part of the shipped pipeline. Reads the generated tracker's literal cells
# (course rows, categories) and re-implements the Progress / GPA / TOTAL
# semantics in Python, so the numbers Excel displays can be checked against an
# independent calculation.
#
# Why this rather than evaluating the workbook directly: the offline `formulas`
# package cannot parse structured table references (Courses[CH]), which the
# tracker uses throughout because they survive row insertion. Excel handles
# them fine — the original hand-built template used the same references.

import sys

import requirements_io
import tracker_reader

# Deliberately spelled out again rather than imported from requirements_io or
# tracker_writer. This file is the INDEPENDENT oracle Patrick's manual Excel
# comparison is measured against; sharing the tables with the code under test
# would let one mistake satisfy both sides.
RANK = {"A+": 9, "A": 8, "A-": 7, "B+": 6, "B": 5, "B-": 4,
        "C+": 3, "C": 2, "D": 1}
POINTS = {"A+": 4.3, "A": 4.0, "A-": 3.7, "B+": 3.3, "B": 3.0,
          "B-": 2.7, "C+": 2.3, "C": 2.0, "D": 1.0}
NON_GPA = {"", "W", "IP", "CR"}


def bar_for(row, req, label):
    """The passing bar for a row, from the category the advisor assigned.

    An uncategorised row falls back to the file's default — same as the
    IFERROR fallback in the tracker's Pass column, and it matters because the
    block TOTAL counts uncategorised rows.
    """
    return req.pass_grade_for(row.get("cat_" + label))


def counts_toward_progress(row, bar):
    """Mirror of the Progress SUMPRODUCT's inner test.

    `bar` is the grade letter this row has to reach, per its category.
    """
    grade = row["grade"]
    override = (row.get("ovr") or "").lower()
    rank = RANK.get(grade, 0)
    threshold = RANK.get(bar, 0)

    if override == "excl.":
        # Forced out. The rescue below cannot also apply — the advisor cannot
        # have said both things about one row.
        return False
    if grade == "CR" or rank >= threshold:
        return True
    # Forced in: any real grade that fell short. Never an F, and never a
    # blank, W or IP — those rank 0 and are not grades at all. Reaching here
    # already means the row fell short, which is what stops an "incl." on a
    # course that clears its bar from being counted a second time.
    return rank > 0 and override == "incl."


def counts_as_planned(row):
    """Mirror of the Planned SUMPRODUCT's inner test.

    A blank grade is a course registered or intended; "IP" is one in progress
    this term. Neither has been marked, so neither has earned anything — but
    both are hours the student is on course to hold, which is what the Planned
    column reports.

    "excl." takes a row out here exactly as it does for Progress. "incl." does
    nothing: it rescues a real grade that fell short, and a blank is not a
    grade.
    """
    grade = row["grade"]
    override = (row.get("ovr") or "").lower()
    if override == "excl.":
        return False
    return grade in ("", "IP")


def gpa(rows, flag=None):
    numerator = 0.0
    denominator = 0.0
    for row in rows:
        if flag and not row.get(flag):
            continue
        grade = row["grade"]
        if grade in NON_GPA:
            continue
        hours = row["hours"] or 0
        numerator += POINTS.get(grade, 0) * hours
        denominator += hours
    if not denominator:
        return "N/A"
    return round(numerator / denominator, 2)


def main():
    tracker_path, req_path = sys.argv[1], sys.argv[2]
    track = sys.argv[3] if len(sys.argv) > 3 else None

    req = requirements_io.load_requirements(req_path)
    data = tracker_reader.read_tracker(tracker_path)
    rows = data["rows"]
    track = track or data["track"] or req.tracks[0]

    label = req.label
    print("tracker :", tracker_path)
    print("track   :", track)
    print("rows    :", len(rows))
    print()

    print("Expected GPA values:")
    print("  D3  Overall GPA :", gpa(rows))
    print("  gpa1            :", gpa(rows, "gpa1"))
    print("  gpa2            :", gpa(rows, "gpa2"))
    print()

    print("Expected Completed / Planned columns "
          "(earned, then earned+planned, over required):")
    total_required = req.totals[track]
    fixed_sum = 0
    lines = []
    for category in req.categories:
        # None means "no target" — a catch-all. It contributes nothing to the
        # sum of what the categories claim, which is the point of it.
        required = category.req_ch.get(track, 0)
        if required is not None:
            fixed_sum += required
        bar = req.pass_grade_for(category.code)
        in_category = [r for r in rows
                       if r.get("cat_" + label) == category.code]
        earned = sum(r["hours"] or 0 for r in in_category
                     if counts_toward_progress(r, bar))
        planned = sum(r["hours"] or 0 for r in in_category
                      if counts_as_planned(r))
        lines.append((category.code, earned, planned, required, bar))

    for code, earned, planned, required, bar in lines:
        if required is None:
            # Rendered the way the tracker renders it: a count, no
            # denominator, and no verdict either way.
            print("  {:10s} {:>3d} ch        {:>3d} ch        "
                  "(pass {})  (no target)".format(
                      code, int(earned), int(earned + planned), bar))
            continue
        met = "  <- met" if required > 0 and earned >= required else ""
        # Amber in the tracker: not met today, met once the plan lands. The
        # "earned < required" bound is what keeps it off the rows that are
        # already green.
        if required > 0 and earned < required \
                and earned + planned >= required:
            met = "  <- on plan"
        greyed = "  (not on this track)" if required == 0 else ""
        print("  {:10s} {:>3d} / {:<3d} ch  {:>3d} / {:<3d} ch  "
              "(pass {}){}{}".format(
                  code, int(earned), int(required),
                  int(earned + planned), int(required), bar, met, greyed))

    # The TOTAL row is not filtered by category, so every row is measured
    # against its own bar — including the uncategorised ones, which fall back
    # to the file's default.
    grand_earned = sum(r["hours"] or 0 for r in rows
                       if counts_toward_progress(r, bar_for(r, req, label)))
    grand_planned = sum(r["hours"] or 0 for r in rows
                        if counts_as_planned(r))
    print("  {:10s} {:>3d} / {:<3d} ch  {:>3d} / {:<3d} ch".format(
        "TOTAL", int(grand_earned), int(total_required),
        int(grand_earned + grand_planned), int(total_required)))
    print()
    # The categories are no longer expected to reconcile to the total — the
    # difference IS the free electives. What must still hold is that they do
    # not over-claim, which is what requirements_io refuses outright.
    print("Sanity check: categories claim", fixed_sum, "ch of the declared",
          total_required, "ch;", total_required - fixed_sum,
          "ch are free ->",
          "OK" if fixed_sum <= total_required else "OVER-CLAIMED")
    print("  catch-all categories:", ", ".join(req.no_target_codes()) or "none")


if __name__ == "__main__":
    main()
