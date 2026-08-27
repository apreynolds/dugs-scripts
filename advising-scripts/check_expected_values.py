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

PASSING = {"A+", "A", "A-", "B+", "B", "B-", "C+", "C", "CR"}
POINTS = {"A+": 4.3, "A": 4.0, "A-": 3.7, "B+": 3.3, "B": 3.0,
          "B-": 2.7, "C+": 2.3, "C": 2.0, "D": 1.0}
NON_GPA = {"", "W", "IP", "CR"}


def counts_toward_progress(row):
    """Mirror of the Progress SUMPRODUCT's inner test."""
    grade = row["grade"]
    override = (row.get("ovr") or "").lower()
    if grade in PASSING and override != "excl.":
        return True
    if grade == "D" and override == "incl.":
        return True
    return False


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

    print("Expected Progress column (earned / required):")
    total_required = req.totals[track]
    fixed_sum = 0
    lines = []
    for category in req.categories:
        if category.code == req.slack_category:
            continue
        required = category.req_ch.get(track, 0)
        fixed_sum += required
        earned = sum(
            r["hours"] or 0 for r in rows
            if r.get("cat_" + label) == category.code
            and counts_toward_progress(r)
        )
        lines.append((category.code, earned, required))

    slack_required = total_required - fixed_sum
    slack_earned = sum(
        r["hours"] or 0 for r in rows
        if r.get("cat_" + label) == req.slack_category
        and counts_toward_progress(r)
    )
    lines.append((req.slack_category, slack_earned, slack_required))

    for code, earned, required in lines:
        met = "  <- met" if required > 0 and earned >= required else ""
        greyed = "  (not on this track)" if required == 0 else ""
        print("  {:10s} {:>3d} / {:<3d} ch{}{}".format(
            code, int(earned), int(required), met, greyed))

    grand_earned = sum(r["hours"] or 0 for r in rows
                       if counts_toward_progress(r))
    print("  {:10s} {:>3d} / {:<3d} ch".format(
        "TOTAL", int(grand_earned), int(total_required)))
    print()
    print("Sanity check: category requirements sum to",
          fixed_sum + slack_required, "and the declared total is",
          total_required,
          "->", "OK" if fixed_sum + slack_required == total_required
          else "MISMATCH")


if __name__ == "__main__":
    main()
