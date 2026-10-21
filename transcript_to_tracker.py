# /// script
# dependencies = ["pdfplumber", "openpyxl"]
# ///
# transcript_to_tracker.py — build a degree tracker from a transcript PDF.
# Usage:
#   uv run transcript_to_tracker.py TRANSCRIPT.pdf -r REQUIREMENTS.xlsx -o OUT.xlsx
#   uv run transcript_to_tracker.py TRANSCRIPT.pdf -r REQ.xlsx -u EXISTING.xlsx
#   uv run transcript_to_tracker.py TRANSCRIPT.pdf -r REQ-2020.xlsx -r REQ-2027.xlsx
#   uv run transcript_to_tracker.py TRANSCRIPT.pdf --ad-hoc -o OUT.xlsx
#   uv run transcript_to_tracker.py --dump-text TRANSCRIPT.pdf
#   uv run transcript_to_tracker.py -h | --help
#
#   TRANSCRIPT.pdf        text-based transcript PDF (not a scan)
#   -r, --requirements    degree-requirements workbook; repeat the flag to put
#                         two or more editions side by side in one tracker
#       --ad-hoc          no requirements at all: the Courses table on its
#                         own, with a free-text Cat. column and a passing bar
#                         you set by hand. For looking at what a student has
#                         taken before there is a programme to measure it
#                         against. Mutually exclusive with -r
#   -o, --output          where to write the tracker (default: alongside the
#                         transcript, named <transcript>-tracker.xlsx)
#   -u, --update          an existing tracker to merge into. Grades and new
#                         courses come from the transcript; your Cat., Ovr.,
#                         Notes and gpa flags are preserved. Implies writing
#                         back to that same file unless -o says otherwise.
#   -t, --track           which track to select (default: the first one the
#                         requirements file defines, or the existing
#                         tracker's)
#   --csv                 also write the parsed transcript rows to a CSV
#   --spare-rows N        blank rows to leave for manual additions (default 15)
#   --no-checkboxes       render the gpa1/gpa2 flags as TRUE/FALSE dropdowns
#                         instead of native Excel checkboxes. The checkboxes
#                         need Excel 365; elsewhere they show as TRUE/FALSE
#                         anyway, so this is only needed to avoid them
#   --dump-text           print the transcript's word grid and exit; use this
#                         when a new transcript format needs the parser
#                         adjusting
#
# Cross-platform by design: no shell-outs, no poppler. Run it with `uv run`
# and the dependencies above resolve themselves on macOS, Windows or Linux.

import argparse
import csv
import sys
from pathlib import Path

# The helper modules live in lib/, beside this script.
sys.path.insert(0, str(Path(__file__).resolve().parent / "lib"))

import requirements_io
import tracker_reader
import tracker_writer
import transcript_parser


def row_key(row):
    """Identity of a course row for merging.

    Keyed on (term, course). A retake is naturally distinct because it sits in
    a different term. Bare transfer credits have no course code ("1000-LEVEL
    HISTORY"), so those fall back to the title.
    """
    course = row.get("course", "").strip().upper()
    if not course:
        return (row.get("term", ""), "~" + row.get("title", "").strip().upper())
    return (row.get("term", ""), course)


def sort_key(row):
    """Chronological order, transfers first, then by course code.

    The term format "YYYY-NS" is built so a plain string sort is chronological:
    "0-TRANSFER" < "2024-1W" < "2024-2S" < "2024-3F" < "2025-1W".
    """
    return (row.get("term", ""), row.get("course", ""), row.get("title", ""))


def categorize(row, requirement_sets):
    """Fill in any Cat. columns this row doesn't already have a value for."""
    filled = 0
    for requirements in requirement_sets:
        key = "cat_" + requirements.label
        if row.get(key):
            continue  # never overwrite an advisor's assignment
        match = requirements.match_category(row.get("course", ""))
        if match:
            row[key] = match
            filled += 1
        else:
            row.setdefault(key, "")
    return filled


def merge(existing_rows, transcript_rows, requirement_sets):
    """Combine an existing tracker's rows with a fresh transcript.

    Returns (rows, report). The report counts what happened so the run can
    tell the advisor rather than changing their file silently.
    """
    report = {"added": 0, "updated": 0, "kept": 0, "categorized": 0,
              "grade_changes": [], "kept_rows": [], "kept_planned": []}

    by_key = {}
    order = []
    for row in existing_rows:
        key = row_key(row)
        by_key[key] = dict(row)
        order.append(key)

    seen_in_transcript = set()

    for incoming in transcript_rows:
        key = row_key(incoming)
        seen_in_transcript.add(key)

        if key not in by_key:
            row = dict(incoming)
            report["categorized"] += categorize(row, requirement_sets)
            by_key[key] = row
            order.append(key)
            report["added"] += 1
            continue

        row = by_key[key]
        changed = False

        # Title and credit hours come from the transcript; it is authoritative.
        for field in ("title", "hours"):
            if incoming.get(field) not in ("", None) \
                    and row.get(field) != incoming.get(field):
                row[field] = incoming[field]
                changed = True

        # A grade is only ever written when the transcript actually has one.
        # A course that was graded and now reads blank means the transcript
        # simply doesn't repeat the grade, not that the grade was revoked.
        new_grade = incoming.get("grade", "")
        if new_grade and row.get("grade", "") != new_grade:
            report["grade_changes"].append(
                (key[0], row.get("course") or row.get("title"),
                 row.get("grade", "") or "(blank)", new_grade))
            row["grade"] = new_grade
            changed = True

        # A row that previously had no grade may now be categorizable.
        report["categorized"] += categorize(row, requirement_sets)

        if changed:
            report["updated"] += 1

    for key in order:
        if key not in seen_in_transcript and key in by_key:
            row = by_key[key]
            # Only worth flagging if it came from a previous tracker rather
            # than this transcript.
            if existing_rows:
                report["kept"] += 1
                # Split by grade. A row with no grade yet is a PLANNED course
                # the advisor typed in — expected, and now counted in the
                # tracker's Planned column. A row that carries a grade and has
                # vanished from the transcript is a different thing entirely,
                # and worth looking at. Same rows as before, two headings.
                entry = (key[0], row.get("course") or row.get("title"))
                if row.get("grade", "") in ("", "IP"):
                    report["kept_planned"].append(entry)
                else:
                    report["kept_rows"].append(entry)

    rows = [by_key[key] for key in order]
    rows.sort(key=sort_key)
    return rows, report


def dump_text(pdf_path):
    """Print each word with its x-position, grouped into printed lines.

    This is the view the parser works from. When a transcript in a new format
    doesn't parse, run this first: it shows where the columns actually sit.
    """
    import pdfplumber

    with pdfplumber.open(str(pdf_path)) as pdf:
        for number, page in enumerate(pdf.pages, start=1):
            print("=" * 72)
            print("PAGE", number, " width", round(page.width, 1))
            print("=" * 72)
            for line in transcript_parser.group_words_into_lines(
                    page.extract_words()):
                print("  ".join(
                    word["text"] + "@" + str(round(word["x0"]))
                    for word in line))


def write_csv(rows, path):
    """Write the parsed rows out as CSV, as the original script did."""
    fields = ["term", "course", "title", "grade", "hours"]
    with open(path, "w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields,
                                extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def build_parser():
    parser = argparse.ArgumentParser(
        description="Build or update a degree tracker from a transcript PDF.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="Requirements workbooks are maintained by hand in Excel, in "
               "the advising folder's degree-requirements/ directory.",
    )
    parser.add_argument("transcript", help="text-based transcript PDF")
    parser.add_argument(
        "-r", "--requirements", action="append", default=[], metavar="XLSX",
        help="degree-requirements workbook; repeat for side-by-side editions",
    )
    parser.add_argument(
        "--ad-hoc", dest="ad_hoc", action="store_true",
        help="build a tracker with no degree requirements at all: the "
             "Courses table on its own, free-text Cat., hand-set passing bar",
    )
    parser.add_argument("-o", "--output", metavar="XLSX",
                        help="tracker file to write")
    parser.add_argument("-u", "--update", metavar="XLSX",
                        help="existing tracker to merge into")
    parser.add_argument("-t", "--track", help="track to select (e.g. Honours)")
    parser.add_argument("--csv", metavar="CSV",
                        help="also write the parsed rows to this CSV")
    parser.add_argument("--spare-rows", type=int, default=15, metavar="N",
                        help="blank rows left for manual additions "
                             "(default: 15)")
    parser.add_argument("--no-checkboxes", dest="checkboxes",
                        action="store_false",
                        help="use TRUE/FALSE dropdowns for the gpa flags "
                             "instead of native Excel checkboxes")
    parser.add_argument("--dump-text", action="store_true",
                        help="print the transcript's word grid and exit")
    return parser


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)

    transcript_path = Path(args.transcript)
    if not transcript_path.exists():
        parser.error("no such transcript: " + str(transcript_path))

    if args.dump_text:
        dump_text(transcript_path)
        return 0

    # An ad hoc tracker has to be ASKED for, never fallen into. Treating a
    # missing -r as "ad hoc" would turn a typo in a path — or a forgotten
    # flag — into a tracker that silently measures the student against
    # nothing, which is the same failure mode the requirements loader guards
    # against by making "no target" a dash you have to type.
    if args.ad_hoc and args.requirements:
        parser.error(
            "--ad-hoc means no degree requirements, so it cannot be combined "
            "with --requirements. Drop one."
        )
    if not args.requirements and not args.ad_hoc:
        parser.error(
            "at least one --requirements file is needed — the .xlsx for "
            "the student's programme and calendar year, from the advising "
            "folder's degree-requirements/ directory. Or pass --ad-hoc for a "
            "tracker with no requirements at all."
        )
    if args.ad_hoc and args.track:
        print("Warning: --track is ignored with --ad-hoc; a tracker with no "
              "requirements has no tracks.", file=sys.stderr)

    # --- load inputs --------------------------------------------------------

    try:
        requirement_sets = [
            requirements_io.load_requirements(path)
            for path in args.requirements
        ]
    except requirements_io.RequirementsError as error:
        print("Requirements file problem: " + str(error), file=sys.stderr)
        return 1

    labels = [r.label for r in requirement_sets]
    if len(set(labels)) != len(labels):
        print("Warning: two requirements files share the edition label " +
              repr(labels[0]) + "; their columns will be numbered to keep "
              "them distinct.", file=sys.stderr)

    # The track selector drives every requirements block at once, so only
    # tracks common to all of them can be offered.
    dropped = tracker_writer.dropped_tracks(requirement_sets)
    if dropped:
        offered = ", ".join(tracker_writer.common_tracks(requirement_sets))
        print("Warning: these editions do not define the same tracks, so the "
              "tracker offers only the ones they share (" + offered + ").",
              file=sys.stderr)
        for label, missing in dropped.items():
            print("         " + label + " also defines: " +
                  ", ".join(missing), file=sys.stderr)

    print("Parsing " + str(transcript_path) + " ...")
    transcript_rows, student = transcript_parser.parse_transcript(
        transcript_path)
    print("  found " + str(len(transcript_rows)) + " course rows")

    if args.csv:
        write_csv(transcript_rows, args.csv)
        print("  wrote " + args.csv)

    # --- merge with an existing tracker, if given ---------------------------

    existing_rows = []
    track = None if args.ad_hoc else args.track

    if args.update:
        try:
            existing = tracker_reader.read_tracker(args.update)
        except tracker_reader.TrackerReadError as error:
            print("Could not read " + args.update + ": " + str(error),
                  file=sys.stderr)
            return 1
        existing_rows = existing["rows"]
        print("Merging into " + args.update + " (" +
              str(len(existing_rows)) + " existing rows)")
        # The existing tracker's identity and track win: the advisor has
        # probably corrected the name the parser guessed at.
        if existing["student"].get("name"):
            student["name"] = existing["student"]["name"]
        if existing["student"].get("id"):
            student["id"] = existing["student"]["id"]
        if not track and not args.ad_hoc:
            track = existing.get("track") or None

    rows, report = merge(existing_rows, transcript_rows, requirement_sets)

    # --- write --------------------------------------------------------------

    if args.output:
        output = Path(args.output)
    elif args.update:
        output = Path(args.update)
    else:
        output = transcript_path.with_name(
            transcript_path.stem + "-tracker.xlsx")

    output.parent.mkdir(parents=True, exist_ok=True)
    tracker_writer.build_tracker(
        output, requirement_sets, rows, student,
        track=track, spare_rows=args.spare_rows,
        checkboxes=args.checkboxes,
    )

    # --- report -------------------------------------------------------------

    print()
    print("Wrote " + str(output))
    print("  " + str(len(rows)) + " course rows"
          + (" (" + str(report["added"]) + " added, "
             + str(report["updated"]) + " updated)" if args.update else ""))
    if report["categorized"]:
        print("  auto-assigned " + str(report["categorized"]) +
              " category cells")

    if args.ad_hoc:
        print("  ad hoc: no degree requirements, so nothing was "
              "auto-categorized and nothing is measured against a target")

    blank = sum(
        1 for row in rows
        for requirements in requirement_sets
        if not row.get("cat_" + requirements.label)
    )
    if blank:
        print("  " + str(blank) + " category cells left blank for you "
              "(Breadth, electives, and anything ambiguous)")

    # A category carried across from the old tracker that the workbook no
    # longer has (renamed or removed) counts toward nothing, silently. Say so.
    for requirements in requirement_sets:
        codes = set(requirements.category_codes())
        unknown = sorted({row["cat_" + requirements.label] for row in rows
                          if row.get("cat_" + requirements.label)
                          and row["cat_" + requirements.label] not in codes})
        if unknown:
            print("  WARNING: categories not in " + requirements.label + "'s "
                  "workbook, so those courses count toward nothing: " +
                  ", ".join(unknown) + ". Reassign them, or add the old name "
                  "to RENAMED_CATEGORIES in lib/tracker_reader.py.")

    if report["grade_changes"]:
        print()
        print("  Grades changed:")
        for term, course, old, new in report["grade_changes"]:
            print("    " + term + "  " + str(course) + "  " + old + " -> " + new)

    if report["kept_planned"]:
        print()
        print("  Planned (no grade yet — these are counted in the tracker's "
              "Planned column):")
        for term, course in report["kept_planned"]:
            print("    " + term + "  " + str(course))

    if report["kept_rows"]:
        print()
        print("  Kept, but no longer on the transcript (they carry a grade, "
              "so this is worth a look):")
        for term, course in report["kept_rows"]:
            print("    " + term + "  " + str(course))

    if track:
        print()
        print("  Track: " + track)
    return 0


if __name__ == "__main__":
    sys.exit(main())
