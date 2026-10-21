# /// script
# dependencies = ["openpyxl"]
# ///
# blank_tracker.py — build an empty degree tracker from a requirements file.
# Usage:
#   uv run blank_tracker.py -r REQUIREMENTS.xlsx -o OUT.xlsx
#   uv run blank_tracker.py -r REQ-2020.xlsx -r REQ-2027.xlsx -o OUT.xlsx
#   uv run blank_tracker.py -r REQ.xlsx -t Honours -n "Last, First" --id 1234567
#   uv run blank_tracker.py --ad-hoc -o OUT.xlsx
#   uv run blank_tracker.py -h | --help
#
#   -r, --requirements    degree-requirements workbook; repeat the flag to put
#                         two or more editions side by side in one tracker
#       --ad-hoc          no requirements at all: the Courses table on its
#                         own, with a free-text Cat. column and a passing bar
#                         you set by hand. Needs -o, since there is no
#                         requirements file to name the output after.
#                         Mutually exclusive with -r
#   -o, --output          where to write the tracker (default: alongside the
#                         first requirements file, named <req>-blank.xlsx)
#   -t, --track           which track to select (default: the first one the
#                         requirements file defines)
#   -n, --name            student name to put in B1
#       --id              student number to put in B2
#       --rows N          blank course rows to leave (default 45, roughly a
#                         four-year degree with room to spare)
#       --no-checkboxes   render the gpa1/gpa2 flags as TRUE/FALSE dropdowns
#                         instead of native Excel checkboxes. The checkboxes
#                         need Excel 365; elsewhere they show as TRUE/FALSE
#                         anyway, so this is only needed to avoid them
#
# Same tracker as transcript_to_tracker.py produces — same formulas, dropdowns,
# conditional formatting and requirements blocks — just with no course rows
# filled in. Two uses:
#
#   - a colleague who would rather type a transcript in by hand than run the
#     parser gets a working tracker without touching the pipeline;
#   - the workbook can be inspected on its own, with the transcript side of
#     the pipeline out of the picture.
#
# A tracker made this way is NOT second-class: rows typed into it are read back
# by tracker_reader like any other, so it can later be fed to
# transcript_to_tracker.py's -u and merged with a real transcript.
#
# Deliberately does not import transcript_to_tracker: that would drag in
# pdfplumber, and this script has no PDF to read. The cost is the handful of
# duplicated warning lines below.

import argparse
import sys
from pathlib import Path

# The helper modules live in lib/, beside this script.
sys.path.insert(0, str(Path(__file__).resolve().parent / "lib"))

import requirements_io
import tracker_writer


def build_parser():
    parser = argparse.ArgumentParser(
        description="Build an empty degree tracker from a requirements "
                    "workbook, for filling in by hand.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="Requirements workbooks are maintained by hand in Excel, in "
               "the advising folder's degree-requirements/ directory.",
    )
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
    parser.add_argument("-t", "--track", help="track to select (e.g. Honours)")
    parser.add_argument("-n", "--name", default="", metavar="NAME",
                        help="student name to fill in")
    parser.add_argument("--id", default="", metavar="ID", dest="student_id",
                        help="student number to fill in")
    parser.add_argument("--rows", type=int, default=45, metavar="N",
                        help="blank course rows to leave (default: 45)")
    parser.add_argument("--no-checkboxes", dest="checkboxes",
                        action="store_false",
                        help="use TRUE/FALSE dropdowns for the gpa flags "
                             "instead of native Excel checkboxes")
    return parser


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.rows < 1:
        parser.error("--rows needs to be at least 1")

    # Same rule as transcript_to_tracker: an ad hoc tracker is asked for, not
    # fallen into when -r is forgotten.
    if args.ad_hoc and args.requirements:
        parser.error(
            "--ad-hoc means no degree requirements, so it cannot be combined "
            "with --requirements. Drop one."
        )
    if not args.requirements and not args.ad_hoc:
        parser.error(
            "at least one --requirements file is needed, or --ad-hoc for a "
            "tracker with no requirements at all."
        )
    if args.ad_hoc and not args.output:
        parser.error(
            "--ad-hoc needs -o: the default output name is taken from the "
            "requirements file, and there isn't one."
        )
    if args.ad_hoc and args.track:
        print("Warning: --track is ignored with --ad-hoc; a tracker with no "
              "requirements has no tracks.", file=sys.stderr)

    # --- load the requirements ----------------------------------------------

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

    # The track selector in B3 drives every requirements block at once, so
    # only tracks common to all of them can be offered.
    dropped = tracker_writer.dropped_tracks(requirement_sets)
    if dropped:
        offered = ", ".join(tracker_writer.common_tracks(requirement_sets))
        print("Warning: these editions do not define the same tracks, so the "
              "tracker offers only the ones they share (" + offered + ").",
              file=sys.stderr)
        for label, missing in dropped.items():
            print("         " + label + " also defines: " +
                  ", ".join(missing), file=sys.stderr)

    # --- write ---------------------------------------------------------------

    if args.output:
        output = Path(args.output)
    else:
        first = Path(args.requirements[0])
        output = first.with_name(first.stem + "-blank.xlsx")

    student = {"name": args.name, "id": args.student_id}

    output.parent.mkdir(parents=True, exist_ok=True)
    # No course rows at all; spare_rows is therefore the whole table. The
    # writer already gives spare rows their dropdowns, formatting and GPA
    # flags, so a blank tracker needs no special handling here.
    tracker_writer.build_tracker(
        output, requirement_sets, [], student,
        track=None if args.ad_hoc else args.track, spare_rows=args.rows,
        checkboxes=args.checkboxes,
    )

    print("Wrote " + str(output))
    if args.ad_hoc:
        print("  ad hoc: no degree requirements, no track, free-text Cat.")
        print("  " + str(args.rows) + " blank course rows")
        return 0
    print("  " + str(len(requirement_sets)) + " requirements edition"
          + ("s" if len(requirement_sets) != 1 else "")
          + ": " + ", ".join(labels))
    print("  " + str(args.rows) + " blank course rows")

    tracks = tracker_writer.common_tracks(requirement_sets)
    chosen = args.track if args.track in tracks else tracks[0]
    if args.track and args.track not in tracks:
        print("  Warning: no track named " + repr(args.track) + "; the "
              "tracker offers " + ", ".join(tracks), file=sys.stderr)
    print("  Track: " + chosen + "  (dropdown in B3: " + ", ".join(tracks) +
          ")")
    return 0


if __name__ == "__main__":
    sys.exit(main())
