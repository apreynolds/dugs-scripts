# /// script
# dependencies = ["openpyxl", "pdfplumber"]
# ///
# check_requirements.py — verify the requirements loader and auto-categorization.
# Usage: uv run check_requirements.py [REQUIREMENTS.xlsx TRANSCRIPT.pdf]
#        (with no arguments, builds both from the repo's own sources)
#
# A regression check, not part of the shipped pipeline. Loads a requirements file, prints what it
# parsed, and shows which category each course from a transcript would be
# auto-assigned to.

import sys

import check_support
import requirements_io
import transcript_parser


def main():
    if len(sys.argv) > 2:
        req_path, pdf_path = sys.argv[1], sys.argv[2]
    else:
        _work, pdf_path, req_path = check_support.prepare("_req_test")

    req = requirements_io.load_requirements(req_path)

    print("program :", req.program)
    print("edition :", req.edition)
    print("tracks  :", req.tracks)
    print("totals  :", req.totals)
    print("no target:", req.no_target_codes() or "-")
    print()
    print("categories:")
    for c in req.categories:
        # None prints as the same dash the requirements file uses, so the
        # summary reads back the way the sheet was typed.
        req_ch = "/".join("—" if c.req_ch[t] is None else str(c.req_ch[t])
                          for t in req.tracks)
        patterns = "; ".join(c.course_patterns) or "-"
        print("  {:10s} {:>7s} ch   {}".format(c.code, req_ch, patterns))
    print()
    print("constraints:")
    for k in req.constraints:
        scope = ", ".join(k.tracks) if k.tracks else "all tracks"
        print("  " + str(k.number) + ". [" + scope + "] " + k.text)
        if k.note:
            # Shown indented because it is a hover tooltip in the tracker,
            # not part of the constraint's visible text.
            print("       note: " + k.note)
    print()

    rows, student = transcript_parser.parse_transcript(pdf_path)
    print("student:", student)
    print()
    print("auto-categorization of", len(rows), "transcript rows:")
    assigned = 0
    for r in rows:
        cat = req.match_category(r["course"])
        if cat:
            assigned += 1
        print("  {:12s} {:12s} {:34s} -> {}".format(
            r["term"], r["course"] or "(none)", r["title"][:34], cat or "—"))
    print()
    print("auto-assigned", assigned, "of", len(rows),
          "(" + str(round(100 * assigned / len(rows))) + "%)")


if __name__ == "__main__":
    main()
