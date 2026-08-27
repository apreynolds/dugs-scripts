# /// script
# dependencies = ["openpyxl", "pdfplumber", "reportlab"]
# ///
# check_parser_parity.py — check the new coordinate parser against the old one.
# Usage: uv run check_parser_parity.py TRANSCRIPT.pdf
#
# A comparison tool rather than a pass/fail check, and the one script here
# that needs a REAL transcript plus poppler (pdftotext) for the legacy parser.
#
# Its purpose was to prove the coordinate rewrite lost nothing: on the real
# UNB sample the two parsers agree on all 43 rows. Run it against the
# synthetic sample instead and the LEGACY parser under-reads — it misses every
# transfer credit, because its "^\s{2,}" anchor is sensitive to leading
# whitespace that varies with the font. That brittleness is the reason the
# rewrite happened, so a divergence here is information, not a failure.
#
# A regression check, not part of the shipped pipeline. Runs the new pdfplumber parser and the old
# pdftotext/regex parser on the same PDF and diffs the results, so we can be
# confident the rewrite didn't lose or invent rows.

import sys
from pathlib import Path

import transcript_parser
import transcript_to_csv  # the old implementation


def old_rows(pdf_path):
    text = transcript_to_csv.extract_layout_text(str(pdf_path))
    rows = []
    for r in transcript_to_csv.parse_transcript(text):
        rows.append((
            r["term"],
            r["course"].replace("*", " "),
            r["title"],
            r["grade"],
            float(r["hours"]),
        ))
    return rows


def new_rows(pdf_path):
    parsed, _student = transcript_parser.parse_transcript(pdf_path)
    return [
        (r["term"], r["course"], r["title"], r["grade"], float(r["hours"]))
        for r in parsed
    ]


def main():
    if len(sys.argv) < 2:
        print(__doc__ or "", end="")
        print("Usage: uv run check_parser_parity.py TRANSCRIPT.pdf\n\n"
              "Point this at a real transcript. The legacy parser it compares\n"
              "against is whitespace-based and under-reads the synthetic\n"
              "sample, so comparing the two there proves nothing.",
              file=sys.stderr)
        return 2
    pdf_path = Path(sys.argv[1])
    old = old_rows(pdf_path)
    new = new_rows(pdf_path)

    print("old parser:", len(old), "rows")
    print("new parser:", len(new), "rows")
    print()

    only_old = [r for r in old if r not in new]
    only_new = [r for r in new if r not in old]

    if not only_old and not only_new:
        print("IDENTICAL — all", len(new), "rows match.")
        return

    if only_old:
        print("--- only in OLD (" + str(len(only_old)) + ") ---")
        for r in only_old:
            print("   ", r)
    if only_new:
        print("--- only in NEW (" + str(len(only_new)) + ") ---")
        for r in only_new:
            print("   ", r)


if __name__ == "__main__":
    main()
