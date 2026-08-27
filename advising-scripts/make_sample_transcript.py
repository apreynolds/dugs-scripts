#!/usr/bin/env -S uv run --script
# /// script
# dependencies = ["reportlab"]
# ///
# make_sample_transcript.py — generate a fake UNB-format transcript PDF.
# Usage:
#   ./make_sample_transcript.py OUTPUT.pdf
#   ./make_sample_transcript.py -h | --help
#
# The checks need a transcript to parse, and a real one can never be committed
# — it is a student record. This draws a wholly invented transcript in the
# same layout, so a fresh clone is self-testing and there is safe demo data to
# hand a colleague.
#
# Everything here is fictional: the student, the ID, the courses and the
# grades. The LAYOUT is what matters — text is placed at the x-coordinates a
# real UNB transcript uses, because transcript_parser.py reads its column
# geometry from the page rather than from whitespace:
#
#   transfer block   COURSE@61  TITLE@131                HRS@298  DEG@520
#   term block       COURSE@77  TITLE@147   GRADE@385    HRS@428  POINTS@482
#
# The column-header lines ("COURSE TITLE HRS DEG CREDITS" and
# "GRADE HRS POINTS TRANSFERS") are drawn too, since the parser uses them to
# locate the columns.
#
# Deliberately included awkward cases, each of which has broken a parser
# before or could plausibly do so:
#   - a transfer credit with NO course code (bare "1000-LEVEL ELECTIVE")
#   - a title long enough to crowd the credit-hours column
#   - titles containing "." ":" "&" "'" and brackets
#   - a D, an F and a W
#   - the same course taken twice, failed then passed
#   - a final term with no grades at all (in progress)
#   - a page break in the middle of the record

import argparse
import sys
from pathlib import Path

from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

PAGE_WIDTH, PAGE_HEIGHT = letter

FONT = "Courier"
SIZE = 7
LINE = 11          # vertical distance between printed lines

# Column x-positions, measured off a real transcript.
TRANSFER_COLUMNS = {"course": 61, "title": 131, "hours": 298, "deg": 520}
TRANSFER_HEADER = {"COURSE": 61, "TITLE": 131, "HRS": 304, "DEG": 520,
                   "CREDITS": 541}
TERM_COLUMNS = {"course": 77, "title": 147, "grade": 385, "hours": 428,
                "points": 482}
TERM_HEADER = {"GRADE": 379, "HRS": 433, "POINTS": 476, "TRANSFERS": 530}

STUDENT_ID = "9900001"
STUDENT_NAME = "Sample, Student"

# (course code or "", title, credit hours)
TRANSFER_CREDITS = [
    ("", "1000-LEVEL ELECTIVE", "3.00"),
    ("", "2000-LEVEL SOCIOLOGY", "3.00"),
    ("CHEM*1982", "INTRO CHEMISTRY LAB I", "2.00"),
    # Long enough to run right up against the credit-hours column. The old
    # whitespace-based parser needed its separator loosened to \s+ for exactly
    # this case, which then made other lines ambiguous.
    ("ESCI*1005", "INTRO OCEANOGRAPHY & PALEOCLIMATOLOGY", "4.00"),
    ("MATH*1003", "CALCULUS I: DIFFERENTIAL CALC", "3.00"),
    # Full stops inside the title, with no space before "FOR".
    ("MATH*2513", "MULTIVARIABLE CALC.FOR ENG.", "4.00"),
]

# (term, [(course, title, grade, hours), ...])
# A grade of "" means the course is in progress and carries no grade yet.
TERMS = [
    ("2024/FA", [
        ("MATH*1013", "CALCULUS II: INTEGRAL CALC", "A", "3.00"),
        ("CS*1073", "INTR COMP PROG I (IN JAVA)", "B+", "4.00"),
        ("PHYS*1051", "INTRODUCTORY PHYSICS I", "A-", "3.00"),
        ("PHYS*1061", "INTRO PHYSICS LAB I", "A", "1.00"),
        ("STAT*2593", "PROBABILITY AND STATISTICS", "D", "3.00"),
    ]),
    ("2025/WI", [
        ("MATH*2003", "CALCULUS III", "A+", "3.00"),
        ("MATH*2213", "LINEAR ALGEBRA I", "B", "3.00"),
        # Same course as 2024/FA above, this time failed.
        ("STAT*2593", "PROBABILITY AND STATISTICS", "F", "3.00"),
        ("CHEM*1012", "INTRO CHEMISTRY II", "C+", "3.00"),
    ]),
    ("2025/FA", [
        ("MATH*2203", "DISCRETE MATHEMATICS", "A", "3.00"),
        ("MATH*3033", "GROUP THEORY", "B-", "3.00"),
        ("STAT*3083", "PROB & MATH STAT I", "A", "3.00"),
        ("MATH*3063", "GEOMETRY", "W", "3.00"),
    ]),
    # No grades: registered but not yet assessed.
    ("2026/WI", [
        ("MATH*3103", "ANALYSIS I", "", "3.00"),
        ("MATH*3213", "LINEAR ALGEBRA II", "", "3.00"),
        ("STAT*3093", "PROB & MATH STAT II", "", "3.00"),
        ("MATH*4100", "HONOURS THESIS", "", "6.00"),
    ]),
]

# Which term starts on page 2. A page break mid-record exercises the parser's
# carrying of section state across pages.
PAGE_BREAK_BEFORE = "2025/FA"


class Sheet:
    """A page being drawn, tracking the current vertical position."""

    def __init__(self, pdf):
        self.pdf = pdf
        self.y = PAGE_HEIGHT - 60

    def line(self, items):
        """Draw one printed line. `items` is a list of (x, text)."""
        for x, text in items:
            if text:
                self.pdf.drawString(x, self.y, str(text))
        self.y -= LINE

    def blank(self):
        self.y -= LINE

    def rule(self):
        self.line([(71, "_" * 80)])

    def banner(self, continued_from=None, continued_to=None):
        self.line([(185, "UNOFFICIAL"), (244, "TRANSCRIPT")])
        if continued_from:
            self.line([(266, "(Continued"), (325, "from"), (352, "page"),
                       (379, str(continued_from) + ")")])
        self.line([(266, STUDENT_ID), (325, STUDENT_NAME.split(", ")[0] + ","),
                   (358, STUDENT_NAME.split(", ")[1])])
        self.blank()


def draw_transfer_block(sheet):
    sheet.line([(50, "Credits"), (93, "transferred:")])
    sheet.line([(x, name) for name, x in TRANSFER_HEADER.items()])

    for course, title, hours in TRANSFER_CREDITS:
        sheet.line([
            (TRANSFER_COLUMNS["course"], course),
            (TRANSFER_COLUMNS["title"], title),
            (TRANSFER_COLUMNS["hours"], hours),
            (TRANSFER_COLUMNS["deg"], "BSC"),
        ])
    sheet.blank()


def draw_term_block(sheet, term, courses, cumulative):
    sheet.line([(50, term), (109, "BSC"), (174, "MATH"), (271, "Fredericton")])
    sheet.line([(x, name) for name, x in TERM_HEADER.items()])

    for course, title, grade, hours in courses:
        # Grade points are cosmetic — the parser ignores the POINTS column —
        # but a plausible number keeps the page looking like a transcript.
        points = ""
        if grade and grade not in ("W", "F"):
            points = "{:.2f}".format(float(hours) * 3.5)
        sheet.line([
            (TERM_COLUMNS["course"], course),
            (TERM_COLUMNS["title"], title),
            (TERM_COLUMNS["grade"], grade),
            (TERM_COLUMNS["hours"], hours),
            (TERM_COLUMNS["points"], points),
        ])

    # A summary line the parser must NOT mistake for a course. Its numbers sit
    # far to the left of the credit-hours column, which is exactly what makes
    # it rejectable.
    sheet.line([
        (50, "Program"), (93, "Credit"), (131, "Hours:"), (179, "Attempted"),
        (244, "{:.2f}".format(cumulative)), (293, "Passed"),
        (341, "{:.2f}".format(cumulative)), (390, "Cumulative"),
        (449, "GPA..."), (487, "3.1"),
    ])
    sheet.rule()


def build(path):
    pdf = canvas.Canvas(str(path), pagesize=letter)
    pdf.setFont(FONT, SIZE)

    sheet = Sheet(pdf)
    sheet.banner()
    sheet.line([(240, "Date"), (267, "of"), (283, "Birth"),
                (331, "Date"), (358, "of"), (374, "Issue")])
    sheet.line([(283, "MM/YY"), (341, "01"), (358, "Sep"), (379, "2026")])
    sheet.blank()

    draw_transfer_block(sheet)

    cumulative = sum(float(h) for _c, _t, h in TRANSFER_CREDITS)
    page = 1

    for term, courses in TERMS:
        if term == PAGE_BREAK_BEFORE:
            sheet.line([(185, "UNOFFICIAL"), (244, "TRANSCRIPT"),
                        (395, "(Continued"), (455, "on"), (471, "page"),
                        (498, "2)")])
            pdf.showPage()
            pdf.setFont(FONT, SIZE)
            page += 1
            sheet = Sheet(pdf)
            sheet.banner(continued_from=page - 1)

        cumulative += sum(float(h) for _c, _t, _g, h in courses)
        draw_term_block(sheet, term, courses, cumulative)

    sheet.blank()
    sheet.line([(50, "Awards"), (88, "Granted:")])
    sheet.line([(50, "2025-26"), (115, "Entrance"), (174, "Scholarship")])
    sheet.blank()
    sheet.line([(50, "End"), (71, "of"), (88, "record")])
    sheet.line([(185, "UNOFFICIAL"), (244, "TRANSCRIPT")])

    pdf.showPage()
    pdf.save()


def expected_row_count():
    """How many course rows a correct parse should produce."""
    return len(TRANSFER_CREDITS) + sum(len(c) for _t, c in TERMS)


def main():
    parser = argparse.ArgumentParser(
        description="Generate a fake UNB-format transcript PDF for testing.",
    )
    parser.add_argument("output", help="path of the .pdf to write")
    args = parser.parse_args()

    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    build(output)
    print("Wrote " + str(output) + " (" + str(expected_row_count()) +
          " course rows expected)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
