# transcript_parser.py — parse a text-based transcript PDF into course rows.
# Usage: imported by transcript_to_tracker.py; not run directly.
#        (Use `transcript_to_tracker.py --dump-text FILE.pdf` to eyeball the
#         coordinate grid this module works from.)
#
# Why coordinates instead of regex-on-text:
#
# The previous version of this parser ran `pdftotext -layout` and matched
# regexes against the resulting space-padded text. That works until a long
# course title crowds its neighbour — "INTRO OCEAN & PALEOCEANOGRAPHY" left
# only ONE space before the credit hours, so a `\s{3,}` separator failed and
# had to be loosened to `\s+`, which then made other lines ambiguous.
#
# Those are all symptoms of flattening a spatially-arranged table into text.
# pdfplumber gives us each word's actual x/y position on the page, so we can
# recover the real columns and never have to guess from whitespace. It is also
# pure Python, so there is no poppler/pdftotext dependency to install on
# Windows.

import re

import pdfplumber

# --- patterns ---------------------------------------------------------------

# A course code as the transcript prints it: "MATH*1013", "CS*1003".
# The asterisk is how the student information system exports codes; the
# tracker wants a space instead ("MATH 1013"), which normalize_course does.
COURSE_CODE_RE = re.compile(r"^([A-Z]{2,5})\*(\d{4})$")

# Credit hours always print with exactly two decimals ("3.00", "4.00").
# Note the anchors: "4.1" (a GPA) must NOT match, or the summary lines at the
# foot of each term block would be mistaken for course rows.
HOURS_RE = re.compile(r"^\d+\.\d{2}$")

# A letter grade, or one of the non-letter outcomes the transcript can print.
GRADE_RE = re.compile(r"^(?:[A-F][+-]?|IP|CR|W|P|NC)$")

# A term heading: "2024/FA", "2025/WI", "2026/SM".
TERM_RE = re.compile(r"^(\d{4})/([A-Z]{2})$")

# Season code as printed -> sortable "NS" suffix. N is the term's number
# within the calendar year and S is the season letter, so a plain alphanumeric
# sort of "YYYY-NS" comes out in chronological order.
SEASON_TO_SORTABLE = {
    "WI": "1W",  # Winter
    "SM": "2S",  # Spring/Summer
    "FA": "3F",  # Fall
}

# First words that mark a line as transcript furniture rather than a course.
# Checked before anything else so summary and banner lines are cheap to drop.
SKIP_FIRST_WORDS = {
    "UNOFFICIAL", "Program", "GRADE", "COURSE", "Dean's", "Awards",
    "End", "Date", "Credits", "(Continued", "FROM", "TITLE",
}


def normalize_term(raw):
    """Convert a printed term identifier to the tracker's sortable form.

        normalize_term("TRANSFER")  -> "0-TRANSFER"   (always sorts first)
        normalize_term("2024/WI")   -> "2024-1W"
        normalize_term("2024/SM")   -> "2024-2S"
        normalize_term("2024/FA")   -> "2024-3F"

    Anything unrecognised is returned unchanged, so a surprising value stays
    visible for the advisor to notice rather than being silently mangled.
    """
    if raw == "TRANSFER":
        return "0-TRANSFER"
    m = TERM_RE.match(raw)
    if m:
        year, season = m.groups()
        return year + "-" + SEASON_TO_SORTABLE.get(season, season)
    return raw


def normalize_course(raw):
    """Turn the transcript's "MATH*1013" into the tracker's "MATH 1013"."""
    m = COURSE_CODE_RE.match(raw)
    if m:
        return m.group(1) + " " + m.group(2)
    return raw


def normalize_hours(raw):
    """Credit hours as a number: "3.00" -> 3, "1.50" -> 1.5.

    Whole numbers come back as int so the spreadsheet shows "3" rather than
    "3.0"; genuine half-credits stay float.
    """
    value = float(raw)
    if value == int(value):
        return int(value)
    return value


# --- page -> lines ----------------------------------------------------------

def group_words_into_lines(words, tolerance=3.0):
    """Cluster pdfplumber words into printed lines by their vertical position.

    Words on one printed line share a 'top' coordinate to within a point or
    two (PDF text isn't perfectly aligned). Rounding 'top' into buckets of
    `tolerance` points groups them reliably.

    Returns a list of lines, each a list of words sorted left-to-right, with
    the lines themselves in top-to-bottom order.
    """
    buckets = {}
    for word in words:
        key = round(word["top"] / tolerance)
        buckets.setdefault(key, []).append(word)

    lines = []
    for key in sorted(buckets):
        lines.append(sorted(buckets[key], key=lambda w: w["x0"]))
    return lines


def text_of(line):
    """Join a line's words back into plain text (for matching furniture)."""
    return " ".join(word["text"] for word in line)


# --- column geometry --------------------------------------------------------

class ColumnLayout:
    """Where the columns sit, in page x-coordinates.

    Rather than hardcoding "hours live at x=428", we read the positions out of
    the column-header line that the transcript prints above each block. If the
    registrar shifts the layout, the headers move with it and this keeps
    working.

    Two block types have different headers:
      transfer block: "COURSE  TITLE  HRS  DEG  CREDITS"
      term block:     "GRADE  HRS  POINTS  TRANSFERS"
    """

    def __init__(self, hours_x, grade_x=None, ignore_from_x=None):
        self.hours_x = hours_x            # centre of the credit-hours column
        self.grade_x = grade_x            # centre of the grade column, if any
        self.ignore_from_x = ignore_from_x  # everything at/right of here is
        #                                     another column we don't want
        #                                     (POINTS, DEG, TRANSFERS)

    @classmethod
    def from_header_line(cls, line):
        """Build a layout from a column-header line, or return None."""
        positions = {word["text"]: word["x0"] for word in line}

        if "GRADE" in positions and "HRS" in positions:
            # Term block. POINTS sits to the right of hours; everything from
            # the midpoint between HRS and POINTS onward is ignored.
            grade_x = positions["GRADE"]
            hours_x = positions["HRS"]
            points_x = positions.get("POINTS")
            ignore_from = (hours_x + points_x) / 2 if points_x else None
            return cls(hours_x=hours_x, grade_x=grade_x,
                       ignore_from_x=ignore_from)

        if "COURSE" in positions and "HRS" in positions:
            # Transfer block. There is no grade column; DEG sits to the right.
            hours_x = positions["HRS"]
            deg_x = positions.get("DEG")
            ignore_from = (hours_x + deg_x) / 2 if deg_x else None
            return cls(hours_x=hours_x, grade_x=None,
                       ignore_from_x=ignore_from)

        return None


# Fallback geometry, used only if a block's header line is missing (e.g. a
# term block continued across a page break). Derived from the sample
# transcript; the header-driven path above is preferred whenever available.
FALLBACK_TERM_LAYOUT = ColumnLayout(hours_x=433, grade_x=379,
                                    ignore_from_x=454)
FALLBACK_TRANSFER_LAYOUT = ColumnLayout(hours_x=304, grade_x=None,
                                        ignore_from_x=412)


# --- row extraction ---------------------------------------------------------

def extract_course_row(line, layout):
    """Try to read one course out of a line. Returns a dict, or None.

    A line counts as a course row only if it carries a credit-hours token in
    the hours column. That single test rejects every piece of transcript
    furniture on its own:

      "Program Credit Hours: Attempted 51.00 Passed 51.00 ..." — its numbers
      sit far to the LEFT of the hours column, so none of them qualify.

      "2024/25 Assessment Year GPA... 4.1" — "4.1" has one decimal place, so
      it fails HOURS_RE.

      The vertical "FROM Eastern Univ" label printed down the middle of the
      transfer block arrives as single characters with no numbers at all.
    """
    hours_word = None
    best_distance = None

    for word in line:
        if not HOURS_RE.match(word["text"]):
            continue
        # Among numeric tokens, take the one nearest the hours column centre.
        # In a term row that picks 3.00@428 over the grade points 11.10@482.
        distance = abs(word["x0"] - layout.hours_x)
        if best_distance is None or distance < best_distance:
            best_distance = distance
            hours_word = word

    # No numeric token, or the nearest one is too far from the hours column
    # to plausibly BE the hours column.
    if hours_word is None or best_distance > 40:
        return None

    hours_x0 = hours_word["x0"]

    # Grade: a grade-shaped token sitting left of hours and right of the title.
    grade = ""
    if layout.grade_x is not None:
        for word in line:
            if word["x0"] >= hours_x0:
                continue
            if abs(word["x0"] - layout.grade_x) <= 25 \
                    and GRADE_RE.match(word["text"]):
                grade = word["text"]
                break

    # Everything left of the grade/hours columns is the course code and title.
    left_cutoff = hours_x0
    if layout.grade_x is not None:
        left_cutoff = min(left_cutoff, layout.grade_x - 10)
    left_words = [w for w in line if w["x0"] < left_cutoff]

    if not left_words:
        return None

    # A leading course code is optional — bare transfer credits like
    # "1000-LEVEL EDUCATION" have a title but no code.
    course = ""
    if COURSE_CODE_RE.match(left_words[0]["text"]):
        course = normalize_course(left_words[0]["text"])
        left_words = left_words[1:]

    title = " ".join(word["text"] for word in left_words).strip()
    if not title and not course:
        return None

    return {
        "course": course,
        "title": title,
        "grade": grade,
        "hours": normalize_hours(hours_word["text"]),
    }


def extract_student(lines):
    """Best-effort read of the student's ID and name from the page banner.

    The banner prints as "<id>  Last,  First". Real transcripts fill these in;
    the anonymised sample carries the literal text "Last, First", so whatever
    comes back is offered to the advisor to correct rather than trusted.
    Returns (student_id, name), either of which may be "".
    """
    for line in lines[:12]:
        words = [w["text"] for w in line]
        if not words:
            continue
        if re.match(r"^\d{6,9}$", words[0]) and len(words) >= 2:
            return words[0], " ".join(words[1:]).strip()
    return "", ""


# --- main entry point -------------------------------------------------------

def parse_transcript(pdf_path):
    """Parse a transcript PDF.

    Returns (rows, student), where rows is a list of dicts with keys
    term / course / title / grade / hours, and student is a dict with
    'id' and 'name'.
    """
    rows = []
    student_id = ""
    student_name = ""

    # Section state lives outside the page loop on purpose: a term block can
    # be split by a page break, and the rows after the break still belong to
    # the term (and the column layout) established before it.
    current_term = ""
    layout = None
    in_transfer = False

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_index, page in enumerate(pdf.pages):
            lines = group_words_into_lines(page.extract_words())

            if page_index == 0:
                student_id, student_name = extract_student(lines)

            for line in lines:
                words = [w["text"] for w in line]
                if not words:
                    continue
                first = words[0]

                # A column-header line redefines the geometry for what follows.
                header_layout = ColumnLayout.from_header_line(line)
                if header_layout is not None:
                    layout = header_layout
                    continue

                # "Credits transferred:" opens the transfer block.
                if first == "Credits" and "transferred:" in words:
                    in_transfer = True
                    current_term = "TRANSFER"
                    layout = FALLBACK_TRANSFER_LAYOUT
                    continue

                # A "YYYY/SS" heading opens an academic term block.
                if TERM_RE.match(first):
                    # Guard against the "2024/25 Assessment Year" summary,
                    # which also starts with something slash-shaped.
                    if "Assessment" not in words:
                        current_term = first
                        in_transfer = False
                        layout = FALLBACK_TERM_LAYOUT
                        continue

                if first in SKIP_FIRST_WORDS:
                    continue

                # Rule-off separators between term blocks.
                if set(first) == {"_"}:
                    continue

                if layout is None or not current_term:
                    continue

                row = extract_course_row(line, layout)
                if row is None:
                    continue

                # Transfer credits carry no GPA-bearing grade; the tracker
                # represents them as "CR", which counts toward progress and
                # total credit hours but is excluded from every GPA formula.
                if in_transfer:
                    row["grade"] = "CR"

                row["term"] = normalize_term(current_term)
                rows.append(row)

    student = {"id": student_id, "name": student_name}
    return rows, student
