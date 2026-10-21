# requirements_io.py — load and validate a degree-requirements workbook.
# Usage: imported by transcript_to_tracker.py; not run directly.
#        (Requirements workbooks are maintained by hand; copy one of the
#        _TEMPLATE-*.xlsx files in degree-requirements/ to start a new one.)
#
# A degree-requirements file is the DATA behind a tracker: which categories a
# program has, how many credit hours each needs, and which courses satisfy
# them. Keeping it as a separate Excel file means:
#
#   - a colleague can maintain BA Math / Math Econ / Math Phys without
#     touching any code;
#   - a new calendar edition is a new file, not a new column bolted onto the
#     tracker sheet;
#   - the tracker itself becomes a generated artifact, so its formula ranges
#     are computed rather than hand-drawn (and therefore can't drift).
#
# The file is deliberately "dumb": plain values, no formulas. Every formula
# lives in the generated tracker.
#
# ---------------------------------------------------------------------------
# SHEET LAYOUT
# ---------------------------------------------------------------------------
#
# Sheet "Requirements" — one row per category, header row first:
#
#   Category | Description | Courses | <Track1> | <Track2> | ... | Pass | Note
#
#   Category     short unique code, e.g. "Calc"    (appears in the tracker's
#                Cat. dropdown)
#   Description  human text shown in the tracker,  e.g. "MATH 1003, MATH 1013"
#   Courses      MACHINE-READABLE list of satisfying courses, separated by
#                semicolons, e.g. "MATH 1003; MATH 1013". A "*" stands for one
#                digit, so "MATH 3***" means any 3000-level MATH course.
#                Leave blank when only an advisor can decide (Breadth, etc.).
#   <Track...>   one column per program track. The COLUMN HEADINGS ARE THE
#                TRACK NAMES — add a column to add a track. Values are the
#                required credit hours for that track. Three spellings, three
#                meanings:
#
#                  12       needs 12 ch on this track
#                  0/blank  not part of this track (greyed and struck out)
#                  —        NO TARGET: a catch-all that reports what has been
#                           banked but is never "met" or "unmet"
#
#                The dash has to be typed deliberately, and that is the point.
#                Blank cannot mean "no target" because blank already means 0,
#                so a colleague who merely forgot a cell would silently get a
#                category that can never be flagged unmet. A forgotten cell
#                fails safe; only a dash opts out.
#   Pass         OPTIONAL. The lowest grade that earns credit in this
#                category, e.g. "C" or "D". Blank falls back to the Meta
#                sheet's "Default passing grade", which is itself "C" if
#                absent — so a file with neither behaves exactly as every
#                file did before this column existed.
#   Note         optional hover-tooltip text shown in the tracker.
#   Plan         OPTIONAL. When the category's courses are suggested in the
#                program: "1st year", "2nd year", "3rd/4th year", "4th year",
#                and so on, optionally followed by a parenthesis for the
#                advisor, e.g. "2nd year (but 1st year is possible)". Only the
#                words before the parenthesis are read. The tracker ignores
#                it; timetable_grid.py uses it to put each course on the page
#                for its suggested year.
#
#   The categories are NOT required to add up to the declared total. The
#   difference is free electives, and the "—" category is where they are
#   recorded. Earlier versions computed that difference and gave it to a
#   "slack" category as a requirement, which made the arithmetic reconcile but
#   produced a row that could contradict the TOTAL row below it: the TOTAL
#   counts uncategorised courses and a category row does not, so a finished
#   student could show a red free-elective row under a green total. The
#   loader now WARNS about an unexplained gap instead of manufacturing one.
#
# The Pass column is what lets one programme differ from another on grades.
# A BSc refuses a D everywhere, so it needs neither the column nor the Meta
# row. A BA accepts a D outside the major, so it sets the default to D and
# overrides the math/stats categories back to C.
#
# Sheet "Meta" — a two-column key/value list:
#
#   Program                BSc Mathematics
#   Edition                2020
#   Default passing grade  C          (optional; "C" when absent)
#   Total CH: Major        129
#   Total CH: Honours      132
#
# Sheet "Constraints" — rules a human must attest to, since they're unreliable
# to compute (e.g. whether 9 ch of Breadth form a coherent discipline):
#
#   # | Constraint | Tracks
#   1 | Max 6 ch STAT counted toward Breadth | (blank = applies to all)
#   2 | Min 3 ch 4000-level MATH             | Honours

import re
import sys
from pathlib import Path

import openpyxl

# Column headings in the Requirements sheet that are NOT track names.
#
# Every heading NOT in here becomes a track, so a new reserved column must be
# added to this set at the same moment the column appears in a workbook.
# Otherwise "Pass" is read as a track called Pass, and the file then fails
# validation for having no "Total CH: Pass" row — loudly, but for a reason
# that reads as nonsense.
RESERVED_HEADINGS = {"category", "description", "courses", "note", "notes",
                     "pass", "plan"}

# The Plan column: the year(s) of the program a category's courses are
# suggested for, e.g. "2nd year" or "3rd/4th year (but 2nd year is possible)".
# Only the words before any parenthesis are read; the parenthesis is for the
# advisor. The tracker ignores the column; timetable_grid.py sorts its pages by it.
PLAN_YEARS = {"1st": 1, "2nd": 2, "3rd": 3, "4th": 4}
PLAN_PATTERN = re.compile(r"(1st|2nd|3rd|4th)(\s*/\s*(1st|2nd|3rd|4th))*\s+years?",
                          flags=re.IGNORECASE)

META_PROGRAM = "program"
META_EDITION = "edition"
META_DEFAULT_PASS = "default passing grade"
META_TOTAL_PREFIX = "total ch:"

# A Meta key that used to matter and no longer does. Kept here only so the
# loader can say "delete this row" rather than ignoring it in silence, which
# is exactly how a dead row survives in a hand-maintained workbook for years.
META_RETIRED = {"slack category": (
    "the slack mechanism was removed; mark the category's track cells "
    '"—" on the Requirements sheet instead'
)}

# What a track cell holds when the category has no target at all. Accepts the
# three dashes a person might actually type — em, en, and plain hyphen — since
# Excel's autocorrect turns one into another depending on settings.
NO_TARGET_LITERALS = {"—", "–", "-"}

# --- the grade ladder -------------------------------------------------------
#
# An ORDINAL rank, deliberately separate from the 4.3 GPA scale in
# tracker_writer. Two reasons to keep them apart:
#
#   * "does this earn credit?" and "what is this worth toward a GPA?" are
#     different regulations. Deriving the first from the second would let an
#     edit to the GPA scale silently move the passing bar.
#   * the tracker compares two ranks inside a formula. Integers compare
#     exactly; 2.3 >= 2.3 in floating point is a question nobody should have
#     to think about while reading an Excel formula.
#
# Anything absent from this table ranks 0: F, W, IP, blank — and CR, which is
# transfer credit and always counts, handled as its own term in the formula
# rather than by giving it a rank.
GRADE_RANK = {"A+": 9, "A": 8, "A-": 7, "B+": 6, "B": 5, "B-": 4,
              "C+": 3, "C": 2, "D": 1}

# What a file means when it says nothing. "C" reproduces the behaviour every
# tracker had before the Pass column existed.
DEFAULT_PASS_GRADE = "C"


class RequirementsError(Exception):
    """Raised when a requirements workbook is malformed.

    The message is shown directly to whoever ran the script, so it should name
    the file and say what to fix.
    """


class Category:
    """One requirement bucket, e.g. "9 ch of third-year MATH"."""

    def __init__(self, code, description, course_patterns, req_ch, note,
                 pass_grade="", plan_years=()):
        self.code = code
        self.description = description
        self.course_patterns = course_patterns  # list of str, may be empty
        # dict: track name -> int, or None for "no target on this track".
        # None and 0 are different: 0 means "not part of this track" and
        # greys the row out, None means "counts, but nothing to meet".
        self.req_ch = req_ch
        self.note = note
        # "" means "no opinion" — the file's default applies. Resolve it
        # through Requirements.pass_grade_for() rather than reading this
        # directly, or a blank cell silently becomes a rank of 0 and every
        # grade passes.
        self.pass_grade = pass_grade
        # tuple of ints, e.g. (3, 4) for "3rd/4th year"; () when the Plan
        # cell is blank or the column is absent.
        self.plan_years = plan_years

    def __repr__(self):
        return "<Category " + self.code + ">"


class Constraint:
    """A rule the advisor ticks off by hand rather than one we compute."""

    def __init__(self, number, text, tracks, note=""):
        self.number = number
        self.text = text
        self.tracks = tracks  # empty list = applies to every track
        self.note = note      # optional; becomes a hover tooltip

    def applies_to(self, track):
        return not self.tracks or track in self.tracks


class Requirements:
    """A whole degree-requirements file, loaded and validated."""

    def __init__(self, program, edition, tracks,
                 categories, totals, constraints, source_path,
                 default_pass_grade=DEFAULT_PASS_GRADE):
        self.program = program
        self.edition = edition
        self.tracks = tracks              # list of track names, in file order
        self.categories = categories      # list of Category, in file order
        self.totals = totals              # dict: track name -> total CH
        self.constraints = constraints
        self.source_path = Path(source_path)
        self.default_pass_grade = default_pass_grade

    @property
    def label(self):
        """Short name for the tracker's table heading, e.g. "2020"."""
        return str(self.edition)

    def category_codes(self):
        return [c.code for c in self.categories]

    def no_target_codes(self):
        """Categories that carry no target on ANY track.

        A category is a catch-all or it is not; a file that gave one a target
        on Major and none on Honours would be describing two different things
        under one code, and `_validate` refuses it. So "any" and "all" agree
        here, and one list serves every track.
        """
        return [c.code for c in self.categories
                if any(v is None for v in c.req_ch.values())]

    def pass_grade_for(self, code):
        """The lowest grade that earns credit in category `code`.

        Always returns a real grade letter, never "". An unknown code — which
        is what a blank Cat. cell amounts to — gets the file's default, since
        the block TOTAL counts uncategorised rows too and they still need a
        bar to clear.
        """
        for category in self.categories:
            if category.code == code and category.pass_grade:
                return category.pass_grade
        return self.default_pass_grade

    def match_category(self, course_code):
        """Which category does `course_code` satisfy? Returns a code or None.

        The rule is **explicit matches beat wildcard matches**. Without that,
        MATH 3033 would match both "3rdMath" (listed explicitly) and
        "MathElec" (via the wildcard "MATH 3***"), and an ambiguity rule would
        blank out nearly every math course.

        Returns None when nothing matches, or when two categories match at the
        same level of specificity — those are left for the advisor rather than
        guessed at.
        """
        if not course_code:
            return None

        exact_hits = []
        wildcard_hits = []

        for category in self.categories:
            for pattern in category.course_patterns:
                if "*" in pattern:
                    if _wildcard_matches(pattern, course_code):
                        wildcard_hits.append(category.code)
                        break
                elif pattern.upper() == course_code.upper():
                    exact_hits.append(category.code)
                    break

        for hits in (exact_hits, wildcard_hits):
            unique = sorted(set(hits))
            if len(unique) == 1:
                return unique[0]
            if len(unique) > 1:
                return None  # genuinely ambiguous; advisor decides
        return None


def _wildcard_matches(pattern, course_code):
    """Match a course code against a pattern where "*" means one digit.

    "MATH 3***" matches "MATH 3033" but not "MATH 2033" or "CS 3033".
    """
    regex = "^" + re.escape(pattern).replace(r"\*", r"\d") + "$"
    return re.match(regex, course_code, flags=re.IGNORECASE) is not None


def _cell_text(value):
    """Normalise a cell to a trimmed string ("" for empty)."""
    if value is None:
        return ""
    return str(value).strip()


def _cell_int(value, where, path):
    """Read a cell that must hold a whole number of credit hours."""
    text = _cell_text(value)
    if text == "":
        return 0
    try:
        number = float(text)
    except ValueError:
        raise RequirementsError(
            str(path) + ": " + where + " should be a number, got " + repr(text)
        )
    if number != int(number):
        raise RequirementsError(
            str(path) + ": " + where + " should be a whole number of credit "
            "hours, got " + repr(text)
        )
    return int(number)


def _cell_req_ch(value, where, path):
    """Read a track-column cell: a whole number, or None for "no target".

    Deliberately NOT folded into _cell_int as a flag. _cell_int also reads the
    Meta sheet's "Total CH:" rows, and a declared total must stay strictly
    numeric — keeping them separate means the dash can never leak into it.
    """
    if _cell_text(value) in NO_TARGET_LITERALS:
        return None
    return _cell_int(value, where, path)


def _cell_pass_grade(value, where, path):
    """Read a cell that must hold a passing-grade letter. "" if empty.

    Only grades that carry a rank are allowed. "F" is rejected because it
    ranks 0, which would set the bar at zero and let blanks, withdrawals and
    failures all count. "CR" and "W"/"IP" are rejected for the same reason —
    they are not points on the scale, so they cannot be a threshold on it.
    """
    text = _cell_text(value).upper()
    if text == "":
        return ""
    if text not in GRADE_RANK:
        raise RequirementsError(
            str(path) + ": " + where + " should be a passing grade (" +
            ", ".join(sorted(GRADE_RANK, key=GRADE_RANK.get)) +
            "), got " + repr(_cell_text(value))
        )
    return text


def _cell_plan(value, where, path):
    """Read a Plan cell: "3rd/4th year (but ...)" -> (3, 4). () if empty.

    A cell that does not read as one of the agreed phrases is WARNED about and
    treated as blank, not rejected: the Plan column only places courses on the
    timetable grids, and a typo there must never stop a tracker being built.
    """
    text = _cell_text(value).split("(")[0].strip()
    if text == "":
        return ()
    if not PLAN_PATTERN.fullmatch(text):
        sys.stderr.write(
            "  note: " + str(path) + ": " + where + " " + repr(_cell_text(value)) +
            " is not read; write e.g. \"2nd year\" or \"3rd/4th year\",\n"
            "        with anything else in parentheses after it.\n"
        )
        return ()
    return tuple(sorted({PLAN_YEARS[y.lower()] for y in re.findall(r"1st|2nd|3rd|4th", text, re.I)}))


def _split_patterns(text):
    """Split the Courses cell into a list of course patterns.

    Accepts semicolons or commas as separators, since colleagues will
    reasonably type either.
    """
    if not text:
        return []
    parts = re.split(r"[;,]", text)
    return [p.strip().upper() for p in parts if p.strip()]


def _require_sheet(workbook, name, path):
    if name not in workbook.sheetnames:
        raise RequirementsError(
            str(path) + ": missing required sheet " + repr(name) +
            ". Found: " + ", ".join(workbook.sheetnames)
        )
    return workbook[name]


def _load_meta(sheet, path):
    """Read the Meta sheet's key/value pairs."""
    program = ""
    edition = ""
    default_pass = DEFAULT_PASS_GRADE
    totals = {}

    for row in sheet.iter_rows(values_only=True):
        if not row:
            continue
        key = _cell_text(row[0]).lower()
        value = _cell_text(row[1]) if len(row) > 1 else ""
        if not key:
            continue

        if key == META_PROGRAM:
            program = value
        elif key == META_EDITION:
            edition = value
        elif key in META_RETIRED:
            # Say so rather than ignoring it. The if/elif chain below has no
            # else, so an unknown key is silently dropped — which is fine for
            # a colleague's own annotations, but wrong for a key that used to
            # change behaviour and no longer does.
            sys.stderr.write(
                "  note: " + str(path) + " Meta sheet still has '" +
                _cell_text(row[0]) + "', which is no longer read.\n"
                "        " + META_RETIRED[key] + ".\n"
            )
        elif key == META_DEFAULT_PASS:
            default_pass = _cell_pass_grade(
                value, "Meta '" + _cell_text(row[0]) + "'", path
            ) or DEFAULT_PASS_GRADE
        elif key.startswith(META_TOTAL_PREFIX):
            track = key[len(META_TOTAL_PREFIX):].strip()
            # Recover the track's original capitalisation from column A.
            original = _cell_text(row[0])[len(META_TOTAL_PREFIX):].strip()
            totals[original or track] = _cell_int(
                value, "Meta '" + _cell_text(row[0]) + "'", path
            )

    return program, edition, default_pass, totals


def _load_requirements_sheet(sheet, path):
    """Read the Requirements sheet into track names + Category objects."""
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        raise RequirementsError(str(path) + ": Requirements sheet is empty.")

    header = [_cell_text(cell) for cell in rows[0]]
    lowered = [h.lower() for h in header]

    for needed in ("category", "description", "courses"):
        if needed not in lowered:
            raise RequirementsError(
                str(path) + ": Requirements sheet needs a " + repr(needed) +
                " column. Found: " + ", ".join(h for h in header if h)
            )

    idx_category = lowered.index("category")
    idx_description = lowered.index("description")
    idx_courses = lowered.index("courses")
    idx_note = lowered.index("note") if "note" in lowered else None
    idx_pass = lowered.index("pass") if "pass" in lowered else None
    idx_plan = lowered.index("plan") if "plan" in lowered else None

    # Every remaining non-empty heading is a track name. This is what lets a
    # colleague add a "Co-op" track by adding one column.
    track_columns = {}
    for i, heading in enumerate(header):
        if heading and heading.lower() not in RESERVED_HEADINGS:
            track_columns[heading] = i

    if not track_columns:
        raise RequirementsError(
            str(path) + ": Requirements sheet has no track columns. Add at "
            "least one column (e.g. 'Major') holding required credit hours."
        )

    tracks = list(track_columns.keys())
    categories = []

    for row_number, row in enumerate(rows[1:], start=2):
        def cell(index):
            return _cell_text(row[index]) if index < len(row) else ""

        # Stop at the first entirely blank row. Anything below it is a
        # colleague's own notes, not data — reading on would turn a stray
        # comment into a phantom category.
        if not any(_cell_text(value) for value in row):
            break

        code = cell(idx_category)
        if not code:
            continue  # spacer row with content elsewhere; skip it

        req_ch = {}
        for track, column in track_columns.items():
            value = row[column] if column < len(row) else None
            req_ch[track] = _cell_req_ch(
                value,
                "Requirements row " + str(row_number) + " track " +
                repr(track), path,
            )

        categories.append(Category(
            code=code,
            description=cell(idx_description),
            course_patterns=_split_patterns(cell(idx_courses)),
            req_ch=req_ch,
            note=cell(idx_note) if idx_note is not None else "",
            pass_grade=_cell_pass_grade(
                cell(idx_pass),
                "Requirements row " + str(row_number) + " Pass", path,
            ) if idx_pass is not None else "",
            plan_years=_cell_plan(
                cell(idx_plan),
                "Requirements row " + str(row_number) + " Plan", path,
            ) if idx_plan is not None else (),
        ))

    if not categories:
        raise RequirementsError(
            str(path) + ": Requirements sheet has no category rows."
        )
    return tracks, categories


def _load_constraints(workbook, path, tracks):
    """Read the optional Constraints sheet."""
    if "Constraints" not in workbook.sheetnames:
        return []

    sheet = workbook["Constraints"]
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return []

    header = [_cell_text(c).lower() for c in rows[0]]
    idx_text = header.index("constraint") if "constraint" in header else 1
    idx_tracks = header.index("tracks") if "tracks" in header else None
    # Optional, exactly as on the Requirements sheet. Absent in every workbook
    # written before this column existed, so it has to stay optional.
    idx_note = header.index("note") if "note" in header else None

    constraints = []
    for row in rows[1:]:
        if not row:
            continue
        text = _cell_text(row[idx_text]) if idx_text < len(row) else ""
        if not text:
            continue

        applies = []
        if idx_tracks is not None and idx_tracks < len(row):
            raw = _cell_text(row[idx_tracks])
            applies = [t.strip() for t in re.split(r"[;,]", raw) if t.strip()]

        note = ""
        if idx_note is not None and idx_note < len(row):
            note = _cell_text(row[idx_note])

        for track in applies:
            if track not in tracks:
                raise RequirementsError(
                    str(path) + ": Constraints references unknown track " +
                    repr(track) + ". Known tracks: " + ", ".join(tracks)
                )

        constraints.append(
            Constraint(number=len(constraints) + 1, text=text, tracks=applies,
                       note=note)
        )
    return constraints


def _validate(requirements, path):
    """Cross-checks that catch a colleague's edit before it reaches Excel."""
    codes = requirements.category_codes()

    duplicates = sorted({c for c in codes if codes.count(c) > 1})
    if duplicates:
        raise RequirementsError(
            str(path) + ": duplicate category codes: " + ", ".join(duplicates)
        )

    # A category is a catch-all or it is not. Half-and-half would mean the
    # same code describes a target on one track and a free-for-all on
    # another, which no reader could hold in their head — and it would make
    # the tracker's row change kind as B3 changes, since the Req. CH cell
    # would resolve to a number for one track and a dash for another.
    for category in requirements.categories:
        values = list(category.req_ch.values())
        if any(v is None for v in values) and not all(v is None
                                                      for v in values):
            marked = sorted(t for t, v in category.req_ch.items()
                            if v is None)
            raise RequirementsError(
                str(path) + ": category " + repr(category.code) + " is "
                'marked "—" (no target) on ' + ", ".join(marked) +
                " but carries a number on the other tracks. Mark every "
                "track, or none of them."
            )

    for track in requirements.tracks:
        if track not in requirements.totals:
            raise RequirementsError(
                str(path) + ": Meta sheet is missing 'Total CH: " + track +
                "'. Every track column needs a total."
            )
        # Only the categories that actually make a claim. A "—" category has
        # no number to add, which is the whole point of it.
        fixed = sum(c.req_ch.get(track) or 0 for c in requirements.categories
                    if c.req_ch.get(track) is not None)
        if fixed > requirements.totals[track]:
            raise RequirementsError(
                str(path) + ": track " + repr(track) + " requires " +
                str(fixed) + " ch across its categories, which exceeds the "
                "declared total of " + str(requirements.totals[track]) + " ch."
            )

        # An unexplained shortfall is a WARNING, not an error: a programme is
        # entitled to leave hours unaccounted for. But it is nearly always a
        # missing catch-all, so say which numbers disagree. This replaces the
        # old slack mechanism's guarantee that the categories reconciled —
        # that guarantee was true by construction and so never caught
        # anything, whereas this fires exactly when a file is incomplete.
        gap = requirements.totals[track] - fixed
        if gap > 0 and not requirements.no_target_codes():
            sys.stderr.write(
                "  note: " + str(path) + " track " + repr(track) + " has "
                "categories totalling " + str(fixed) + " ch against a "
                "declared " + str(requirements.totals[track]) + " ch, and no "
                "category is marked \"—\".\n"
                "        " + str(gap) + " ch are unaccounted for. Add a "
                "catch-all category (e.g. FreeElec) with \"—\" in its track "
                "cells.\n"
            )


def _warn_about_cell_comments(workbook):
    """Point out Excel cell comments, which this loader does not read.

    An Excel "Note" (the yellow sticky, stored in xl/comments1.xml) lives
    nowhere near the cell's value, so a note added that way is invisible here
    and simply never reaches the tracker — silently, which is how it cost an
    afternoon once. The Requirements and Constraints sheets both take a real
    Note column; say so rather than dropping the text on the floor.
    """
    for name in ("Meta", "Requirements", "Constraints"):
        if name not in workbook.sheetnames:
            continue
        for row in workbook[name].iter_rows():
            for cell in row:
                if cell.comment is None:
                    continue
                sys.stderr.write(
                    "  note: " + name + "!" + cell.coordinate + " carries an "
                    "Excel cell comment, which is not read.\n"
                    "        Put the text in the sheet's Note column instead: "
                    + repr(cell.comment.text.strip()) + "\n"
                )


def load_requirements(path):
    """Load and validate a requirements workbook. Returns a Requirements."""
    path = Path(path)
    if not path.exists():
        raise RequirementsError("No such requirements file: " + str(path))

    workbook = openpyxl.load_workbook(str(path), data_only=True)
    _warn_about_cell_comments(workbook)

    meta_sheet = _require_sheet(workbook, "Meta", path)
    req_sheet = _require_sheet(workbook, "Requirements", path)

    program, edition, default_pass, totals = _load_meta(meta_sheet, path)
    tracks, categories = _load_requirements_sheet(req_sheet, path)
    constraints = _load_constraints(workbook, path, tracks)

    requirements = Requirements(
        program=program,
        edition=edition or path.stem,
        tracks=tracks,
        categories=categories,
        totals=totals,
        constraints=constraints,
        source_path=path,
        default_pass_grade=default_pass,
    )
    _validate(requirements, path)
    return requirements
