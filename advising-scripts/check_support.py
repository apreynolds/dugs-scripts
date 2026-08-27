# check_support.py — shared setup for the check_*.py scripts.
# Usage: imported by the checks; not run directly.
#
# Every check needs two inputs: a transcript to parse, and a requirements
# workbook to build against. Both are fixtures carried by the repo:
#
#   sample-transcript.pdf        an ANONYMIZED real transcript. Its student is
#                                "Last, First", number 1234567, born "MM/YY".
#                                It is the only .pdf allowed in this history —
#                                see the exception in .gitignore.
#   requirements/*.tsv           requirements data as tab-separated text,
#                                rebuilt into a workbook below.
#
# The .tsv rather than the .xlsx because .xlsx is opaque to git: a fixture you
# cannot read in a diff is a fixture nobody will notice going stale. The real
# degree-requirements workbooks are NOT here — they are maintained by hand in
# the advising data folder and passed to the CLI with -r.
#
# This also sidesteps a trap: the checks locate themselves with
# Path(__file__).parent, which resolves into the REPO even when they are
# invoked through a symlink from a working folder elsewhere. Owning their
# inputs means the checks behave the same however they are reached.

import shutil
import subprocess
import sys
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
REQUIREMENTS_TSV = HERE / "requirements" / "bsc-math-2020.tsv"
SAMPLE_TRANSCRIPT = HERE / "sample-transcript.pdf"


def fresh_dir(name):
    """Create (or empty) a scratch directory beside the checks."""
    path = HERE / name
    if path.exists():
        shutil.rmtree(path)
    path.mkdir(parents=True)
    return path


def run_script(script_name, *args):
    """Run one of the repo's scripts with the current interpreter."""
    result = subprocess.run(
        [sys.executable, str(HERE / script_name)] + [str(a) for a in args],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        sys.stderr.write(result.stdout)
        sys.stderr.write(result.stderr)
        raise SystemExit(script_name + " failed")
    return result.stdout


# --- requirements .tsv -> .xlsx ---------------------------------------------
#
# requirements_io.load_requirements() reads a workbook and nothing else, so the
# committed .tsv has to become one before any check can use it. This is a
# TEST-ONLY, one-direction builder: content, no formatting. requirements_io
# reads exactly three sheets (Meta / Requirements / Constraints) and no styling,
# so a fixture needs nothing more. The hand-maintained workbooks in the advising
# folder carry the styling and the "How to use" sheet; those are not built here.


def _is_comment(line):
    """Is this line a comment rather than data?

    The Constraints sheet's first column is literally headed "#", so a naive
    "starts with #" test swallows that header row. A comment is a "#" followed
    by whitespace or nothing; "#<tab>Constraint<tab>Tracks" is data.

    This is not hypothetical: getting it wrong silently dropped the "Max 6 ch
    STAT counted toward Breadth" constraint from the fixture.
    """
    if not line.startswith("#"):
        return False
    return len(line) == 1 or line[1] not in "\t"


def read_tsv(tsv_path):
    """Parse a requirements .tsv into {section name: [[cell, ...], ...]}."""
    sections = {}
    current = None

    for number, raw in enumerate(
            Path(tsv_path).read_text(encoding="utf-8").splitlines(), start=1):
        line = raw.rstrip("\n")
        if not line.strip() or _is_comment(line.lstrip()):
            continue
        stripped = line.strip()
        if stripped.startswith("[") and stripped.endswith("]"):
            current = stripped[1:-1]
            sections[current] = []
            continue
        if current is None:
            raise ValueError(
                str(tsv_path) + " line " + str(number) +
                ": data before any [Section] heading"
            )
        sections[current].append(line.split("\t"))

    missing = [s for s in ("Meta", "Requirements") if s not in sections]
    if missing:
        raise ValueError(
            str(tsv_path) + ": missing required section(s) " +
            ", ".join("[" + m + "]" for m in missing)
        )
    return sections


def _without_header(rows, first_heading):
    """Drop the header row if it's there, rather than assuming rows[0] is it.

    Blindly slicing [1:] silently eats the first data row whenever the header
    goes missing for any reason — which is exactly how the "#" comment clash
    above lost a constraint.
    """
    if rows and rows[0] and rows[0][0].strip() == first_heading:
        return rows[1:]
    return rows


def _as_number(text):
    """Convert a cell to int/float where it looks numeric, else leave it."""
    try:
        value = float(text)
    except (TypeError, ValueError):
        return text
    return int(value) if value == int(value) else value


def workbook_from_tsv(tsv_path):
    """Rebuild a requirements workbook from tab-separated text."""
    sections = read_tsv(tsv_path)

    req_rows = sections["Requirements"]
    header = req_rows[0]
    body = req_rows[1:]

    # Track names come from the header, exactly as requirements_io reads them.
    reserved = {"category", "description", "courses", "note", "notes"}
    tracks = [h for h in header[3:] if h and h.lower() not in reserved]

    # Pad short rows so every row has a cell per column, and convert the
    # credit-hour columns back to numbers so Excel stores them as numbers.
    width = len(header)
    rows = []
    for cells in body:
        padded = list(cells) + [""] * (width - len(cells))
        for index, name in enumerate(header):
            if name in tracks:
                padded[index] = _as_number(padded[index]) if padded[index] \
                    else 0
        rows.append(padded[:width])

    meta = []
    for cells in _without_header(sections["Meta"], "Field"):
        key = cells[0] if cells else ""
        value = cells[1] if len(cells) > 1 else ""
        if key:
            meta.append([key, _as_number(value)])

    constraints = []
    for cells in _without_header(sections.get("Constraints", []), "#"):
        number = _as_number(cells[0]) if cells else ""
        text = cells[1] if len(cells) > 1 else ""
        applies = cells[2] if len(cells) > 2 else ""
        if text:
            constraints.append([number, text, applies])

    workbook = openpyxl.Workbook()

    requirements = workbook.active
    requirements.title = "Requirements"
    requirements.append(["Category", "Description", "Courses"] + tracks +
                        ["Note"])
    for row in rows:
        requirements.append(list(row))

    meta_sheet = workbook.create_sheet("Meta")
    meta_sheet.append(["Field", "Value"])
    for row in meta:
        meta_sheet.append(row)

    constraints_sheet = workbook.create_sheet("Constraints")
    constraints_sheet.append(["#", "Constraint", "Tracks"])
    for row in constraints:
        constraints_sheet.append(row)

    return workbook


# --- the fixtures themselves ------------------------------------------------


def sample_transcript(work_dir=None, name=None):
    """The committed anonymized transcript.

    Takes no work_dir in practice — the fixture is read-only and every check
    parses it in place. The argument is kept so callers read symmetrically
    with sample_requirements().
    """
    if not SAMPLE_TRANSCRIPT.exists():
        raise SystemExit(
            "missing fixture " + str(SAMPLE_TRANSCRIPT) + " — it is committed, "
            "so a fresh clone should have it"
        )
    return SAMPLE_TRANSCRIPT


def sample_requirements(work_dir, name="bsc-math-2020.xlsx"):
    """Build the BSc Math requirements workbook from the committed .tsv."""
    path = Path(work_dir) / name
    workbook_from_tsv(REQUIREMENTS_TSV).save(str(path))
    return path


def prepare(work_dir_name):
    """The usual setup: a scratch dir, the transcript, and a requirements file.

    Returns (work_dir, transcript_pdf, requirements_xlsx).
    """
    work = fresh_dir(work_dir_name)
    return work, sample_transcript(), sample_requirements(work)
