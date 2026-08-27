# check_support.py — shared setup for the check_*.py scripts.
# Usage: imported by the checks; not run directly.
#
# Every check needs a transcript to parse and a requirements file to build
# against. Neither can be committed: a real transcript is a student record,
# and .xlsx is binary. So each check builds its own inputs from the two things
# the repo DOES carry — make_sample_transcript.py and a requirements .tsv —
# into a scratch directory it owns.
#
# That also sidesteps a trap: the checks locate themselves with
# Path(__file__).parent, which resolves into the REPO even when they are
# invoked through a symlink from a working folder elsewhere. Generating inputs
# rather than looking for them beside the transcript means the checks behave
# the same however they are reached.

import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
REQUIREMENTS_TSV = HERE / "requirements" / "bsc-math-2020.tsv"


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


def sample_transcript(work_dir, name="sample-transcript.pdf"):
    """Generate the synthetic transcript PDF into `work_dir`."""
    import make_sample_transcript

    path = Path(work_dir) / name
    make_sample_transcript.build(path)
    return path


def sample_requirements(work_dir, name="bsc-math-2020.xlsx"):
    """Build the BSc Math requirements workbook from the committed .tsv."""
    path = Path(work_dir) / name
    run_script("make_requirements.py", path, "--from-tsv", REQUIREMENTS_TSV,
               "--force")
    return path


def prepare(work_dir_name):
    """The usual setup: a scratch dir with a transcript and a requirements file.

    Returns (work_dir, transcript_pdf, requirements_xlsx).
    """
    work = fresh_dir(work_dir_name)
    return work, sample_transcript(work), sample_requirements(work)
