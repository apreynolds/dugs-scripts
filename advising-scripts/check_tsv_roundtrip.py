# /// script
# dependencies = ["openpyxl"]
# ///
# check_tsv_roundtrip.py — verify requirements survive .xlsx <-> .tsv.
# Usage: uv run check_tsv_roundtrip.py [REQUIREMENTS.tsv]
#
# A regression check, not part of the shipped pipeline.
#
# The .tsv is what the repo version-controls; the .xlsx is what gets edited in
# Excel. If the conversion between them loses anything, a requirement quietly
# disappears from a student's tracker — so this rebuilds the workbook from the
# committed text, exports it back out, and compares both the parsed content
# and the text itself.
#
# It has already earned its keep: the Constraints sheet's first column is
# headed "#", which the TSV reader was treating as a comment marker. The
# header row vanished, the header-skip then ate the first real constraint, and
# "Max 6 ch STAT counted toward Breadth" was silently lost.

import sys
from pathlib import Path

import check_support
import make_requirements
import requirements_io

HERE = Path(__file__).resolve().parent


def snapshot(requirements):
    """Everything about a requirements file that the tracker depends on."""
    return {
        "program": requirements.program,
        "edition": requirements.edition,
        "slack": requirements.slack_category,
        "tracks": requirements.tracks,
        "totals": requirements.totals,
        "categories": [
            (c.code, c.description, c.course_patterns, c.req_ch, c.note)
            for c in requirements.categories
        ],
        "constraints": [
            (k.number, k.text, k.tracks) for k in requirements.constraints
        ],
    }


def main():
    source_tsv = Path(sys.argv[1]) if len(sys.argv) > 1 \
        else check_support.REQUIREMENTS_TSV

    if not source_tsv.exists():
        print("No such requirements .tsv: " + str(source_tsv), file=sys.stderr)
        return 1

    work = check_support.fresh_dir("_tsv_test")
    print("source:", source_tsv)

    # .tsv -> .xlsx -> .tsv
    first_xlsx = work / "first.xlsx"
    make_requirements.workbook_from_tsv(source_tsv).save(str(first_xlsx))

    second_tsv = work / "second.tsv"
    make_requirements.export_tsv(first_xlsx, second_tsv)

    second_xlsx = work / "second.xlsx"
    make_requirements.workbook_from_tsv(second_tsv).save(str(second_xlsx))

    failures = []

    # 1. The parsed content must be identical.
    first = snapshot(requirements_io.load_requirements(first_xlsx))
    second = snapshot(requirements_io.load_requirements(second_xlsx))
    if first != second:
        for key in first:
            if first[key] != second[key]:
                failures.append("content differs in " + key + ":\n"
                                "      first:  " + repr(first[key]) + "\n"
                                "      second: " + repr(second[key]))
    else:
        print("   OK  content identical:",
              len(first["categories"]), "categories,",
              len(first["tracks"]), "tracks,",
              len(first["constraints"]), "constraints")

    # 2. The text must be stable, so committing an export never churns the diff.
    original = source_tsv.read_text(encoding="utf-8")
    reexported = second_tsv.read_text(encoding="utf-8")
    if original != reexported:
        first_lines = original.splitlines()
        second_lines = reexported.splitlines()
        for number, (a, b) in enumerate(zip(first_lines, second_lines), 1):
            if a != b:
                failures.append("text differs at line " + str(number) + ":\n"
                                "      committed: " + repr(a) + "\n"
                                "      exported:  " + repr(b))
                break
        if len(first_lines) != len(second_lines):
            failures.append("line count differs: " + str(len(first_lines)) +
                            " committed vs " + str(len(second_lines)) +
                            " exported")
    else:
        print("   OK  re-export is byte-identical to the committed .tsv")

    # 3. Every constraint must survive — the specific bug this check was
    #    written for.
    if len(first["constraints"]) == 0:
        failures.append("no constraints survived the round-trip")

    print()
    if failures:
        print("FAILED:")
        for failure in failures:
            print("  -", failure)
        return 1
    print("PASSED — requirements round-trip losslessly through .tsv.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
