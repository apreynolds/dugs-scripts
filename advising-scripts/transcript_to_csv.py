#!/usr/bin/env python3
"""
transcript_to_csv.py — convert a text-based transcript PDF to CSV.

Uses pdftotext -layout (from poppler) to preserve column alignment,
then parses term blocks and course rows.

Usage:
    python3 transcript_to_csv.py input.pdf output.csv
"""

import csv
import re
import subprocess
import sys


def extract_layout_text(pdf_path):
    """Run pdftotext on the PDF and return the extracted text as a string.

    The -layout flag tells pdftotext to preserve the spatial arrangement of
    text on the page using spaces, instead of just dumping lines in reading
    order. This keeps the table columns aligned, which makes parsing much
    easier.

    The final '-' argument tells pdftotext to write to stdout (the terminal)
    rather than a file — subprocess then captures that output for us.
    """
    result = subprocess.run(
        ["pdftotext", "-layout", pdf_path, "-"],
        capture_output=True,
        text=True,
        check=True,
    )
    return result.stdout


# Season code (as it appears in the transcript) -> sortable "NS" suffix.
# N is the term number in the calendar year, S is the season letter, so an
# alphanumeric sort of "YYYY-NS" preserves chronological order.
SEASON_TO_SORTABLE = {
    "WI": "1W",  # Winter
    "SM": "2S",  # Spring/Summer
    "FA": "3F",  # Fall
}


def normalize_term(raw):
    """Convert a raw transcript term identifier to the sortable tracker form.

    The tracker stores terms in a format that sorts alphanumerically into
    chronological order:
        normalize_term('TRANSFER') -> '0-TRANSFER'   # always sorts first
        normalize_term('2024/WI')  -> '2024-1W'      # Winter 2024
        normalize_term('2024/SM')  -> '2024-2S'      # Spring/Summer 2024
        normalize_term('2024/FA')  -> '2024-3F'      # Fall 2024

    The leading '0-' on TRANSFER ensures transfer credits sort before any
    academic term. Within a year, '1W' < '2S' < '3F' lexically matches
    the Winter → Spring/Summer → Fall calendar order.

    Anything that doesn't match a known pattern is returned unchanged, so
    surprising values stay visible for the advisor to review rather than
    being silently mangled.
    """
    if raw == "TRANSFER":
        return "0-TRANSFER"
    m = re.match(r"^(\d{4})/(\w{2})$", raw)
    if m:
        yyyy, season = m.groups()
        return yyyy + "-" + SEASON_TO_SORTABLE.get(season, season)
    return raw


def parse_transcript(text):
    """Parse the layout text and return a list of course records.

    Each record is a dict with keys: term, course, title, grade, hours.
    This mirrors the columns in the output CSV.

    The transcript has two kinds of sections:
      - A "Credits transferred:" block at the top listing transfer credits
      - One block per academic term (e.g. 2024/FA, 2025/WI) listing courses

    We walk through the lines one at a time, keeping track of which section
    we're currently in, and try to match each line against the appropriate
    regex pattern.
    """
    rows = []
    lines = text.splitlines()

    term_re = re.compile(r"^(\d{4}/\w{2})\s")
    graded_re = re.compile(
        r"^\s{4,}([A-Z]+\*\d+)\s{2,}(.+?)\s{2,}([A-Z][+-]?)\s+([\d]+\.\d+)"
    )
    ungraded_re = re.compile(
        r"^\s{4,}([A-Z]+\*\d+)\s{2,}(.+?)\s{2,}([\d]+\.\d+)\s*$"
    )
    transfer_re = re.compile(
        r"^\s{2,}([A-Z]+\*\d+)?\s{0,4}([A-Z0-9][A-Z0-9 &\':.*-]+?)\s+([\d]+\.\d+)\s+\S"
    )

    current_term = ""
    in_transfer = False

    for line in lines:
        # Skip page headers / continuation markers / column-header rows.
        if "UNOFFICIAL TRANSCRIPT" in line or "Continued" in line:
            continue
        if re.match(r"^\s*COURSE\s+TITLE", line):
            continue

        # Enter the transfer-credit section.
        if "Credits transferred:" in line:
            in_transfer = True
            current_term = "TRANSFER"
            continue

        # A new academic term heading (e.g. "2024/FA ...").
        m = term_re.match(line)
        if m:
            current_term = m.group(1)
            in_transfer = False
            continue

        if in_transfer:
            m = transfer_re.match(line)
            if m:
                rows.append({
                    "term": normalize_term(current_term),
                    "course": (m.group(1) or "").strip(),
                    "title": m.group(2).strip(),
                    "grade": "CR",
                    "hours": m.group(3),
                })
            continue

        # A graded course row.
        m = graded_re.match(line)
        if m:
            rows.append({
                "term": normalize_term(current_term),
                "course": m.group(1),
                "title": m.group(2).strip(),
                "grade": m.group(3),
                "hours": m.group(4),
            })
            continue

        # An ungraded (in-progress / no-grade) course row.
        m = ungraded_re.match(line)
        if m:
            rows.append({
                "term": normalize_term(current_term),
                "course": m.group(1),
                "title": m.group(2).strip(),
                "grade": "",
                "hours": m.group(3),
            })
            continue

    return rows


def write_csv(rows, csv_path):
    """Write the list of row dicts to a CSV file.

    csv.DictWriter handles quoting automatically — if a title contains a comma
    (e.g. "APPL IN BIOLOGY, PART II") it will be wrapped in double quotes so
    the comma isn't misread as a column separator.
    """
    fieldnames = ["term", "course", "title", "grade", "hours"]
    with open(csv_path, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def main():
    if len(sys.argv) != 3:
        print("Usage: python3 transcript_to_csv.py input.pdf output.csv")
        sys.exit(1)

    pdf_path = sys.argv[1]
    csv_path = sys.argv[2]

    print("Extracting text from " + pdf_path + " ...")
    text = extract_layout_text(pdf_path)
    print("Parsing transcript ...")
    rows = parse_transcript(text)
    print("Writing " + str(len(rows)) + " rows to " + csv_path + " ...")
    write_csv(rows, csv_path)
    print("Done.")


if __name__ == "__main__":
    main()
