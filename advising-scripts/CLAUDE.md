# advising-scripts

**Read `NOTES.md` first** — it covers the commands, the file layout, the
design decisions and the two Excel traps. This file only adds what Claude
needs on top of it.

## Hard rules

- **Never commit a `.pdf` or `.xlsx`.** A transcript is a student record; the
  spreadsheets are all generated. `.gitignore` enforces this — do not add
  exceptions. If a test needs a transcript, generate one with
  `make_sample_transcript.py`.
- **Never commit anything derived from a real transcript** either — parsed
  CSVs, text dumps, a populated tracker. The original `output.txt` was kept
  out of this history because it contained a student number the sample's
  anonymisation had missed.
- Requirements data is edited as `.xlsx` but **committed as `.tsv`**. After
  changing a requirements workbook, export before committing:
  `make_requirements.py requirements/NAME.tsv --from-xlsx NAME.xlsx --force`

## Before changing `tracker_writer.py`

The generated workbook is only fully testable in Excel, and two classes of bug
are invisible from the XML — a rejected `dxf` (Excel silently discards all
conditional formatting) and a non-array-safe grade test (every GPA reads
0.00). Both shipped once. See the "Two Excel traps" section of `NOTES.md`.

Run `check_formulas.py` after any formula change; it guards the second class
specifically. Then open the workbook and check a computed cell.

## Running the checks

`uv run check_*.py`. They generate their own inputs, so they work from a bare
clone. `check_parser_parity.py` is the exception — it needs a real transcript
and poppler.

## Style

Patrick reads Python more than he writes it: prefer explained, annotated code
over terse code, and comment the *why* rather than the *what*. Every script
carries a header comment with a `Usage:` line, supports `-h`, and declares its
dependencies in a PEP 723 inline block for `uv`.

Cross-platform matters — this is meant to be handed to Windows-using
colleagues. No shell-outs, no system dependencies, `pathlib` for paths.
