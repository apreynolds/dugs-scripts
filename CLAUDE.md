# dugs-scripts: guidance for Claude

Read `README.md` first: it says what each script is for and who it is for.
Then read the doc for the part you are touching: `docs/courses-and-timetable.md` or `docs/degree-tracker.md` (the tracker's invariants, Excel traps and how to verify a change).

## Layout

- The scripts a person runs sit at the top level. Keep that list short: a successor as DUGS may not be a programmer, and the useful scripts should be easy to find.
- Shared helper code goes in `lib/`; each top-level script puts `lib/` on `sys.path`.
- Tests and their fixtures go in `checks/`.
- One-off analysis (for example, finding alternative times for a timetable) does not belong in this repo. Write helpers for it in the scratchpad or beside the data.

## Rules

- **Data paths come from `lib/paths.py`** and nowhere else. The data lives in the Teams folders on OneDrive.
- **Never commit a `.pdf`, `.xlsx` or `.csv`**, nor anything derived from a transcript. `checks/sample-transcript.pdf` is the one anonymized exception; do not add another.
- **Never read anything in `TEAM-ADVISING/student-records/`** unless Patrick names the file.
- Generated workbooks are rebuilt from scratch by their script, never edited in place by one. Typed (yellow-header) columns are carried across.
- After changing the tracker, run the checks (see `docs/degree-tracker.md`). After changing a course script, rebuild into a scratch file with `-o` and compare it with the real workbook before writing over the real one.
- Each script keeps its `Usage:` header and `-h` accurate, and declares its dependencies in a PEP 723 block for `uv run`.
