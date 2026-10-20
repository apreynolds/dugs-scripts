# dugs-scripts

Scripts that keep the DUGS's records in step, for the Department of Mathematics and Statistics at UNB.
Started by Patrick Reynolds (DUGS from 2026) with Claude Code.

Most of what the department records about its courses can be regenerated from the registrar's data.
These scripts do that regenerating, so that the yearly upkeep is "download the new data, run the script" rather than retyping spreadsheets.
None of the data is in this repository: it lives in the Teams folders on OneDrive.

## The scripts

| Script | What it does | When |
|---|---|---|
| `build_courses.py` | Rebuilds the master list of MATH/STAT courses, `COURSES-math-stat.xlsx` (Curriculum team) | Summer, after new Informer exports |
| `build_advising.py` | Rebuilds the advising timetable, `advising-timetable.xlsx` (Advising team): every course our students take, with its usual slot | Right after `build_courses.py` |
| `timetable_grid.py` | Draws a year's timetable as week grids: one for the department, one per degree program (Timetable team) | Fall (draft), December, March, August |
| `transcript_to_tracker.py` | Turns a student's transcript into a degree-progress tracker | Whenever advising |
| `blank_tracker.py` | An empty tracker for a degree, with no transcript | Whenever needed |
| `mirror_calendar.py` | Saves a copy of the online undergraduate calendar | Summer, once the new calendar is up |
| `sweep_references.py` | Finds every mention of a MATH/STAT course in that copy: who else requires our courses | Right after `mirror_calendar.py` |
| `course_history.py` | Lists which calendar years had an entry for each MATH/STAT course | Summer, after the sweep |

`check_timetable.py` is on its way out; its useful parts are moving into `timetable_grid.py`.

Every script prints its options with `-h`, and starts with a comment block that explains it in full.

## Using them

You need [uv](https://docs.astral.sh/uv/), which installs the right Python and each script's packages by itself, on macOS or Windows.
Then, from this folder:

```
uv run build_courses.py
uv run build_advising.py
```

- **Not a git user?** Download this repository as a ZIP and unpack it anywhere; that is all.
- **A git user?** Clone it, and feel free to commit.
- **Where the data is** is set in one file, `lib/paths.py`. If your OneDrive or Teams folders are named differently, change it there.
- **Close a workbook in Excel before rebuilding it**; the scripts refuse to run while it is open.
- Columns with a **yellow header** are typed by a person and survive every rebuild; everything else is regenerated, so type only in yellow columns.
- A rebuild first copies the workbook it replaces to `NAME-bak.xlsx` beside it.
- `timetable_grid.py` turns its LaTeX into PDFs with Patrick's own `pr-latexmk`. Without it, run with `--no-pdf` and compile the `.tex` files with `pdflatex`.

## What else is here

- `docs/courses-and-timetable.md`: the course, calendar and timetable scripts in detail, with the yearly routine.
- `docs/degree-tracker.md`: the degree tracker in detail.
- `lib/`: helper code the scripts share; nothing to run.
- `checks/`: tests for the degree tracker, run after changing it (see `docs/degree-tracker.md`).
