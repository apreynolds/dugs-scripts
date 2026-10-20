# Courses, calendar and timetable

Scripts behind the department's course records, kept by Patrick Reynolds (DUGS, Mathematics and Statistics, UNB) with Claude Code.
The data they read and write is not in this repository; it lives in the team folders on OneDrive, which `lib/paths.py` names.

Every script is run with `uv run --script NAME.py`, and prints its options with `-h`.
Each one starts with a comment block that explains it in more detail than this file does.

## The two workbooks

| Script | Reads | Writes |
|---|---|---|
| `build_courses.py` | The Informer exports, and the typed columns of the workbook it is about to replace | `curriculum/COURSES-math-stat.xlsx`, the master course list |
| `build_advising.py` | The master list, the degree-requirements workbooks, the timetable exports | `TEAM-ADVISING/resources/advising-timetable.xlsx` |

The routine rebuild, master list first because the advising timetable reads it:

```
uv run --script ~/3work/dugs-scripts/build_courses.py && uv run --script ~/3work/dugs-scripts/build_advising.py
```

- Columns with yellow headers are typed by a person and carried across every rebuild; everything else is regenerated.
- Both scripts refuse to run while their workbook is open in Excel.
- Each rebuild first copies the workbook it replaces to `NAME-bak.xlsx` beside it, replacing the last such copy; older versions are in OneDrive's version history.
- The master list's `Required by others` column is generated from the newest calendar sweep (see below): the units whose calendar text names the course. If the build lists a program page with no unit, add a rule for it to `PAGE_UNITS` in `build_courses.py`.
- A course on the master list's `removed` sheet gets no row in the advising timetable, even if a degree-requirements workbook still names it; the build prints which workbook that is.
- If the advising build fails with "Operation timed out" and "File is not a zip file", OneDrive was still syncing the workbook after a recent write. Wait a minute and run it again.

## The timetable grids

`timetable_grid.py YEAR` draws a year's timetable as week grids: landscape ledger PDFs, via `pr-latexmk`, with days across, each split by course level, half-hour rows, coloured by level.
It writes into the Timetable team's folder for that year, `TEAM-TIMETABLE/General/YEAR/`, which must exist:

- `departmental-grids/`: one page per term with every course of ours, and in small grey type under each the programs that require it: ours from the degree-requirements workbooks, other units' from the "Required, by course" table in `COURSES-required-by-other-units.md`.
- `program-grids-by-level/`: one PDF per degree program, a page per level (1000, 2000, 3000/4000) per term, each day split into Required and Options (named choices, every upper-level MATH/STAT course, and the first-year science courses, whose list is `FIRST_YEAR_SCIENCE` in the script).

The times come from the newest export of the year in `timetable-exports-by-year`, so the routine is: download the export, then run the script.
Four passes a year:

| When | Command | What it shows |
|---|---|---|
| Fall | `uv run timetable_grid.py 2027-28 --draft` | Next year as a draft for the committee: this year's courses at this year's times, files named `-draft` |
| December | `uv run timetable_grid.py 2026-27` | This year again, with any winter changes |
| March | `uv run timetable_grid.py 2027-28` | Next year, once the registrar has released it |
| August | `uv run timetable_grid.py 2027-28` | Next year again, nearer the start of term |

- `--proposal WORKBOOK` draws the committee's working timetable instead: our courses from the `Course` and `Suggested time` columns of its `Fall YYYY` and `Winter YYYY` sheets, files named `-proposed`. Add `--alternatives` to draw every course that has an `Alternative time` there, struck out in grey where it was and in red where it goes.
- `--only department` or `--only programs` draws one kind; `--no-pdf` writes the `.tex` only.
- Whether next year's timetable reaches Informer as soon as the registrar releases it is not known yet (October 2026).

## The academic calendar

The undergraduate calendar is only online from 2026-27, so it is copied to disk and searched there.
Everything below lives under `curriculum/academic-calendar`, one folder per calendar year.

| Script | What it does |
|---|---|
| `mirror_calendar.py YEARDIR` | Saves the online calendar's Fredericton program and course pages into `YEARDIR/html` and `YEARDIR/text`. Needs network access to `www.unb.ca`. |
| `sweep_references.py YEARDIR` | Finds every MATH, STAT or DATA course number in a mirrored year and writes three tab-separated files into `YEARDIR`. No network. |
| `course_history.py` | Reads every saved calendar and writes `course-history.tsv`: one row per MATH/STAT course, with an x under each year it has an entry. |

The three files `sweep_references.py` writes:

- `sweep-by-course.tsv`: one row per course of ours, with the other units' programs and courses that name it.
- `sweep-programs.tsv`: every mention in a program regulation, with the page, the section and the full line of calendar text.
- `sweep-courses.tsv`: every mention inside another unit's course entry, with a `Role` column (prerequisite, co-requisite, recommended, exclusion, cross-listed, unclear) read from the wording on the line.

## The yearly routine

Once the new calendar is online (it has appeared at the "current" address each summer):

1. Mirror it into a new year folder:
   `uv run --script ~/3work/dugs-scripts/mirror_calendar.py "$HOME/OneDrive - University of New Brunswick/TEAM-CURRICULUM/General/academic-calendar/2027-28"`
2. Sweep it:
   `uv run --script ~/3work/dugs-scripts/sweep_references.py "$HOME/OneDrive - University of New Brunswick/TEAM-CURRICULUM/General/academic-calendar/2027-28"`
3. Compare the three `sweep-*.tsv` files with last year's. Only the differences need reading.
   There is no script for this comparison yet.
4. Update `curriculum/COURSES-required-by-other-units.md`, the department's reference for which unit requires which of our courses, and change the calendar year in its subtitle and its date.
5. Rebuild the course history, `uv run --script ~/3work/dugs-scripts/course_history.py`, and add any course that has dropped out of the calendar to the master list's `removed` sheet, with its last calendar year and the reason.
6. Put wording that ought to be corrected into the proposed-changes file for the relevant year under `academic-calendar`, not into the reference document.

## Things to know before trusting the output

- For other units' **courses**, the role is read by the script from the wording, so it is dependable.
- For **program** pages, whether a course is "required", "one of" or on an option list is not decided by the script; a person reads the matched line in `sweep-programs.tsv`. Quote the calendar's words when passing one of these to another unit.
- The calendar's own index has had broken links, so some program pages are missing from a mirror; `MIRROR-LOG.txt` in the year folder lists the failures.
- A requirement written without a course number ("an approved course in Statistics") cannot be found by searching for numbers.
- In `course-history.tsv`, an isolated blank between two x's is a reading miss in that year's PDF, not a year without the course.
- The calendar PDFs list the Saint John courses before the Fredericton ones; the scripts use only Fredericton's.
- The calendars for 2015-16 to 2017-18, and anything before 2014-15, are not saved.
