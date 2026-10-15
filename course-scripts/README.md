# course-scripts

Scripts behind the department's course records, kept by Patrick Reynolds (DUGS, Mathematics and Statistics, UNB) with Claude Code.
The data they read and write is not in this repository; it lives in `~/Documents/3work/WIP-DUGS-related/curriculum` and `~/Documents/3work/TEAM-Advising`.

Every script is run with `uv run --script NAME.py`, and prints its options with `-h`.
Each one starts with a comment block that explains it in more detail than this file does.

## The two workbooks

| Script | Reads | Writes |
|---|---|---|
| `build_courses.py` | The Informer exports, and the typed columns of the workbook it is about to replace | `curriculum/COURSES-math-stat.xlsx`, the master course list |
| `build_advising.py` | The master list, the degree-requirements workbooks, the timetable exports | `TEAM-Advising/resources/advising-timetable.xlsx` |

The routine rebuild, master list first because the advising timetable reads it:

```
uv run --script ~/3work/course-scripts/build_courses.py && uv run --script ~/3work/course-scripts/build_advising.py
```

- Columns with yellow headers are typed by a person and carried across every rebuild; everything else is regenerated.
- Both scripts refuse to run while their workbook is open in Excel.
- Each rebuild keeps a dated copy of the workbook it replaced in a `zzzzz` folder beside it.
- The master list's `Required by others` column is generated from the newest calendar sweep (see below): the units whose calendar text names the course. If the build lists a program page with no unit, add a rule for it to `PAGE_UNITS` in `build_courses.py`.
- A course on the master list's `removed` sheet gets no row in the advising timetable, even if a degree-requirements workbook still names it; the build prints which workbook that is.
- If the advising build fails with "Operation timed out" and "File is not a zip file", OneDrive was still syncing the workbook after a recent write. Wait a minute and run it again.

## The timetable check

`check_timetable.py` checks the timetable committee's proposed timetable against the degree programs.
It reads the term sheets of `timetable-committee/proposed-timetable-2027-28.xlsx`, the degree-requirements workbooks, and the timetable exports (for other units' times), and writes three sheets into the same workbook: `Check - clashes`, `Check - alternatives` and `Check - about`.
Those three are regenerated on every run; the term sheets and `Notes` are never touched.

```
uv run --script ~/3work/course-scripts/check_timetable.py
uv run --script ~/3work/course-scripts/check_timetable.py -n --try "MATH 3213=TTh 1:00-2:20"
```

- Rerun it after changing any time in the workbook. `-n` prints the findings without writing; `--try` tests a time that is not in the workbook.
- A clash is reported only when no choice of sections avoids it, and is graded by how firmly the programs need both courses: Both required, Required + option, Both options, Elective, or Ours only.
- Other units' times are last year's, since their new ones are not known yet.
- For next year's timetable, point `-w` at the new workbook; it needs `Fall YYYY` and `Winter YYYY` sheets with `Course`, `Suggested time`, `Alternative time` and `Tutorial` columns.

`timetable_grid.py` draws one term sheet of the same workbook as a one-page week grid (landscape ledger PDF, via `pr-latexmk`): days across, each split by course level, half-hour rows, coloured by level, with the programs that require each course in small grey type under it: ours from the degree-requirements workbooks, other units' from the "Required, by course" table in `curriculum/COURSES-required-by-other-units.md`.
It writes `timetable-grid-fall-2027.tex` and `.pdf` beside the workbook; `-s "Winter 2028"` draws the other term.
With `--programs` it instead writes one PDF per degree-requirements workbook into `program-grids/`, a page per term, showing the courses that program requires: ours at their suggested times, other units' (in italic) at their most recent time in that term.

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
   `uv run --script ~/3work/course-scripts/mirror_calendar.py ~/Documents/3work/WIP-DUGS-related/curriculum/academic-calendar/2027-28`
2. Sweep it:
   `uv run --script ~/3work/course-scripts/sweep_references.py ~/Documents/3work/WIP-DUGS-related/curriculum/academic-calendar/2027-28`
3. Compare the three `sweep-*.tsv` files with last year's. Only the differences need reading.
   There is no script for this comparison yet.
4. Update `curriculum/COURSES-required-by-other-units.md`, the department's reference for which unit requires which of our courses, and change the calendar year in its subtitle and its date.
5. Rebuild the course history, `uv run --script ~/3work/course-scripts/course_history.py`, and add any course that has dropped out of the calendar to the master list's `removed` sheet, with its last calendar year and the reason.
6. Put wording that ought to be corrected into the proposed-changes file for the relevant year under `academic-calendar`, not into the reference document.

## Things to know before trusting the output

- For other units' **courses**, the role is read by the script from the wording, so it is dependable.
- For **program** pages, whether a course is "required", "one of" or on an option list is not decided by the script; a person reads the matched line in `sweep-programs.tsv`. Quote the calendar's words when passing one of these to another unit.
- The calendar's own index has had broken links, so some program pages are missing from a mirror; `MIRROR-LOG.txt` in the year folder lists the failures.
- A requirement written without a course number ("an approved course in Statistics") cannot be found by searching for numbers.
- In `course-history.tsv`, an isolated blank between two x's is a reading miss in that year's PDF, not a year without the course.
- The calendar PDFs list the Saint John courses before the Fredericton ones; the scripts use only Fredericton's.
- The calendars for 2015-16 to 2017-18, and anything before 2014-15, are not saved.
