# advising-scripts

Turns a student's transcript PDF into a degree-progress tracker spreadsheet,
and keeps it current year after year. Built for UNB math degrees; the degree
requirements are **data**, so other programmes are new data files, not new
code.

This file is both the user documentation and the working guidance: how to run
it, the invariants, the traps, and how to verify a change.

---

## Running it

```bash
cd ~/Documents/3work/WIP-DUGS-related/advising

uv run ~/3work/advising-scripts/transcript_to_tracker.py \
    student-records/NAME/transcript.pdf \
    -r degree-requirements/bsc-math-2020.xlsx \
    -o student-records/NAME/tracker.xlsx -t Honours
```

Next year, re-run with `-u` pointing at the existing tracker. Grades and new
courses come from the transcript; the `Cat.`, `Ovr.`, `Notes` and GPA flags
already entered are preserved.

```bash
uv run ~/3work/advising-scripts/transcript_to_tracker.py \
    student-records/NAME/transcript.pdf \
    -r degree-requirements/bsc-math-2020.xlsx \
    -u student-records/NAME/tracker.xlsx
```

Other things it does:

```bash
# Two calendar editions side by side in one tracker — repeat -r
... -r degree-requirements/bsc-math-2020.xlsx -r degree-requirements/bsc-math-2027.xlsx

# When a transcript won't parse, look at the coordinate grid first
uv run transcript_to_tracker.py --dump-text transcript.pdf
```

`uv` puts the script's own directory on `sys.path`, so the sibling imports
resolve inside the repo while the working directory stays wherever the data
is. Only the paths typed as *arguments* are relative to the shell.

---

## Hard rules

- **Never commit a `.pdf` or an `.xlsx`.** A transcript is a student record;
  the trackers are all generated. `.gitignore` enforces this.

  The single exception is `sample-transcript.pdf`, the anonymized fixture the
  checks run on — its student is "Last, First", number `1234567`, born
  `MM/YY`. Do not add a second exception.

- **Never commit anything *derived* from a real transcript** either — parsed
  CSVs, text dumps, a populated tracker. This history deliberately excludes an
  early `output.txt` because it contained a student number the sample's
  anonymisation had missed.

- **The requirements workbooks are not tracked here.** They live in
  `advising/degree-requirements/`, are maintained by hand in Excel, and reach
  the pipeline through `-r`. The `requirements/*.tsv` in this repo is a *test
  fixture only* — text so it is readable in a diff.

- **No shell-outs, no system dependencies, `pathlib` for paths.** Dependencies
  go in each script's PEP 723 inline block so `uv run` resolves them anywhere.

---

## Architecture

```
requirements .xlsx (hand-maintained) ─> [requirements_io] ──┐
                                                            ├─> tracker .xlsx
transcript .pdf ──────> [transcript_parser] ──> course rows ┤
                                                            │
existing tracker .xlsx ─> [tracker_reader] ──> advisor's edits ─┘
```

| Module | Responsibility |
|---|---|
| `transcript_to_tracker.py` | CLI, and the merge logic |
| `transcript_parser.py` | PDF → course rows. **Change here for a new transcript format.** |
| `requirements_io.py` | load + validate a requirements workbook |
| `tracker_writer.py` | generate the tracker; every Excel formula lives here |
| `tracker_reader.py` | read an existing tracker back as data |
| `check_support.py` | shared setup for the checks, incl. the fixture `.tsv` → workbook builder |

### Invariants — don't break these

1. **The tracker workbook is always written from scratch.** "Merging" reads the
   old one purely as a *data source*. Never read-modify-write: openpyxl
   silently drops parts it does not understand, so a round-trip would degrade
   the file a little more every year.
2. **Merge preserves the advisor's work.** Rows key on `(term, course)`.
   Grades, titles and hours come from the transcript; `Cat.`, `Ovr.`, `Notes`
   and the GPA flags never get overwritten. A row present in the tracker but
   absent from the transcript is kept and reported — usually a planned course.
3. **Requirements are the source of truth; the tracker is derived.** Formula
   ranges are computed from the data, never hardcoded. That is what makes the
   old hand-built template's bugs impossible here — its category dropdown
   leaked onto the Notes column and included the TOTAL row.
4. **Parse by coordinate, not by whitespace.** `transcript_parser` reads column
   geometry from each block's own header line. Do not reintroduce
   regex-on-layout-text; that is what the rewrite removed.

---

## Requirements files

A requirements file says which categories a programme has, how many credit
hours each needs on each track, and which courses satisfy them. It has four
sheets, and carries a **"How to use"** sheet documenting its own schema.

The three things worth knowing:

- **Track columns are the track names.** Add a column headed `Co-op`, add a
  matching `Total CH: Co-op` row on the `Meta` sheet, and that track exists.
  No code changes.
- **An exact course code beats a wildcard.** With `MATH 3033` listed under
  `3rdMath` and `MATH 3***` under `MathElec`, MATH 3033 goes to `3rdMath`.
- **Blank `Courses` cells are deliberate.** Where only a human can decide
  (Breadth, the science lecture/lab split) leave the cell empty and the
  tracker leaves the call to you — about 40% of a typical transcript.
- **`Note` columns become hover tooltips.** Both the `Requirements` and the
  `Constraints` sheet take an optional `Note` column, rendered as a cell
  comment on the description in the generated tracker. Optional and resolved
  by header name, so a workbook without one is unaffected and the column can
  sit anywhere.

  An Excel **cell comment is not a substitute**. It lives in
  `xl/comments1.xml`, nowhere near the cell's value, so the loader cannot see
  it — a note added that way silently never reaches the tracker. That happened
  once; `requirements_io` now warns when it finds one.

---

## How the tracker works

- **Rows are keyed on `(term, course)`.** A retake is naturally distinct.
- **`B3` is a track dropdown.** `Req. CH` cells resolve it via INDEX/MATCH
  against a hidden `_TrackData` sheet. The grey-and-strikethrough formatting
  keys off "this row requires 0 ch on the selected track", so it works for any
  track a requirements file defines.
- **`Ovr.` is `incl.` / `excl.`** — force a course to count toward progress,
  or force it not to. It never affects GPA.
- **No native Excel checkboxes.** They are a Microsoft 365-only feature that
  does not render in Office 2019/2021 or LibreOffice. `gpa1`/`gpa2` and the
  constraint ticks are TRUE/FALSE cells with a dropdown.

---

## The two Excel traps

Both shipped once in working-looking code. **Neither is visible in the
generated XML** — the files were well-formed and the formulas syntactically
valid. Read this before touching `tracker_writer.py`.

### 1. A `dxf` is not a cell style

The differential format a conditional-formatting rule points at has its own
rules. A solid fill must put its colour in **`bgColor`** with `fgColor` left
automatic (`indexed="64"`). openpyxl's usual
`PatternFill("solid", fgColor=<colour>)` produces a dxf Excel rejects — and
Excel responds by repairing the file and discarding **every** CF rule on the
sheet, not just the offending one.

Two smaller faults ride along:

- a 6-digit hex colour gets padded by openpyxl to `00`-alpha (transparent),
  not `FF`. Always write 8-digit ARGB.
- CF formulas do **not** accept array constants. `MATCH($B$3,{"Honours"},0)`
  has to become a chain of `<>` tests.

CF formulas are also written for the **top-left cell of the range** and
re-evaluated relatively, so references must be column-absolute/row-relative
(`$N6`, not `$N$6`) or every row tests row 6's condition.

*Technique:* diff the generated `xl/styles.xml` dxfs against those in a
workbook Excel itself wrote. That generalises to any openpyxl feature Excel
rejects.

### 2. Grade tests inside SUMPRODUCT must be array-safe

SUMPRODUCT feeds a whole column in. Two idioms silently fail to operate
elementwise:

- `IFERROR(CHOOSE(MATCH(Courses[Gr.],{...},0),4.3,...),0)` — CHOOSE rejects an
  array index, returns `#VALUE!`, and the IFERROR collapses the entire factor
  to the **scalar 0**. The GPA numerator became 0 while the denominator
  computed normally, so every GPA read `0.00`.
- `ISNUMBER(MATCH(...))` — collapses to one TRUE/FALSE for the whole column.
  This fails *quietly*: Progress counted every course in a category regardless
  of grade, failures included, without looking wrong.

Write sums of explicit `(Courses[Gr.]="A+")` comparisons instead. Verbose, but
elementwise in every Excel version. `check_formulas.py` fails if either idiom
returns.

---

## Verifying a change

```bash
for c in check_requirements check_merge check_two_editions check_formulas; do
    echo "### $c"; uv run $c.py; done
```

The checks build their own inputs from the two fixtures the repo carries —
`sample-transcript.pdf` and `requirements/bsc-math-2020.tsv` — into scratch
directories, so they run from a bare clone with nothing else present.

**What cannot be checked here:** whether Excel accepts the workbook, and
whether the formulas evaluate. The offline `formulas` package cannot parse
structured table references (`Courses[CH]`), which the tracker uses throughout
because they survive row insertion. `check_expected_values.py` recomputes the
numbers independently in Python instead.

So after any change to `tracker_writer.py`: run `check_formulas.py`, then
**ask Patrick to open the workbook and spot-check a computed cell.** That step
is load-bearing, not belt-and-braces — it is how both traps above were caught,
and structural XML inspection missed both.

`check_formulas.py` plants an `F`, a `W` and two `D` grades into the tracker
before saving it, for exactly that reason. The sample transcript is a strong
student — every grade A-range or `CR` — and against a transcript like that, a
Progress column that wrongly counts failures looks identical to a correct one.
The planted grades give the manual comparison something to disagree about.
`check_merge.py` plants a retake for the same reason.

---

## Things that look like bugs but aren't

- **A category reading `8 / 4 ch` in green.** If a student took both CS 1003
  and CS 1073, auto-categorization assigns both to `CS`. The advisor moves one
  to `FreeElec`. Working as designed.
- **~40% of `Cat.` cells left blank.** Breadth and the science lecture/lab
  split need human judgement, so the `Courses` column in the requirements file
  is deliberately blank for those categories.
- **The same course in two terms.** A retake. The merge key is
  `(term, course)`, so these are distinct rows, not duplicates.
- **`gpa1`/`gpa2` showing `N/A`.** Nothing flagged yet. `0.00` would be the
  bug — see trap 2.

---

## Style

Patrick reads Python more than he writes it. Prefer explained, annotated code
over terse code, and comment the **why**, not the what — especially where a
non-obvious workaround exists, since most of them here encode an Excel
behaviour that is not otherwise discoverable.

Every script carries a header comment with a `Usage:` line, supports `-h` via
argparse, and declares its dependencies in a PEP 723 block. Keep the `Usage:`
line accurate when flags change.

---

## Data lives elsewhere

This repo holds no real transcripts and no spreadsheets. Patrick keeps those in
`~/Documents/3work/WIP-DUGS-related/advising/` and will point you at a specific
file when a task needs one — don't go looking. Real student records are never
to be read unless he explicitly asks for a named file.
