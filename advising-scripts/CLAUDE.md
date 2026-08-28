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

### A blank tracker, with no transcript

For a colleague who would rather type a transcript in by hand than run the
parser — and for looking at the generated workbook with the PDF side of the
pipeline out of the picture:

```bash
uv run ~/3work/advising-scripts/blank_tracker.py \
    -r degree-requirements/bsc-math-2020.xlsx \
    -o student-records/NAME/tracker.xlsx -t Honours \
    -n "Last, First" --id 1234567
```

Same workbook the full pipeline produces — same formulas, dropdowns,
conditional formatting, requirements blocks — with an empty Courses table.
`--rows N` sets how many blank rows to leave (default 45); `-r` repeats for
side-by-side editions exactly as above.

It is not a second-class tracker: rows typed into it read back through
`tracker_reader` like any others, so a hand-filled one can later be passed to
`-u` and merged with a real transcript. Only openpyxl is needed, no pdfplumber
— a colleague can run it without the PDF stack.

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

`blank_tracker.py` is the top edge of that diagram on its own: requirements in,
tracker out, no course rows. It calls the same `tracker_writer.build_tracker`
with an empty row list, so there is no second generator to keep in step.

| Module | Responsibility |
|---|---|
| `transcript_to_tracker.py` | CLI, and the merge logic |
| `blank_tracker.py` | CLI for an empty tracker: requirements in, no transcript |
| `transcript_parser.py` | PDF → course rows. **Change here for a new transcript format.** |
| `requirements_io.py` | load + validate a requirements workbook |
| `tracker_writer.py` | generate the tracker; every Excel formula lives here |
| `checkbox_xlsx.py` | post-process the saved .xlsx into native Excel checkboxes; the only place that touches raw OOXML |
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
- **A track cell has three spellings, not two.** `12` needs 12 ch; `0` or
  blank means "not part of this track" (greyed and struck out); **`—` means
  no target at all** — a catch-all that reports what has been banked and is
  never met or unmet. `FreeElec` is the case it exists for.

  Blank cannot mean "no target", because blank already means 0. A colleague
  who merely forgot a cell would otherwise get a category that can never be
  flagged unmet — so the dash has to be typed deliberately, and a forgotten
  cell fails safe. `-`, `–` and `—` are all accepted, since Excel's
  autocorrect swaps one for another.

  Mark **every** track or none: a category numbered on Major and `—` on
  Honours describes two different things under one code, and `requirements_io`
  refuses it.

  **The categories are not expected to add up to the declared total.** The
  difference is the free electives. Earlier versions computed that difference
  and handed it to a "slack" category as a requirement — the arithmetic
  reconciled, but the row could contradict the TOTAL row directly beneath it,
  because TOTAL counts uncategorised courses and a category row does not. A
  finished student could show a red `12 / 26` under a green `132 / 132`. The
  loader now *warns* about an unexplained gap instead of manufacturing one,
  so deleting the catch-all is still caught.
- **`Pass` sets the lowest grade that earns credit, per category.** Optional
  column on the `Requirements` sheet, holding a grade letter. Blank falls back
  to the `Meta` sheet's optional `Default passing grade`, which is itself `C`
  when absent — so a file with neither behaves exactly as every file did
  before the column existed, and `bsc-math-2020.xlsx` needs no edit.

  This is what lets one programme differ from another on grades. A **BSc**
  refuses a D everywhere, so it says nothing and gets `C` throughout. A **BA**
  accepts a D outside the major, so it sets `Default passing grade` to `D` and
  writes `C` on the eight MATH/STAT categories.

  Write it that way round — default `D`, override the major — rather than
  default `C` with `D` on the electives. It mirrors the regulation's own shape
  ("a D counts, except in the major"), and it fails safe: a category added
  later inherits `D`, which is right for the electives that are most of the
  degree.

  Only `A+ A A- B+ B B- C+ C D` are accepted. `F` is refused because it ranks
  0, and a bar of zero would wave through blanks, withdrawals and failures
  alike. `requirements/ba-math-2020.tsv` is the fixture that exercises all of
  this; the BSc fixture deliberately has no `Pass` column so the default path
  stays covered.

  Adding the column to a workbook **before** the code supports it breaks that
  workbook: every non-reserved heading is read as a track, so `Pass` becomes a
  track with no `Total CH:` row. Loud, but the message reads as nonsense. The
  workbook edit and `RESERVED_HEADINGS` have to land together.

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
- **`Pass` is derived, not entered.** One column of the Courses table,
  between `Gr.` and `Ovr.`, holding an `INDEX/MATCH` that resolves the row's
  `Cat.` to its category's bar, defaulting to the file's default when `Cat.`
  is blank. It is grey where `Gr.` is bold: the grade is the fact, the bar
  beside it is only a reference.

  The lookup has to be a real cell and not inlined into the Progress
  `SUMPRODUCT` — see trap 2. With two editions the **first** one wins; two
  editions normally means two calendar years of one degree, whose bars agree.
  A BA block beside a BSc block would need two verdicts on one row and is not
  a supported layout.

- **`Ovr.` is `incl.` / `excl.`** — force a course to count toward progress,
  or force it not to. It never affects GPA. `incl.` rescues any real grade
  that fell short of its bar **except F**; a blank, `W` or `IP` is not a grade
  and is never rescued, so an `incl.` left on a planned course does nothing.
  Both values are coloured — green and light red — so a tracker carrying
  several exceptions looks different at a glance from one carrying none.

  **The rescue term must stay bounded below the bar.** Without
  `rank < bar` on it, an `incl.` on a course that already passes satisfies
  both terms of the Progress sum and its credit hours are counted **twice**.
  The old D-only rule was safe from this by accident — `D` was not in the
  passing list, so the terms could not overlap. `check_formulas.py` asserts
  the bound is present, because nothing offline can evaluate the formula and
  the generated XML looks perfectly well-formed either way.

- **`gpa1`/`gpa2` are native Excel checkboxes; the constraint ticks are not.**
  This reverses an earlier decision. Checkboxes are a Microsoft 365 feature,
  but the cell underneath stays an ordinary boolean — Office 2019/2021 and
  LibreOffice just show `TRUE`/`FALSE`, which is still correct and still reads
  back through `tracker_reader`. `--no-checkboxes` on either CLI forces the
  old dropdowns.

  The constraint ticks deliberately keep theirs: those are attestations being
  signed off, and looking different from two scratch comparison flags is the
  point.

  openpyxl cannot write them, so `checkbox_xlsx.py` post-processes the saved
  archive. Read its header before touching it — a template cannot work, for
  three separate reasons, and each one is a dead end that looks live.

  **Never read-modify-write a tracker that has them.** openpyxl rebuilds the
  archive from its object model and silently drops the parts. That is already
  invariant 1, but this is a second, quieter reason for it;
  `check_checkboxes.py` pins the behaviour so it is discovered here rather
  than in a stripped workbook.

- **Nothing wraps.** No cell in the tracker or the requirements blocks sets
  `wrap_text`. A constraint too long to fit is one that wants shortening, with
  the rest moved into its `Note` column, which becomes a hover tooltip.
  Growing the row hides the problem instead of showing it.

- **Grey means one thing: "not on this track".** The Courses table has no
  banded row stripes, because the stripe colour and the struck-out `OFF_FILL`
  grey were near enough identical that the second stopped registering.
  Anything new that wants a neutral fill should pick a different one.

- **A category with no target gets its own colour.** A `—` category shows
  `26 ch` — a count, no denominator — on the pale blue `NO_TARGET_FILL`, and
  can go neither green nor grey. Leaving it *unfilled* was the obvious cheap
  option and is wrong: in a column of green rows, the absence of green reads
  as "unmet", which is the exact false alarm the row exists to stop raising.
  So the state is stated positively, in a third colour.

  **Which row is which is a question about the cell's TYPE.** The `Req. CH`
  cell holds text for a no-target row and a number for every other, so the
  rules test `ISNUMBER` / `ISTEXT` rather than counting rows. That is what
  lets the ranges stay contiguous and lets a requirements file put its
  catch-all anywhere, or have several. See trap 4 for why the guard is not
  optional.

---

## The Excel traps

Each of these shipped once in working-looking code. **None of them is visible
in the generated XML** — the files were well-formed and the formulas
syntactically valid. Read this before touching `tracker_writer.py` or
`checkbox_xlsx.py`.

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

The same trap is why the passing bar is a **column** rather than a lookup
inside the sum. `MATCH(Courses[Cat. 2020], …)` handed a whole column collapses
to one scalar, so every row would be measured against row 6's bar — silently,
and only for students whose categories happen to differ. A per-row cell
computing its own `INDEX/MATCH` is a scalar lookup and is safe.

### 3. A cell format is shared, so clone it before changing it

openpyxl **dedupes** cell formats: every cell given the same font, fill,
border and alignment ends up pointing at one `<xf>` in `xl/styles.xml`. In the
tracker that means `gpa1`, `gpa2`, `Term`, `Course`, `CH`, `Gr.`, `Pass` and
`Ovr.` all share a single entry.

So attaching the checkbox extension to "the gpa cells' format" turns **eight
columns** into checkboxes. `checkbox_xlsx.py` appends a *clone* carrying the
extension and repoints only the target cells; `check_checkboxes.py` asserts
that no other column came along.

Two smaller faults ride along, both worth a repair prompt on their own:

- `<extLst>` must be the **last** child of `<xf>`, after `<alignment>` and
  `<protection>`;
- a part, its `[Content_Types].xml` override and its relationship have to
  arrive together, with an `rId` that is not already taken.

*Technique:* when splitting XML with a regex, assert the split accounts for
everything — `len(parts) == text.count("<xf")`, and rejoining reproduces the
input. The first version of that pattern let `[^>]*` eat the `/` of a
self-closing tag and silently merged two cell formats into one; the guard
catches it instantly, and reading the output does not.

### 4. In Excel, text is greater than every number

`"—" > 0` is **TRUE**. Excel's sort order puts every text value above every
number, so a comparison meant as "does this row need any hours?" silently
answers yes for a row whose `Req. CH` holds a dash.

That is why the `Req. CH` cell can carry either a number or `—` at all, and
why both the "met" and the "not on this track" rules are written
`AND(ISNUMBER($O6), …)` rather than testing the value alone. Without the
guard, the met rule fires on a category that has no target and cannot be met.

The nasty part is that it *looks* fixed either way. `AND($O6>0,$Q6>=$O6)` also
comes out FALSE on a dash — the second term compares a number to text and
fails — so the row renders correctly while resting entirely on Excel's type
ordering, one edit away from breaking. Every offline check passes and the XML
is well-formed.

Say what is meant: ask whether the cell is a number, then ask what the number
is. `check_no_target.py` asserts the guard is present on both rules, and
mutation-testing it (delete the guard, confirm the check fails) is the only
way to know the assertion works.

---

## Verifying a change

```bash
for c in check_requirements check_merge check_two_editions check_formulas \
         check_blank check_passing_grades check_checkboxes check_no_target; do
    echo "### $c"; uv run $c.py; done
```

The checks build their own inputs from the fixtures the repo carries —
`sample-transcript.pdf`, `requirements/bsc-math-2020.tsv` and
`requirements/ba-math-2020.tsv` — into scratch directories, so they run from a
bare clone with nothing else present.

**What cannot be checked here:** whether Excel accepts the workbook, and
whether the formulas evaluate. The offline `formulas` package cannot parse
structured table references (`Courses[CH]`), which the tracker uses throughout
because they survive row insertion. `check_expected_values.py` recomputes the
numbers independently in Python instead.

So after any change to `tracker_writer.py`: run `check_formulas.py`, then
**ask Patrick to open the workbook and spot-check a computed cell.** That step
is load-bearing, not belt-and-braces — it is how both traps above were caught,
and structural XML inspection missed both.

`check_formulas.py` plants six edge-case grades into the tracker before saving
it, for exactly that reason — `F`, `W`, a bare `D`, `D + incl.`, `F + incl.`
and `B + incl.` (the double-count case). The sample transcript is a strong
student — every grade A-range or `CR` — and against a transcript like that, a
Progress column that wrongly counts failures looks identical to a correct one.
The planted grades give the manual comparison something to disagree about.
`check_merge.py` plants a retake for the same reason.

---

## Things that look like bugs but aren't

- **A category reading `8 / 4 ch` in green.** If a student took both CS 1003
  and CS 1073, auto-categorization assigns both to `CS`. The advisor moves one
  to `FreeElec`. Working as designed.
- **`FreeElec` reading `26 ch` on blue, never going green.** It has no target
  — that is the point. The verdict on whether the hours are there is the
  TOTAL row, which counts every passing course including the uncategorised
  ones. A `FreeElec` row that *could* go green would be re-deriving TOTAL and
  would sometimes disagree with it.
- **The category rows not adding up to the TOTAL row.** Expected: they claim
  106 of 132 ch on Honours, and the 26 ch of difference are the free
  electives. If the categories ever *over*-claim, the loader refuses the file
  outright.
- **~40% of `Cat.` cells left blank.** Breadth and the science lecture/lab
  split need human judgement, so the `Courses` column in the requirements file
  is deliberately blank for those categories.
- **The same course in two terms.** A retake. The merge key is
  `(term, course)`, so these are distinct rows, not duplicates.
- **A whole `Pass` column reading `C`.** Expected on a BSc, where every
  category has the same bar. It is grey precisely so it reads as background.
- **A `D` counting toward a BA.** Also expected, outside the major — that is
  the regulation. Check the row's `Pass` cell: `D` means it counts, `C` means
  it does not.
- **`Pass` reading `C` on a BA row.** The category in `Cat.` is a MATH/STAT
  one, or `Cat.` is blank and the file's default is `C`. Recategorising the
  row changes the bar with it.
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
