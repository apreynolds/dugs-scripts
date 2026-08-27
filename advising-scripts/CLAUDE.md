# advising-scripts

Turns a student's transcript PDF into a degree-progress tracker spreadsheet,
and keeps it current year after year. Built for UNB math degrees; the degree
requirements are **data**, so other programmes are new data files, not new
code.

`NOTES.md` is the user-facing documentation — commands, requirements-file
schema, how to hand this to a colleague. This file is the working guidance:
the invariants, the traps, and how to verify a change.

---

## Hard rules

- **Never commit a `.pdf` or `.xlsx`.** A transcript is a student record; the
  spreadsheets are all generated. `.gitignore` enforces this — do not add
  exceptions. If something needs a transcript, generate one with
  `make_sample_transcript.py`.
- **Never commit anything *derived* from a real transcript** either — parsed
  CSVs, text dumps, a populated tracker. This history deliberately excludes an
  early `output.txt` because it contained a student number the sample's
  anonymisation had missed.
- **Requirements data is edited as `.xlsx` but committed as `.tsv`.** After
  changing a workbook, export before committing:
  `make_requirements.py requirements/NAME.tsv --from-xlsx NAME.xlsx --force`
- **No shell-outs, no system dependencies, `pathlib` for paths.** This is meant
  to be handed to Windows-using colleagues. Dependencies go in each script's
  PEP 723 inline block so `uv run` resolves them anywhere.

---

## Architecture

```
requirements/*.tsv ──> [make_requirements] ──> requirements .xlsx ─┐
                                                                   ├─> tracker .xlsx
transcript .pdf ──────> [transcript_parser] ──> course rows ───────┤
                                                                   │
existing tracker .xlsx ─> [tracker_reader] ──> advisor's edits ────┘
```

| Module | Responsibility |
|---|---|
| `transcript_to_tracker.py` | CLI, and the merge logic |
| `transcript_parser.py` | PDF → course rows. **Change here for a new transcript format.** |
| `requirements_io.py` | load + validate a requirements workbook |
| `tracker_writer.py` | generate the tracker; every Excel formula lives here |
| `tracker_reader.py` | read an existing tracker back as data |
| `make_requirements.py` | build/convert requirements files, incl. the TSV round-trip |
| `make_sample_transcript.py` | generate a fake transcript PDF |
| `check_support.py` | shared setup for the checks |

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

**What can be checked here:** `uv run check_*.py`. They build their own inputs
(a synthetic transcript, and a requirements workbook from the committed
`.tsv`) into scratch directories, so they run from a bare clone.

`check_parser_parity.py` is the exception — it compares against the legacy
`transcript_to_csv.py`, so it needs a **real** transcript and poppler. On the
synthetic sample the legacy parser under-reads, which is itself why the parser
was rewritten.

**What cannot be checked here:** whether Excel accepts the workbook, and
whether the formulas evaluate. The offline `formulas` package cannot parse
structured table references (`Courses[CH]`), which the tracker uses throughout
because they survive row insertion. `check_expected_values.py` recomputes the
numbers independently in Python instead.

So after any change to `tracker_writer.py`: run `check_formulas.py`, then
**ask Patrick to open the workbook and spot-check a computed cell.** That step
is load-bearing, not belt-and-braces — it is how both traps above were caught,
and structural XML inspection missed both.

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

This repo holds no transcripts and no spreadsheets. Patrick keeps those in a
separate advising folder and will point you at it when a task needs it — don't
go looking. Real student records are never to be read unless he explicitly
asks for a specific file.
