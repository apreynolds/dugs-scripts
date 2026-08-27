# advising-scripts

Tooling for UNB math degree advising: turn a student's transcript PDF into a
degree-progress tracker spreadsheet, and keep it up to date year after year.

Built for BSc Mathematics first, but nothing here is specific to it — the
degree requirements are **data**, so BA Math, Math Econ, Math Phys and Math
Bio are new data files rather than new code.

---

## Quick start

You need [uv](https://docs.astral.sh/uv/). Nothing else — every script
declares its own dependencies inline and `uv` fetches them on first run.

```bash
# 1. Build a requirements workbook from the version-controlled data
uv run make_requirements.py bsc-math-2020.xlsx \
    --from-tsv requirements/bsc-math-2020.tsv

# 2. Make a tracker from a transcript
uv run transcript_to_tracker.py TRANSCRIPT.pdf \
    -r bsc-math-2020.xlsx -o student.xlsx -t Honours

# 3. Next year, re-run to merge. Grades and new courses come from the
#    transcript; the Cat./Ovr./Notes/GPA flags you entered are preserved.
uv run transcript_to_tracker.py TRANSCRIPT.pdf \
    -r bsc-math-2020.xlsx -u student.xlsx
```

No transcript to hand? Generate a safe fake one:

```bash
uv run make_sample_transcript.py sample-transcript.pdf
```

Other things you may want:

```bash
# Two calendar editions side by side in one tracker
uv run transcript_to_tracker.py T.pdf -r req-2020.xlsx -r req-2027.xlsx -o s.xlsx

# A blank requirements file for a new programme
uv run make_requirements.py ba-math-2020.xlsx

# When a transcript won't parse, look at the coordinate grid first
uv run transcript_to_tracker.py --dump-text TRANSCRIPT.pdf
```

---

## What's here

| File | Role |
|---|---|
| `transcript_to_tracker.py` | the CLI; also holds the merge logic |
| `transcript_parser.py` | PDF → course rows. **Adjust here for a new transcript format.** |
| `requirements_io.py` | loads and validates a requirements workbook |
| `tracker_writer.py` | generates the tracker; every formula lives here |
| `tracker_reader.py` | reads an existing tracker back as data, for merging |
| `make_requirements.py` | build / convert requirements files (`--bsc-math`, `--from-tsv`, `--from-xlsx`) |
| `make_sample_transcript.py` | generate a fake transcript PDF for testing |
| `requirements/*.tsv` | **the requirements data.** Version-controlled source. |
| `check_*.py` | regression checks — run them after touching anything above |
| `transcript_to_csv.py` | the original pdftotext parser, kept as a reference |
| `timetable-to-csv.sh` | unrelated side-tool: UNB timetable HTML → CSV |
| `bscmath-tracker-design-summary.md` | spec of the original hand-built tracker |

### Nothing binary is tracked

No `.xlsx` and no `.pdf`, by policy and by `.gitignore`:

- **Transcripts are student records.** They must never enter this history.
  Generate a fake one instead.
- **Spreadsheets are generated.** A requirements workbook is rebuilt from its
  `.tsv`; a tracker is rebuilt from a requirements file plus a transcript.

That is also why requirements data lives in `.tsv`. Edit the `.xlsx` in Excel,
then export before committing, so a calendar change shows up as a readable
diff rather than "binary files differ":

```bash
uv run make_requirements.py requirements/bsc-math-2020.tsv \
    --from-xlsx bsc-math-2020.xlsx --force
```

---

## Requirements files

A requirements file says which categories a programme has, how many credit
hours each needs on each track, and which courses satisfy them. It has four
sheets, and carries a **"How to use"** sheet documenting its own schema, so a
colleague can maintain one without reading any of this.

The two things worth knowing up front:

- **Track columns are the track names.** Add a column headed `Co-op`, add a
  matching `Total CH: Co-op` row on the `Meta` sheet, and that track exists.
  No code changes.
- **An exact course code beats a wildcard.** With `MATH 3033` listed under
  `3rdMath` and `MATH 3***` under `MathElec`, MATH 3033 goes to `3rdMath`.
  Leave the `Courses` cell blank where only a human can decide (Breadth, the
  science lecture/lab split) and the tracker leaves those to you — about 40%
  of a typical transcript.

---

## How the tracker works

- **The workbook is always written from scratch.** "Merging" reads the old
  tracker as a *data source* and generates a new file. Nothing is ever
  read-modify-written, because openpyxl silently drops parts it does not
  understand and a round-trip would degrade the file a little each year.
- **Rows are keyed on `(term, course)`.** A retake is naturally distinct. A
  row in the tracker that the transcript does not mention is kept and
  reported — usually a planned course.
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

## Two Excel traps, both hit in practice

Worth reading before touching `tracker_writer.py`. `git show` on the two fix
commits has the full story.

**1. A `dxf` is not a cell style.** A solid fill in a differential format must
put its colour in `bgColor` with `fgColor` left automatic (`indexed="64"`).
openpyxl's usual `PatternFill("solid", fgColor=...)` produces a dxf that Excel
rejects — and it responds by repairing the file and discarding **every**
conditional-formatting rule on the sheet, not just the bad one. Also: a
6-digit hex colour gets padded to `00`-alpha (transparent), and CF formulas do
not accept array constants.

**2. Grade tests must be array-safe.** SUMPRODUCT feeds a whole column into
them. `CHOOSE` rejects an array index, so
`IFERROR(CHOOSE(MATCH(Courses[Gr.],{...},0),4.3,...),0)` collapsed to the
scalar `0` and every GPA read `0.00`. `ISNUMBER(MATCH(...))` fails more
quietly, collapsing to one TRUE/FALSE for the whole column and counting failed
courses toward progress. Both are now sums of explicit
`(Courses[Gr.]="A+")` comparisons. `check_formulas.py` fails if either idiom
returns.

Neither is visible in the generated XML — the files were well-formed and the
formulas syntactically valid. **Spot-check computed cells in Excel after
changing the writer.** The offline `formulas` package cannot parse structured
table references (`Courses[CH]`), so it cannot evaluate these workbooks;
`check_expected_values.py` recomputes the numbers in Python instead.

---

## Running the checks

```bash
for c in check_tsv_roundtrip check_requirements check_merge \
         check_two_editions check_formulas; do
    echo "### $c"; uv run $c.py; done
```

They build their own inputs — a synthetic transcript and a requirements
workbook — into scratch directories, so a fresh clone can run them with no
other files present.

`check_parser_parity.py` is the exception: it compares the current parser
against the legacy `transcript_to_csv.py`, so it needs a **real** transcript
and `pdftotext` (poppler). On the synthetic sample the legacy parser
under-reads — its whitespace anchors miss every transfer credit — which is
itself why the parser was rewritten.

---

## Working alongside a data folder

The repo holds no transcripts or spreadsheets, so testing happens wherever
your data lives. Symlink the repo in and run from there:

```bash
cd ~/Documents/3work/WIP-DUGS-related/advising/TEMP-conversion-script
ln -s ~/3work/advising-scripts scripts

uv run scripts/transcript_to_tracker.py sample-transcript.pdf \
    -r requirements/bsc-math-2020.xlsx -o out.xlsx -t Honours
```

`uv` puts the script's own directory on `sys.path`, so the sibling imports
resolve inside the repo while your working directory stays put and relative
paths to local files keep working. Nothing to copy, nothing to go stale.

---

## Handing this to a colleague

```bash
git clone git@github.com:apreynolds/advising-scripts.git
cd advising-scripts
uv run make_sample_transcript.py sample.pdf
uv run make_requirements.py req.xlsx --from-tsv requirements/bsc-math-2020.tsv
uv run transcript_to_tracker.py sample.pdf -r req.xlsx -o tracker.xlsx
```

That works on macOS, Windows and Linux — there are no shell-outs and no
system dependencies beyond `uv` (the one exception is
`check_parser_parity.py`, which wants poppler).

To cover a different programme, they copy `requirements/bsc-math-2020.tsv`,
edit the copy — or build a starter workbook with `make_requirements.py` and
edit it in Excel — and pass it with `-r`. No Python required.
