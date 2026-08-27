# Degree Tracker — Spreadsheet Specification

A single-sheet course log + degree-progress tracker for an undergraduate degree program. Built originally for a UNB BSc Mathematics (Honours) student, but the structure generalizes to any degree with course categories, credit hour requirements, and per-discipline GPA filters.

**Sheet name:** `Tracker` (single sheet, single workbook).

**Roles:** Read-only for the student. Maintained by an academic advisor (or whoever owns the workbook).

**Audience for this doc:** Reuse-as-a-template. If you're building a similar tracker for another program, copy the structure and substitute the program-specific values (categories, descriptions, required CH per category, total CH).

---

## Layout overview

```
   A         B            C             D         E   F   G   H   I       J          K          L   M           N             O       P
1  NAME      <Name>
2  ID        <ID>                                                          gpa1       gpa2
3  Honours   ☑            Overall GPA   <calc>                             <calc>     <calc>
4
5  Term  Course   Title  CH  Gr.  Ovr.  Cat.2020  Cat.ALT  Notes  gpa1  gpa2
6  ──── (Courses table, A6:K66) ────                                                            DEGREE REQUIREMENTS (2020)        DEGREE REQUIREMENTS (ALT)
7                                                                                               Category  Description  Req.CH  Progress
...
66 ──── (last row of Courses table) ────
```

Column blocks:
- **A:K** — Courses table (the course log)
- **L** — gap
- **M:P** — Degree Requirements table for the primary edition (e.g., 2020 calendar)
- **Q** — gap
- **R:U** — Degree Requirements table for an alternative edition (e.g., a different calendar year)
- **A1:B3** — top-level identifiers (Name / ID / Honours toggle)
- **C3:D3** — Overall GPA label + value
- **J2:K3** — GPA 1 / GPA 2 sub-calc labels (row 2) + values (row 3), side-by-side above the Courses table's gpa1 / gpa 2 checkbox columns

> **Tip:** Side-by-side editions are useful when comparing degree options (different calendar years, or major-vs-honours). The Courses table has a separate `Cat.` column for each edition so a single course can satisfy a different requirement category in each edition.

---

## 1. Top section — student identity and Honours toggle

| Cell | Field | Type | Notes |
|---|---|---|---|
| A1 | `NAME` (label) | text | Bold green Arial 10pt |
| B1 | Student name | text input | Left-aligned |
| A2 | `ID` (label) | text | Bold green |
| B2 | Student number | number input | |
| A3 | `Honours` (label) | text | Bold green |
| B3 | Honours toggle | **native checkbox** | Drives Honours-conditional logic across the sheet |

The Honours checkbox at `B3` is the single source of truth for whether the student is in the Honours program. All Honours-conditional formulas reference `$B$3` (TRUE = Honours).

---

## 2. Courses table (A5:K66)

A native Excel **Table** named `Courses`. Columns:

| # | Column header | Type | Purpose |
|---|---|---|---|
| 0 | `Term` | text (e.g. "0-TRANSFER", "2024-1W", "2024-2S", "2024-3F") | Sortable term identifier — `YYYY-NS` where `N` is the term number in the calendar year (1=Winter, 2=Spring/Summer, 3=Fall) and `S` is the season letter, so an alphanumeric sort preserves chronology. `0-TRANSFER` is used for transfer credits so they sort before any academic term. Used only for human reading; not referenced by formulas. |
| 1 | `Course` | text (e.g. "MATH 3033") | Course code |
| 2 | `Title` | text | Course title |
| 3 | `CH` | number | Credit hours |
| 4 | `Gr.` | text (data validation: A+,A,A-,B+,B,B-,C+,C,D,F,W,IP,CR, blank) | Final letter grade. Blank means in-progress or not started. `CR` = transfer credit (counts toward Progress, excluded from all GPA calculations). |
| 5 | `Ovr.` | text (data validation: yes/no/blank) | **Override** — adjusts Progress counting (see below) |
| 6 | `Cat. 2020` | text (DV from `$M$6:$M$20`) | Category code for the primary edition |
| 7 | `Cat. ALT` | text (DV from `$R$6:$R$20`) | Category code for the alternate edition |
| 8 | `Notes` | text | Free-form notes per course |
| 9 | `gpa1` | **checkbox** | Mark TRUE to include this course in GPA 1 sub-calc |
| 10 | `gpa 2` | **checkbox** | Mark TRUE to include in GPA 2 sub-calc |

**Override semantics** (column F, `Ovr.`):
- **blank** — default rule: `C` or better (and `CR`) count toward Progress; `D`, `F`, `W`, `IP` do not.
- **yes** — force-include this course in Progress (e.g., a `D` that the advisor has approved to count).
- **no** — force-exclude (e.g., a duplicate of a course taken twice; only one should count).
- Override does **not** affect GPA calculations. Only Progress.

A note is attached to `F5` (the Override header) explaining this for advisors.

---

## 3. Degree Requirements tables (M:P and R:U)

Two parallel tables — one per degree-edition the student is comparing.

### Structure
- **Header row** (row 4): "DEGREE REQUIREMENTS (<edition>)"
- **Column headers** (row 5): `Category`, `Description`, `Req. CH`, `Progress`
- **Category rows** (rows 6–20): one per requirement bucket
- **TOTAL row** (row 21): grand total
- **Constraints subsection** (rows 23–25): see §4

### Categories (rows 6–20 — 15 buckets total)

| Row | Category | Description (UNB BSc Math example) | Req. CH (Majors) | Req. CH (Honours) |
|---:|---|---|---:|---:|
| 6 | `Calc` | MATH 1003, MATH 1013 | 6 | 6 |
| 7 | `Sci1001` | Sci 1001 | 1 | 1 |
| 8 | `SciLec` | First Year Science Lectures | 18 | 18 |
| 9 | `SciLab` | First Year Science Labs | 8 | 8 |
| 10 | `2ndMath` | MATH 2003, 2013, 2203, 2213 | 12 | 12 |
| 11 | `Stat` | STAT 3083, STAT 3093 | 6 | 6 |
| 12 | `CS` | CS 1003 or 1073 | 4 | 4 |
| 13 | `3rdMath` | MATH 3033, 3213, 3243 | 9 | 9 |
| 14 | `Analys` | MATH 3003 or 3103 | 3 | 3 |
| 15 | `Breadth1` | Breadth (3000/4000-level), same discipline | 9 | 9 |
| 16 | `Breadth2` | Breadth (3000/4000-level) | 6 | 6 |
| 17 | `MathElec` | Upper Year MATH Electives | 15 | 15 |
| 18 | `AnalysH` | MATH 3113 (Honours-only) | 0 | 3 |
| 19 | `Thesis` | MATH 4100 (Honours-only) | 0 | 6 |
| 20 | `FreeElec` | Free Electives | (computed) | (computed) |
| 21 | `TOTAL` | — | 129 | 132 |

**Honours-conditional cells** (rows 18, 19): `Req. CH` is a formula `=IF($B$3, <honours value>, 0)`. When `B3` is FALSE, these zero out and conditional formatting grays the row + applies strikethrough.

**FreeElec** is computed: `=IF($B$3,132,129)-SUMIF($M$6:$M$19,"<>FreeElec",$O$6:$O$19)`. This makes the totals always reconcile; FreeElec absorbs the slack.

**TOTAL row** sums all qualifying CH and shows it as `<earned> / <total>` text.

### Progress formulas (column P / U)

Each category row's Progress cell is a SUMPRODUCT over the Courses table that computes `<earned CH> / <required CH> ch`. Pattern:

```
=TEXT(
  SUMPRODUCT(
    ((Courses[Cat. 2020]=M<row>)
     * ((ISNUMBER(MATCH(Courses[Gr.],{"A+","A","A-","B+","B","B-","C+","C","CR"},0))
         * (Courses[Ovr.]<>"no"))
        + ((Courses[Gr.]="D") * (Courses[Ovr.]="yes")))
    ) * Courses[CH]
  ),
  "0"
) & " / " & TEXT(O<row>,"0") & " ch"
```

Reads as: "for courses where the category matches **and** ((grade is C-or-better, or `CR`, and not force-excluded) **or** (grade is D and force-included)), sum the CH."

The TOTAL row is the same SUMPRODUCT without the category filter, with the denominator switching between 132/129 by `$B$3`.

### Conditional formatting (per table)

1. **Met-target green** on the Progress column (`P6:P21`, `U6:U21`): when the earned CH ≥ required CH, fill green (`#C6EFCE`), font dark green (`#006100`). Parses the `"X / Y ch"` string with FIND/MID/VALUE.

2. **Honours-only gray + strikethrough** on rows 18–19 (`M18:P19`, `R18:U19`): when `$B$3` is FALSE, fill `#E7E6E6`, font `#7F7F7F`, strikethrough on. Visually disables Honours-only rows for non-Honours students.

3. **Gridline-style borders** on the entire table block (`M5:P21` etc.): always-on light gray (`#D4D4D4`) borders that survive the CF fills layered on top — Excel hides default gridlines under any cell fill, so explicit thin borders are needed to keep the "default" gridline look.

---

## 4. Constraints subsection (rows 23–25 of each table)

Below each Degree Requirements table is a small Constraints block:

| Row | Col M (or R) | Col N (or S) | Col O (or T) | Col P (or U) |
|---|---|---|---|---|
| 23 | **CONSTRAINTS** | | | |
| 24 | 1. | Max 6 ch STAT counted toward Breadth (any combination of Breadth1/Breadth2) | | ☐ checkbox |
| 25 | 2. | Min 3 ch 4000-level MATH (Honours) | | ☐ checkbox (gray+strike when not Honours) |

The advisor toggles the checkboxes to attest that constraints are satisfied. These are not auto-computed — they encode rules that are hard or unreliable to check via formula (e.g., "discipline coherence" of the 9 ch Breadth1 requires a human judgment call).

Row 25 has the same Honours-only gray+strike CF as the AnalysH/Thesis rows.

---

## 5. GPA block (top of sheet)

Three GPA calculations, placed in the top section rather than below the Degree Requirements tables:

| Cell | Field |
|---|---|
| C3 | `Overall GPA` (label) |
| D3 | **Overall GPA** — average of all completed courses (excludes blank/W/IP) |
| J2 / J3 | `gpa1` label (row 2) and **GPA 1** value (row 3) — average over courses where the `gpa1` checkbox is TRUE |
| K2 / K3 | `gpa2` label (row 2) and **GPA 2** value (row 3) — average over courses where the `gpa 2` checkbox is TRUE |

The `gpa1` / `gpa2` value cells sit directly above the Courses table's matching checkbox columns (J / K), so each sub-GPA is visually aligned with the column of checkboxes that drives it.

Formula pattern (4.3 scale):
```
=IFERROR(
  SUMPRODUCT(
    (Courses[gpa1]=TRUE) *
    (Courses[Gr.]<>"") * (Courses[Gr.]<>"W") * (Courses[Gr.]<>"IP") * (Courses[Gr.]<>"CR") *
    IFERROR(CHOOSE(MATCH(Courses[Gr.], {"A+","A","A-","B+","B","B-","C+","C","D"}, 0),
                   4.3, 4, 3.7, 3.3, 3, 2.7, 2.3, 2, 1), 0) *
    Courses[CH]
  ) / SUMPRODUCT(
    (Courses[gpa1]=TRUE) *
    (Courses[Gr.]<>"") * (Courses[Gr.]<>"W") * (Courses[Gr.]<>"IP") * (Courses[Gr.]<>"CR") *
    Courses[CH]
  ),
  "N/A"
)
```

`CR` (transfer credit) is excluded from GPA calculations alongside blank, `W`, and `IP` — transfer credits count toward Progress and the credit-hour total but carry no GPA-bearing grade.

The **Overall GPA** formula omits the `(Courses[gpa1]=TRUE)` factor.

The `gpa1` and `gpa 2` columns in the Courses table are native Excel **checkboxes**. The student/advisor clicks to toggle inclusion in each sub-GPA. Use cases: GPA in the major (math courses only), GPA in a recent term, etc.

---

## 6. Notes (legacy yellow notes, not threaded comments)

Hover-tooltips attached to header/description cells to surface program-specific subtleties:

| Cell | Note |
|---|---|
| F5 | Override semantics (yes/no/blank rules) |
| N11, S11 | "STAT 2593 can count for 3083 with a grade of B or better" |
| N12, S12 | "CS 1003 = Python, CS 1073 = Java. Recommended to take in 1st or 2nd year." |
| N15, S15 | "Cannot be STAT, as a max of 6ch can count as Breadth." |
| N17, S17 | "For Honours: at least 3 ch must be 4000-level. (See Constraints section below.)" |

Use **legacy notes** (`worksheet.notes.add(addr, content)`), not threaded comments — they're for static reference info, not collaboration. Threaded comments add author avatars and reply chains that clutter a documentation use case.

---

## 7. Styling conventions

- **Font:** Arial 10pt throughout
- **Header fills:** `#1F4E79` (dark blue) with white bold text
- **Borders:** light gray `#D4D4D4` "gridline-style" thin solid borders on all data blocks (because cell fills hide Excel's default gridlines)
- **Honours-disabled CF:** gray fill `#E7E6E6`, font `#7F7F7F`, strikethrough
- **Met-target CF:** green fill `#C6EFCE`, font `#006100`
- **Top-section labels** (NAME/ID/Honours, GPA labels): bold Arial 10pt, color `#375623` (dark green)
- **Top-section values:** bold black Arial 10pt, centered (except B1 left-aligned for long names)

---

## 8. How to repurpose for a different program

1. **Replace the categories** (rows 6–19 in M:P and R:U). Pick category codes that are short and unique (the Cat. column data validation pulls from this range).
2. **Set Req. CH per category.** Honours-only rows use `=IF($B$3, <honours val>, 0)`; non-conditional rows are static numbers.
3. **Set TOTAL CH** in the FreeElec formula (`132/129` becomes whatever your program requires).
4. **Update notes** (N11, N12, N15, N17, S11, S12, S15, S17) with program-specific subtleties — or remove if not applicable.
5. **Update Constraints** (rows 24–25) with rules that need advisor attestation.
6. **If the program has only one calendar edition**, delete columns Q:U (and the `Cat. ALT` column from the Courses table). Otherwise keep both editions and rename the table titles in M4 / R4.
7. **GPA filters** (gpa1/gpa2) are generic — repurpose them for whatever sub-GPA you want (major-only, recent-term, etc.).

The Progress and TOTAL formulas use structured references (`Courses[Cat. 2020]`, `Courses[CH]`, etc.) so they survive table row insertions automatically. Just keep the Cat. column name aligned with what the formulas reference.

---

## 9. Known limitations

- **Honours toggle is binary.** Programs with three or more tracks (e.g., Honours / Co-op / Major) need more than one toggle, or a single dropdown driving multi-way `IF`s.
- **No discipline-coherence check** for Breadth1's "9 ch from one coherent field" requirement — relies on the advisor reviewing the 3 courses tagged Breadth1.
- **No 4000-level check** for the Honours upper-level math requirement — encoded as a constraint checkbox instead of a formula. (Formula-checking would require parsing the course code, which is brittle.)
- **No prerequisite chain validation** — the tracker assumes the courses listed are validly taken. Advisor's job to catch chronological problems.
- **Notes carry a small embedded format block from Excel.** Programmatically-added notes inherit the workbook author's display name as a "Patrick Reynolds:" prefix in bold; strip it manually if you don't want it.
