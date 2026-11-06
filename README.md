# grade-automation-project

Automates final-grade computation and student notification for a Calculus course.
Excel stays the input (scores entered per exam sheet); Python validates, computes, reconciles against the Excel formulas, and sends the grade emails.

## Files

| File | Purpose |
|---|---|
| `Calculus_Dashboard Template.xlsx` | Blank gradebook template (Main + LE1–LE4, FE, SW sheets) |
| `Calculus_Grades_AY2526_2nd_sem.xlsx` | Anonymized sample gradebook used for testing |
| `grade_automation/` | Python package: `config` (settings), `load` (read Main sheet), `grading` (policy in Python), `validate` (checks), `notify` (render + send), `__main__` (CLI) |
| `tests/` | pytest suite run against the sample workbook |
| `.env.example` | Template for credentials and run settings |

## Grading policy (as implemented in the Excel template)

```
Final Grade = 0.15 × Σ(LE_i / LE_i max)          # 4 long exams, 15% each
            + (0.30 + 0.15 × excused) × FE / FE max   # FE is 30%, or 45% if one LE is excused
            + 0.10 × SW%                           # seatwork/homework
```

Final grade is rounded to 2 decimals (half-up) before mapping to the UP equivalent (1.00–5.00).

## Roadmap

1. **Validate** ✅ — missing/invalid scores, scores over max, duplicate students, Main↔exam-sheet mismatches (matched on student number, not row position), and a Python recomputation of every final grade compared against Excel's.
2. **Compute** ✅ — Python is the source of truth for Final Grade / Equiv (`grading.py`, reconciles 1:1 with the Excel template). `compute` writes `output/final_grades.{csv,xlsx}` with a Discrepancies sheet; emails use the Python values unless `GRADE_SOURCE=excel`.
3. **Notify** ✅ (basic) — HTML preview, send log, duplicate-send protection, explicit confirm step. Remaining: final approval summary before a live send.

## Validation checks

| Check | Severity | What it catches |
|---|---|---|
| `duplicate_student` | error | same student number twice (Main or any exam sheet) |
| `missing_score` / `invalid_status` | error | blank cell, or text other than `Excused` / `Absent` / `LOA` |
| `score_out_of_range` / `item_out_of_range` | error | score below 0 or above the max in the maxima row |
| `multiple_excused` | error | more than one long exam excused |
| `main_sheet_mismatch` | error | Main's LE/FE/SW value ≠ that student's Total on the exam sheet — i.e. the `='LE1'!D:D` row references drifted |
| `student_missing_in_sheet` / `_in_main` | error | roster differs between Main and an exam sheet |
| `recompute_mismatch` | error | Excel's Final Grade / Rounded_FG / Orig_Equiv / Equiv ≠ Python's |
| `fe_not_numeric`, `final_grade_blank`, `total_without_items`, `item_max_missing`, `missing_first_name`, `bad_email` | warning | review recommended |

`send` refuses to run while there are errors (`--skip-validation` overrides, not recommended).

### Known template quirk

`Rounded_FG` in the Main sheet is `=ROUND(E3,2)`, which shows `0.00` for students with no final grade (Absent/LOA). It is cosmetic — `Equiv` still comes out blank — but `=IF(E3="","",ROUND(E3,2))` is the cleaner formula. The Python side treats "no grade on both sides" as consistent, so this is not reported as a discrepancy.

## Data handling

Never commit real student data. Only the anonymized sample and the blank template are tracked; real workbooks go in `data/` (gitignored).

## Setup

```bash
pip install -r requirements.txt
cp .env.example .env   # then fill in
```

## Usage

Everything is a dry run unless you say otherwise.

```bash
# 0. Check the workbook; writes output/validation_report.csv
python -m grade_automation validate

# 0b. Final grades from Python -> output/final_grades.csv + .xlsx
#     (sheets: Final Grades, Discrepancies vs Excel, Summary). Exit code 1 if anything differs.
python -m grade_automation compute

# 1. Render one HTML file per student to output/previews/ — sends nothing
python -m grade_automation preview

# 2. Show who would receive what (still sends nothing)
python -m grade_automation send

# 3. Send. With TEST_MODE=true in .env, every email goes to TEST_EMAIL.
python -m grade_automation send --confirm

# 4. For the real send: set TEST_MODE=false in .env, then
python -m grade_automation send --confirm
```

Safety features:

- Students without an email, without a final grade (Absent/LOA), or with a blank equivalent are listed as skipped with the reason, never silently dropped.
- Every send is appended to `output/sent_log.csv`. Re-running skips anyone already logged as sent; use `--resend` to override.
- Sending refuses to start if `.env` still has placeholder values.
- Maxima (LE/FE out-of) are read from row 2 of the Main sheet, not hardcoded.

Run the tests with `python -m pytest`.
