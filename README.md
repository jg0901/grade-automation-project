# grade-automation-project

Automates final-grade computation and student notification for a Calculus course.
Excel stays the input (scores entered per exam sheet); Python validates, computes, reconciles against the Excel formulas, and sends the grade emails.

## Files

| File | Purpose |
|---|---|
| `Calculus_Dashboard Template.xlsx` | Blank gradebook template (Main + LE1–LE4, FE, SW sheets) |
| `Calculus_Grades_AY2526_2nd_sem.xlsx` | Anonymized sample gradebook used for testing |
| `grade_automation/` | Python package: `config` (settings), `load` (read workbook), `notify` (render + send), `__main__` (CLI) |
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

1. **Validate** — detect missing/invalid scores, scores over max, duplicate students, Main↔exam-sheet mismatches (matched on student number, not row position).
2. **Compute** — reimplement the grading rules in Python; reconcile against Excel results; flag discrepancies.
3. **Notify** — HTML preview, send log, duplicate-send protection, explicit confirm step.

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
