# Calculus Grade Automation

[![Tests](https://github.com/jg0901/grade-automation-project/actions/workflows/tests.yml/badge.svg)](https://github.com/jg0901/grade-automation-project/actions/workflows/tests.yml)

A Python-based workflow for **validating Calculus gradebooks, calculating final grades, reconciling results with Excel, and preparing personalized grade notifications**.

Built from a grading and email-notification process used across two academic terms. Excel remains the familiar score-entry interface; Python adds reproducible calculations, quality checks, reporting, and safeguards before any email is sent.

## How it works

```text
Excel gradebook (Main + LE1–LE4, FE, SW)
             |
             v
   Load and validate records
   - score ranges and missing values
   - student-number matching across sheets
   - Excel vs. Python grade reconciliation
             |
             v
     Compute grades in Python
             |
             +--> CSV / Excel reports
             |
             +--> HTML email previews
                        |
                        v
                Review recipients
                        |
                        v
             Explicitly approved send
                 (test or live)
```

### Key features

- **Configurable exam maxima:** each long exam can have a different total (e.g., 30, 35, or 40 points); Python reads these from the workbook rather than assuming every exam is out of 30.
- **Policy-aware grading:** accounts for a single excused long exam by transferring its 15% weight to the final exam.
- **Independent reconciliation:** compares Python-computed grades and equivalents with the **saved/cached Excel results** and reports discrepancies.
- **Student-level checks:** detects duplicate student numbers, invalid scores, missing records, and mismatches between `Main` and individual assessment sheets, matched by student number rather than row position.
- **Safer notifications:** generates HTML previews, defaults to a no-send recipient review, supports test-inbox routing, logs send outcomes, and prevents accidental repeat live sends.

## Grading policy

For each student, the final grade is computed as a fraction from 0 to 1:

```text
Final grade =
    0.15 * sum(LE_i score / LE_i maximum)       # numeric long exams
  + (0.30 + 0.15 * excused_count) * (FE / FE maximum)
  + 0.10 * (SW score / SW maximum)
```

| Component | Weight |
| --- | ---: |
| Long exams (LE1–LE4) | 15% each |
| Final exam | 30% normally; 45% if one LE is excused |
| Seatwork/homework | 10% |

An `Excused` LE contributes no direct score; its weight moves to the final exam. The current policy allows **at most one** excused long exam. Incomplete or ineligible records (including `Absent` / `LOA` cases) do not receive a computed final grade.

The Python engine uses **half-up rounding on the grade fraction to two decimal places** (matching the template's Excel rounding rule) before converting it to the UP equivalent (1.00–5.00). Equivalents are determined by thresholds defined in `grade_automation/grading.py`.

## Quick start

**Requirements:** Python 3.11+ and a compatible Excel gradebook.

```bash
git clone https://github.com/jg0901/grade-automation-project.git
cd grade-automation-project
python -m pip install -r requirements.txt
cp .env.example .env
```

Edit `.env` to select your workbook and output directory:

```dotenv
GRADES_FILE=Calculus_Grades_AY2526_2nd_sem.xlsx
GRADES_SHEET=Main
OUTPUT_DIR=output
GRADE_SOURCE=python
TEST_MODE=true
```

The included anonymized workbook is intended for demonstration and tests. For actual use, place private workbooks in `data/` (ignored by Git) and set `GRADES_FILE=data/your_gradebook.xlsx`.

**Important:** The loader reads Excel's *saved formula values*; it does not recalculate Excel formulas. If assessment totals or maximums were changed, open the workbook in Excel, recalculate, and save it before running the commands below.

## Commands

Run these from the repository root after configuring `.env`.

| Command | Result | Sends email? |
| --- | --- | :---: |
| `python -m grade_automation validate` | Writes `output/validation_report.csv`; blocks on validation errors | No |
| `python -m grade_automation compute` | Writes `output/final_grades.csv` and `output/final_grades.xlsx` | No |
| `python -m grade_automation preview` | Creates one HTML preview per eligible student in `output/previews/` | No |
| `python -m grade_automation send` | Shows a recipient list and how many messages would be sent | No |
| `python -m grade_automation send --confirm` | Sends to `TEST_EMAIL` when `TEST_MODE=true`; otherwise initiates live-send approval | **Yes** |

The exported Excel report includes **Final Grades**, **Discrepancies**, and **Summary** worksheets. `compute` exits with a nonzero status if Python and Excel grade values differ.

By default, `GRADE_SOURCE=python` means notification content uses Python-computed values. `GRADE_SOURCE=excel` switches it to Excel's saved values, while validation still compares both calculations.

### Sending safely

1. Run `validate` and investigate any errors or warnings.
2. Run `compute` and review discrepancies, if any.
3. Run `preview` to inspect the emails, then `send` to inspect the recipient list.
4. For a test delivery, fill in `SENDER_EMAIL`, `APP_PASSWORD`, `SENDER_NAME`, and `TEST_EMAIL` in `.env`, leave `TEST_MODE=true`, and run `send --confirm`. **This sends real emails to your test inbox**, not to students.
5. For a live delivery, set `TEST_MODE=false`, review the summary, and run `send --confirm` in an interactive terminal. The program requires an exact `SEND <n>` confirmation, where `n` is the number of pending emails.

Additional protections:

- Sending stops on validation errors unless `--skip-validation` is explicitly supplied (**not recommended**).
- Records without a usable email, computed grade, or equivalent are reported as skipped.
- Live sends are logged in `output/sent_log.csv`; already-sent student numbers are skipped on subsequent live runs. `--resend` overrides this protection.
- Placeholder sender credentials are rejected. Test-mode deliveries are **not** deduplicated, so repeated confirmed tests will send again.
- The live-send approval step is rejected in non-interactive environments.

## Workbook format

The repository contains:

- `Calculus_Dashboard Template.xlsx` — an Excel gradebook template with `Main` and assessment worksheets.
- `Calculus_Grades_AY2526_2nd_sem.xlsx` — an anonymized example used by the test suite.

The `Main` sheet has headers in row 1, **assessment maximums in row 2**, and student records from row 3 onward. Supported score states include numeric values and `Excused`, `Absent`, or `LOA` where applicable. Assessment sheets are cross-checked against `Main` using student numbers.

The current implementation expects the template's column layout and sheet names; it is not a general-purpose gradebook importer.

## Validation and tests

Validation distinguishes **errors** (which block a send) from **warnings** (which need review). It covers:

- Missing, invalid, negative, or above-maximum scores; more than one excused LE
- Duplicate student numbers and inconsistent class rosters across worksheets
- Mismatched scores between `Main` and assessment sheets
- Item-level scores outside their allowed ranges or missing item maxima
- Differences between Python-calculated and saved Excel grades/equivalents
- Missing first names, malformed emails, or incomplete grading information

Run the automated test suite:

```bash
python -m pip install pytest
python -m pytest
```

GitHub Actions runs the tests and sample-workbook CLI checks on Python **3.11, 3.12, and 3.13**.

## Repository structure

| Path | Responsibility |
| --- | --- |
| `grade_automation/load.py` | Read workbook data, student records, and assessment maxima |
| `grade_automation/grading.py` | Compute grades, apply rounding and equivalent thresholds |
| `grade_automation/validate.py` | Validate scores, rosters, and reconciliation results |
| `grade_automation/compute.py` | Generate grade exports and discrepancy summaries |
| `grade_automation/notify.py` | Render email content, send via SMTP, and log delivery outcomes |
| `grade_automation/__main__.py` | Command-line interface |
| `tests/` | Unit and CLI tests, including simulated email delivery |
| `.env.example` | Non-secret example of runtime and email settings |

## Privacy and intended use

**Do not commit real student records, credentials, generated reports, or email logs.** The `data/` and `output/` directories and local `.env` are Git-ignored. Keep any real gradebooks in a private location, and review files and Git history before making a repository public.

This tool supports an instructor's grading workflow; it does not replace review of grading policy, disputed marks, or the final decision to release grades.
