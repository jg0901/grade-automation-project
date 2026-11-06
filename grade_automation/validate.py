"""Phase 1: validate the gradebook before anything is computed or sent.

Checks (E = error, blocks sending; W = warning, review recommended):

  Main sheet
    E  duplicate student number
    E  student row with no student number but with scores
    W  missing first name
    W  email present but malformed
    E  LE/FE cell is neither numeric nor an allowed status (Excused/Absent/LOA)
    E  numeric score < 0 or > max (from row 2)
    E  more than one Excused long exam
    W  FE is non-numeric (grade will be blank)
    W  final grade blank although all inputs look complete
    E  Excel final grade / equivalents differ from the Python recomputation

  Exam sheets (LE1–LE4, FE, SW)
    E  student in Main missing from sheet / student in sheet missing from Main
    E  Main's score for a student != that student's Total on the exam sheet
       (catches the row-position-based ='LE1'!D:D references going out of sync)
    E  per-item score < 0 or > item max
    W  Total is numeric but every item is blank (scores not entered?)
    W  item max missing for a column that has scores
"""

from __future__ import annotations

import csv
import math
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pandas as pd

from .grading import grade_summary
from .load import EXCUSED, LONG_EXAMS, Gradebook, StudentRecord

ALLOWED_STATUSES = {EXCUSED.lower(), "absent", "loa"}
TOL = 1e-6  # floating tolerance for Excel-vs-Python comparisons

SHEET_FOR_SCORE = {"LE1": "LE1", "LE2": "LE2", "LE3": "LE3", "LE4": "LE4", "FE": "FE", "SW": "SW"}


@dataclass
class Issue:
    severity: str          # "error" | "warning"
    check: str             # short machine-readable name
    sheet: str
    row: int | None        # 1-based Excel row, if known
    student_number: str
    message: str

    def as_row(self) -> dict[str, Any]:
        return {
            "severity": self.severity, "check": self.check, "sheet": self.sheet,
            "row": "" if self.row is None else self.row,
            "student_number": self.student_number, "message": self.message,
        }


@dataclass
class ValidationReport:
    issues: list[Issue] = field(default_factory=list)

    def error(self, check, sheet, row, sn, msg):   # noqa: D102
        self.issues.append(Issue("error", check, sheet, row, sn, msg))

    def warn(self, check, sheet, row, sn, msg):    # noqa: D102
        self.issues.append(Issue("warning", check, sheet, row, sn, msg))

    @property
    def errors(self) -> list[Issue]:
        return [i for i in self.issues if i.severity == "error"]

    @property
    def warnings(self) -> list[Issue]:
        return [i for i in self.issues if i.severity == "warning"]

    @property
    def ok(self) -> bool:
        return not self.errors

    def sorted(self) -> list[Issue]:
        """Errors before warnings; within that, grouped by student, then sheet."""
        sev = {"error": 0, "warning": 1}
        return sorted(self.issues, key=lambda i: (sev[i.severity], i.student_number, i.sheet, i.row or 0))

    def write_csv(self, path: Path) -> Path:
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=["severity", "check", "sheet", "row", "student_number", "message"])
            w.writeheader()
            for i in self.sorted():
                w.writerow(i.as_row())
        return path


# --------------------------------------------------------------------------
# exam-sheet loader
# --------------------------------------------------------------------------

@dataclass
class ExamRow:
    row: int                       # 1-based Excel row
    student_number: str
    total: Any                     # float or status str or None
    items: dict[str, Any]          # item header -> float | None


@dataclass
class ExamSheet:
    name: str
    item_max: dict[str, float | None]
    total_max: float | None
    rows: dict[str, ExamRow]       # keyed by student number
    duplicates: list[str]


def _blank(v: Any) -> bool:
    return v is None or (isinstance(v, float) and math.isnan(v)) or (isinstance(v, str) and not v.strip())


def _num(v: Any) -> float | None:
    if _blank(v) or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip())
    except ValueError:
        return None


def _cell(v: Any) -> Any:
    if _blank(v):
        return None
    n = _num(v)
    return n if n is not None else str(v).strip()


def load_exam_sheet(path: str | Path, sheet: str) -> ExamSheet:
    """Parse an exam sheet regardless of whether the maxima row is above or
    below the header row (LE sheets: header row 1, maxima row 2; FE/SW:
    maxima row 1, header row 2)."""
    raw = pd.read_excel(path, sheet_name=sheet, header=None)

    header_idx = None
    for i in range(min(3, len(raw))):
        first = raw.iat[i, 0]
        if isinstance(first, str) and first.strip().lower() in {"sn#", "student number", "sn"}:
            header_idx = i
            break
    if header_idx is None:
        raise ValueError(f"{sheet}: could not find the header row (expected 'SN#' or 'Student Number' in column A)")

    headers = [None if _blank(h) else str(h).strip() for h in raw.iloc[header_idx].tolist()]
    try:
        total_col = next(i for i, h in enumerate(headers) if h and h.lower() == "total")
    except StopIteration:
        raise ValueError(f"{sheet}: no 'Total' column") from None

    # Maxima row: the neighbour of the header row whose Total is numeric and SN is blank.
    cand = [i for i in (header_idx - 1, header_idx + 1) if 0 <= i < len(raw)]
    max_idx = next(
        (i for i in cand if _blank(raw.iat[i, 0]) and _num(raw.iat[i, total_col]) is not None),
        None,
    )

    # Item columns: everything after Total, skipping '%' style derived columns.
    item_cols = [i for i, h in enumerate(headers) if i > total_col and h and h != "%"]
    item_max = {headers[i]: (_num(raw.iat[max_idx, i]) if max_idx is not None else None) for i in item_cols}
    total_max = _num(raw.iat[max_idx, total_col]) if max_idx is not None else None

    rows: dict[str, ExamRow] = {}
    duplicates: list[str] = []
    for i in range(header_idx + 1, len(raw)):
        if i == max_idx:
            continue
        sn = raw.iat[i, 0]
        if _blank(sn):
            continue
        sn = str(sn).strip()
        er = ExamRow(
            row=i + 1,
            student_number=sn,
            total=_cell(raw.iat[i, total_col]),
            items={headers[c]: _num(raw.iat[i, c]) for c in item_cols},
        )
        if sn in rows:
            duplicates.append(sn)
        rows[sn] = er
    return ExamSheet(name=sheet, item_max=item_max, total_max=total_max, rows=rows, duplicates=duplicates)


# --------------------------------------------------------------------------
# checks
# --------------------------------------------------------------------------

def _check_main(gb: Gradebook, rep: ValidationReport) -> None:
    seen: dict[str, int] = {}
    for s in gb.students:
        sn = s.student_number
        if sn in seen:
            rep.error("duplicate_student", "Main", s.row_number, sn,
                      f"student number also appears on row {seen[sn]}")
        seen.setdefault(sn, s.row_number)

        if not s.first_name:
            rep.warn("missing_first_name", "Main", s.row_number, sn, "first name is blank")
        if s.email and ("@" not in s.email or " " in s.email or "." not in s.email.split("@")[-1]):
            rep.warn("bad_email", "Main", s.row_number, sn, f"email looks malformed: {s.email!r}")

        excused = 0
        complete = True
        for exam in (*LONG_EXAMS, "FE"):
            v = s.scores.get(exam)
            mx = gb.maxima.for_exam(exam)
            if v is None:
                complete = False
                rep.error("missing_score", "Main", s.row_number, sn, f"{exam} is blank")
            elif isinstance(v, str):
                if v.lower() not in ALLOWED_STATUSES:
                    rep.error("invalid_status", "Main", s.row_number, sn,
                              f"{exam} = {v!r} is not a number or one of Excused/Absent/LOA")
                    complete = False
                elif v.lower() == EXCUSED.lower():
                    if exam == "FE":
                        rep.warn("fe_not_numeric", "Main", s.row_number, sn,
                                 "FE is 'Excused'; the policy has no rule for this and the grade will be blank")
                        complete = False
                    else:
                        excused += 1
                else:
                    complete = False  # Absent / LOA -> no grade, by design
            else:
                if v < 0 or v > mx:
                    rep.error("score_out_of_range", "Main", s.row_number, sn,
                              f"{exam} = {v:g} is outside 0..{mx:g}")
        if excused > 1:
            rep.error("multiple_excused", "Main", s.row_number, sn,
                      f"{excused} long exams marked Excused; policy allows at most 1")
            complete = False
        if s.sw is None:
            complete = False
            rep.error("missing_score", "Main", s.row_number, sn, "SW is blank")
        elif not 0 <= s.sw <= max(1.0, gb.maxima.SW):
            rep.error("score_out_of_range", "Main", s.row_number, sn, f"SW = {s.sw:g} is outside 0..1")

        # Final grade presence / recomputation
        if complete and s.final_grade is None:
            rep.warn("final_grade_blank", "Main", s.row_number, sn,
                     "all inputs look complete but Final Grade is blank (formula not recalculated?)")
        if s.final_grade is not None:
            py = grade_summary(s.scores, s.sw, gb.maxima)
            for key, excel_val in (("final_grade", s.final_grade), ("rounded_fg", s.rounded_fg),
                                   ("orig_equiv", s.orig_equiv), ("equiv", s.equiv)):
                pv = py[key]
                if pv is None or excel_val is None:
                    if pv != excel_val:
                        rep.error("recompute_mismatch", "Main", s.row_number, sn,
                                  f"{key}: Excel={excel_val} Python={pv}")
                elif abs(pv - excel_val) > TOL:
                    rep.error("recompute_mismatch", "Main", s.row_number, sn,
                              f"{key}: Excel={excel_val:.6f} Python={pv:.6f}")


def _check_exam_sheet(gb: Gradebook, es: ExamSheet, main_key: str, rep: ValidationReport) -> None:
    main_by_sn: dict[str, StudentRecord] = {s.student_number: s for s in gb.students}

    for sn in es.duplicates:
        rep.error("duplicate_student", es.name, es.rows[sn].row, sn, "student number appears more than once")

    for sn, s in main_by_sn.items():
        if sn not in es.rows:
            rep.error("student_missing_in_sheet", es.name, None, sn,
                      f"present in Main (row {s.row_number}) but not in {es.name}")
    for sn, er in es.rows.items():
        if sn not in main_by_sn:
            rep.error("student_missing_in_main", es.name, er.row, sn, f"present in {es.name} but not in Main")

    for h, mx in es.item_max.items():
        if mx is None and any(er.items.get(h) is not None for er in es.rows.values()):
            rep.warn("item_max_missing", es.name, None, "", f"column {h!r} has scores but no maximum")

    for sn, er in es.rows.items():
        # per-item range
        for h, v in er.items.items():
            mx = es.item_max.get(h)
            if v is None:
                continue
            if v < 0:
                rep.error("item_out_of_range", es.name, er.row, sn, f"{h} = {v:g} is negative")
            elif mx is not None and v > mx + TOL:
                rep.error("item_out_of_range", es.name, er.row, sn, f"{h} = {v:g} exceeds max {mx:g}")

        all_blank = all(v is None for v in er.items.values())
        if isinstance(er.total, float) and all_blank and es.name != "SW":
            rep.warn("total_without_items", es.name, er.row, sn,
                     f"Total = {er.total:g} but no item scores entered")

        # Main <-> sheet agreement, matched on student number
        s = main_by_sn.get(sn)
        if s is None:
            continue
        if main_key == "SW":
            main_val = s.sw
            sheet_val = None
            if isinstance(er.total, float) and es.total_max:
                sheet_val = er.total / es.total_max
            if main_val is not None and sheet_val is not None and abs(main_val - sheet_val) > TOL:
                rep.error("main_sheet_mismatch", "Main", s.row_number, sn,
                          f"SW in Main = {main_val:.6f} but SW sheet Total/max = {sheet_val:.6f} (row {er.row})")
            continue

        main_val = s.scores.get(main_key)
        sheet_val = er.total
        same = (
            (main_val is None and sheet_val is None)
            or (isinstance(main_val, float) and isinstance(sheet_val, float) and abs(main_val - sheet_val) <= TOL)
            or (isinstance(main_val, str) and isinstance(sheet_val, str) and main_val.lower() == sheet_val.lower())
        )
        if not same:
            rep.error("main_sheet_mismatch", "Main", s.row_number, sn,
                      f"{main_key} in Main = {main_val!r} but {es.name} Total = {sheet_val!r} (row {er.row}); "
                      "Main may be referencing the wrong row")


def validate(gb: Gradebook, sheets: dict[str, ExamSheet] | None = None) -> ValidationReport:
    """Run all checks. `sheets` may be passed for tests; otherwise they are
    loaded from the gradebook's workbook."""
    rep = ValidationReport()
    _check_main(gb, rep)

    if sheets is None:
        sheets = {}
        for key, name in SHEET_FOR_SCORE.items():
            try:
                sheets[key] = load_exam_sheet(gb.source, name)
            except ValueError as exc:
                rep.error("sheet_unreadable", name, None, "", str(exc))
    for key, es in sheets.items():
        _check_exam_sheet(gb, es, key, rep)

    return rep


def format_report(rep: ValidationReport, limit: int = 60) -> str:
    lines = [f"Validation: {len(rep.errors)} error(s), {len(rep.warnings)} warning(s)"]
    for i in rep.sorted()[:limit]:
        row = f"row {i.row}" if i.row else ""
        lines.append(f"  {i.severity[:4].upper():<4} {i.sheet:<5} {row:<8} {i.student_number:<12} "
                     f"{i.check:<26} {i.message}")
    if len(rep.issues) > limit:
        lines.append(f"  ... {len(rep.issues) - limit} more (see the CSV report)")
    return "\n".join(lines)
