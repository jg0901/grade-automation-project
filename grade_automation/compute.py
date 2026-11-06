"""Phase 2: compute final grades in Python and make them the source of truth.

`compute_grades` applies the policy in grading.py to every student, keeps
Excel's cached values alongside for comparison, and returns rows that can be
exported (CSV / XLSX) or fed to the email renderer.
"""

from __future__ import annotations

import csv
from collections import Counter
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .grading import grade_summary
from .load import LONG_EXAMS, Gradebook, StudentRecord

TOL = 1e-6


@dataclass
class ComputedGrade:
    student: StudentRecord
    final_grade: float | None
    rounded_fg: float | None
    orig_equiv: float | None
    equiv: float | None
    reason: str | None            # why no grade, if final_grade is None

    @property
    def student_number(self) -> str:
        return self.student.student_number

    @property
    def graded(self) -> bool:
        return self.final_grade is not None

    def excel_values(self) -> dict[str, float | None]:
        s = self.student
        return {"final_grade": s.final_grade, "rounded_fg": s.rounded_fg,
                "orig_equiv": s.orig_equiv, "equiv": s.equiv}

    def python_values(self) -> dict[str, float | None]:
        return {"final_grade": self.final_grade, "rounded_fg": self.rounded_fg,
                "orig_equiv": self.orig_equiv, "equiv": self.equiv}

    def discrepancies(self) -> list[tuple[str, float | None, float | None]]:
        """(field, excel, python) for every field where they differ.

        When neither side has a final grade the derived fields are not
        compared: the template's `=ROUND(E3,2)` yields 0 on a blank grade,
        which is a display quirk, not a policy difference."""
        if self.final_grade is None and self.student.final_grade is None:
            return []
        out = []
        for k, ev in self.excel_values().items():
            pv = self.python_values()[k]
            if ev is None and pv is None:
                continue
            if ev is None or pv is None or abs(ev - pv) > TOL:
                out.append((k, ev, pv))
        return out

    def as_student(self) -> StudentRecord:
        """StudentRecord whose grade fields are the Python-computed ones."""
        return replace(self.student, final_grade=self.final_grade, rounded_fg=self.rounded_fg,
                       orig_equiv=self.orig_equiv, equiv=self.equiv)


def _no_grade_reason(s: StudentRecord) -> str:
    le = [s.scores.get(e) for e in LONG_EXAMS]
    statuses = [f"{e}={v}" for e, v in zip(LONG_EXAMS, le) if isinstance(v, str) and v.lower() != "excused"]
    fe = s.scores.get("FE")
    if isinstance(fe, str):
        statuses.append(f"FE={fe}")
    if statuses:
        return "no grade: " + ", ".join(statuses)
    blanks = [e for e, v in zip((*LONG_EXAMS, "FE"), (*le, fe)) if v is None]
    if s.sw is None:
        blanks.append("SW")
    if blanks:
        return "incomplete: " + ", ".join(blanks) + " blank"
    excused = sum(1 for v in le if isinstance(v, str) and v.lower() == "excused")
    if excused > 1:
        return f"invalid: {excused} long exams excused"
    return "no grade (unknown reason)"


def compute_grades(gb: Gradebook) -> list[ComputedGrade]:
    out = []
    for s in gb.students:
        g = grade_summary(s.scores, s.sw, gb.maxima)
        out.append(ComputedGrade(
            student=s,
            final_grade=g["final_grade"], rounded_fg=g["rounded_fg"],
            orig_equiv=g["orig_equiv"], equiv=g["equiv"],
            reason=None if g["final_grade"] is not None else _no_grade_reason(s),
        ))
    return out


def graded_gradebook(gb: Gradebook, computed: list[ComputedGrade] | None = None) -> Gradebook:
    """A copy of `gb` whose students carry Python-computed grade fields."""
    computed = computed or compute_grades(gb)
    return Gradebook(maxima=gb.maxima, students=[c.as_student() for c in computed], source=gb.source)


# --------------------------------------------------------------------------
# export
# --------------------------------------------------------------------------

EXPORT_COLUMNS = (
    "Student Number", "Family Name", "First Name", "Email",
    "LE1", "LE2", "LE3", "LE4", "FE", "SW",
    "Excused", "Final Grade", "Rounded FG", "Orig Equiv", "Equiv",
    "Excel Final Grade", "Excel Equiv", "Matches Excel", "Note",
)


def _fmt(v: Any) -> Any:
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return "" if v is None else v


def export_rows(computed: list[ComputedGrade]) -> list[dict[str, Any]]:
    rows = []
    for c in computed:
        s = c.student
        rows.append({
            "Student Number": s.student_number,
            "Family Name": s.family_name,
            "First Name": s.first_name,
            "Email": s.email or "",
            **{e: _fmt(s.scores.get(e)) for e in (*LONG_EXAMS, "FE")},
            "SW": "" if s.sw is None else round(s.sw, 6),
            "Excused": s.excused_exam or "",
            "Final Grade": "" if c.final_grade is None else round(c.final_grade, 6),
            "Rounded FG": _fmt(c.rounded_fg),
            "Orig Equiv": _fmt(c.orig_equiv),
            "Equiv": _fmt(c.equiv),
            "Excel Final Grade": "" if s.final_grade is None else round(s.final_grade, 6),
            "Excel Equiv": _fmt(s.equiv),
            "Matches Excel": "yes" if not c.discrepancies() else "NO",
            "Note": c.reason or "",
        })
    return rows


def write_csv(computed: list[ComputedGrade], path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=EXPORT_COLUMNS)
        w.writeheader()
        w.writerows(export_rows(computed))
    return path


def write_xlsx(computed: list[ComputedGrade], path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    head_font = Font(bold=True)
    head_fill = PatternFill("solid", fgColor="DDDDDD")
    bad_fill = PatternFill("solid", fgColor="F8CBAD")

    # Sheet 1: Final Grades
    ws = wb.active
    ws.title = "Final Grades"
    ws.append(list(EXPORT_COLUMNS))
    for cell in ws[1]:
        cell.font = head_font
        cell.fill = head_fill
    rows = export_rows(computed)
    for r in rows:
        ws.append([r[k] for k in EXPORT_COLUMNS])
    match_col = EXPORT_COLUMNS.index("Matches Excel") + 1
    for i, r in enumerate(rows, start=2):
        if r["Matches Excel"] == "NO":
            for cell in ws[i]:
                cell.fill = bad_fill
        ws.cell(i, EXPORT_COLUMNS.index("Final Grade") + 1).number_format = "0.00%"
        ws.cell(i, EXPORT_COLUMNS.index("Excel Final Grade") + 1).number_format = "0.00%"
        ws.cell(i, EXPORT_COLUMNS.index("Rounded FG") + 1).number_format = "0.00%"
        ws.cell(i, EXPORT_COLUMNS.index("SW") + 1).number_format = "0.00%"
        for k in ("Orig Equiv", "Equiv", "Excel Equiv"):
            ws.cell(i, EXPORT_COLUMNS.index(k) + 1).number_format = "0.00"
    ws.freeze_panes = "B2"
    for col in range(1, len(EXPORT_COLUMNS) + 1):
        width = max(len(str(ws.cell(r, col).value or "")) for r in range(1, ws.max_row + 1))
        ws.column_dimensions[get_column_letter(col)].width = min(max(8, width + 2), 40)
    _ = match_col

    # Sheet 2: Discrepancies
    wd = wb.create_sheet("Discrepancies")
    wd.append(["Student Number", "Field", "Excel", "Python", "Difference"])
    for cell in wd[1]:
        cell.font = head_font
        cell.fill = head_fill
    n = 0
    for c in computed:
        for field, ev, pv in c.discrepancies():
            diff = "" if ev is None or pv is None else pv - ev
            wd.append([c.student_number, field, _fmt(ev), _fmt(pv), diff])
            n += 1
    if n == 0:
        wd.append(["(none)", "", "", "", ""])
    for col in range(1, 6):
        wd.column_dimensions[get_column_letter(col)].width = 16

    # Sheet 3: Summary
    wsum = wb.create_sheet("Summary")
    summ = summary(computed)
    wsum.append(["Metric", "Value"])
    for cell in wsum[1]:
        cell.font = head_font
        cell.fill = head_fill
    wsum.append(["Students", summ["students"]])
    wsum.append(["Graded", summ["graded"]])
    wsum.append(["No grade", summ["ungraded"]])
    wsum.append(["Discrepancies vs Excel", summ["discrepancies"]])
    wsum.append([])
    wsum.append(["Equivalent", "Count"])
    for cell in wsum[wsum.max_row]:
        cell.font = head_font
    for eq, cnt in summ["distribution"]:
        wsum.append([f"{eq:.2f}", cnt])
    wsum.column_dimensions["A"].width = 26
    wsum["A1"].alignment = Alignment(horizontal="left")

    wb.save(path)
    return path


def summary(computed: list[ComputedGrade]) -> dict[str, Any]:
    dist = Counter(c.equiv for c in computed if c.graded)
    return {
        "students": len(computed),
        "graded": sum(1 for c in computed if c.graded),
        "ungraded": sum(1 for c in computed if not c.graded),
        "discrepancies": sum(len(c.discrepancies()) for c in computed),
        "distribution": sorted(dist.items()),
    }


def format_summary(computed: list[ComputedGrade]) -> str:
    s = summary(computed)
    lines = [
        f"Computed grades: {s['graded']} graded, {s['ungraded']} without a grade, "
        f"{s['discrepancies']} field(s) differ from Excel",
    ]
    if s["distribution"]:
        lines.append("Equivalent distribution: " + "  ".join(f"{eq:.2f}×{n}" for eq, n in s["distribution"]))
    for c in computed:
        if not c.graded:
            lines.append(f"  {c.student_number:<12} {c.reason}")
    for c in computed:
        for field, ev, pv in c.discrepancies():
            lines.append(f"  DIFF {c.student_number:<12} {field}: Excel={ev} Python={pv}")
    return "\n".join(lines)
