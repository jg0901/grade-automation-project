"""Read the gradebook's Main sheet into typed per-student records.

Layout of the Main sheet (see Calculus_Dashboard Template.xlsx):

    row 1  headers
    row 2  maxima  (LE1..LE4 max, FE max, SW max=1)  -- no student number
    row 3+ one student per row

Scores for the long exams / final exam are numeric, or one of the status
strings "Excused", "Absent", "LOA" (anything non-numeric is kept verbatim so
it can be shown in the email and flagged by validation).
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pandas as pd

LONG_EXAMS = ("LE1", "LE2", "LE3", "LE4")
EXCUSED = "Excused"

# Canonical column name -> accepted header spellings in the workbook.
# The notebook used "HW/SW" while the template header is "SW"; accept both.
COLUMN_ALIASES: dict[str, tuple[str, ...]] = {
    "email": ("Email Address:", "Email Address", "Email"),
    "student_number": ("Student Number", "SN#", "SN"),
    "family_name": ("Family Name", "Last Name"),
    "first_name": ("First Name",),
    "final_grade": ("Final Grade",),
    "orig_equiv": ("Orig_Equiv",),
    "rounded_fg": ("Rounded_FG",),
    "equiv": ("Equiv",),
    "LE1": ("LE1",),
    "LE2": ("LE2",),
    "LE3": ("LE3",),
    "LE4": ("LE4",),
    "FE": ("FE",),
    "SW": ("SW", "HW/SW", "HWSW", "SW/HW"),
}


class GradebookError(Exception):
    """Raised when the workbook can't be read in the expected layout."""


@dataclass
class Maxima:
    LE1: float
    LE2: float
    LE3: float
    LE4: float
    FE: float
    SW: float = 1.0

    def for_exam(self, name: str) -> float:
        return getattr(self, name)


@dataclass
class StudentRecord:
    row_number: int                 # 1-based Excel row, for error messages
    student_number: str
    first_name: str
    family_name: str
    email: str | None
    final_grade: float | None       # fraction, e.g. 0.8376
    orig_equiv: float | None
    rounded_fg: float | None        # fraction rounded to 2 dp, e.g. 0.84
    equiv: float | None
    scores: dict[str, Any] = field(default_factory=dict)  # LE1..LE4, FE: float or status str
    sw: float | None = None         # fraction, e.g. 0.74

    @property
    def excused_exam(self) -> str | None:
        for exam in LONG_EXAMS:
            if _is_status(self.scores.get(exam), EXCUSED):
                return exam
        return None

    @property
    def has_email(self) -> bool:
        return bool(self.email) and "@" in self.email

    def skip_reason(self) -> str | None:
        """Why this student should NOT receive a grade email, or None if OK."""
        if not self.has_email:
            return "no email address"
        if self.final_grade is None:
            return "no final grade (Absent/LOA or incomplete scores)"
        if self.equiv is None or self.orig_equiv is None:
            return "equivalent grade is blank"
        return None


@dataclass
class Gradebook:
    maxima: Maxima
    students: list[StudentRecord]
    source: Path

    def sendable(self) -> list[StudentRecord]:
        return [s for s in self.students if s.skip_reason() is None]

    def skipped(self) -> list[tuple[StudentRecord, str]]:
        return [(s, r) for s in self.students if (r := s.skip_reason())]


# --------------------------------------------------------------------------
# helpers
# --------------------------------------------------------------------------

def _is_blank(v: Any) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and math.isnan(v):
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    return False


def _is_status(v: Any, status: str) -> bool:
    return isinstance(v, str) and v.strip().lower() == status.lower()


def _to_float(v: Any) -> float | None:
    """Numeric value or None. Non-numeric strings (Excused/Absent/LOA) -> None."""
    if _is_blank(v):
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip())
    except ValueError:
        return None


def _score(v: Any) -> Any:
    """Keep numbers as float, keep status strings verbatim, blank -> None."""
    if _is_blank(v):
        return None
    f = _to_float(v)
    return f if f is not None else str(v).strip()


def _resolve_columns(columns: list[str]) -> dict[str, str]:
    """Map canonical names -> actual header present in the sheet."""
    normalized = {str(c).strip().lower(): c for c in columns}
    resolved: dict[str, str] = {}
    missing: list[str] = []
    for canonical, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias.lower() in normalized:
                resolved[canonical] = normalized[alias.lower()]
                break
        else:
            missing.append(f"{canonical} (tried: {', '.join(aliases)})")
    if missing:
        raise GradebookError("Missing columns in Main sheet: " + "; ".join(missing))
    return resolved


# --------------------------------------------------------------------------
# public API
# --------------------------------------------------------------------------

def load_gradebook(path: str | Path, sheet: str = "Main") -> Gradebook:
    path = Path(path)
    if not path.exists():
        raise GradebookError(f"Gradebook not found: {path}")

    df = pd.read_excel(path, sheet_name=sheet)
    cols = _resolve_columns(list(df.columns))
    c = lambda name: cols[name]  # noqa: E731

    # Row 2 in Excel (index 0 here) holds the maxima and has no student number.
    first = df.iloc[0]
    if not _is_blank(first[c("student_number")]):
        raise GradebookError(
            "Expected row 2 of the Main sheet to hold the maximum scores "
            "(no student number), but it has a student number."
        )
    maxima_vals = {k: _to_float(first[c(k)]) for k in (*LONG_EXAMS, "FE")}
    bad = [k for k, v in maxima_vals.items() if v is None or v <= 0]
    if bad:
        hint = ""
        if len(bad) == 5 and all(_is_blank(first[c(k)]) for k in bad):
            hint = (" All five are blank, which usually means the workbook was last saved by a "
                    "tool that does not calculate formulas (e.g. openpyxl). Open it in Excel, "
                    "let it recalculate, and save again.")
        raise GradebookError(f"Row 2 maxima missing or non-positive for: {', '.join(bad)}.{hint}")
    maxima = Maxima(**maxima_vals, SW=_to_float(first[c("SW")]) or 1.0)

    students: list[StudentRecord] = []
    for idx, row in df.iloc[1:].iterrows():
        sn = row[c("student_number")]
        if _is_blank(sn):
            continue  # trailing empty rows
        email = row[c("email")]
        students.append(
            StudentRecord(
                row_number=int(idx) + 2,  # pandas index 0 == Excel row 2
                student_number=str(sn).strip(),
                first_name="" if _is_blank(row[c("first_name")]) else str(row[c("first_name")]).strip(),
                family_name="" if _is_blank(row[c("family_name")]) else str(row[c("family_name")]).strip(),
                email=None if _is_blank(email) else str(email).strip(),
                final_grade=_to_float(row[c("final_grade")]),
                orig_equiv=_to_float(row[c("orig_equiv")]),
                rounded_fg=_to_float(row[c("rounded_fg")]),
                equiv=_to_float(row[c("equiv")]),
                scores={k: _score(row[c(k)]) for k in (*LONG_EXAMS, "FE")},
                sw=_to_float(row[c("SW")]),
            )
        )

    if not students:
        raise GradebookError("No student rows found in the Main sheet.")

    return Gradebook(maxima=maxima, students=students, source=path)
