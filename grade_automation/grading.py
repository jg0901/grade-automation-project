"""The grading policy, as implemented in the Excel template, in Python.

    final = 0.15 * Σ(LE_i / LE_i max)   over the (up to 4) numeric long exams
          + (0.30 + 0.15 * excused) * FE / FE max
          + 0.10 * SW                     (SW is already a fraction of 1)

    rounded = ROUND(final, 2)   -- Excel ROUND is half-away-from-zero
    equiv   = thresholds applied to `rounded`

Rules the Excel formula enforces (grade is blank otherwise):
  * exactly 4 long-exam cells are numeric-or-"Excused", at most 1 Excused
  * FE and SW are numeric

Rounding note: Python's round() is banker's rounding on binary floats
(round(0.845, 2) == 0.84); Excel gives 0.85. We use Decimal + ROUND_HALF_UP
on the repr of the float, which matches Excel for the values that occur here.
"""

from __future__ import annotations

from decimal import ROUND_HALF_UP, Decimal
from typing import Any

from .load import EXCUSED, LONG_EXAMS, Maxima

# (lower bound inclusive on the ROUNDED final grade, equivalent)
EQUIVALENT_THRESHOLDS: tuple[tuple[float, float], ...] = (
    (0.93, 1.00),
    (0.90, 1.25),
    (0.87, 1.50),
    (0.84, 1.75),
    (0.80, 2.00),
    (0.75, 2.25),
    (0.70, 2.50),
    (0.65, 2.75),
    (0.60, 3.00),
    (0.55, 4.00),
    (0.00, 5.00),
)

LE_WEIGHT = 0.15
FE_BASE_WEIGHT = 0.30
SW_WEIGHT = 0.10


def round_half_up(x: float, places: int = 2) -> float:
    """Excel-style ROUND(x, places)."""
    q = Decimal(1).scaleb(-places)  # 10**-places
    return float(Decimal(repr(x)).quantize(q, rounding=ROUND_HALF_UP))


def equivalent(rounded_final: float | None) -> float | None:
    """Map a (2-dp rounded) final grade fraction to the 1.00–5.00 scale."""
    if rounded_final is None:
        return None
    for lower, equiv in EQUIVALENT_THRESHOLDS:
        if rounded_final >= lower:
            return equiv
    return None  # negative input: not a valid grade


def _is_excused(v: Any) -> bool:
    return isinstance(v, str) and v.strip().lower() == EXCUSED.lower()


def compute_final_grade(scores: dict[str, Any], sw: float | None, maxima: Maxima) -> float | None:
    """Return the final grade as a fraction, or None if inputs are incomplete
    (mirrors the `valid` guard in the Excel LET formula)."""
    le_vals = [scores.get(e) for e in LONG_EXAMS]
    excused = sum(1 for v in le_vals if _is_excused(v))
    numeric = [v for v in le_vals if isinstance(v, (int, float)) and not isinstance(v, bool)]
    fe = scores.get("FE")

    if len(numeric) + excused != 4 or excused > 1:
        return None
    if not isinstance(fe, (int, float)) or isinstance(fe, bool):
        return None
    if sw is None:
        return None

    le_part = LE_WEIGHT * sum(
        float(v) / maxima.for_exam(e)
        for e, v in zip(LONG_EXAMS, le_vals)
        if isinstance(v, (int, float)) and not isinstance(v, bool)
    )
    fe_part = (FE_BASE_WEIGHT + LE_WEIGHT * excused) * float(fe) / maxima.FE
    sw_part = SW_WEIGHT * float(sw) / (maxima.SW or 1.0)
    return le_part + fe_part + sw_part


def grade_summary(scores: dict[str, Any], sw: float | None, maxima: Maxima) -> dict[str, float | None]:
    """final, rounded, orig_equiv (from unrounded), equiv (from rounded)."""
    final = compute_final_grade(scores, sw, maxima)
    if final is None:
        return {"final_grade": None, "rounded_fg": None, "orig_equiv": None, "equiv": None}
    rounded = round_half_up(final, 2)
    return {
        "final_grade": final,
        "rounded_fg": rounded,
        "orig_equiv": equivalent(final),
        "equiv": equivalent(rounded),
    }
