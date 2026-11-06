from pathlib import Path

import pytest

from grade_automation.grading import compute_final_grade, equivalent, grade_summary, round_half_up
from grade_automation.load import Maxima, load_gradebook

SAMPLE = Path(__file__).resolve().parents[1] / "Calculus_Grades_AY2526_2nd_sem.xlsx"
MAX = Maxima(LE1=30, LE2=30, LE3=30, LE4=30, FE=60, SW=1)


def test_round_half_up_matches_excel_not_python():
    assert round(0.845, 2) == 0.84            # Python banker's / binary float
    assert round_half_up(0.845, 2) == 0.85    # Excel ROUND
    assert round_half_up(0.925, 2) == 0.93
    assert round_half_up(0.9349, 2) == 0.93
    assert round_half_up(0.6464, 2) == 0.65


@pytest.mark.parametrize("fg,eq", [
    (0.93, 1.0), (0.9299, 1.25), (0.90, 1.25), (0.87, 1.5), (0.84, 1.75), (0.80, 2.0),
    (0.75, 2.25), (0.70, 2.5), (0.65, 2.75), (0.60, 3.0), (0.55, 4.0), (0.5499, 5.0), (0.0, 5.0),
])
def test_equivalent_thresholds(fg, eq):
    assert equivalent(fg) == eq


def test_equivalent_none_passthrough():
    assert equivalent(None) is None


def test_compute_normal():
    scores = {"LE1": 30, "LE2": 30, "LE3": 30, "LE4": 30, "FE": 60}
    assert compute_final_grade(scores, 1.0, MAX) == pytest.approx(1.0)
    scores = {"LE1": 15, "LE2": 15, "LE3": 15, "LE4": 15, "FE": 30}
    assert compute_final_grade(scores, 0.5, MAX) == pytest.approx(0.5)


def test_compute_excused_shifts_weight_to_fe():
    scores = {"LE1": 30, "LE2": 30, "LE3": "Excused", "LE4": 30, "FE": 60}
    # 0.15*3 + 0.45*1 + 0.10*1 = 1.0
    assert compute_final_grade(scores, 1.0, MAX) == pytest.approx(1.0)
    scores = {"LE1": 30, "LE2": 30, "LE3": "Excused", "LE4": 30, "FE": 0}
    assert compute_final_grade(scores, 1.0, MAX) == pytest.approx(0.55)


@pytest.mark.parametrize("scores,sw", [
    ({"LE1": 30, "LE2": "Absent", "LE3": 30, "LE4": 30, "FE": 60}, 1.0),      # Absent
    ({"LE1": 30, "LE2": "Excused", "LE3": "Excused", "LE4": 30, "FE": 60}, 1.0),  # 2 excused
    ({"LE1": 30, "LE2": 30, "LE3": 30, "LE4": 30, "FE": "LOA"}, 1.0),         # FE not numeric
    ({"LE1": 30, "LE2": 30, "LE3": 30, "LE4": 30, "FE": 60}, None),           # SW blank
    ({"LE1": 30, "LE2": 30, "LE3": 30, "LE4": None, "FE": 60}, 1.0),          # LE blank
])
def test_compute_incomplete_returns_none(scores, sw):
    assert compute_final_grade(scores, sw, MAX) is None


def test_recompute_matches_excel_for_every_sample_student():
    gb = load_gradebook(SAMPLE)
    graded = 0
    for s in gb.students:
        py = grade_summary(s.scores, s.sw, gb.maxima)
        if s.final_grade is None:
            assert py["final_grade"] is None, s.student_number
            continue
        graded += 1
        assert py["final_grade"] == pytest.approx(s.final_grade, abs=1e-9), s.student_number
        assert py["rounded_fg"] == pytest.approx(s.rounded_fg, abs=1e-9), s.student_number
        assert py["orig_equiv"] == s.orig_equiv, s.student_number
        assert py["equiv"] == s.equiv, s.student_number
    assert graded == 22
