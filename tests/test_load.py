from pathlib import Path

import pytest

from grade_automation.load import GradebookError, load_gradebook

SAMPLE = Path(__file__).resolve().parents[1] / "Calculus_Grades_AY2526_2nd_sem.xlsx"


@pytest.fixture(scope="module")
def gb():
    return load_gradebook(SAMPLE, "Main")


def test_maxima_from_row_2(gb):
    assert (gb.maxima.LE1, gb.maxima.LE2, gb.maxima.LE3, gb.maxima.LE4) == (30, 30, 30, 30)
    assert gb.maxima.FE == 60
    assert gb.maxima.SW == 1


def test_student_count_excludes_maxima_and_blank_rows(gb):
    assert len(gb.students) == 23
    assert all(s.student_number.startswith("SN262-") for s in gb.students)
    assert gb.students[0].row_number == 3  # first student is Excel row 3


def test_sw_column_resolves_despite_notebook_alias(gb):
    # The notebook crashed with KeyError('HW/SW'); the header is actually 'SW'.
    s = gb.students[0]
    assert s.sw == pytest.approx(0.602326, abs=1e-6)


def test_excused_detection(gb):
    by_sn = {s.student_number: s for s in gb.students}
    assert by_sn["SN262-001"].excused_exam == "LE3"
    assert by_sn["SN262-013"].excused_exam == "LE2"
    assert by_sn["SN262-002"].excused_exam is None


def test_absent_loa_student_is_skipped_not_crashed(gb):
    by_sn = {s.student_number: s for s in gb.students}
    s = by_sn["SN262-006"]
    assert s.final_grade is None
    assert s.scores["LE1"] == "Absent" and s.scores["LE4"] == "LOA"
    assert s.skip_reason() is not None


def test_sample_has_no_emails_so_everyone_is_skipped(gb):
    # The anonymized sample has no email addresses; the old notebook would
    # silently send nothing. Now every student has an explicit skip reason.
    assert gb.sendable() == []
    reasons = {r for _, r in gb.skipped()}
    assert "no email address" in reasons


def test_missing_file_raises():
    with pytest.raises(GradebookError):
        load_gradebook("does-not-exist.xlsx")
