from dataclasses import replace
from pathlib import Path

import openpyxl
import pytest

from grade_automation.compute import (
    compute_grades,
    export_rows,
    graded_gradebook,
    summary,
    write_csv,
    write_xlsx,
)
from grade_automation.load import load_gradebook
from grade_automation.notify import render_email
from tests.test_notify import make_settings

SAMPLE = Path(__file__).resolve().parents[1] / "Calculus_Grades_AY2526_2nd_sem.xlsx"


@pytest.fixture(scope="module")
def gb():
    return load_gradebook(SAMPLE)


@pytest.fixture(scope="module")
def computed(gb):
    return compute_grades(gb)


def test_counts_and_no_discrepancies_on_sample(computed):
    s = summary(computed)
    assert s["students"] == 23 and s["graded"] == 22 and s["ungraded"] == 1
    assert s["discrepancies"] == 0
    assert dict(s["distribution"])[1.25] == 3


def test_ungraded_reason_is_specific(computed):
    c = next(c for c in computed if c.student_number == "SN262-006")
    assert not c.graded
    assert c.reason.startswith("no grade:") and "LE4=LOA" in c.reason and "FE=LOA" in c.reason


def test_ungraded_rounded_fg_is_blank_in_fixed_template(computed):
    # Template now uses =IF(E="","",ROUND(E,2)); both sides blank for an ungraded student.
    c = next(c for c in computed if c.student_number == "SN262-006")
    assert c.student.rounded_fg is None and c.rounded_fg is None
    assert c.discrepancies() == []


def test_old_template_rounded_fg_zero_quirk_not_flagged(gb):
    # Workbooks built from the old template have =ROUND(E,2), which yields 0 for a
    # blank grade. That must not be reported as a discrepancy.
    s = next(s for s in gb.students if s.student_number == "SN262-006")
    old_style = replace(gb, students=[replace(s, rounded_fg=0.0)])
    c = compute_grades(old_style)[0]
    assert c.student.rounded_fg == 0.0 and c.rounded_fg is None
    assert c.discrepancies() == []


def test_discrepancy_detected_when_excel_differs(gb):
    s = gb.students[1]
    tampered = replace(gb, students=[replace(s, equiv=1.0)])
    c = compute_grades(tampered)[0]
    fields = [f for f, _, _ in c.discrepancies()]
    assert fields == ["equiv"]


def test_graded_gradebook_carries_python_values(gb, computed):
    g = graded_gradebook(gb, computed)
    assert len(g.students) == len(gb.students)
    for raw, new in zip(gb.students, g.students):
        assert new.student_number == raw.student_number
        assert new.scores == raw.scores
    ungraded = next(s for s in g.students if s.student_number == "SN262-006")
    assert ungraded.rounded_fg is None  # Python value, not Excel's 0


def test_email_identical_from_either_source(gb, computed, tmp_path):
    settings = make_settings(tmp_path)
    g = graded_gradebook(gb, computed)
    for raw, new in zip(gb.students, g.students):
        if raw.final_grade is None:
            continue
        raw_e = render_email(replace(raw, email="x@example.edu"), gb.maxima, settings)
        new_e = render_email(replace(new, email="x@example.edu"), gb.maxima, settings)
        assert raw_e.html == new_e.html, raw.student_number


def test_export_rows_columns(computed):
    rows = export_rows(computed)
    assert len(rows) == 23
    r = next(r for r in rows if r["Student Number"] == "SN262-001")
    assert r["Excused"] == "LE3" and r["Equiv"] == 4 and r["Matches Excel"] == "yes"
    r6 = next(r for r in rows if r["Student Number"] == "SN262-006")
    assert r6["Final Grade"] == "" and r6["Note"].startswith("no grade")


def test_write_csv_and_xlsx(computed, tmp_path):
    csv_path = write_csv(computed, tmp_path / "o" / "final_grades.csv")
    assert csv_path.exists() and csv_path.read_text().count("\n") == 24  # header + 23

    xlsx_path = write_xlsx(computed, tmp_path / "o" / "final_grades.xlsx")
    wb = openpyxl.load_workbook(xlsx_path)
    assert wb.sheetnames == ["Final Grades", "Discrepancies", "Summary"]
    ws = wb["Final Grades"]
    assert ws.max_row == 24
    assert ws["A1"].value == "Student Number"
    assert wb["Discrepancies"]["A2"].value == "(none)"
