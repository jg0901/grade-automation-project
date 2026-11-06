import copy
from dataclasses import replace
from pathlib import Path

import pytest

from grade_automation.load import load_gradebook
from grade_automation.validate import SHEET_FOR_SCORE, load_exam_sheet, validate

SAMPLE = Path(__file__).resolve().parents[1] / "Calculus_Grades_AY2526_2nd_sem.xlsx"


@pytest.fixture(scope="module")
def clean():
    gb = load_gradebook(SAMPLE)
    sheets = {k: load_exam_sheet(SAMPLE, v) for k, v in SHEET_FOR_SCORE.items()}
    return gb, sheets


@pytest.fixture
def gb_sheets(clean):
    gb, sheets = clean
    return copy.deepcopy(gb), copy.deepcopy(sheets)


def checks(rep, severity=None):
    return [i.check for i in rep.issues if severity is None or i.severity == severity]


# ---- exam-sheet loader -----------------------------------------------------

def test_loader_handles_both_layouts(clean):
    _, sheets = clean
    le1, fe, sw = sheets["LE1"], sheets["FE"], sheets["SW"]
    assert le1.total_max == 30 and fe.total_max == 60 and sw.total_max == 215
    assert le1.item_max["I.1"] == 3 and fe.item_max["VII"] == 8 and sw.item_max["SW1.1"] == 10
    assert len(le1.rows) == len(fe.rows) == len(sw.rows) == 23
    assert "%" not in sw.item_max  # derived column skipped
    assert sheets["LE3"].rows["SN262-001"].total == "Excused"


# ---- clean sample ----------------------------------------------------------

def test_sample_is_clean(clean):
    gb, sheets = clean
    rep = validate(gb, sheets)
    assert rep.ok and rep.issues == []


# ---- Main checks -----------------------------------------------------------

def test_duplicate_student(gb_sheets):
    gb, sheets = gb_sheets
    gb.students[2] = replace(gb.students[2], student_number=gb.students[1].student_number)
    rep = validate(gb, sheets)
    assert "duplicate_student" in checks(rep, "error")


def test_invalid_status_and_out_of_range(gb_sheets):
    gb, sheets = gb_sheets
    s = gb.students[1]
    s.scores["LE1"] = "excuse"      # typo
    s.scores["LE2"] = 31            # > max 30
    s.scores["FE"] = -1
    rep = validate(gb, sheets)
    errs = checks(rep, "error")
    assert "invalid_status" in errs
    assert errs.count("score_out_of_range") == 2


def test_multiple_excused(gb_sheets):
    gb, sheets = gb_sheets
    s = next(x for x in gb.students if x.student_number == "SN262-001")  # LE3 already Excused
    s.scores["LE1"] = "Excused"
    rep = validate(gb, sheets)
    assert "multiple_excused" in checks(rep, "error")


def test_fe_excused_is_warning(gb_sheets):
    gb, sheets = gb_sheets
    gb.students[1].scores["FE"] = "Excused"
    gb.students[1].final_grade = None
    rep = validate(gb, sheets)
    assert "fe_not_numeric" in checks(rep, "warning")


def test_blank_final_grade_with_complete_inputs(gb_sheets):
    gb, sheets = gb_sheets
    gb.students[1].final_grade = None
    rep = validate(gb, sheets)
    assert "final_grade_blank" in checks(rep, "warning")


def test_recompute_mismatch_detects_tampered_excel_value(gb_sheets):
    gb, sheets = gb_sheets
    s = gb.students[1]
    s.final_grade = s.final_grade + 0.01   # Excel says something the policy doesn't
    rep = validate(gb, sheets)
    assert "recompute_mismatch" in checks(rep, "error")


def test_recompute_mismatch_detects_wrong_equivalent(gb_sheets):
    gb, sheets = gb_sheets
    s = gb.students[1]
    s.equiv = 1.0
    rep = validate(gb, sheets)
    msgs = [i.message for i in rep.errors if i.check == "recompute_mismatch"]
    assert any(m.startswith("equiv:") for m in msgs)


# ---- Main <-> exam sheet checks -------------------------------------------

def test_row_misalignment_is_caught_by_student_number_match(gb_sheets):
    gb, sheets = gb_sheets
    # Simulate ='LE1'!D:D pointing one row off: Main shows student 2's total for student 1.
    a, b = gb.students[1], gb.students[2]
    a.scores["LE1"], b.scores["LE1"] = b.scores["LE1"], a.scores["LE1"]
    rep = validate(gb, sheets)
    mism = [i for i in rep.errors if i.check == "main_sheet_mismatch"]
    assert {i.student_number for i in mism} == {a.student_number, b.student_number}


def test_student_missing_from_sheet_and_from_main(gb_sheets):
    gb, sheets = gb_sheets
    del sheets["LE2"].rows["SN262-005"]
    sheets["FE"].rows["SN999-999"] = replace(sheets["FE"].rows["SN262-001"], student_number="SN999-999")
    rep = validate(gb, sheets)
    errs = checks(rep, "error")
    assert "student_missing_in_sheet" in errs
    assert "student_missing_in_main" in errs


def test_item_over_max(gb_sheets):
    gb, sheets = gb_sheets
    sheets["LE1"].rows["SN262-002"].items["I.1"] = 3.5   # max 3
    rep = validate(gb, sheets)
    assert "item_out_of_range" in checks(rep, "error")


def test_total_without_items_warns(gb_sheets):
    gb, sheets = gb_sheets
    er = sheets["LE4"].rows["SN262-002"]
    for k in er.items:
        er.items[k] = None
    er.total = 0.0
    gb_student = next(s for s in gb.students if s.student_number == "SN262-002")
    gb_student.scores["LE4"] = 0.0   # keep Main consistent so only this check fires
    rep = validate(gb, sheets)
    assert "total_without_items" in checks(rep, "warning")


def test_sw_mismatch(gb_sheets):
    gb, sheets = gb_sheets
    gb.students[1].sw = 0.5
    rep = validate(gb, sheets)
    mism = [i for i in rep.errors if i.check == "main_sheet_mismatch" and "SW" in i.message]
    assert mism


def test_write_csv(gb_sheets, tmp_path):
    gb, sheets = gb_sheets
    gb.students[1].scores["LE1"] = 99
    rep = validate(gb, sheets)
    p = rep.write_csv(tmp_path / "r" / "validation.csv")
    text = p.read_text()
    assert "severity,check,sheet,row,student_number,message" in text
    assert "score_out_of_range" in text
