from dataclasses import replace
from pathlib import Path

import pytest

from grade_automation.config import Settings
from grade_automation.load import load_gradebook
from grade_automation.notify import (
    already_sent,
    fmt_percent,
    formula_note,
    render_email,
    send_emails,
)

SAMPLE = Path(__file__).resolve().parents[1] / "Calculus_Grades_AY2526_2nd_sem.xlsx"


def make_settings(tmp_path: Path, **overrides) -> Settings:
    base = dict(
        grades_file=SAMPLE,
        grades_sheet="Main",
        sender_email="teacher@school.edu",
        app_password="secret",
        sender_name="Teacher",
        test_mode=False,
        test_email="",
        smtp_host="localhost",
        smtp_port=0,
        course_name="Calculus",
        email_subject="Calculus Final Grade",
        clarification_deadline="June 8, 2026",
        exam_review_schedule="June 8, 2026 (Monday), 11:00 AM to 2:00 PM",
        crs_upload_date="June 9, 2026",
        output_dir=tmp_path / "out",
    )
    base.update(overrides)
    return Settings(**base)


@pytest.fixture(scope="module")
def gb():
    return load_gradebook(SAMPLE, "Main")


def with_email(student, email="student@example.edu"):
    return replace(student, email=email)


def test_fmt_percent_treats_values_as_fractions():
    assert fmt_percent(0.556483) == "55.65%"
    assert fmt_percent(0.56) == "56.00%"
    assert fmt_percent(1.0) == "100.00%"   # old heuristic misread 1.0
    assert fmt_percent(None) == ""


def test_formula_note_uses_sheet_maxima_and_excused_weight(gb):
    normal = formula_note(None, gb.maxima)
    assert "(LE1÷30)×0.15" in normal and "(FE÷60)×0.30" in normal
    excused = formula_note("LE3", gb.maxima)
    assert "LE3" not in excused and "(FE÷60)×0.45" in excused


def test_render_excused_student(gb, tmp_path):
    s = with_email(next(x for x in gb.students if x.student_number == "SN262-001"))
    e = render_email(s, gb.maxima, make_settings(tmp_path))
    assert e.recipient == "student@example.edu"
    assert "Final Exam (45%)" in e.html
    assert "LE3: <strong>Excused</strong>" in e.html
    assert "55.65%" in e.html and "4.00" in e.html


def test_render_shows_rounding_only_when_it_helped(gb, tmp_path):
    by_sn = {x.student_number: x for x in gb.students}
    # SN262-008: 0.6464 -> rounds to 0.65 -> equiv 2.75 vs orig 3.00
    helped = render_email(with_email(by_sn["SN262-008"]), gb.maxima, make_settings(tmp_path))
    assert "rounded to" in helped.html and "Original Equivalent" in helped.html
    # SN262-002: 0.6867 -> 0.69 -> equiv unchanged 2.75
    same = render_email(with_email(by_sn["SN262-002"]), gb.maxima, make_settings(tmp_path))
    assert "rounded to" not in same.html


def test_test_mode_redirects_recipient(gb, tmp_path):
    s = with_email(gb.students[1])
    e = render_email(s, gb.maxima, make_settings(tmp_path, test_mode=True, test_email="me@school.edu"))
    assert e.recipient == "me@school.edu"
    assert e.student_email == "student@example.edu"


def test_render_refuses_student_without_final_grade(gb, tmp_path):
    s = with_email(next(x for x in gb.students if x.student_number == "SN262-006"))
    with pytest.raises(ValueError):
        render_email(s, gb.maxima, make_settings(tmp_path))


class FakeSMTP:
    """Records messages instead of sending them."""
    sent: list = []

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def login(self, *a):
        pass

    def send_message(self, msg):
        FakeSMTP.sent.append(msg)


def test_send_logs_and_blocks_duplicates(gb, tmp_path):
    settings = make_settings(tmp_path)
    students = [with_email(x, f"{x.student_number}@example.edu") for x in gb.students[1:4]]
    emails = [render_email(s, gb.maxima, settings) for s in students]

    FakeSMTP.sent = []
    first = send_emails(emails, settings, smtp_factory=FakeSMTP, log=lambda m: None)
    assert first.sent == 3 and first.skipped_duplicate == 0
    assert len(FakeSMTP.sent) == 3
    assert already_sent(settings.sent_log) == {s.student_number for s in students}

    # Re-running the same batch sends nothing new.
    FakeSMTP.sent = []
    second = send_emails(emails, settings, smtp_factory=FakeSMTP, log=lambda m: None)
    assert second.sent == 0 and second.skipped_duplicate == 3
    assert FakeSMTP.sent == []

    # --resend overrides the log.
    third = send_emails(emails, settings, resend=True, smtp_factory=FakeSMTP, log=lambda m: None)
    assert third.sent == 3


def test_send_refuses_incomplete_settings(gb, tmp_path):
    settings = make_settings(tmp_path, app_password="xxxx xxxx xxxx xxxx")
    emails = [render_email(with_email(gb.students[1]), gb.maxima, settings)]
    with pytest.raises(RuntimeError):
        send_emails(emails, settings, smtp_factory=FakeSMTP, log=lambda m: None)
