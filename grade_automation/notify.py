"""Render and send per-student grade emails.

Safety model:
  * Rendering is pure (no network) so it can be previewed and unit-tested.
  * Sending requires an explicit opt-in from the CLI and refuses to run if
    settings are incomplete.
  * Every send is appended to a CSV log keyed on student number; students
    already logged as sent are skipped unless `resend=True`.
  * TEST_MODE routes every message to TEST_EMAIL instead of the student.
"""

from __future__ import annotations

import csv
import html
import smtplib
import uuid
from dataclasses import dataclass
from datetime import datetime, timezone
from email.message import EmailMessage
from pathlib import Path
from typing import Any, Callable, Iterable

from .config import Settings
from .load import LONG_EXAMS, Gradebook, Maxima, StudentRecord

# --------------------------------------------------------------------------
# formatting
# --------------------------------------------------------------------------

def fmt_percent(fraction: float | None, dp: int = 2) -> str:
    """0.8376 -> '83.76%'. All grade columns are fractions of 1."""
    if fraction is None:
        return ""
    return f"{fraction * 100:.{dp}f}%"


def fmt_equiv(x: float | None) -> str:
    return "" if x is None else f"{x:.2f}"


def fmt_score(v: Any) -> str:
    """Numeric score -> trimmed number; status strings verbatim; blank -> '—'."""
    if v is None:
        return "—"
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else f"{v:g}"
    return html.escape(str(v))


def _num(v: float) -> str:
    return str(int(v)) if float(v).is_integer() else f"{v:g}"


# --------------------------------------------------------------------------
# rendering
# --------------------------------------------------------------------------

@dataclass
class RenderedEmail:
    student_number: str
    student_email: str
    recipient: str
    subject: str
    text: str
    html: str


def formula_note(excused_exam: str | None, maxima: Maxima) -> str:
    """Human-readable version of the grading formula actually applied."""
    included = [e for e in LONG_EXAMS if e != excused_exam]
    le_terms = " + ".join(f"({e}÷{_num(maxima.for_exam(e))})×0.15" for e in included)
    fe_weight = 0.45 if excused_exam else 0.30
    return (
        f"Final Grade(%) = [{le_terms} + (SW/HW)×0.10 "
        f"+ (FE÷{_num(maxima.FE)})×{fe_weight:.2f}]×100"
    )


def render_email(student: StudentRecord, maxima: Maxima, settings: Settings) -> RenderedEmail:
    reason = student.skip_reason()
    if reason:
        raise ValueError(f"{student.student_number}: cannot render ({reason})")

    first_name = html.escape(student.first_name or "Student")
    excused = student.excused_exam
    fe_weight_text = "Final Exam (45%)" if excused else "Final Exam (30%)"

    # Rounding to 2 dp can only ever improve the equivalent (thresholds are at
    # 2 dp), so show the rounding step only when it changed the result.
    rounding_helped = (
        student.rounded_fg is not None
        and student.equiv is not None
        and student.orig_equiv is not None
        and student.equiv < student.orig_equiv
    )
    if rounding_helped:
        grade_section = f"""
<ul>
  <li>Final Grade: <strong>{fmt_percent(student.final_grade)}</strong> rounded to <strong>{fmt_percent(student.rounded_fg)}</strong></li>
  <li>Original Equivalent: <strong>{fmt_equiv(student.orig_equiv)}</strong></li>
  <li><strong>Final Equivalent: {fmt_equiv(student.equiv)}</strong></li>
</ul>"""
    else:
        grade_section = f"""
<ul>
  <li>Final Grade: <strong>{fmt_percent(student.final_grade)}</strong></li>
  <li>Equivalent: <strong>{fmt_equiv(student.equiv)}</strong></li>
</ul>"""

    le_items = "\n".join(
        f"  <li>{e}: <strong>{fmt_score(student.scores.get(e))}</strong> out of {_num(maxima.for_exam(e))}</li>"
        for e in LONG_EXAMS
    )

    sender_name = html.escape(settings.sender_name)
    course = html.escape(settings.course_name)

    html_body = f"""<html>
<body>
<p>Good day, {first_name},</p>

<p>I hope you are doing well.</p>

<p>Your {course} <strong>final grade</strong> is:</p>
{grade_section}

<p>Below is a breakdown of your final grade:</p>

<p><strong>Long Exams (15% each):</strong></p>
<ul>
{le_items}
</ul>

<p><strong>{fe_weight_text}:</strong> <strong>{fmt_score(student.scores.get("FE"))}</strong> out of {_num(maxima.FE)}</p>

<p><strong>HW and SW (10%):</strong> <strong>{fmt_percent(student.sw)}</strong></p>

<p><strong>Note:</strong><br>{html.escape(formula_note(excused, maxima))}</p>

<p>If you have any questions or clarifications regarding your grade or scores,
please email me by {html.escape(settings.clarification_deadline)}.</p>

<p>If you would like to review your final exam paper, you can visit me at
<strong>my office on {html.escape(settings.exam_review_schedule)}</strong>.
I will upload your grade in CRS on {html.escape(settings.crs_upload_date)}.
Once uploaded, you cannot request for any changes.</p>

<p>Please take note that, as per the institute's policy, you cannot receive a copy
of your final exam papers. You may only view it for the purpose of checking for corrections.</p>

<p>I wish you all the best in your future subjects. Thank you very much!</p>

<p>Best regards,<br>{sender_name}</p>
</body>
</html>
"""

    text_body = (
        f"Good day, {student.first_name or 'Student'},\n\n"
        f"Your {settings.course_name} final grade is {fmt_percent(student.final_grade)} "
        f"(equivalent {fmt_equiv(student.equiv)}).\n\n"
        "Please view this email in HTML format for the full breakdown.\n\n"
        f"Best regards,\n{settings.sender_name}\n"
    )

    recipient = settings.test_email if settings.test_mode else student.email
    return RenderedEmail(
        student_number=student.student_number,
        student_email=student.email,
        recipient=recipient,
        subject=settings.email_subject,
        text=text_body,
        html=html_body,
    )


def write_previews(emails: Iterable[RenderedEmail], preview_dir: Path) -> list[Path]:
    preview_dir.mkdir(parents=True, exist_ok=True)
    paths = []
    for e in emails:
        p = preview_dir / f"{e.student_number}.html"
        p.write_text(e.html, encoding="utf-8")
        paths.append(p)
    return paths


# --------------------------------------------------------------------------
# send log / duplicate protection
# --------------------------------------------------------------------------

LOG_FIELDS = ("run_id", "timestamp_utc", "student_number", "student_email",
              "recipient", "test_mode", "status", "detail")


def already_sent(log_path: Path) -> set[str]:
    """Student numbers with a status=sent row in the log (real sends only)."""
    if not log_path.exists():
        return set()
    with log_path.open(newline="", encoding="utf-8") as f:
        return {
            r["student_number"]
            for r in csv.DictReader(f)
            if r.get("status") == "sent" and r.get("test_mode") == "False"
        }


def _append_log(log_path: Path, row: dict[str, Any]) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    new = not log_path.exists()
    with log_path.open("a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=LOG_FIELDS)
        if new:
            w.writeheader()
        w.writerow(row)


# --------------------------------------------------------------------------
# sending
# --------------------------------------------------------------------------

@dataclass
class SendSummary:
    run_id: str
    sent: int = 0
    skipped_duplicate: int = 0
    failed: int = 0


def build_message(e: RenderedEmail, sender_email: str) -> EmailMessage:
    msg = EmailMessage()
    msg["Subject"] = e.subject
    msg["From"] = sender_email
    msg["To"] = e.recipient
    msg.set_content(e.text)
    msg.add_alternative(e.html, subtype="html")
    return msg


def send_emails(
    emails: list[RenderedEmail],
    settings: Settings,
    *,
    resend: bool = False,
    smtp_factory: Callable[[], smtplib.SMTP] | None = None,
    log: Callable[[str], None] = print,
) -> SendSummary:
    """Send rendered emails over SMTP with duplicate protection and logging.

    `smtp_factory` is injectable for tests; default is Gmail SMTP over SSL.
    """
    problems = settings.validate_for_sending()
    if problems:
        raise RuntimeError("Refusing to send; fix settings first: " + "; ".join(problems))

    run_id = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ") + "-" + uuid.uuid4().hex[:6]
    summary = SendSummary(run_id=run_id)
    done = set() if (resend or settings.test_mode) else already_sent(settings.sent_log)

    def _row(e: RenderedEmail, status: str, detail: str = "") -> dict[str, Any]:
        return {
            "run_id": run_id,
            "timestamp_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
            "student_number": e.student_number,
            "student_email": e.student_email,
            "recipient": e.recipient,
            "test_mode": str(settings.test_mode),
            "status": status,
            "detail": detail,
        }

    if smtp_factory is None:
        smtp_factory = lambda: smtplib.SMTP_SSL(settings.smtp_host, settings.smtp_port)  # noqa: E731

    with smtp_factory() as smtp:
        smtp.login(settings.sender_email, settings.app_password)
        for e in emails:
            if e.student_number in done:
                summary.skipped_duplicate += 1
                _append_log(settings.sent_log, _row(e, "skipped", "already sent in a previous run"))
                log(f"SKIP (already sent): {e.student_number}")
                continue
            try:
                smtp.send_message(build_message(e, settings.sender_email))
            except Exception as exc:  # log and continue with the rest
                summary.failed += 1
                _append_log(settings.sent_log, _row(e, "failed", repr(exc)))
                log(f"FAIL: {e.student_number} -> {e.recipient}: {exc}")
                continue
            summary.sent += 1
            _append_log(settings.sent_log, _row(e, "sent"))
            mode = " [TEST MODE]" if settings.test_mode else ""
            log(f"Sent{mode}: {e.student_number} -> {e.recipient}")

    return summary


def render_all(gradebook: Gradebook, settings: Settings) -> list[RenderedEmail]:
    return [render_email(s, gradebook.maxima, settings) for s in gradebook.sendable()]
