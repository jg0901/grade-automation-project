"""Runtime settings, loaded from environment variables / a .env file.

Nothing secret lives in code. Copy .env.example to .env and fill it in.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path

from dotenv import load_dotenv


def _env_bool(name: str, default: bool) -> bool:
    raw = os.getenv(name)
    if raw is None or raw.strip() == "":
        return default
    return raw.strip().lower() in {"1", "true", "yes", "y", "on"}


@dataclass(frozen=True)
class Settings:
    # Input
    grades_file: Path
    grades_sheet: str

    # Sending
    sender_email: str
    app_password: str
    sender_name: str
    test_mode: bool
    test_email: str
    smtp_host: str
    smtp_port: int

    # Email content
    course_name: str
    email_subject: str
    clarification_deadline: str   # e.g. "June 8, 2026"
    exam_review_schedule: str     # e.g. "June 8, 2026 (Monday), from 11:00 AM to 2:00 PM"
    crs_upload_date: str          # e.g. "June 9, 2026"

    # Which numbers go into the emails: "python" (policy in grading.py, the
    # default) or "excel" (the workbook's cached Final Grade / Equiv columns).
    grade_source: str = "python"

    # Output
    output_dir: Path = field(default_factory=lambda: Path("output"))

    @property
    def preview_dir(self) -> Path:
        return self.output_dir / "previews"

    @property
    def sent_log(self) -> Path:
        return self.output_dir / "sent_log.csv"

    def validate_for_sending(self) -> list[str]:
        """Return a list of problems that must be fixed before a real send."""
        problems = []
        if not self.sender_email or "@" not in self.sender_email or "example.com" in self.sender_email:
            problems.append("SENDER_EMAIL is not set to a real address")
        if not self.app_password or self.app_password.startswith("xxxx"):
            problems.append("APP_PASSWORD is not set")
        if self.test_mode and ("@" not in self.test_email or "example.com" in self.test_email):
            problems.append("TEST_MODE is on but TEST_EMAIL is not a real address")
        if "[insert" in self.sender_name.lower() or not self.sender_name.strip():
            problems.append("SENDER_NAME is not set")
        if self.grade_source not in {"python", "excel"}:
            problems.append(f"GRADE_SOURCE must be 'python' or 'excel', got {self.grade_source!r}")
        return problems


def load_settings(env_file: str | os.PathLike | None = None) -> Settings:
    """Load settings from `env_file` (default: .env in the current directory)."""
    load_dotenv(env_file, override=False)

    return Settings(
        grades_file=Path(os.getenv("GRADES_FILE", "Calculus_Grades_AY2526_2nd_sem.xlsx")),
        grades_sheet=os.getenv("GRADES_SHEET", "Main"),
        sender_email=os.getenv("SENDER_EMAIL", ""),
        app_password=os.getenv("APP_PASSWORD", ""),
        sender_name=os.getenv("SENDER_NAME", ""),
        test_mode=_env_bool("TEST_MODE", True),
        test_email=os.getenv("TEST_EMAIL", ""),
        smtp_host=os.getenv("SMTP_HOST", "smtp.gmail.com"),
        smtp_port=int(os.getenv("SMTP_PORT", "465")),
        course_name=os.getenv("COURSE_NAME", "Calculus"),
        email_subject=os.getenv("EMAIL_SUBJECT", "Calculus Final Grade"),
        clarification_deadline=os.getenv("CLARIFICATION_DEADLINE", "[date]"),
        exam_review_schedule=os.getenv("EXAM_REVIEW_SCHEDULE", "[schedule]"),
        crs_upload_date=os.getenv("CRS_UPLOAD_DATE", "[date]"),
        grade_source=os.getenv("GRADE_SOURCE", "python").strip().lower(),
        output_dir=Path(os.getenv("OUTPUT_DIR", "output")),
    )
