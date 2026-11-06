"""End-to-end CLI tests on a values-only copy of the sample with emails filled in."""

import io
from pathlib import Path

import pandas as pd
import pytest

from grade_automation import __main__ as cli
from grade_automation import notify

SAMPLE = Path(__file__).resolve().parents[1] / "Calculus_Grades_AY2526_2nd_sem.xlsx"


@pytest.fixture(scope="module")
def workbook_with_emails(tmp_path_factory):
    """Values-only copy of every sheet (formulas resolved), with emails filled in."""
    d = tmp_path_factory.mktemp("wb")
    x = pd.ExcelFile(SAMPLE)
    sheets = {n: pd.read_excel(x, sheet_name=n, header=None) for n in x.sheet_names}
    main = sheets["Main"]
    main[0] = main[0].astype(object)
    for i in range(2, len(main)):  # row index 2 == Excel row 3 == first student
        sn = main.iat[i, 1]
        if isinstance(sn, str):
            main.iat[i, 0] = sn.lower() + "@example.edu"
    p = d / "grades.xlsx"
    with pd.ExcelWriter(p) as w:
        for n, df in sheets.items():
            df.to_excel(w, sheet_name=n, header=False, index=False)
    return p


def write_env(tmp_path, workbook, **over):
    vals = {
        "GRADES_FILE": str(workbook), "OUTPUT_DIR": str(tmp_path / "out"),
        "SENDER_EMAIL": "teacher@school.edu", "APP_PASSWORD": "secret", "SENDER_NAME": "Teacher",
        "TEST_MODE": "false", "TEST_EMAIL": "teacher@school.edu",
    }
    vals.update(over)
    p = tmp_path / "t.env"
    p.write_text("\n".join(f"{k}={v}" for k, v in vals.items()))
    return str(p)


class FakeSMTP:
    sent = []

    def __init__(self, *a, **k):
        pass

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def login(self, *a):
        pass

    def send_message(self, msg):
        FakeSMTP.sent.append(msg)


@pytest.fixture
def fake_smtp(monkeypatch):
    FakeSMTP.sent = []
    monkeypatch.setattr(notify.smtplib, "SMTP_SSL", FakeSMTP)
    return FakeSMTP


@pytest.fixture(autouse=True)
def isolate_env(monkeypatch):
    # .env in the repo (if any) must not leak into tests; dotenv doesn't override set vars.
    for k in ("GRADES_FILE", "OUTPUT_DIR", "SENDER_EMAIL", "APP_PASSWORD", "SENDER_NAME",
              "TEST_MODE", "TEST_EMAIL", "GRADE_SOURCE"):
        monkeypatch.delenv(k, raising=False)


def test_dry_run_sends_nothing(tmp_path, workbook_with_emails, fake_smtp, capsys):
    env = write_env(tmp_path, workbook_with_emails)
    assert cli.main(["--env", env, "send"]) == 0
    out = capsys.readouterr().out
    assert "DRY RUN: would send 22 email(s)" in out
    assert fake_smtp.sent == []


def test_live_send_refuses_without_tty(tmp_path, workbook_with_emails, fake_smtp, monkeypatch):
    env = write_env(tmp_path, workbook_with_emails)
    monkeypatch.setattr(cli.sys, "stdin", io.StringIO(""))   # not a tty
    assert cli.main(["--env", env, "send", "--confirm"]) == 6
    assert fake_smtp.sent == []


class TtyInput(io.StringIO):
    def isatty(self):
        return True


def test_live_send_aborts_on_wrong_phrase(tmp_path, workbook_with_emails, fake_smtp, monkeypatch, capsys):
    env = write_env(tmp_path, workbook_with_emails)
    monkeypatch.setattr(cli.sys, "stdin", TtyInput("yes\n"))
    monkeypatch.setattr("builtins.input", lambda prompt="": "yes")
    assert cli.main(["--env", env, "send", "--confirm"]) == 0
    assert "Aborted" in capsys.readouterr().out
    assert fake_smtp.sent == []


def test_live_send_proceeds_with_exact_phrase_then_blocks_duplicates(
    tmp_path, workbook_with_emails, fake_smtp, monkeypatch, capsys
):
    env = write_env(tmp_path, workbook_with_emails)
    monkeypatch.setattr(cli.sys, "stdin", TtyInput(""))
    monkeypatch.setattr("builtins.input", lambda prompt="": "SEND 22")
    assert cli.main(["--env", env, "send", "--confirm"]) == 0
    assert len(fake_smtp.sent) == 22
    assert fake_smtp.sent[0]["To"] == "sn262-001@example.edu"

    # Second run: everyone is already in the log -> prompt shows 0, nothing sent.
    fake_smtp.sent.clear()
    monkeypatch.setattr("builtins.input", lambda prompt="": "SEND 0")
    assert cli.main(["--env", env, "send", "--confirm"]) == 0
    out = capsys.readouterr().out
    assert "Already sent (skip) : 22" in out
    assert fake_smtp.sent == []


def test_test_mode_needs_no_approval(tmp_path, workbook_with_emails, fake_smtp, monkeypatch):
    env = write_env(tmp_path, workbook_with_emails, TEST_MODE="true")
    monkeypatch.setattr("builtins.input", lambda prompt="": pytest.fail("should not prompt in TEST_MODE"))
    assert cli.main(["--env", env, "send", "--confirm"]) == 0
    assert len(fake_smtp.sent) == 22
    assert all(m["To"] == "teacher@school.edu" for m in fake_smtp.sent)


def test_validate_and_compute_commands(tmp_path, workbook_with_emails, capsys):
    env = write_env(tmp_path, workbook_with_emails)
    assert cli.main(["--env", env, "validate"]) == 0
    assert "0 error(s), 0 warning(s)" in capsys.readouterr().out
    assert cli.main(["--env", env, "compute"]) == 0
    assert (tmp_path / "out" / "final_grades.xlsx").exists()
