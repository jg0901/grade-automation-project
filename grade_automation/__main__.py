"""Command-line entry point.

    python -m grade_automation validate           # check the workbook, write a CSV report
    python -m grade_automation preview            # render HTML previews, send nothing
    python -m grade_automation send --confirm     # actually send (TEST_MODE still honoured)
    python -m grade_automation send --confirm --resend   # ignore the sent log

Dry run is the default everywhere. `send` without `--confirm` only prints
what it would do, and `send` refuses to run while validation reports errors.
"""

from __future__ import annotations

import argparse
import smtplib
import sys

from .config import load_settings
from .load import GradebookError, load_gradebook
from .notify import render_all, send_emails, write_previews
from .validate import format_report, validate


def _run_validation(gradebook, settings, *, quiet=False) -> bool:
    rep = validate(gradebook)
    path = rep.write_csv(settings.output_dir / "validation_report.csv")
    if not quiet or rep.issues:
        print(format_report(rep))
        if rep.issues:
            print(f"Report: {path}")
    return rep.ok


def cmd_validate(args) -> int:
    settings = load_settings(args.env)
    gradebook = load_gradebook(settings.grades_file, settings.grades_sheet)
    _print_summary(gradebook)
    print()
    return 0 if _run_validation(gradebook, settings) else 1


def _print_summary(gradebook) -> None:
    sendable = gradebook.sendable()
    skipped = gradebook.skipped()
    print(f"Gradebook: {gradebook.source}")
    print(f"Students : {len(gradebook.students)}  (sendable {len(sendable)}, skipped {len(skipped)})")
    if skipped:
        print("Skipped:")
        for s, reason in skipped:
            print(f"  row {s.row_number:>4}  {s.student_number:<12} {reason}")
    if not sendable:
        print("WARNING: nothing to send — check the Email Address column.")


def cmd_preview(args) -> int:
    settings = load_settings(args.env)
    gradebook = load_gradebook(settings.grades_file, settings.grades_sheet)
    _print_summary(gradebook)
    emails = render_all(gradebook, settings)
    paths = write_previews(emails, settings.preview_dir)
    print(f"Wrote {len(paths)} preview(s) to {settings.preview_dir}/")
    return 0


def cmd_send(args) -> int:
    settings = load_settings(args.env)
    gradebook = load_gradebook(settings.grades_file, settings.grades_sheet)
    _print_summary(gradebook)
    print()
    if not _run_validation(gradebook, settings, quiet=True):
        if args.skip_validation:
            print("WARNING: continuing despite validation errors (--skip-validation).")
        else:
            print("Refusing to send while validation has errors. Fix them, or pass --skip-validation.")
            return 5
    emails = render_all(gradebook, settings)
    if not emails:
        return 1

    mode = "TEST MODE -> all to " + settings.test_email if settings.test_mode else "LIVE -> real student addresses"
    print(f"\nMode: {mode}")
    problems = settings.validate_for_sending()
    if problems:
        print("Cannot send; fix these in .env:")
        for p in problems:
            print("  -", p)
        return 2

    if not args.confirm:
        print(f"\nDRY RUN: would send {len(emails)} email(s). Re-run with --confirm to send.")
        for e in emails:
            print(f"  {e.student_number:<12} -> {e.recipient}")
        return 0

    try:
        summary = send_emails(emails, settings, resend=args.resend)
    except smtplib.SMTPAuthenticationError:
        print("ERROR: SMTP login failed. Check SENDER_EMAIL / APP_PASSWORD "
              "(Gmail needs an App Password, not your account password).", file=sys.stderr)
        return 4
    except (OSError, smtplib.SMTPException) as exc:
        print(f"ERROR: could not reach {settings.smtp_host}:{settings.smtp_port} — {exc}", file=sys.stderr)
        return 4
    print(f"\nRun {summary.run_id}: sent {summary.sent}, "
          f"skipped as duplicate {summary.skipped_duplicate}, failed {summary.failed}")
    print(f"Log: {settings.sent_log}")
    return 0 if summary.failed == 0 else 3


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog="grade_automation", description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--env", default=None, help="path to .env (default: ./.env)")
    sub = p.add_subparsers(dest="command", required=True)

    sv = sub.add_parser("validate", help="check the workbook; writes output/validation_report.csv")
    sv.set_defaults(func=cmd_validate)

    sp = sub.add_parser("preview", help="render HTML previews; sends nothing")
    sp.set_defaults(func=cmd_preview)

    ss = sub.add_parser("send", help="send emails (dry run unless --confirm)")
    ss.add_argument("--confirm", action="store_true", help="actually send")
    ss.add_argument("--resend", action="store_true", help="ignore the sent log")
    ss.add_argument("--skip-validation", action="store_true",
                    help="send even if validation reports errors (not recommended)")
    ss.set_defaults(func=cmd_send)

    args = p.parse_args(argv)
    try:
        return args.func(args)
    except GradebookError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
