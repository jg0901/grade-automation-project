"""Command-line entry point.

    python -m grade_automation validate           # check the workbook, write a CSV report
    python -m grade_automation compute            # final grades from Python -> output/final_grades.{csv,xlsx}
    python -m grade_automation preview            # render HTML previews, send nothing
    python -m grade_automation send --confirm     # actually send (TEST_MODE still honoured)
    python -m grade_automation send --confirm --resend   # ignore the sent log

Dry run is the default everywhere. `send` without `--confirm` only prints
what it would do, and `send` refuses to run while validation reports errors.
"""

from __future__ import annotations

import argparse
from collections import Counter
import smtplib
import sys

from .config import load_settings
from .load import GradebookError, load_gradebook
from .compute import compute_grades, format_summary, graded_gradebook, write_csv, write_xlsx
from .notify import already_sent, render_all, send_emails, write_previews
from .validate import format_report, validate


def _load(settings):
    """Load the workbook and, unless GRADE_SOURCE=excel, replace the grade
    columns with the Python-computed values. Returns (raw, for_emails)."""
    raw = load_gradebook(settings.grades_file, settings.grades_sheet)
    if settings.grade_source == "excel":
        return raw, raw
    return raw, graded_gradebook(raw)


def cmd_compute(args) -> int:
    settings = load_settings(args.env)
    raw = load_gradebook(settings.grades_file, settings.grades_sheet)
    computed = compute_grades(raw)
    print(f"Gradebook: {raw.source}")
    print(format_summary(computed))
    csv_path = write_csv(computed, settings.output_dir / "final_grades.csv")
    xlsx_path = write_xlsx(computed, settings.output_dir / "final_grades.xlsx")
    print(f"Wrote {csv_path} and {xlsx_path}")
    return 0 if all(not c.discrepancies() for c in computed) else 1


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
        if len(skipped) <= max(10, len(gradebook.students) // 2):
            print("Skipped:")
            for s, reason in skipped:
                print(f"  row {s.row_number:>4}  {s.student_number:<12} {reason}")
        else:
            # Most of the class is skipped: summarise by reason instead of listing everyone.
            by_reason = Counter(reason for _, reason in skipped)
            print("Skipped: " + "; ".join(f"{reason} ({n})" for reason, n in by_reason.most_common()))
    if not sendable:
        print("WARNING: nothing to send — check the Email Address column.")


def cmd_preview(args) -> int:
    settings = load_settings(args.env)
    _, gradebook = _load(settings)
    _print_summary(gradebook)
    print(f"Grade source: {settings.grade_source}")
    emails = render_all(gradebook, settings)
    paths = write_previews(emails, settings.preview_dir)
    print(f"Wrote {len(paths)} preview(s) to {settings.preview_dir}")
    return 0


def cmd_send(args) -> int:
    settings = load_settings(args.env)
    raw, gradebook = _load(settings)
    _print_summary(gradebook)
    print(f"Grade source: {settings.grade_source}")
    print()
    # Validation always runs on the raw workbook (it compares Excel vs Python).
    validation_ok = _run_validation(raw, settings, quiet=True)
    if not validation_ok:
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

    done = set() if (args.resend or settings.test_mode) else already_sent(settings.sent_log)
    to_send = [e for e in emails if e.student_number not in done]

    if not args.confirm:
        print(f"\nDRY RUN: would send {len(to_send)} email(s)"
              + (f", skip {len(done)} already sent" if done else "")
              + ". Re-run with --confirm to send.")
        for e in to_send:
            print(f"  {e.student_number:<12} -> {e.recipient}")
        return 0

    if not settings.test_mode:
        # Live send: show exactly what is about to happen and require a typed approval.
        print("\n" + "=" * 60)
        print("LIVE SEND — this goes to real students")
        print("=" * 60)
        print(f"  Emails to send      : {len(to_send)}")
        print(f"  Already sent (skip) : {len(done)}")
        print(f"  No grade (skip)     : {len(gradebook.skipped())}")
        print(f"  Grade source        : {settings.grade_source}")
        print(f"  Validation          : {'passed' if validation_ok else 'ERRORS (overridden)'}")
        print(f"  From                : {settings.sender_name} <{settings.sender_email}>")
        print(f"  Subject             : {settings.email_subject}")
        print("  First recipients    :")
        for e in to_send[:3]:
            print(f"      {e.student_number:<12} {e.recipient}")
        if len(to_send) > 3:
            print(f"      ... and {len(to_send) - 3} more")
        print(f"  Previews            : {settings.preview_dir} (run 'preview' if you haven't checked them)")
        print()
        if not sys.stdin.isatty():
            print("Refusing live send: no interactive terminal to approve from.")
            return 6
        answer = input(f"Type SEND {len(to_send)} to continue, anything else to abort: ").strip()
        if answer != f"SEND {len(to_send)}":
            print("Aborted. Nothing was sent.")
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

    sc = sub.add_parser("compute", help="compute final grades in Python; writes output/final_grades.{csv,xlsx}")
    sc.set_defaults(func=cmd_compute)

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
