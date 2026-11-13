"""Send one test email using the first valid contact in test_emails.xlsx.

Without --send, this creates a preview .eml file and does not contact SMTP.
Set SMTP_HOST, SMTP_PORT, SMTP_USERNAME, and SMTP_PASSWORD before sending.
"""

from __future__ import annotations

import argparse
import os
import smtplib
from email.message import EmailMessage
from pathlib import Path

from email_automation import (
    DEFAULT_CV,
    DEFAULT_SHEET,
    build_message,
    email_address,
    load_rows,
    make_subject,
    personalize,
)


DEFAULT_WORKBOOK = Path("test_emails.xlsx")
DEFAULT_TEMPLATE = Path("email.txt")
DEFAULT_OUTPUT = Path("test_first_email.eml")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Send an email to only the first valid row in test_emails.xlsx."
    )
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK)
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE)
    parser.add_argument("--cv", type=Path, default=DEFAULT_CV)
    parser.add_argument("--sheet", default=DEFAULT_SHEET)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument(
        "--send",
        action="store_true",
        help="Send through SMTP. Without this flag, only create a preview.",
    )
    return parser.parse_args()


def send(message: EmailMessage) -> None:
    host = os.getenv("SMTP_HOST")
    username = os.getenv("SMTP_USERNAME")
    password = os.getenv("SMTP_PASSWORD")
    port = int(os.getenv("SMTP_PORT", "587"))
    if not host or not username or not password:
        raise RuntimeError(
            "Set SMTP_HOST, SMTP_USERNAME, and SMTP_PASSWORD before using --send."
        )
    with smtplib.SMTP(host, port, timeout=30) as smtp:
        smtp.starttls()
        smtp.login(username, password)
        smtp.send_message(message, from_addr=username)


def main() -> None:
    args = parse_args()
    for path in (args.workbook, args.template, args.cv):
        if not path.is_file():
            raise FileNotFoundError(f"Required file was not found: {path}")

    rows = load_rows(args.workbook, args.sheet)
    first_row = next((row for row in rows if "@" in email_address(row)), None)
    if first_row is None:
        raise ValueError("No row with a valid email address was found.")

    recipient = email_address(first_row)
    message = build_message(
        recipient=recipient,
        body=personalize(args.template.read_text(encoding="utf-8"), first_row),
        cv_path=args.cv,
        subject=make_subject(first_row),
        sender=os.getenv("SMTP_USERNAME"),
    )
    args.output.write_bytes(bytes(message))

    if args.send:
        send(message)
        print(f"Sent one test email to {recipient}.")
    else:
        print(f"Preview created for {recipient}: {args.output}")
        print("No email was sent. Add --send to send it.")


if __name__ == "__main__":
    main()
