"""Create personalized email drafts from an Excel contact list.

The script does not send anything unless --send is supplied. For sending,
configure SMTP_HOST, SMTP_PORT, SMTP_USERNAME, and SMTP_PASSWORD in the
environment. SMTP_USERNAME is also used as the From address.
"""

from __future__ import annotations

import argparse
import csv
import os
import re
import smtplib
from email.message import EmailMessage
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


DEFAULT_WORKBOOK = Path("final.xlsx")
DEFAULT_TEMPLATE = Path("email.txt")
DEFAULT_CV = Path("Muhammad_Ilyas_Academic_CV.pdf")
DEFAULT_OUTPUT = Path("drafts")
DEFAULT_SHEET = "Enhanced Email List"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Create personalized email drafts from an Excel workbook."
    )
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK)
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE)
    parser.add_argument("--cv", type=Path, default=DEFAULT_CV)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--sheet", default=DEFAULT_SHEET)
    parser.add_argument(
        "--send",
        action="store_true",
        help="Send emails through SMTP instead of only creating .eml drafts.",
    )
    parser.add_argument(
        "--limit",
        type=int,
        help="Process at most this many rows (useful for testing).",
    )
    return parser.parse_args()


def clean_value(value: Any) -> str:
    return "" if value is None else str(value).strip()


def recipient_name(value: str) -> str:
    """Convert common spreadsheet titles to a natural email salutation."""
    name = re.sub(r"\s*\([^)]*\)", "", value).strip()
    name = re.sub(r"^(professor|prof\.|dr\.?)\s+", "", name, flags=re.I)
    return f"Professor {name}" if name else "Professor"


def email_address(row: dict[str, Any]) -> str:
    preferred = clean_value(row.get("Preferred Email"))
    fallback = clean_value(row.get("Email"))
    return preferred or fallback


def personalize(template: str, row: dict[str, Any]) -> str:
    university = clean_value(row.get("University"))
    professor = recipient_name(clean_value(row.get("Professor Name")))
    research_area = clean_value(row.get("Research Area / Department"))
    values = {
        "university": university,
        "professor_name": professor,
        "research_area": research_area or "Artificial Intelligence and Software Engineering",
    }
    try:
        return template.format_map(values)
    except KeyError as exc:
        raise ValueError(
            f"Template contains unsupported placeholder {exc.args[0]!r}. "
            "Supported placeholders: {university}, {professor_name}, {research_area}."
        ) from exc


def build_message(
    recipient: str,
    body: str,
    cv_path: Path,
    subject: str,
    sender: str | None,
) -> EmailMessage:
    message = EmailMessage()
    message["To"] = recipient
    message["Subject"] = subject
    if sender:
        message["From"] = sender
    message.set_content(body)
    message.add_attachment(
        cv_path.read_bytes(),
        maintype="application",
        subtype="pdf",
        filename=cv_path.name,
    )
    return message


def make_subject(row: dict[str, Any]) -> str:
    university = clean_value(row.get("University"))
    return f"Prospective Master's Student - CSC Scholarship 2027 - {university}"


def load_rows(workbook_path: Path, sheet_name: str) -> list[dict[str, Any]]:
    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    if sheet_name not in workbook.sheetnames:
        raise ValueError(
            f"Worksheet {sheet_name!r} was not found. "
            f"Available sheets: {', '.join(workbook.sheetnames)}"
        )
    worksheet = workbook[sheet_name]
    rows = worksheet.iter_rows(values_only=True)
    headers = [clean_value(value) for value in next(rows)]
    if not {"University", "Professor Name"}.issubset(headers):
        raise ValueError(
            "The worksheet must contain 'University' and 'Professor Name' columns."
        )
    return [dict(zip(headers, row)) for row in rows if any(row)]


def send_messages(messages: list[EmailMessage]) -> None:
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
        for message in messages:
            smtp.send_message(message, from_addr=username)


def main() -> None:
    args = parse_args()
    for path in (args.workbook, args.template, args.cv):
        if not path.is_file():
            raise FileNotFoundError(f"Required file was not found: {path}")

    template = args.template.read_text(encoding="utf-8")
    rows = load_rows(args.workbook, args.sheet)
    if args.limit is not None:
        if args.limit < 1:
            raise ValueError("--limit must be at least 1.")
        rows = rows[: args.limit]

    sender = os.getenv("SMTP_USERNAME")
    messages: list[EmailMessage] = []
    log_rows: list[dict[str, str]] = []
    args.output.mkdir(parents=True, exist_ok=True)

    for index, row in enumerate(rows, start=2):
        recipient = email_address(row)
        if not recipient or "@" not in recipient:
            log_rows.append({"row": str(index), "email": recipient, "status": "skipped: invalid email"})
            continue
        message = build_message(
            recipient=recipient,
            body=personalize(template, row),
            cv_path=args.cv,
            subject=make_subject(row),
            sender=sender,
        )
        messages.append(message)
        draft_path = args.output / f"{index:04d}-{recipient.replace('@', '_at_')}.eml"
        draft_path.write_bytes(bytes(message))
        log_rows.append({"row": str(index), "email": recipient, "status": f"drafted: {draft_path.name}"})

    if args.send:
        send_messages(messages)
        for log_row in log_rows:
            if log_row["status"].startswith("drafted:"):
                log_row["status"] = "sent"

    with (args.output / "email_log.csv").open("w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=["row", "email", "status"])
        writer.writeheader()
        writer.writerows(log_rows)

    action = "sent" if args.send else "drafted"
    print(f"{len(messages)} email(s) {action}.")
    print(f"Log: {args.output / 'email_log.csv'}")


if __name__ == "__main__":
    main()
