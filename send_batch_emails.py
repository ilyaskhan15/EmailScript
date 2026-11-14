"""Send one daily batch of personalized emails from an Excel workbook."""

from __future__ import annotations

import argparse
import os
from datetime import date, datetime
from pathlib import Path
from zoneinfo import ZoneInfo

from email_automation import (
    DEFAULT_CV,
    DEFAULT_SHEET,
    build_message,
    email_address,
    load_rows,
    make_subject,
    personalize,
    send_messages,
)


DEFAULT_WORKBOOK = Path("final.xlsx")
DEFAULT_TEMPLATE = Path("email.txt")
DEFAULT_BATCH_SIZE = 50
PAKISTAN_TIME = ZoneInfo("Asia/Karachi")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Send today's 50-contact email batch from final.xlsx."
    )
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK)
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE)
    parser.add_argument("--cv", type=Path, default=DEFAULT_CV)
    parser.add_argument("--sheet", default=DEFAULT_SHEET)
    parser.add_argument("--batch-size", type=int, default=DEFAULT_BATCH_SIZE)
    parser.add_argument(
        "--start-date",
        type=date.fromisoformat,
        required=True,
        help="Pakistan date on which batch 1 should be sent (YYYY-MM-DD).",
    )
    parser.add_argument(
        "--send",
        action="store_true",
        help="Send through SMTP. Without this flag, print the selected batch only.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if args.batch_size < 1:
        raise ValueError("--batch-size must be at least 1.")
    for path in (args.workbook, args.template, args.cv):
        if not path.is_file():
            raise FileNotFoundError(f"Required file was not found: {path}")

    today = datetime.now(PAKISTAN_TIME).date()
    day_number = (today - args.start_date).days
    if day_number < 0:
        print(f"Batch schedule starts on {args.start_date.isoformat()}; no emails sent.")
        return

    rows = [
        row
        for row in load_rows(args.workbook, args.sheet)
        if "@" in email_address(row)
    ]
    start = day_number * args.batch_size
    batch = rows[start : start + args.batch_size]
    if not batch:
        print(f"All {len(rows)} valid contacts have already been processed; no emails sent.")
        return

    template = args.template.read_text(encoding="utf-8")
    messages = [
        build_message(
            recipient=email_address(row),
            body=personalize(template, row),
            cv_path=args.cv,
            subject=make_subject(row),
            sender=os.getenv("SMTP_USERNAME"),
        )
        for row in batch
    ]

    print(
        f"Pakistan date {today.isoformat()}: selected contacts "
        f"{start + 1}-{start + len(batch)} of {len(rows)}."
    )
    if args.send:
        send_messages(messages)
        print(f"Sent {len(messages)} email(s).")
    else:
        print("Preview only; no emails sent. Add --send to send this batch.")
        for message in messages:
            print(message["To"])


if __name__ == "__main__":
    main()
