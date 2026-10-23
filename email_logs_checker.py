#!/usr/bin/env python3
"""
email_logs_checker.py

Checks the freshness of one email log per weekday and writes an Excel
report with the last modification timestamp and status of each file.

Status values:
    OK       -> log was updated within the allowed age
    STALE    -> log exists but its last update is older than the threshold
    MISSING  -> log file not found in the target directory
    ERROR    -> log file exists but could not be read

Exit code:
    0 -> all logs OK
    1 -> at least one log is STALE, MISSING or ERROR

Usage:
    python email_logs_checker.py --path ./logs
    python email_logs_checker.py --path ./logs --max-age-hours 36
"""

import argparse
import logging
import os
import sys
from datetime import datetime

import openpyxl

# Example mapping. Replace with the real filenames of your environment.
EMAIL_LOGS = {
    "Monday": "email_monday.log",
    "Tuesday": "email_tuesday.log",
    "Wednesday": "email_wednesday.log",
    "Thursday": "email_thursday.log",
    "Friday": "email_friday.log",
    "Saturday": "email_saturday.log",
}


def parse_args():
    parser = argparse.ArgumentParser(
        description="Check freshness of daily email log files and generate an Excel report."
    )
    parser.add_argument("--path", default=".", help="Directory containing the log files.")
    parser.add_argument(
        "--output", default="email_logs_report.xlsx", help="Output Excel file."
    )
    parser.add_argument(
        "--max-age-hours",
        type=float,
        default=36.0,
        help="Maximum acceptable age of a log file, in hours (default: 36).",
    )
    return parser.parse_args()


def classify(age_hours, max_age_hours):
    if age_hours is None:
        return "MISSING"
    return "STALE" if age_hours > max_age_hours else "OK"


def main():
    args = parse_args()
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Email Logs"
    sheet.append(
        ["Weekday", "File", "Last Modified (date)", "Last Modified (time)", "Age (h)", "Status"]
    )

    now = datetime.now()
    issues = 0

    for weekday, filename in EMAIL_LOGS.items():
        full_path = os.path.join(args.path, filename)

        if not os.path.exists(full_path):
            logging.warning("Missing: %s (%s)", filename, weekday)
            sheet.append([weekday, filename, "-", "-", "-", "MISSING"])
            issues += 1
            continue

        try:
            mtime = datetime.fromtimestamp(os.path.getmtime(full_path))
        except OSError as exc:
            logging.error("Cannot read %s: %s", filename, exc)
            sheet.append([weekday, filename, "ERROR", "ERROR", "-", "ERROR"])
            issues += 1
            continue

        age_hours = (now - mtime).total_seconds() / 3600
        status = classify(age_hours, args.max_age_hours)
        if status != "OK":
            issues += 1

        sheet.append(
            [
                weekday,
                filename,
                mtime.strftime("%d/%m/%Y"),
                mtime.strftime("%H:%M"),
                round(age_hours, 1),
                status,
            ]
        )

    workbook.save(args.output)
    logging.info("Report saved to %s", args.output)

    if issues:
        logging.warning("%d issue(s) found.", issues)
        sys.exit(1)


if __name__ == "__main__":
    main()