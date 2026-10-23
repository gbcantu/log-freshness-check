# log-freshness-check

Two small Python utilities for checking whether log files are being updated
as expected, and producing an Excel report with the last modification time
and a freshness status for each file.

This is useful when logs are an indirect indicator of service health: if a
system stops writing its log, it may have stopped working.

## Scripts

### `system_logs_checker.py`

Checks a set of system log files and writes a consolidated report.
The default list covers different kinds of machines and roles, such as:

- email archive and current email logs
- a finance workstation log
- HR workstation and NAS logs
- a lab workstation log
- a maintenance workstation log
- generic user workstation logs

### `email_logs_checker.py`

Checks one email log per weekday (Monday to Saturday) and writes a weekly
report.

Both scripts are independent and can be used separately.

## Status values

- `OK` — log was updated within the allowed age.
- `STALE` — log exists but its last update is older than the threshold.
- `MISSING` — log file was not found in the target directory.
- `ERROR` — log file exists but could not be read.

## Usage

Install dependencies:

```bash
pip install -r requirements.txt