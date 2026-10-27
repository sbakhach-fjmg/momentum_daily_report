# momentum_daily_report

Downloads the CSV attachment from the newest daily report email in an Office 365 inbox (via Microsoft Graph)
and saves it as parquet in `raw_files/`, named `<received timestamp>_<attachment name>.parquet`.
All columns are kept as text, plus an `email_received_at` column.

## Setup

Requires [uv](https://docs.astral.sh/uv/). Python 3.12 is pinned in `.python-version`.

```
uv sync
```

Create `credentials.env` next to the script (it is git-ignored):

```
TENANT_ID=...
CLIENT_ID=...
CLIENT_SECRET=...
USER_EMAIL=timedata@fjcorp.net        # mailbox to read (optional, this is the default)
REPORT_SENDER=sender@example.com      # address the daily report comes from
REPORT_SUBJECT=Text in the subject    # substring matched against the subject
```

## Run

```
uv run get_latest_report.py
```
