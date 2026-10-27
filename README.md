# momentum_daily_report

Downloads the CSV attachments from the daily report emails of the last N days in an Office 365 inbox (via Microsoft Graph)
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
uv run download_reports.py            # last 7 days (default)
uv run download_reports.py --days 30  # look further back
```

Every matching email in the window is processed. Re-running is safe: an attachment whose parquet file
already exists in `raw_files/` is skipped without being downloaded again.
