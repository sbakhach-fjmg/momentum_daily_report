#!/usr/bin/env python
# coding: utf-8

import argparse
import os
import sys
import time
import base64
import io
from pathlib import Path

import pandas as pd
import pyarrow as pa
import pyarrow.parquet as pq
import requests
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / 'credentials.env')

TENANT_ID = os.getenv('TENANT_ID')
CLIENT_ID = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
# Office 365 user email
USER_EMAIL = os.getenv('USER_EMAIL', 'timedata@fjcorp.net')

# URLs for Azure OAuth2 token and Microsoft Graph API
TOKEN_URL = f'https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token'
RESOURCE_URL = 'https://graph.microsoft.com/v1.0'

# Placeholders until the new report's sender/subject are known; override in credentials.env
PLACEHOLDER_SENDER = 'CHANGE_ME@example.com'
PLACEHOLDER_SUBJECT = 'CHANGE ME'
SENDER = os.getenv('REPORT_SENDER', PLACEHOLDER_SENDER)
SUBJECT_MARKER = os.getenv('REPORT_SUBJECT', PLACEHOLDER_SUBJECT)
OUTPUT_DIR = BASE_DIR / 'raw_files'

MAX_RETRIES = 4


def get_access_token(client_id, client_secret, tenant_id):
    data = {
        'client_id': client_id,
        'scope': 'https://graph.microsoft.com/.default',
        'client_secret': client_secret,
        'grant_type': 'client_credentials',
    }
    response = requests.post(TOKEN_URL, data=data)
    response.raise_for_status()
    return response.json().get('access_token')


def _paged_get(url, headers):
    """Yield JSON pages following @odata.nextLink, retrying transient failures."""
    while url:
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                resp = requests.get(url, headers=headers, timeout=120)
                resp.raise_for_status()
                data = resp.json()
                break
            except (requests.ConnectionError, requests.Timeout,
                    requests.exceptions.ContentDecodingError, requests.HTTPError) as e:
                status = getattr(getattr(e, 'response', None), 'status_code', None)
                if attempt == MAX_RETRIES or (status is not None and status < 500 and status != 429):
                    raise
                print(f"  retry {attempt}/{MAX_RETRIES - 1} after error: {e.__class__.__name__}")
                time.sleep(5 * attempt)
        yield data
        url = data.get('@odata.nextLink')


def find_messages(headers, user_email, since):
    """Yield inbox messages from SENDER received at or after `since` whose subject contains SUBJECT_MARKER."""
    # Graph requires $orderby properties to also appear first in $filter
    filter_clause = (
        f"receivedDateTime ge {since:%Y-%m-%dT%H:%M:%SZ}"
        f" and from/emailAddress/address eq '{SENDER}'"
        f" and hasAttachments eq true"
    )
    url = (
        f"{RESOURCE_URL}/users/{user_email}/mailFolders/inbox/messages"
        f"?$filter={filter_clause}"
        f"&$orderby=receivedDateTime desc"
        f"&$select=id,subject,receivedDateTime"
        f"&$top=50"
    )
    # Subject matching is done client-side
    for page in _paged_get(url, headers):
        for message in page.get('value', []):
            if SUBJECT_MARKER in (message.get('subject') or ''):
                yield message


def reshape_report(csv_bytes):
    """Turn the wide report (Day N / date / BDC|Sales header rows) into one row per location, date and type."""
    raw = pd.read_csv(io.BytesIO(csv_bytes), header=None, dtype=str, encoding='utf-8-sig')
    # Row 0 is "Location, Day 1, Day 1, ..."; rows 1-2 hold each column's date and appointment type
    dates = pd.to_datetime(raw.iloc[1, 1:], format='ISO8601').dt.date
    types = raw.iloc[2, 1:].str.strip()
    body = raw.iloc[3:].set_index(0)
    body.index.name = 'Location'
    body.columns = pd.MultiIndex.from_arrays([dates, types], names=['date', 'type'])
    body = body[body.index.str.strip().str.lower() != 'totals']

    long = body.stack(['date', 'type']).rename('appts').reset_index()
    long['appts'] = pd.to_numeric(long['appts']).astype('Int64')
    return long


def write_parquet(df, path):
    """Write `df` as parquet with plain utf8 strings; pandas 3 emits large_utf8, which Alteryx can't read."""
    table = pa.Table.from_pandas(df, preserve_index=False)
    schema = pa.schema([
        field.with_type(pa.string()) if pa.types.is_large_string(field.type) else field
        for field in table.schema
    ], metadata=table.schema.metadata)
    pq.write_table(table.cast(schema), path)


def process_message(headers, user_email, message):
    """Save each CSV attachment of `message` as parquet, skipping ones already on disk. Returns (saved, skipped)."""
    received_at = pd.to_datetime(message['receivedDateTime'], utc=True)
    saved = skipped = 0
    # List attachment metadata only, so already-processed files are never downloaded
    attach_url = (f"{RESOURCE_URL}/users/{user_email}/messages/{message['id']}/attachments"
                  f"?$select=id,name,isInline&$top=50")
    for attachments_page in _paged_get(attach_url, headers):
        for attachment in attachments_page.get('value', []):
            # Only save real CSV file attachments (skip inline/linked items);
            # match on extension since CSVs often arrive as application/octet-stream
            if attachment.get('@odata.type') != '#microsoft.graph.fileAttachment':
                continue
            if attachment.get('isInline', False):
                continue
            file_name = attachment['name']
            if not file_name.lower().endswith('.csv'):
                continue

            # Prefix with received timestamp so repeated same-name reports don't collide;
            # the name is deterministic, so an existing file means this attachment is done
            target_path = OUTPUT_DIR / f"{received_at:%Y-%m-%d_%H%M%S}_{Path(file_name).stem}.parquet"
            if target_path.exists():
                skipped += 1
                print(f"  already have {target_path.name}")
                continue

            content_url = f"{RESOURCE_URL}/users/{user_email}/messages/{message['id']}/attachments/{attachment['id']}"
            content = next(_paged_get(content_url, headers))
            df = reshape_report(base64.b64decode(content['contentBytes']))
            df['email_received_at'] = received_at

            # Write to a temp file first so an interrupted run never leaves a partial file that looks done
            tmp_path = target_path.with_suffix('.parquet.tmp')
            write_parquet(df, tmp_path)
            tmp_path.replace(target_path)
            saved += 1
            print(f"  saved {target_path.name} ({len(df):,} rows)")
    return saved, skipped


def download_recent(access_token, user_email, days):
    headers = {
        'Authorization': 'Bearer ' + access_token,
        'Content-Type': 'application/json',
        'Accept-Encoding': 'identity',  # avoid intermittent gzip decode errors on large attachments
    }
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    since = pd.Timestamp.now(tz='UTC') - pd.Timedelta(days=days)
    total_saved = total_skipped = messages = 0
    for message in find_messages(headers, user_email, since):
        messages += 1
        received_at = pd.to_datetime(message['receivedDateTime'], utc=True)
        print(f"'{message['subject']}' received {received_at}")
        saved, skipped = process_message(headers, user_email, message)
        if not saved and not skipped:
            print("  no .csv attachment")
        total_saved += saved
        total_skipped += skipped

    if not messages:
        sys.exit(f"No inbox email from {SENDER} with '{SUBJECT_MARKER}' in the subject "
                 f"was received in the last {days} day(s).")
    print(f"Done: {messages} email(s), {total_saved} new file(s) saved, {total_skipped} already present.")


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Download daily report emails from the last N days as parquet.')
    parser.add_argument('--days', type=int, default=7, help='how many days back to look (default: 7)')
    args = parser.parse_args()

    missing = [name for name in ('TENANT_ID', 'CLIENT_ID', 'CLIENT_SECRET') if not os.getenv(name)]
    if missing:
        sys.exit(f"Missing required values in credentials.env: {', '.join(missing)}")
    if SENDER == PLACEHOLDER_SENDER or SUBJECT_MARKER == PLACEHOLDER_SUBJECT:
        sys.exit("Set REPORT_SENDER and REPORT_SUBJECT in credentials.env for the new daily report.")
    token = get_access_token(CLIENT_ID, CLIENT_SECRET, TENANT_ID)
    download_recent(token, USER_EMAIL, args.days)
