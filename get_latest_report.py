#!/usr/bin/env python
# coding: utf-8

import os
import sys
import time
import base64
import io
from pathlib import Path

import pandas as pd
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


def find_latest_message(headers, user_email):
    """Return the newest inbox message from SENDER whose subject contains SUBJECT_MARKER."""
    # Graph requires $orderby properties to also appear first in $filter
    filter_clause = (
        f"receivedDateTime ge 1900-01-01T00:00:00Z"
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
    # Subject matching is done client-side; newest-first order means the first hit wins
    for page in _paged_get(url, headers):
        for message in page.get('value', []):
            if SUBJECT_MARKER in (message.get('subject') or ''):
                return message
    return None


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


def download_latest(access_token, user_email):
    headers = {
        'Authorization': 'Bearer ' + access_token,
        'Content-Type': 'application/json',
        'Accept-Encoding': 'identity',  # avoid intermittent gzip decode errors on large attachments
    }
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    message = find_latest_message(headers, user_email)
    if message is None:
        sys.exit(f"No inbox email from {SENDER} with '{SUBJECT_MARKER}' in the subject was found.")

    subject = message['subject']
    received_at = pd.to_datetime(message['receivedDateTime'], utc=True)
    print(f"Latest match: '{subject}' received {received_at}")

    saved = 0
    attach_url = f"{RESOURCE_URL}/users/{user_email}/messages/{message['id']}/attachments?$top=50"
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

            decoded = base64.b64decode(attachment['contentBytes'])
            df = reshape_report(decoded)
            df['email_received_at'] = received_at

            # Prefix with received timestamp so repeated same-name reports don't collide
            target_path = OUTPUT_DIR / f"{received_at:%Y-%m-%d_%H%M%S}_{Path(file_name).stem}.parquet"
            df.to_parquet(target_path, index=False)
            saved += 1
            print(f"Saved {target_path.name} ({len(df):,} rows)")

    if not saved:
        sys.exit("The latest matching email had no .csv attachment.")


if __name__ == '__main__':
    missing = [name for name in ('TENANT_ID', 'CLIENT_ID', 'CLIENT_SECRET') if not os.getenv(name)]
    if missing:
        sys.exit(f"Missing required values in credentials.env: {', '.join(missing)}")
    if SENDER == PLACEHOLDER_SENDER or SUBJECT_MARKER == PLACEHOLDER_SUBJECT:
        sys.exit("Set REPORT_SENDER and REPORT_SUBJECT in credentials.env for the new daily report.")
    token = get_access_token(CLIENT_ID, CLIENT_SECRET, TENANT_ID)
    download_latest(token, USER_EMAIL)
