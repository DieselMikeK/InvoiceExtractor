"""Gmail API client for fetching emails and downloading attachments."""
import os
import io
import csv
import base64
import pickle
import re
import time
from html import unescape
from email.utils import parseaddr
from datetime import datetime
from google.auth.exceptions import RefreshError
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

SCOPES = [
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/drive',
]

PROCESSED_LABEL_NAME = "InvoiceExtractor-Processed"


class WrongAuthorizedAccountError(Exception):
    """Raised when OAuth succeeds with a disallowed Gmail account."""


def retry_with_backoff(func, max_retries=3, base_delay=2, status_callback=None):
    """Execute a function with exponential backoff retry on failure."""
    for attempt in range(max_retries + 1):
        try:
            return func()
        except HttpError as e:
            if attempt == max_retries:
                raise
            delay = base_delay * (2 ** attempt)
            if status_callback:
                status_callback(
                    f"API error (attempt {attempt + 1}/{max_retries}): {e}. "
                    f"Retrying in {delay}s...", "warning"
                )
            time.sleep(delay)
        except Exception as e:
            if attempt == max_retries:
                raise
            delay = base_delay * (2 ** attempt)
            if status_callback:
                status_callback(
                    f"Error (attempt {attempt + 1}/{max_retries}): {e}. "
                    f"Retrying in {delay}s...", "warning"
                )
            time.sleep(delay)


def _extract_sender_email(from_header):
    """Normalize a Gmail From header down to the sender email address."""
    raw = str(from_header or '').strip()
    if not raw:
        return ''
    parsed = parseaddr(raw)[1].strip().lower()
    return parsed


def _extract_email_addresses(text):
    if not text:
        return []
    return re.findall(
        r'[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}',
        str(text),
        re.IGNORECASE,
    )


def _decode_gmail_body_data(data):
    if not data:
        return ''
    raw = str(data).strip()
    if not raw:
        return ''
    padding = (-len(raw)) % 4
    if padding:
        raw += '=' * padding
    try:
        return base64.urlsafe_b64decode(raw.encode('utf-8')).decode('utf-8', errors='replace')
    except Exception:
        return ''


def _html_to_text(value):
    text = str(value or '')
    if not text:
        return ''
    text = re.sub(r'(?i)<br\s*/?>', '\n', text)
    text = re.sub(r'(?i)</(?:p|div|li|tr|table|section|h[1-6])>', '\n', text)
    text = re.sub(r'(?s)<[^>]+>', ' ', text)
    text = unescape(text)
    text = re.sub(r'\r\n?', '\n', text)
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()


def _collect_message_text_parts(payload):
    if not isinstance(payload, dict):
        return []

    parts = []
    mime_type = str(payload.get('mimeType', '') or '').lower()
    body = payload.get('body', {}) or {}
    data = body.get('data')

    if data and mime_type in ('text/plain', 'text/html'):
        decoded = _decode_gmail_body_data(data)
        if decoded:
            parts.append(decoded if mime_type == 'text/plain' else _html_to_text(decoded))

    for part in payload.get('parts', []) or []:
        parts.extend(_collect_message_text_parts(part))

    return parts


def _extract_forwarded_sender(payload, snippet=''):
    """Extract the original sender from a forwarded Gmail message body when present."""
    text_chunks = _collect_message_text_parts(payload)
    if snippet:
        text_chunks.append(unescape(str(snippet)))
    combined = '\n'.join(chunk for chunk in text_chunks if chunk).strip()
    if not combined:
        return '', ''

    lines = [re.sub(r'^\s*>+\s*', '', line).strip() for line in combined.splitlines()]
    marker_indexes = [
        idx for idx, line in enumerate(lines)
        if re.search(r'forwarded message|begin forwarded message', line, re.IGNORECASE)
    ]
    if not marker_indexes:
        return '', ''

    for start_idx in marker_indexes:
        for idx in range(start_idx + 1, min(start_idx + 20, len(lines))):
            line = lines[idx]
            match = re.match(r'^From:\s*(.+)$', line, re.IGNORECASE)
            if not match:
                continue

            header_parts = [match.group(1).strip()]
            for next_idx in range(idx + 1, min(idx + 4, len(lines))):
                next_line = lines[next_idx].strip()
                if not next_line:
                    break
                if re.match(r'^(?:to|cc|bcc|subject|date|reply-to)\s*:', next_line, re.IGNORECASE):
                    break
                if re.match(r'^[A-Za-z][A-Za-z\-]+\s*:', next_line):
                    break
                header_parts.append(next_line)

            header_value = re.sub(r'\s+', ' ', ' '.join(header_parts)).strip()
            emails = _extract_email_addresses(header_value)
            if emails:
                return emails[0].lower(), header_value
            parsed = _extract_sender_email(header_value)
            if parsed:
                return parsed, header_value

    return '', ''


def _message_internal_timestamp(message):
    """Return Gmail's stored message timestamp as whole seconds."""
    raw = str((message or {}).get('internalDate', '') or '').strip()
    if not raw:
        return None
    try:
        value = int(raw)
    except (TypeError, ValueError):
        return None
    if value > 10_000_000_000:
        value //= 1000
    return value


def _message_matches_time_filter(message, message_time_filter):
    """Check whether a Gmail message falls inside the requested timestamp window."""
    if not message_time_filter:
        return True

    timestamp = _message_internal_timestamp(message)
    if timestamp is None:
        return False

    try:
        start_ts = int(message_time_filter.get('start_ts', 0))
        end_ts = int(message_time_filter.get('end_ts', 0))
    except (AttributeError, TypeError, ValueError):
        return False

    return start_ts <= timestamp < end_ts


class GmailClient:
    def __init__(
        self,
        base_dir,
        status_callback=None,
        data_dir=None,
        invoices_dir=None,
        expected_email=None,
        should_stop=None,
    ):
        self.base_dir = base_dir
        self.data_dir = data_dir or base_dir
        self.status_callback = status_callback or (lambda msg, tag=None: None)
        self.client_secret = os.path.join(self.data_dir, 'client_secret.json')
        self.token_file = os.path.join(self.data_dir, 'token.pickle')
        self.invoices_dir = invoices_dir or os.path.join(self.data_dir, 'invoices')
        self.expected_email = str(expected_email or '').strip().lower()
        self.should_stop = should_stop or (lambda: False)
        self.service = None

        os.makedirs(self.invoices_dir, exist_ok=True)

    def _clear_cached_token(self):
        """Delete cached token file so OAuth can run fresh."""
        if os.path.exists(self.token_file):
            try:
                os.remove(self.token_file)
            except Exception as e:
                self.status_callback(
                    f"Warning: couldn't remove cached token file: {e}",
                    "warning"
                )

    def authenticate(self):
        """Authenticate with Gmail API, caching token for future runs."""
        creds = None
        should_persist_token = False
        if os.path.exists(self.token_file):
            try:
                with open(self.token_file, 'rb') as f:
                    creds = pickle.load(f)
            except Exception:
                self.status_callback(
                    "Cached token file is unreadable; forcing re-authentication.",
                    "warning"
                )
                self._clear_cached_token()
                creds = None

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                self.status_callback("Refreshing authentication token...")
                try:
                    creds.refresh(Request())
                    should_persist_token = True
                except RefreshError as e:
                    # Common case: invalid_grant when refresh token was revoked or expired.
                    self.status_callback(
                        f"Refresh token is invalid/revoked ({e}); "
                        "removing cached token and re-authenticating.",
                        "warning"
                    )
                    self._clear_cached_token()
                    creds = None
                except Exception as e:
                    err_text = str(e).lower()
                    if 'invalid_grant' in err_text or 'expired or revoked' in err_text:
                        self.status_callback(
                            f"Refresh token is invalid/revoked ({e}); "
                            "removing cached token and re-authenticating.",
                            "warning"
                        )
                        self._clear_cached_token()
                        creds = None
                    else:
                        raise

            if not creds or not creds.valid:
                if not os.path.exists(self.client_secret):
                    raise FileNotFoundError(
                        f"client_secret.json not found in {self.data_dir}. "
                        "Please download OAuth credentials from Google Cloud Console."
                    )
                self.status_callback("Opening browser for Gmail authentication...")
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.client_secret, SCOPES
                )
                creds = flow.run_local_server(
                    port=0,
                    prompt='select_account'
                )
                should_persist_token = True

        self.creds = creds
        self.service = build('gmail', 'v1', credentials=creds)

        profile = retry_with_backoff(
            lambda: self.service.users().getProfile(userId='me').execute(),
            status_callback=self.status_callback
        )
        connected_email = str(profile.get('emailAddress', '')).strip()
        if self.expected_email and connected_email.lower() != self.expected_email:
            self.status_callback(
                f"Connected account {connected_email or '(unknown)'} is not allowed.",
                "error"
            )
            self.status_callback(
                f"Please sign in as {self.expected_email}.",
                "error"
            )
            self._clear_cached_token()
            raise WrongAuthorizedAccountError(
                f"Signed in as {connected_email or '(unknown)'}; "
                f"expected {self.expected_email}."
            )

        if should_persist_token:
            with open(self.token_file, 'wb') as f:
                pickle.dump(creds, f)

        self.status_callback(f"Connected to: {connected_email}", "success")

        # Ensure our processed label exists
        self.processed_label_id = self._get_or_create_label(PROCESSED_LABEL_NAME)

        return profile

    def _get_or_create_label(self, label_name):
        """Get the ID of a label, creating it if it doesn't exist."""
        # List existing labels
        results = retry_with_backoff(
            lambda: self.service.users().labels().list(userId='me').execute(),
            status_callback=self.status_callback
        )
        labels = results.get('labels', [])

        # Check if label already exists
        for label in labels:
            if label['name'] == label_name:
                self.status_callback(f"Using existing label: {label_name}")
                return label['id']

        # Create the label
        label_body = {
            'name': label_name,
            'labelListVisibility': 'labelShow',
            'messageListVisibility': 'show',
        }
        created = retry_with_backoff(
            lambda: self.service.users().labels().create(
                userId='me', body=label_body
            ).execute(),
            status_callback=self.status_callback
        )
        self.status_callback(f"Created new label: {label_name}", "success")
        return created['id']

    def _add_label_to_message(self, msg_id, label_id):
        """Add a label to a message."""
        body = {'addLabelIds': [label_id]}
        retry_with_backoff(
            lambda: self.service.users().messages().modify(
                userId='me', id=msg_id, body=body
            ).execute(),
            status_callback=self.status_callback
        )

    def fetch_all_message_ids(self, query=None):
        """Fetch message IDs from Gmail, optionally filtered by a search query."""
        all_messages = []
        page_token = None

        while True:
            def list_messages(pt=page_token):
                kwargs = {'userId': 'me', 'maxResults': 500}
                if pt:
                    kwargs['pageToken'] = pt
                if query:
                    kwargs['q'] = query
                return self.service.users().messages().list(**kwargs).execute()

            results = retry_with_backoff(
                lambda: list_messages(page_token),
                status_callback=self.status_callback
            )
            messages = results.get('messages', [])
            all_messages.extend(messages)

            page_token = results.get('nextPageToken')
            if not page_token:
                break

        return all_messages

    def get_message_details(self, msg_id):
        """Get full message details including attachment info."""
        return retry_with_backoff(
            lambda: self.service.users().messages().get(
                userId='me', id=msg_id, format='full'
            ).execute(),
            status_callback=self.status_callback
        )

    def download_attachment(self, msg_id, attachment_id, filename):
        """Download a single attachment and save to invoices folder."""
        result = retry_with_backoff(
            lambda: self.service.users().messages().attachments().get(
                userId='me', messageId=msg_id, id=attachment_id
            ).execute(),
            status_callback=self.status_callback
        )
        file_data = base64.urlsafe_b64decode(result['data'])

        # Avoid filename collisions by prepending msg_id prefix
        safe_filename = filename.replace('/', '_').replace('\\', '_')
        filepath = os.path.join(self.invoices_dir, safe_filename)

        # If file already exists, add a numeric suffix
        if os.path.exists(filepath):
            name, ext = os.path.splitext(safe_filename)
            counter = 1
            while os.path.exists(filepath):
                safe_filename = f"{name}_{counter}{ext}"
                filepath = os.path.join(self.invoices_dir, safe_filename)
                counter += 1

        with open(filepath, 'wb') as f:
            f.write(file_data)

        return safe_filename

    def find_attachments_in_parts(self, parts, msg_id):
        """Recursively find attachments in message parts."""
        attachments = []
        if not parts:
            return attachments

        for part in parts:
            filename = part.get('filename', '')
            body = part.get('body', {})
            attachment_id = body.get('attachmentId')

            if filename and attachment_id:
                attachments.append({
                    'filename': filename,
                    'attachment_id': attachment_id,
                    'msg_id': msg_id,
                    'mime_type': part.get('mimeType', ''),
                    'size': body.get('size', 0),
                })

            # Recurse into nested parts (multipart messages)
            nested_parts = part.get('parts', [])
            if nested_parts:
                attachments.extend(
                    self.find_attachments_in_parts(nested_parts, msg_id)
                )

        return attachments

    def fetch_and_download_new_attachments(self, query=None, message_time_filter=None):
        """Main method: fetch all emails, download new attachments.

        Returns:
            tuple: (
                list of downloaded attachment metadata dicts,
                total emails checked,
                total new emails,
            )
        """
        self.status_callback("Fetching email list...")
        all_messages = self.fetch_all_message_ids(query=query)
        total_emails = len(all_messages)
        if query:
            self.status_callback(f"Found {total_emails} email(s) matching filter.")
        else:
            self.status_callback(f"Found {total_emails} total emails in account.")

        # Use filtered results directly (label filtering is handled via Gmail query)
        new_messages = all_messages
        new_count = len(new_messages)

        if new_count == 0:
            self.status_callback("No emails to process.", "success")
            return [], total_emails, 0

        self.status_callback(f"Processing {new_count} new emails...")

        downloaded_attachments = []

        for i, msg_data in enumerate(new_messages, 1):
            if self.should_stop():
                self.status_callback(
                    "Stop requested during Gmail download; leaving remaining emails untagged.",
                    "warning",
                )
                break
            msg_id = msg_data['id']
            self.status_callback(f"Checking email {i}/{new_count}...")

            try:
                msg = self.get_message_details(msg_id)
                if not _message_matches_time_filter(msg, message_time_filter):
                    self.status_callback(
                        "  Skipped: Gmail timestamp is outside the requested time window."
                    )
                    continue
                payload = msg.get('payload', {})

                # Get subject for logging
                headers = {
                    h['name']: h['value']
                    for h in payload.get('headers', [])
                }
                subject = headers.get('Subject', '(no subject)')
                from_header = headers.get('From', '')
                sender_email = _extract_sender_email(from_header)
                sender_header = from_header
                forwarded_sender_email, forwarded_sender_header = _extract_forwarded_sender(
                    payload,
                    msg.get('snippet', ''),
                )
                if forwarded_sender_email:
                    sender_email = forwarded_sender_email
                    sender_header = forwarded_sender_header or sender_header
                    self.status_callback(
                        f"  Forwarded sender detected: {sender_email}"
                    )

                # Find attachments
                parts = payload.get('parts', [])
                attachments = self.find_attachments_in_parts(parts, msg_id)

                # Also check top-level body (single-part messages)
                if not parts and payload.get('filename') and payload.get('body', {}).get('attachmentId'):
                    attachments.append({
                        'filename': payload['filename'],
                        'attachment_id': payload['body']['attachmentId'],
                        'msg_id': msg_id,
                        'mime_type': payload.get('mimeType', ''),
                        'size': payload.get('body', {}).get('size', 0),
                    })

                if attachments:
                    self.status_callback(
                        f"  Email: \"{subject}\" - {len(attachments)} attachment(s)"
                    )
                    download_completed = True
                    for att in attachments:
                        if self.should_stop():
                            self.status_callback(
                                "  Stop requested before email finished downloading; this email will remain untagged.",
                                "warning",
                            )
                            download_completed = False
                            break
                        saved_name = self.download_attachment(
                            msg_id, att['attachment_id'], att['filename']
                        )
                        self.status_callback(
                            f"    Downloaded: {saved_name}", "success"
                        )
                        downloaded_attachments.append({
                            'filename': saved_name,
                            'sender_email': sender_email,
                            'sender_header': sender_header,
                            'subject': subject,
                            'message_id': msg_id,
                        })
                    if download_completed:
                        try:
                            self._add_label_to_message(msg_id, self.processed_label_id)
                        except Exception as label_err:
                            self.status_callback(
                                f"    Warning: couldn't add label: {label_err}", "warning"
                            )
                    else:
                        break

            except Exception as e:
                self.status_callback(
                    f"  Error processing email {msg_id}: {e}", "error"
                )

        self.status_callback(
            f"Download complete: {len(downloaded_attachments)} attachments from "
            f"{new_count} new emails.", "success"
        )
        return downloaded_attachments, total_emails, new_count


HISTORY_FILENAME = 'invoice_history.csv'
HISTORY_FIELDNAMES = ['bill_no', 'po_number', 'vendor', 'invoice_date', 'downloaded_at', 'source_file']


class DriveHistoryClient:
    """Reads and writes the shared invoice_history.csv on Google Drive."""

    def __init__(self, creds, status_callback=None):
        self.service = build('drive', 'v3', credentials=creds, static_discovery=False)
        self.status_callback = status_callback or (lambda msg, tag=None: None)
        self._file_id = None

    def _get_or_create_file_id(self):
        if self._file_id:
            return self._file_id
        # Search for existing file in Drive root
        results = self.service.files().list(
            q=f"name='{HISTORY_FILENAME}' and 'root' in parents and trashed=false",
            spaces='drive',
            fields='files(id, name)',
        ).execute()
        files = results.get('files', [])
        if files:
            self._file_id = files[0]['id']
        else:
            # Create empty CSV with header
            content = ','.join(HISTORY_FIELDNAMES) + '\n'
            media = MediaIoBaseUpload(
                io.BytesIO(content.encode('utf-8')),
                mimetype='text/csv',
                resumable=False
            )
            meta = {'name': HISTORY_FILENAME, 'parents': ['root']}
            f = self.service.files().create(body=meta, media_body=media, fields='id').execute()
            self._file_id = f['id']
        return self._file_id

    def download_rows(self):
        """Download invoice_history.csv from Drive and return list of dicts."""
        try:
            file_id = self._get_or_create_file_id()
            request = self.service.files().get_media(fileId=file_id)
            buf = io.BytesIO()
            downloader = MediaIoBaseDownload(buf, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
            buf.seek(0)
            reader = csv.DictReader(io.TextIOWrapper(buf, encoding='utf-8'))
            return [{k: (v or '').strip() for k, v in row.items()} for row in reader]
        except Exception as e:
            self.status_callback(f"Warning: could not read remote invoice history ({e})", "warning")
            return []

    def upload_rows(self, rows):
        """Upload rows (list of dicts) as invoice_history.csv to Drive, replacing the existing file."""
        try:
            file_id = self._get_or_create_file_id()
            buf = io.StringIO()
            writer = csv.DictWriter(buf, fieldnames=HISTORY_FIELDNAMES)
            writer.writeheader()
            for row in rows:
                writer.writerow({k: row.get(k, '') for k in HISTORY_FIELDNAMES})
            media = MediaIoBaseUpload(
                io.BytesIO(buf.getvalue().encode('utf-8')),
                mimetype='text/csv',
                resumable=False
            )
            self.service.files().update(fileId=file_id, media_body=media).execute()
        except Exception as e:
            self.status_callback(f"Warning: could not update remote invoice history ({e})", "warning")
