"""Gmail API client for fetching emails and downloading attachments."""
import os
import base64
import pickle
import time
from datetime import datetime
from google.auth.exceptions import RefreshError
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ['https://www.googleapis.com/auth/gmail.modify']

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


class GmailClient:
    def __init__(
        self,
        base_dir,
        status_callback=None,
        data_dir=None,
        invoices_dir=None,
        expected_email=None
    ):
        self.base_dir = base_dir
        self.data_dir = data_dir or base_dir
        self.status_callback = status_callback or (lambda msg, tag=None: None)
        self.client_secret = os.path.join(self.data_dir, 'client_secret.json')
        self.token_file = os.path.join(self.data_dir, 'token.pickle')
        self.invoices_dir = invoices_dir or os.path.join(self.data_dir, 'invoices')
        self.expected_email = str(expected_email or '').strip().lower()
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
                creds = flow.run_local_server(port=0)
                should_persist_token = True

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

    def fetch_and_download_new_attachments(self, query=None):
        """Main method: fetch all emails, download new attachments.

        Returns:
            tuple: (list of new filenames downloaded, total emails checked, total new emails)
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

        downloaded_files = []

        for i, msg_data in enumerate(new_messages, 1):
            msg_id = msg_data['id']
            self.status_callback(f"Checking email {i}/{new_count}...")

            try:
                msg = self.get_message_details(msg_id)
                payload = msg.get('payload', {})

                # Get subject for logging
                headers = {
                    h['name']: h['value']
                    for h in payload.get('headers', [])
                }
                subject = headers.get('Subject', '(no subject)')

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
                    for att in attachments:
                        saved_name = self.download_attachment(
                            msg_id, att['attachment_id'], att['filename']
                        )
                        self.status_callback(
                            f"    Downloaded: {saved_name}", "success"
                        )
                        downloaded_files.append(saved_name)

                # Mark email as processed via Gmail label
                try:
                    self._add_label_to_message(msg_id, self.processed_label_id)
                except Exception as label_err:
                    self.status_callback(
                        f"    Warning: couldn't add label: {label_err}", "warning"
                    )

            except Exception as e:
                self.status_callback(
                    f"  Error processing email {msg_id}: {e}", "error"
                )
                # Best effort label to avoid retrying broken emails forever
                try:
                    self._add_label_to_message(msg_id, self.processed_label_id)
                except Exception:
                    pass  # Best effort

        self.status_callback(
            f"Download complete: {len(downloaded_files)} attachments from "
            f"{new_count} new emails.", "success"
        )
        return downloaded_files, total_emails, new_count
