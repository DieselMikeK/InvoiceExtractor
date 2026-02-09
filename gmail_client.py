"""Gmail API client for fetching emails and downloading attachments."""
import os
import base64
import json
import pickle
import time
from datetime import datetime
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ['https://www.googleapis.com/auth/gmail.modify']

PROCESSED_LABEL_NAME = "InvoiceExtractor-Processed"


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
    def __init__(self, base_dir, status_callback=None, data_dir=None, log_file=None, invoices_dir=None):
        self.base_dir = base_dir
        self.data_dir = data_dir or base_dir
        self.status_callback = status_callback or (lambda msg, tag=None: None)
        self.client_secret = os.path.join(self.data_dir, 'client_secret.json')
        self.token_file = os.path.join(self.data_dir, 'token.pickle')
        self.invoices_dir = invoices_dir or os.path.join(self.data_dir, 'invoices')
        self.log_file = log_file or os.path.join(self.base_dir, 'processed_log.json')
        self.service = None

        os.makedirs(self.invoices_dir, exist_ok=True)

    def authenticate(self):
        """Authenticate with Gmail API, caching token for future runs."""
        creds = None
        if os.path.exists(self.token_file):
            with open(self.token_file, 'rb') as f:
                creds = pickle.load(f)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                self.status_callback("Refreshing authentication token...")
                creds.refresh(Request())
            else:
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

            with open(self.token_file, 'wb') as f:
                pickle.dump(creds, f)

        self.service = build('gmail', 'v1', credentials=creds)

        profile = retry_with_backoff(
            lambda: self.service.users().getProfile(userId='me').execute(),
            status_callback=self.status_callback
        )
        self.status_callback(f"Connected to: {profile['emailAddress']}", "success")

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

    def load_processed_log(self):
        """Load the tracking log of already-processed emails and invoices."""
        if os.path.exists(self.log_file):
            with open(self.log_file, 'r') as f:
                return json.load(f)
        return {"processed_emails": {}, "processed_invoices": {}}

    def save_processed_log(self, log):
        """Save the tracking log."""
        with open(self.log_file, 'w') as f:
            json.dump(log, f, indent=2)

    def fetch_all_message_ids(self):
        """Fetch all message IDs from the inbox."""
        all_messages = []
        page_token = None

        while True:
            def list_messages(pt=page_token):
                kwargs = {'userId': 'me', 'maxResults': 500}
                if pt:
                    kwargs['pageToken'] = pt
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

    def fetch_and_download_new_attachments(self):
        """Main method: fetch all emails, download new attachments.

        Returns:
            tuple: (list of new filenames downloaded, total emails checked, total new emails)
        """
        log = self.load_processed_log()
        processed_emails = log.get("processed_emails", {})

        self.status_callback("Fetching email list...")
        all_messages = self.fetch_all_message_ids()
        total_emails = len(all_messages)
        self.status_callback(f"Found {total_emails} total emails in account.")

        # Filter out already-processed emails
        new_messages = [m for m in all_messages if m['id'] not in processed_emails]
        new_count = len(new_messages)

        if new_count == 0:
            self.status_callback("No new emails to process.", "success")
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

                # Mark email as processed (local log + Gmail label)
                processed_emails[msg_id] = datetime.now().isoformat()
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
                # Still mark it processed to avoid retrying broken emails forever
                processed_emails[msg_id] = datetime.now().isoformat()
                try:
                    self._add_label_to_message(msg_id, self.processed_label_id)
                except Exception:
                    pass  # Best effort

        # Save updated log
        log["processed_emails"] = processed_emails
        self.save_processed_log(log)

        self.status_callback(
            f"Download complete: {len(downloaded_files)} attachments from "
            f"{new_count} new emails.", "success"
        )
        return downloaded_files, total_emails, new_count
