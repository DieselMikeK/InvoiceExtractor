"""Minimal test to verify Gmail API credentials work."""
import os
import pickle
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CLIENT_SECRET = os.path.join(BASE_DIR, 'client_secret.json')
TOKEN_FILE = os.path.join(BASE_DIR, 'token.pickle')


def authenticate():
    creds = None
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, 'rb') as f:
            creds = pickle.load(f)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'wb') as f:
            pickle.dump(creds, f)

    return creds


def main():
    print("Authenticating with Gmail API...")
    creds = authenticate()
    print("Authentication successful!")

    service = build('gmail', 'v1', credentials=creds)

    # Get profile info
    profile = service.users().getProfile(userId='me').execute()
    print(f"Connected to: {profile['emailAddress']}")
    print(f"Total messages: {profile['messagesTotal']}")

    # Fetch the 5 most recent messages (just subjects)
    results = service.users().messages().list(userId='me', maxResults=5).execute()
    messages = results.get('messages', [])

    if not messages:
        print("No messages found.")
    else:
        print(f"\nMost recent {len(messages)} emails:")
        for msg_data in messages:
            msg = service.users().messages().get(
                userId='me', id=msg_data['id'], format='metadata',
                metadataHeaders=['Subject', 'From']
            ).execute()
            headers = {h['name']: h['value'] for h in msg['payload'].get('headers', [])}
            subject = headers.get('Subject', '(no subject)')
            sender = headers.get('From', '(unknown)')
            print(f"  - From: {sender}")
            print(f"    Subject: {subject}")

    print("\nGmail API is working correctly!")


if __name__ == '__main__':
    main()
