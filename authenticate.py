#!/usr/bin/env python3
"""
Run this script once to authenticate with Google for NotebookLM.
A browser window will open asking you to log in with your Google account.
Credentials are saved to ~/.claude/notebooklm_token.json for future sessions.
"""
import json
import os
import sys

CREDENTIALS_FILE = os.path.expanduser("~/.claude/notebooklm_token.json")
CLIENT_SECRETS_FILE = os.path.join(os.path.dirname(__file__), "client_secrets.json")

SCOPES = [
    "https://www.googleapis.com/auth/userinfo.email",
    "openid",
]


def authenticate():
    try:
        from google_auth_oauthlib.flow import InstalledAppFlow
        from google.oauth2.credentials import Credentials
        from google.auth.transport.requests import Request
    except ImportError:
        print("Installing required packages...")
        os.system(
            "pip install google-auth google-auth-oauthlib google-auth-httplib2 --quiet"
        )
        from google_auth_oauthlib.flow import InstalledAppFlow
        from google.oauth2.credentials import Credentials
        from google.auth.transport.requests import Request

    # Reuse existing token if still valid
    creds = None
    if os.path.exists(CREDENTIALS_FILE):
        creds = Credentials.from_authorized_user_file(CREDENTIALS_FILE, SCOPES)

    if creds and creds.valid:
        print(f"Already authenticated. Token saved at: {CREDENTIALS_FILE}")
        return

    if creds and creds.expired and creds.refresh_token:
        print("Refreshing expired token...")
        creds.refresh(Request())
    else:
        if not os.path.exists(CLIENT_SECRETS_FILE):
            print(
                "Error: client_secrets.json not found.\n"
                "Download it from Google Cloud Console:\n"
                "  1. Go to https://console.cloud.google.com/apis/credentials\n"
                "  2. Create an OAuth 2.0 Client ID (Desktop app)\n"
                "  3. Download JSON and save as client_secrets.json in this directory"
            )
            sys.exit(1)

        print("Opening browser for Google login...")
        flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRETS_FILE, SCOPES)
        creds = flow.run_local_server(port=0, prompt="consent")

    os.makedirs(os.path.dirname(CREDENTIALS_FILE), exist_ok=True)
    with open(CREDENTIALS_FILE, "w") as f:
        f.write(creds.to_json())
    os.chmod(CREDENTIALS_FILE, 0o600)

    print(f"\nAuthentication successful!")
    print(f"Token saved to: {CREDENTIALS_FILE}")
    print("\nYour NotebookLM sessions will now authenticate automatically.")


if __name__ == "__main__":
    authenticate()
