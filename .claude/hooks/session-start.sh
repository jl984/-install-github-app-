#!/bin/bash
set -euo pipefail

# Only run in remote (web) sessions
if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

echo "Setting up NotebookLM authentication..."

# Install Google auth libraries
pip install google-auth google-auth-oauthlib google-auth-httplib2 requests --quiet

TOKEN_FILE="$HOME/.claude/notebooklm_token.json"
COOKIE_FILE="$CLAUDE_PROJECT_DIR/.claude/notebooklm_cookie"

# Option 1: OAuth token saved by authenticate.py
if [ -f "$TOKEN_FILE" ]; then
  echo "export NOTEBOOKLM_TOKEN_FILE=$TOKEN_FILE" >> "$CLAUDE_ENV_FILE"
  echo "OAuth token loaded from $TOKEN_FILE"

# Option 2: Cookie stored in .claude/notebooklm_cookie file
elif [ -f "$COOKIE_FILE" ]; then
  COOKIE_VALUE=$(cat "$COOKIE_FILE")
  echo "export NOTEBOOKLM_COOKIE=$COOKIE_VALUE" >> "$CLAUDE_ENV_FILE"
  echo "Cookie loaded from $COOKIE_FILE"

# Option 3: Service account JSON passed as environment variable
elif [ -n "${GOOGLE_CREDENTIALS_JSON:-}" ]; then
  CREDS_FILE="/tmp/google_credentials.json"
  echo "$GOOGLE_CREDENTIALS_JSON" > "$CREDS_FILE"
  echo "export GOOGLE_APPLICATION_CREDENTIALS=$CREDS_FILE" >> "$CLAUDE_ENV_FILE"
  echo "Service account credentials configured."

# Option 4: Cookie passed as environment variable
elif [ -n "${NOTEBOOKLM_COOKIE:-}" ]; then
  echo "export NOTEBOOKLM_COOKIE=$NOTEBOOKLM_COOKIE" >> "$CLAUDE_ENV_FILE"
  echo "Cookie-based authentication configured."

else
  echo "Warning: No NotebookLM credentials found."
  echo "Run 'python authenticate.py' to log in via browser, or"
  echo "save your cookie to .claude/notebooklm_cookie"
fi

echo "NotebookLM setup complete."
