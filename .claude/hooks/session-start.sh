#!/bin/bash
set -euo pipefail

# Only run in remote (web) sessions
if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

echo "Setting up NotebookLM authentication..."

# Install Google auth libraries required for NotebookLM authentication
pip install google-auth google-auth-oauthlib google-auth-httplib2 requests --quiet

# Write Google service account credentials to a temp file if provided
if [ -n "${GOOGLE_CREDENTIALS_JSON:-}" ]; then
  CREDS_FILE="/tmp/google_credentials.json"
  echo "$GOOGLE_CREDENTIALS_JSON" > "$CREDS_FILE"
  echo "export GOOGLE_APPLICATION_CREDENTIALS=$CREDS_FILE" >> "$CLAUDE_ENV_FILE"
  echo "Google service account credentials configured."
fi

# Expose the NotebookLM cookie (used by the unofficial API) if provided
if [ -n "${NOTEBOOKLM_COOKIE:-}" ]; then
  echo "export NOTEBOOKLM_COOKIE=$NOTEBOOKLM_COOKIE" >> "$CLAUDE_ENV_FILE"
  echo "NotebookLM cookie authentication configured."
fi

if [ -z "${GOOGLE_CREDENTIALS_JSON:-}" ] && [ -z "${NOTEBOOKLM_COOKIE:-}" ]; then
  echo "Warning: No NotebookLM credentials found."
  echo "Set GOOGLE_CREDENTIALS_JSON (service account JSON) or NOTEBOOKLM_COOKIE in your project secrets."
fi

echo "NotebookLM setup complete."
