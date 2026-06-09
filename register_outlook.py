"""
Register (or refresh) the Outlook -> Task Bot webhook with Microsoft Graph.

This replaces the fiddly Postman steps. Run it locally once your Railway app
is deployed and live. It is safe to re-run: it deletes any existing Task Bot
subscription first, then creates a fresh one.

Microsoft Graph message subscriptions expire after at most ~3 days (4230
minutes), so you must re-run this (or set up automatic renewal) periodically.

Usage:
    pip install httpx
    # set these in your shell first:
    export OUTLOOK_TENANT_ID=...
    export OUTLOOK_CLIENT_ID=...
    export OUTLOOK_CLIENT_SECRET=...
    export RAILWAY_URL=https://task-bot-production.up.railway.app
    export OUTLOOK_USER=you@yourcompany.com   # the mailbox to watch
    python register_outlook.py
"""

import os
import httpx
from datetime import datetime, timedelta, timezone

TENANT = os.environ["OUTLOOK_TENANT_ID"]
CLIENT_ID = os.environ["OUTLOOK_CLIENT_ID"]
CLIENT_SECRET = os.environ["OUTLOOK_CLIENT_SECRET"]
RAILWAY_URL = os.environ["RAILWAY_URL"].rstrip("/")
USER = os.environ["OUTLOOK_USER"]
CLIENT_STATE = os.environ.get("OUTLOOK_CLIENT_STATE", "task-bot-secret-123")

NOTIFY_URL = f"{RAILWAY_URL}/webhook/outlook"

token = httpx.post(
    f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token",
    data={
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
    },
    timeout=30,
).json()["access_token"]

headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

# 1) Remove any old Task Bot subscriptions pointing at our URL.
existing = httpx.get("https://graph.microsoft.com/v1.0/subscriptions", headers=headers).json()
for sub in existing.get("value", []):
    if sub.get("notificationUrl") == NOTIFY_URL:
        httpx.delete(f"https://graph.microsoft.com/v1.0/subscriptions/{sub['id']}", headers=headers)
        print(f"Deleted old subscription {sub['id']}")

# 2) Create the new one. Max lifetime for messages is 4230 minutes.
expiry = (datetime.now(timezone.utc) + timedelta(minutes=4230)).strftime("%Y-%m-%dT%H:%M:%SZ")

resp = httpx.post(
    "https://graph.microsoft.com/v1.0/subscriptions",
    headers=headers,
    json={
        "changeType": "created",
        "notificationUrl": NOTIFY_URL,
        "resource": f"users/{USER}/mailFolders/inbox/messages",
        "expirationDateTime": expiry,
        "clientState": CLIENT_STATE,
    },
    timeout=30,
)

print("Status:", resp.status_code)
print(resp.text)
if resp.status_code == 201:
    print("\nSuccess. Subscription expires at", expiry)
    print("Re-run this script before then to keep it alive.")
