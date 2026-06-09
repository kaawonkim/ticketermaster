"""
Task Bot — watches Slack channels and an Outlook inbox, uses Claude to pull
out actionable tasks, and appends them to a Google Sheet.

Endpoints:
  GET  /                 -> health check
  POST /webhook/slack    -> Slack Events API (handshake + message events)
  POST /webhook/outlook  -> Microsoft Graph change notifications (handshake + new mail)
"""

import os
import json
import base64
import hmac
import hashlib
import time
import logging
from datetime import datetime

from zoneinfo import ZoneInfo

import httpx
import gspread
from google.oauth2.service_account import Credentials
from fastapi import FastAPI, Request, BackgroundTasks, Response

PACIFIC = ZoneInfo("America/Los_Angeles")  # auto-handles PST/PDT

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("task-bot")

app = FastAPI()

# ---------------------------------------------------------------------------
# Configuration (all set as environment variables in Railway)
# ---------------------------------------------------------------------------
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
ANTHROPIC_MODEL = os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-4-6")

SLACK_SIGNING_SECRET = os.environ.get("SLACK_SIGNING_SECRET", "")
SLACK_BOT_TOKEN = os.environ.get("SLACK_BOT_TOKEN", "")  # optional, xoxb-... (for sender names)

OUTLOOK_TENANT_ID = os.environ.get("OUTLOOK_TENANT_ID", "")
OUTLOOK_CLIENT_ID = os.environ.get("OUTLOOK_CLIENT_ID", "")
OUTLOOK_CLIENT_SECRET = os.environ.get("OUTLOOK_CLIENT_SECRET", "")
OUTLOOK_CLIENT_STATE = os.environ.get("OUTLOOK_CLIENT_STATE", "task-bot-secret-123")

GOOGLE_SHEET_ID = os.environ.get("GOOGLE_SHEET_ID", "")
GOOGLE_SERVICE_ACCOUNT_JSON = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "")


# ---------------------------------------------------------------------------
# Google Sheets
# ---------------------------------------------------------------------------
def _service_account_info():
    raw = GOOGLE_SERVICE_ACCOUNT_JSON.strip()
    if not raw:
        raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_JSON is not set")
    # Accept either the raw JSON blob or a base64-encoded blob.
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return json.loads(base64.b64decode(raw).decode("utf-8"))


def append_task(task: dict, source: str, sender: str):
    """Append one task as a row in the Google Sheet."""
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(_service_account_info(), scopes=scopes)
    gc = gspread.authorize(creds)
    ws = gc.open_by_key(GOOGLE_SHEET_ID).sheet1

    row = [
        task.get("description", ""),                              # A  Item
        task.get("assigned_to") or sender or "",                 # B  Person Assigned
        task.get("due_date", ""),                                # C  Due Date
        datetime.now(PACIFIC).strftime("%Y-%m-%d %I:%M %p %Z"),  # D  Time Assigned
        sender or "",                                            # E  Assigner
        task.get("status", "To Do"),                             # F  Action
        source,                                                  # G  Source
    ]
    ws.append_row(row, value_input_option="USER_ENTERED")
    log.info("Added task from %s: %s", source, task.get("description", "")[:80])


# ---------------------------------------------------------------------------
# Claude task extraction
# ---------------------------------------------------------------------------
SYSTEM_PROMPT = (
    "You extract actionable tasks from workplace messages and emails. "
    "Respond with ONLY a JSON object and nothing else — no prose, no code fences. "
    "Schema: "
    '{"is_task": boolean, '
    '"description": string, '
    '"due_date": string (YYYY-MM-DD, or "" if none stated), '
    '"assigned_to": string (the person responsible, or "" if unclear)}. '
    "Set is_task to false for greetings, small talk, pure FYIs, newsletters, "
    "or anything with no clear action someone must take. Resolve relative dates "
    "(e.g. 'Friday', 'tomorrow', 'EOD') against the provided current date."
)


def extract_task(text: str):
    """Return a task dict if the message contains one, else None."""
    if not text or not text.strip():
        return None

    today = datetime.now(PACIFIC).strftime("%Y-%m-%d")
    user_msg = f"Current date: {today}\n\nMessage:\n{text}"

    try:
        resp = httpx.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": ANTHROPIC_API_KEY,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json",
            },
            json={
                "model": ANTHROPIC_MODEL,
                "max_tokens": 400,
                "system": SYSTEM_PROMPT,
                "messages": [{"role": "user", "content": user_msg}],
            },
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        out = "".join(
            b.get("text", "") for b in data.get("content", []) if b.get("type") == "text"
        ).strip()
        # Strip accidental code fences just in case.
        if out.startswith("```"):
            out = out.strip("`")
            out = out[4:].strip() if out.lower().startswith("json") else out.strip()
        parsed = json.loads(out)
        return parsed if parsed.get("is_task") else None
    except Exception as e:
        log.exception("Claude extraction failed: %s", e)
        return None


# ---------------------------------------------------------------------------
# Slack
# ---------------------------------------------------------------------------
def verify_slack_signature(body: bytes, headers) -> bool:
    """Validate the X-Slack-Signature header. Used for real events only."""
    if not SLACK_SIGNING_SECRET:
        return True  # not configured yet — don't block
    ts = headers.get("x-slack-request-timestamp", "")
    sig = headers.get("x-slack-signature", "")
    if not ts or not sig:
        return False
    try:
        if abs(time.time() - int(ts)) > 300:  # reject requests older than 5 min
            return False
    except ValueError:
        return False
    base = f"v0:{ts}:{body.decode('utf-8')}".encode()
    mine = "v0=" + hmac.new(SLACK_SIGNING_SECRET.encode(), base, hashlib.sha256).hexdigest()
    return hmac.compare_digest(mine, sig)


def resolve_slack_user(user_id: str) -> str:
    """Turn a Slack user ID into a real name (needs SLACK_BOT_TOKEN). Falls back to the ID."""
    if not user_id or not SLACK_BOT_TOKEN:
        return user_id or ""
    try:
        r = httpx.get(
            "https://slack.com/api/users.info",
            headers={"Authorization": f"Bearer {SLACK_BOT_TOKEN}"},
            params={"user": user_id},
            timeout=10,
        )
        d = r.json()
        if d.get("ok"):
            profile = d["user"].get("profile", {})
            return profile.get("real_name") or d["user"].get("name") or user_id
    except Exception:
        log.warning("Could not resolve Slack user %s", user_id)
    return user_id


def handle_slack_message(event: dict):
    text = event.get("text", "")
    sender = resolve_slack_user(event.get("user", ""))
    task = extract_task(text)
    if task:
        append_task(task, source="Slack", sender=sender)


@app.post("/webhook/slack")
async def slack_webhook(request: Request, background: BackgroundTasks):
    body = await request.body()
    try:
        payload = json.loads(body)
    except json.JSONDecodeError:
        return Response(status_code=400)

    # 1) URL verification handshake. This is the step that was failing.
    #    We echo the challenge straight back, no signature needed.
    if payload.get("type") == "url_verification":
        return {"challenge": payload.get("challenge", "")}

    # 2) Verify the signature before acting on real events.
    if not verify_slack_signature(body, request.headers):
        log.warning("Rejected Slack request: bad signature")
        return Response(status_code=403)

    # 3) Process message events in the background so Slack gets a fast 200.
    if payload.get("type") == "event_callback":
        event = payload.get("event", {})
        if (
            event.get("type") == "message"
            and not event.get("bot_id")       # ignore other bots
            and not event.get("subtype")      # ignore edits/joins/etc.
        ):
            background.add_task(handle_slack_message, event)

    return Response(status_code=200)


# ---------------------------------------------------------------------------
# Outlook / Microsoft Graph
# ---------------------------------------------------------------------------
def graph_token() -> str:
    r = httpx.post(
        f"https://login.microsoftonline.com/{OUTLOOK_TENANT_ID}/oauth2/v2.0/token",
        data={
            "grant_type": "client_credentials",
            "client_id": OUTLOOK_CLIENT_ID,
            "client_secret": OUTLOOK_CLIENT_SECRET,
            "scope": "https://graph.microsoft.com/.default",
        },
        timeout=30,
    )
    r.raise_for_status()
    return r.json()["access_token"]


def handle_outlook_notification(note: dict):
    try:
        resource = note.get("resource", "").lstrip("/")  # e.g. Users/{id}/Messages/{id}
        if not resource:
            return
        token = graph_token()
        r = httpx.get(
            f"https://graph.microsoft.com/v1.0/{resource}",
            headers={"Authorization": f"Bearer {token}"},
            params={"$select": "subject,bodyPreview,from"},
            timeout=30,
        )
        r.raise_for_status()
        msg = r.json()
        subject = msg.get("subject", "")
        preview = msg.get("bodyPreview", "")
        sender = ((msg.get("from") or {}).get("emailAddress") or {}).get("name", "")
        task = extract_task(f"Subject: {subject}\n\n{preview}")
        if task:
            append_task(task, source="Outlook", sender=sender)
    except Exception as e:
        log.exception("Outlook handling failed: %s", e)


@app.post("/webhook/outlook")
async def outlook_webhook(request: Request, background: BackgroundTasks):
    # 1) Subscription validation handshake — Graph sends ?validationToken=...
    #    when the subscription is created. Echo it back as plain text.
    token = request.query_params.get("validationToken")
    if token:
        return Response(content=token, media_type="text/plain")

    body = await request.body()
    try:
        payload = json.loads(body)
    except json.JSONDecodeError:
        return Response(status_code=202)

    for note in payload.get("value", []):
        # Optional security check: ignore notifications without our clientState.
        if OUTLOOK_CLIENT_STATE and note.get("clientState") != OUTLOOK_CLIENT_STATE:
            log.warning("Skipping Outlook notification with bad clientState")
            continue
        background.add_task(handle_outlook_notification, note)

    return Response(status_code=202)


@app.get("/")
def health():
    # Diagnostic: shows which variables the running app can actually see.
    # Reports only presence + length, never the secret values themselves.
    expected = [
        "ANTHROPIC_API_KEY",
        "SLACK_SIGNING_SECRET",
        "SLACK_BOT_TOKEN",
        "GOOGLE_SHEET_ID",
        "GOOGLE_SERVICE_ACCOUNT_JSON",
    ]
    env_status = {
        k: {"present": bool(os.environ.get(k, "")), "length": len(os.environ.get(k, ""))}
        for k in expected
    }
    return {"status": "ok", "service": "task-bot", "env": env_status}
