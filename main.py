"""
Task Bot — watches Slack channels and an Outlook inbox, uses Claude to pull
out actionable tasks, and appends them to a Google Sheet.

Endpoints:
  GET  /             -> health check
  POST /webhook/slack   -> Slack Events API (handshake + message events)
  POST /webhook/outlook -> Microsoft Graph change notifications (handshake + new mail)
"""

import os
import re
import json
import base64
import hmac
import hashlib
import time
import threading
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

ANTHROPIC_API_KEY     = os.environ.get("ANTHROPIC_API_KEY", "")
ANTHROPIC_MODEL       = os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-4-6")
SLACK_SIGNING_SECRET  = os.environ.get("SLACK_SIGNING_SECRET", "")
SLACK_BOT_TOKEN       = os.environ.get("SLACK_BOT_TOKEN", "")
OUTLOOK_TENANT_ID     = os.environ.get("OUTLOOK_TENANT_ID", "")
OUTLOOK_CLIENT_ID     = os.environ.get("OUTLOOK_CLIENT_ID", "")
OUTLOOK_CLIENT_SECRET = os.environ.get("OUTLOOK_CLIENT_SECRET", "")
OUTLOOK_CLIENT_STATE  = os.environ.get("OUTLOOK_CLIENT_STATE", "task-bot-secret-123")
GOOGLE_SHEET_ID       = os.environ.get("GOOGLE_SHEET_ID", "")
GOOGLE_SERVICE_ACCOUNT_JSON = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "")

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def format_date(date_str: str) -> str:
    """Convert a date string from Claude into 'Monday (6/15/2026)'
    or 'Monday (6/15/2026) @ 11:30 AM PST' if a time is present.
    Shorthand like 'EOD', 'ASAP', 'EOW' are passed through as-is."""
    if not date_str:
        return ""
    # Pass through plain shorthand that shouldn't be parsed as a date
    SHORTHANDS = {"EOD", "ASAP", "EOW", "TBD"}
    if date_str.strip().upper() in SHORTHANDS:
        return date_str.strip().upper()
    # Try with time first (Claude returns 'YYYY-MM-DD h:mm AM/PM')
    for fmt in ("%Y-%m-%d %I:%M %p", "%Y-%m-%d %H:%M"):
        try:
            dt = datetime.strptime(date_str.strip(), fmt)
            month  = str(dt.month)
            day    = str(dt.day)
            year   = str(dt.year)
            hour   = str(dt.hour % 12 or 12)
            minute = dt.strftime("%M")
            ampm   = dt.strftime("%p")
            return f"{dt.strftime('%A')} ({month}/{day}/{year}) @ {hour}:{minute} {ampm} PST"
        except ValueError:
            pass
    # Try date only
    try:
        dt = datetime.strptime(date_str.strip(), "%Y-%m-%d")
        month = str(dt.month)
        day   = str(dt.day)
        year  = str(dt.year)
        return f"{dt.strftime('%A')} ({month}/{day}/{year})"
    except ValueError:
        pass
    return date_str


def format_assigned_time() -> str:
    """Current Pacific time as 'Monday (6/15/2026) @ 11:30 AM PST'."""
    now    = datetime.now(PACIFIC)
    month  = str(now.month)
    day    = str(now.day)
    year   = str(now.year)
    hour   = str(now.hour % 12 or 12)
    minute = now.strftime("%M")
    ampm   = now.strftime("%p")
    return f"{now.strftime('%A')} ({month}/{day}/{year}) @ {hour}:{minute} {ampm} PST"


def format_channel_name(raw: str) -> str:
    """Convert 'general-tasks' to 'General Tasks'."""
    return " ".join(word.capitalize() for word in raw.replace("-", " ").replace("_", " ").split())

# ---------------------------------------------------------------------------
# Google Sheets
# ---------------------------------------------------------------------------

def _service_account_info():
    raw = GOOGLE_SERVICE_ACCOUNT_JSON.strip()
    if not raw:
        raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_JSON is not set")
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return json.loads(base64.b64decode(raw).decode("utf-8"))


def append_task(task: dict, source: str, sender: str, channel: str = ""):
    """Append one task as a row in the Google Sheet."""
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds  = Credentials.from_service_account_info(_service_account_info(), scopes=scopes)
    gc     = gspread.authorize(creds)
    ws     = gc.open_by_key(GOOGLE_SHEET_ID).sheet1

    assigned = task.get("assigned_to") or "Admin"

    row = [
        task.get("description", ""),            # A  Action
        assigned,                               # B  Person Assigned
        format_date(task.get("due_date", "")),  # C  Due Date (PST)
        format_assigned_time(),                 # D  Time Assigned
        sender or "",                           # E  Assigner
        channel,                               # F  Channel
        source,                                # G  Source
        # H (Done checkbox) is left for the user to manage manually
    ]

    ws.append_row(row, value_input_option="USER_ENTERED")
    log.info("Added task from %s: %s", source, task.get("description", "")[:80])

# ---------------------------------------------------------------------------
# Claude task extraction
# ---------------------------------------------------------------------------

SYSTEM_PROMPT = (
    "You extract actionable tasks from workplace messages and emails. "
    "The text may be several messages from one person stitched together, and may "
    "contain zero, one, or several distinct tasks. "
    "Respond with ONLY a JSON object and nothing else — no prose, no code fences. "
    'Schema: {"tasks": [ '
    '{"description": string, "due_date": string, "assigned_to": string} ] }. '
    "Return an empty array only if the message is purely a greeting, a single emoji, "
    "or contains absolutely no actionable content. "
    "Err on the side of capturing tasks. If a message contains a request, a question "
    "directed at someone, something that implies action is needed, or anything that "
    "suggests someone should do something, treat it as a task. "
    "Merge fragments that describe the SAME task into one entry; only create separate "
    "entries for genuinely different tasks. "
    "For due_date: if the message says 'EOD', return exactly 'EOD'; "
    "if it says 'ASAP', return 'ASAP'; if it says 'end of week' or 'EOW', return 'EOW'. "
    "Otherwise, if a specific time is given, use 'YYYY-MM-DD h:mm AM/PM' "
    "(e.g. '2026-06-12 4:00 PM'); if only a day is given, use 'YYYY-MM-DD'; "
    "if no deadline is stated, use ''. "
    "assigned_to: use the name of the person responsible if mentioned or clearly implied. "
    "If the message is directed at a specific person (e.g. '@John can you...'), assign it to them. "
    "If truly unclear, leave assigned_to as '' and the system will assign it to Admin. "
    "Resolve relative dates and times (e.g. 'Friday', 'tomorrow', '4pm Friday', "
    "'EOD') against the provided current date and time."
)


def extract_tasks(text: str):
    """Return a list of task dicts found in the text (possibly empty)."""
    if not text or not text.strip():
        return []

    now      = datetime.now(PACIFIC)
    today    = now.strftime("%A, %Y-%m-%d %I:%M %p %Z")
    user_msg = f"Current date and time: {today}\n\nMessage(s):\n{text}"

    try:
        resp = httpx.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key":         ANTHROPIC_API_KEY,
                "anthropic-version": "2023-06-01",
                "content-type":      "application/json",
            },
            json={
                "model":      ANTHROPIC_MODEL,
                "max_tokens": 800,
                "system":     SYSTEM_PROMPT,
                "messages":   [{"role": "user", "content": user_msg}],
            },
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        out  = "".join(
            b.get("text", "") for b in data.get("content", []) if b.get("type") == "text"
        ).strip()

        if out.startswith("```"):
            out = out.strip("`")
            out = out[4:].strip() if out.lower().startswith("json") else out.strip()

        parsed = json.loads(out)
        tasks  = [t for t in parsed.get("tasks", []) if t.get("description")]
        if not tasks:
            log.info("Claude found no tasks in message: %s", text[:120])
        return tasks

    except Exception as e:
        log.exception("Claude extraction failed: %s", e)
        return []

# ---------------------------------------------------------------------------
# Slack
# ---------------------------------------------------------------------------

def verify_slack_signature(body: bytes, headers) -> bool:
    if not SLACK_SIGNING_SECRET:
        return True
    ts  = headers.get("x-slack-request-timestamp", "")
    sig = headers.get("x-slack-signature", "")
    if not ts or not sig:
        return False
    try:
        if abs(time.time() - int(ts)) > 300:
            return False
    except ValueError:
        return False
    base = f"v0:{ts}:{body.decode('utf-8')}".encode()
    mine = "v0=" + hmac.new(SLACK_SIGNING_SECRET.encode(), base, hashlib.sha256).hexdigest()
    return hmac.compare_digest(mine, sig)


def resolve_slack_user(user_id: str) -> str:
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


def resolve_slack_channel(channel_id: str) -> str:
    """Convert a Slack channel ID to a clean name like 'General Tasks'."""
    if not channel_id or not SLACK_BOT_TOKEN:
        return ""
    try:
        r = httpx.get(
            "https://slack.com/api/conversations.info",
            headers={"Authorization": f"Bearer {SLACK_BOT_TOKEN}"},
            params={"channel": channel_id},
            timeout=10,
        )
        d = r.json()
        if d.get("ok"):
            raw_name = d["channel"].get("name", "")
            if raw_name:
                return format_channel_name(raw_name)
    except Exception:
        log.warning("Could not resolve Slack channel %s", channel_id)
    return ""


_USER_RE    = re.compile(r"<@([UW][A-Z0-9]+)(?:\|[^>]+)?>")
_CHANNEL_RE = re.compile(r"<#[A-Z0-9]+\|([^>]+)>")
_SPECIAL_RE = re.compile(r"<!([^>|]+)(?:\|([^>]+))?>")
_LINK_RE    = re.compile(r"<(https?://[^|>]+)(?:\|([^>]+))?>")


def clean_slack_text(text: str) -> str:
    if not text:
        return ""
    text = _USER_RE.sub(lambda m: f"@{resolve_slack_user(m.group(1))}", text)
    text = _CHANNEL_RE.sub(lambda m: f"#{m.group(1)}", text)
    text = _SPECIAL_RE.sub(lambda m: f"@{m.group(2) or m.group(1)}", text)
    text = _LINK_RE.sub(lambda m: m.group(2) or m.group(1), text)
    return text


# --- Message buffering -----------------------------------------------------

_BUFFER_WINDOW = float(os.environ.get("SLACK_BUFFER_SECONDS", "8"))
_buffers      = {}
_timers       = {}
_buffer_lock  = threading.Lock()


def buffer_slack_message(event: dict):
    text = clean_slack_text(event.get("text", ""))
    if not text.strip():
        return

    channel = event.get("channel", "")
    user_id = event.get("user", "")
    thread  = event.get("thread_ts", event.get("ts", ""))
    key     = f"{channel}:{user_id}:{thread}"

    with _buffer_lock:
        _buffers.setdefault(key, []).append(text)
        if key in _timers:
            _timers[key].cancel()
        timer = threading.Timer(_BUFFER_WINDOW, flush_buffer, args=(key, user_id, channel))
        timer.daemon = True
        _timers[key] = timer
        timer.start()


def flush_buffer(key: str, user_id: str, channel_id: str = ""):
    with _buffer_lock:
        texts = _buffers.pop(key, [])
        _timers.pop(key, None)
    if not texts:
        return
    combined = "\n".join(texts)
    sender   = resolve_slack_user(user_id)
    channel  = resolve_slack_channel(channel_id)
    for task in extract_tasks(combined):
        append_task(task, source="Slack", sender=sender, channel=channel)


@app.post("/webhook/slack")
async def slack_webhook(request: Request, background: BackgroundTasks):
    body = await request.body()
    try:
        payload = json.loads(body)
    except json.JSONDecodeError:
        return Response(status_code=400)

    if payload.get("type") == "url_verification":
        return {"challenge": payload.get("challenge", "")}

    if not verify_slack_signature(body, request.headers):
        log.warning("Rejected Slack request: bad signature")
        return Response(status_code=403)

    if payload.get("type") == "event_callback":
        event      = payload.get("event", {})
        event_type = event.get("type", "")
        is_message = event_type == "message"
        is_reply   = event_type == "message" and event.get("thread_ts") and event.get("thread_ts") != event.get("ts")

        if (
            (is_message or is_reply)
            and not event.get("bot_id")
            and not event.get("subtype")
        ):
            log.info("Queuing Slack message for processing: %s", event.get("text", "")[:80])
            background.add_task(buffer_slack_message, event)

    return Response(status_code=200)

# ---------------------------------------------------------------------------
# Outlook / Microsoft Graph
# ---------------------------------------------------------------------------

def graph_token() -> str:
    r = httpx.post(
        f"https://login.microsoftonline.com/{OUTLOOK_TENANT_ID}/oauth2/v2.0/token",
        data={
            "grant_type":    "client_credentials",
            "client_id":     OUTLOOK_CLIENT_ID,
            "client_secret": OUTLOOK_CLIENT_SECRET,
            "scope":         "https://graph.microsoft.com/.default",
        },
        timeout=30,
    )
    r.raise_for_status()
    return r.json()["access_token"]


def handle_outlook_notification(note: dict):
    try:
        resource = note.get("resource", "").lstrip("/")
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
        msg     = r.json()
        subject = msg.get("subject", "")
        preview = msg.get("bodyPreview", "")
        sender  = ((msg.get("from") or {}).get("emailAddress") or {}).get("name", "")
        for task in extract_tasks(f"Subject: {subject}\n\n{preview}"):
            append_task(task, source="Outlook", sender=sender, channel="")
    except Exception as e:
        log.exception("Outlook handling failed: %s", e)


@app.post("/webhook/outlook")
async def outlook_webhook(request: Request, background: BackgroundTasks):
    token = request.query_params.get("validationToken")
    if token:
        return Response(content=token, media_type="text/plain")

    body = await request.body()
    try:
        payload = json.loads(body)
    except json.JSONDecodeError:
        return Response(status_code=202)

    for note in payload.get("value", []):
        if OUTLOOK_CLIENT_STATE and note.get("clientState") != OUTLOOK_CLIENT_STATE:
            log.warning("Skipping Outlook notification with bad clientState")
            continue
        background.add_task(handle_outlook_notification, note)

    return Response(status_code=202)


@app.get("/")
def health():
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
