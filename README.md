# Task Bot

Watches your Slack channels and Outlook inbox, uses Claude to detect actionable
tasks, and writes them to a Google Sheet with: Status, Due Date, Description,
Assigned To, Source, Sender, Date Added.

## How it actually works (and a note on "MCP")

This is a **webhook** automation, not an MCP server. That distinction matters:
MCP lets *you* ask Claude to go check things on demand. What you want —
tasks captured **automatically** as messages arrive — needs a small always-on
server that Slack and Outlook push events to. That's what this is. The flow:

```
Slack message  ─┐
                 ├─►  Railway server  ─►  Claude (extract task)  ─►  Google Sheet
Outlook email  ─┘
```

---

## THE KEY FIX (why the Slack "challenge" was failing)

Slack verifies a webhook URL by sending a `challenge` value and expecting your
server to echo it back. beeceptor and a blank URL can't do that — only this
deployed app can (see `/webhook/slack` in `main.py`).

**So the order in the original guide was the problem. The correct order is:**

1. Deploy this code to Railway FIRST.
2. THEN paste the Railway URL into Slack's Event Subscriptions.

Also note: current Slack will not let you *save* Event Subscriptions with a
blank Request URL. So "leave it blank for now" no longer works — you must have
the live URL ready, which is exactly why you deploy first.

---

## Setup

### Step 1 — Anthropic API key
Same as before: console.anthropic.com → API Keys → Create Key → save as
`ANTHROPIC_API_KEY`.

### Step 2 — Slack app
1. api.slack.com/apps → Create New App → From Scratch → name it, pick workspace.
2. **Basic Information → App Credentials → Signing Secret → Show.** Save as `SLACK_SIGNING_SECRET`.
3. **OAuth & Permissions → Bot Token Scopes**, add: `channels:history`, `channels:read`, `users:read`.
4. Click **Install to Workspace → Allow.**
5. Back on **OAuth & Permissions**, copy the **Bot User OAuth Token** (starts `xoxb-`). Save as `SLACK_BOT_TOKEN`. (This is new — it lets the bot turn user IDs into real names for the "Sender" column.)
6. **Event Subscriptions → Enable Events ON.** Under **Subscribe to Bot Events** add `message.channels`.
7. **Do NOT try to set the Request URL yet.** Come back after Step 6 below. Leave this tab open.

### Step 3 — Microsoft Azure (Outlook)
Follow the original guide, with two corrections noted in Step 7:
1. portal.azure.com → App registrations → New registration → name `task-bot`, "Accounts in this organizational directory only", no redirect URI → Register.
2. Save **Application (client) ID** → `OUTLOOK_CLIENT_ID` and **Directory (tenant) ID** → `OUTLOOK_TENANT_ID`.
3. **Certificates & secrets → New client secret** → copy the **Value** → `OUTLOOK_CLIENT_SECRET`.
4. **API permissions → Add a permission → Microsoft Graph → Application permissions →** check **Mail.Read** → Add → **Grant admin consent.**
   - Note: `Mail.Read` (application) grants read access to *all* mailboxes in the org. For a tighter scope, look up "Graph Application Access Policy" later. Fine for a personal/small setup.

### Step 4 — Google Sheet + service account
Same as the original guide:
1. Create sheet "Task Tracker" with headers in row 1: Status, Due Date, Description, Assigned To, Source, Sender, Date Added.
2. Save the ID from the URL as `GOOGLE_SHEET_ID`.
3. Cloud console → new project → enable **Google Sheets API** → create a **Service Account** (Editor role) → Keys → Add Key → JSON.
4. Save the entire JSON file contents as `GOOGLE_SERVICE_ACCOUNT_JSON`.
   - Railway tip: if pasting the raw JSON gives you trouble, base64-encode it
     (`base64 -w0 key.json` on Linux/Mac) and paste that instead — the app
     auto-detects base64.
5. Share the sheet (Editor) with the `client_email` from the JSON.

### Step 5 — Deploy to Railway
1. Put these files in a private GitHub repo (`main.py`, `requirements.txt`, `Procfile`, `register_outlook.py`, `.gitignore`).
2. railway.app → New Project → Deploy from GitHub repo → pick the repo.
3. After build: Settings → Networking → **Generate Domain.** Save it as `YOUR_RAILWAY_URL`.
4. **Variables tab** → add every variable:
   - `ANTHROPIC_API_KEY`
   - `SLACK_SIGNING_SECRET`
   - `SLACK_BOT_TOKEN`
   - `OUTLOOK_TENANT_ID`, `OUTLOOK_CLIENT_ID`, `OUTLOOK_CLIENT_SECRET`
   - `GOOGLE_SHEET_ID`, `GOOGLE_SERVICE_ACCOUNT_JSON`
5. Wait for the redeploy to finish. Visit `https://YOUR_RAILWAY_URL/` — you should see `{"status":"ok"}`. If you do, the server is live.

### Step 6 — Verify the Slack URL (the part that was failing)
1. Back in Slack → Event Subscriptions → Request URL, paste:
   `https://YOUR_RAILWAY_URL/webhook/slack`
2. Slack sends the challenge; the live server echoes it; you get a green **Verified**. ✅
3. Save Changes.
4. In Slack, invite the bot to a channel: `/invite @YourBotName`.

### Step 7 — Register the Outlook webhook (corrected)
The original Postman JSON had **two bugs**:
- `"resource": "me/mailFolders/inbox/messages"` — `me` does **not** work with
  application permissions. It must be `users/{your-email}/mailFolders/inbox/messages`.
- `"expirationDateTime": "2026-12-31..."` — too far out. Message subscriptions
  max out at **4230 minutes (~3 days)** and Graph will reject anything longer.

Easiest path — use the included script instead of Postman:
```bash
pip install httpx
export OUTLOOK_TENANT_ID=...   OUTLOOK_CLIENT_ID=...   OUTLOOK_CLIENT_SECRET=...
export RAILWAY_URL=https://YOUR_RAILWAY_URL
export OUTLOOK_USER=you@yourcompany.com
python register_outlook.py
```
A `201` means success. Because subscriptions expire in ~3 days, **re-run this
script every couple of days** (or schedule it) to keep mail flowing.

### Step 8 — Test
- Slack: post `can you review the Q3 report by Friday?` in a channel the bot is in → a row should appear in the sheet within seconds.
- Outlook: email yourself `Please review the budget spreadsheet before EOD tomorrow` → another row appears.

## Troubleshooting
- Slack still says challenge failed → the Railway URL isn't live. Check `https://YOUR_RAILWAY_URL/` returns ok, and that the path ends in `/webhook/slack`.
- Nothing in the sheet → check Railway's **Deploy Logs** for errors (most common: bad `GOOGLE_SERVICE_ACCOUNT_JSON`, or you forgot to share the sheet with the service account email).
- Outlook 403 on register → admin consent wasn't granted in Azure Step 3.4.
