# ms-teams-graph-poc

Minimal Python PoC demonstrating authenticated access to Microsoft Teams conversations via the Microsoft Graph API. Built to validate the data-collection layer of a master's thesis on LLM-assisted social engineering detection in enterprise chat.

## What This Does

1. Authenticates to Microsoft Entra using a single-tenant App Registration (interactive browser login with MSAL device-flow fallback)
2. Calls `GET /me` — verifies login
3. Calls `GET /me/chats` — lists accessible Teams chats
4. Calls `GET /chats/{id}/messages` — fetches messages from the first chat
5. Saves each response as JSON to `output/`

## Prerequisites

- macOS / Python 3.10+
- A Microsoft Entra App Registration with:
  - Account type: single tenant
  - Redirect URI: `http://localhost`
  - Authentication: `Allow public client flows = Yes`
  - Delegated API permissions: `User.Read`, `Chat.Read` (with admin/user consent)

## Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Copy `.env.example` to `.env` and fill in your App Registration values:

```env
CLIENT_ID=your-azure-app-client-id
TENANT_ID=your-azure-tenant-id
SCOPES=User.Read Chat.Read
OUTPUT_DIR=output
MSAL_PORT=53100
```

## Run

```bash
python fetch_teams_chat.py
```

A browser window opens for interactive login. If browser login fails, the script falls back to device-code flow and prints the URL + code in the terminal.

## Expected Output

```
output/
  me.json            GET /me response
  chats.json         GET /me/chats response
  messages_<id>.json messages from first chat
```

Terminal confirms:
- Login success
- User display name
- Number of chats found
- Whether messages were retrieved

## Troubleshooting

| Error | Cause | Fix |
|-------|-------|-----|
| `token 取得失敗` | Wrong CLIENT_ID / TENANT_ID | Re-check App Registration |
| `403` on `/me/chats` | `Chat.Read` consent not granted | Grant consent in Azure portal |
| Empty `value` array | No chats exist | Create a test chat in Teams first |

## Context

This PoC is the data-collection proof-of-concept for a two-stage privacy-preserving pipeline:

```
Teams API → this script → PII filter (thesis-pii-pipeline) → LLM risk scoring → admin-console (thesis-admin-console)
```
