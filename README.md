# OutlookRedirector
## Outlook / Microsoft Graph — Email Access Tools

Two Python scripts to interact with Outlook mailboxes via the **Microsoft Graph API** and **OWA REST API**, with no external dependencies (stdlib only).

---

## Files

| File | Purpose |
|---|---|
| `outlook_token_refresh.py` | Obtains a new access token from a refresh token, automatically testing Microsoft FOCI client IDs |
| `outlook_read_mail.py` | Reads emails, lists folders, searches messages, and manages forwarding rules using an access token |

---

## `outlook_token_refresh.py`

Refreshes the Outlook/Graph access token. Supports two operating modes:

- **FOCI scan** — automatically tests Microsoft first-party client IDs (Office, Outlook Mobile, Teams, OneDrive, Azure CLI, Graph Explorer) against both v1 and v2 endpoints, stopping at the first one that works.
- **Direct** — uses a specific `--client-id` provided by the user.

### Usage

```bash
# Automatic FOCI scan (tries all known clients)
python3 outlook_token_refresh.py --refresh-token <TOKEN>

# Specific client ID
python3 outlook_token_refresh.py --refresh-token <TOKEN> --client-id <CLIENT_ID>

# Token for OWA (outlook.office.com) instead of Graph
python3 outlook_token_refresh.py --refresh-token <TOKEN> --owa

# Specific tenant
python3 outlook_token_refresh.py --refresh-token <TOKEN> --tenant-id contoso.com

# v1 endpoint (uses resource= instead of scope=)
python3 outlook_token_refresh.py --refresh-token <TOKEN> --v1

# Raw JSON output (useful for scripting)
python3 outlook_token_refresh.py --refresh-token <TOKEN> --json
```

### Parameters

| Parameter | Required | Description |
|---|---|---|
| `--refresh-token` | Yes | Previously obtained refresh token |
| `--client-id` | No | Forces a specific client_id; omitting this triggers FOCI scan |
| `--tenant-id` | No | Tenant ID or domain (default: `common`) |
| `--v1` | No | Uses OAuth v1 endpoint (`resource=`) |
| `--owa` | No | Requests a token for `outlook.office.com` instead of Graph |
| `--json` | No | Prints the full response as JSON |

### FOCI Client IDs tested

| Name | Client ID |
|---|---|
| office | `d3590ed6-52b3-4102-aeff-aad2292ab01c` |
| outlook-mobile | `27922004-5251-4030-b22d-91ecd9a37ea4` |
| teams | `1fec8e78-bce4-4aaf-ab1b-5451cc387264` |
| onedrive | `ab9b8c07-8f02-4f72-87fa-80105867a763` |
| azure-cli | `04b07795-8ddb-461a-bbee-02f9e1bf7b46` |
| graph-explorer | `de8bc8b5-d9f9-48b1-a8ad-b748da725064` |

---

## `outlook_read_mail.py`

Accesses and manages emails via the MS Graph API (or OWA REST v2.0) using an access token.

### General usage

```bash
python3 outlook_read_mail.py --token <ACCESS_TOKEN> <command> [options]
```

### Available commands

#### `me` — Authenticated user info
```bash
python3 outlook_read_mail.py --token <TOKEN> me
```

#### `inbox` — List inbox messages
```bash
python3 outlook_read_mail.py --token <TOKEN> inbox
python3 outlook_read_mail.py --token <TOKEN> inbox --top 20
```

#### `read` — Read a full message
```bash
python3 outlook_read_mail.py --token <TOKEN> read --msg-id <MESSAGE_ID>
```

#### `folders` — List mail folders
```bash
python3 outlook_read_mail.py --token <TOKEN> folders
```

#### `search` — Search messages
```bash
python3 outlook_read_mail.py --token <TOKEN> search --query "subject or sender"
python3 outlook_read_mail.py --token <TOKEN> search --query "invoice" --top 5
```

#### `forward-set` — Enable forwarding via inbox rule
Removes any existing rule with the same name before creating a new one.
```bash
python3 outlook_read_mail.py --token <TOKEN> forward-set --forward-to dest@example.com
python3 outlook_read_mail.py --token <TOKEN> forward-set --forward-to dest@example.com --no-copy
```

#### `forward-rule` — Create a forwarding inbox rule (alternative method)
Requires the `MailFolder.ReadWrite` scope.
```bash
python3 outlook_read_mail.py --token <TOKEN> forward-rule --forward-to dest@example.com
python3 outlook_read_mail.py --token <TOKEN> forward-rule --forward-to dest@example.com --rule-name "My rule"
```

#### `forward-list` — List existing inbox rules
```bash
python3 outlook_read_mail.py --token <TOKEN> forward-list
```

#### `forward-remove` — Remove a forwarding rule by ID
```bash
python3 outlook_read_mail.py --token <TOKEN> forward-remove --rule-id <RULE_ID>
```

#### `forward-owa` — Enable forwarding via OWA internal `SetMailbox`
Requires a token with audience `https://outlook.office.com` (generate with `--owa` in the refresh script).
```bash
python3 outlook_read_mail.py --token <OWA_TOKEN> forward-owa --forward-to dest@example.com
```

### Global parameters

| Parameter | Description |
|---|---|
| `--token` | MS Graph or OWA access token (required) |
| `--top N` | Number of messages to return (default: 10) |
| `--msg-id` | Message ID (used with `read`) |
| `--query` | Search string (used with `search`) |
| `--forward-to` | Destination email for forwarding |
| `--rule-name` | Rule name (default: "Redirecionar tudo") |
| `--rule-id` | Rule ID to delete (used with `forward-remove`) |
| `--no-copy` | Do not keep a local copy when forwarding |
| `--owa` | Use OWA REST v2.0 endpoint instead of Graph |

---

## Typical workflow

```bash
# 1. Get a new access token from a refresh token
python3 outlook_token_refresh.py --refresh-token <REFRESH_TOKEN> --json > token.json

# 2. Extract the access token
TOKEN=$(python3 -c "import json; print(json.load(open('token.json'))['access_token'])")

# 3. Check authenticated user
python3 outlook_read_mail.py --token "$TOKEN" me

# 4. List the 10 most recent emails
python3 outlook_read_mail.py --token "$TOKEN" inbox

# 5. Search for emails by keyword
python3 outlook_read_mail.py --token "$TOKEN" search --query "invoice"

# 6. For OWA-based forwarding, generate an OWA token first
python3 outlook_token_refresh.py --refresh-token <REFRESH_TOKEN> --owa --json > token_owa.json
TOKEN_OWA=$(python3 -c "import json; print(json.load(open('token_owa.json'))['access_token'])")
python3 outlook_read_mail.py --token "$TOKEN_OWA" forward-owa --forward-to dest@example.com
```

---

## Requirements

- Python 3.10+
- No external dependencies (uses only `urllib`, `json`, `argparse`, `base64`)

## Required Microsoft Graph scopes

| Feature | Scopes |
|---|---|
| Read emails / profile | `Mail.Read`, `User.Read` |
| Manage forwarding settings | `MailboxSettings.ReadWrite` |
| Create/delete inbox rules | `Mail.ReadWrite` |
