# AIChiefOfStaff — Outlook Edition (AICOS)

> **This is the Windows (PowerShell) edition.** All commands use PowerShell syntax — backtick `` ` `` for line continuation, `$VAR` for variables, `Start-Sleep` instead of `sleep`. For macOS instructions see `README-mac.md`.

Your personal AI Chief of Staff — built on Azure Functions + Microsoft Graph + Claude AI.

AICOS monitors your Outlook inbox 24/7, classifies every email using Claude AI into 8 life dimensions, and takes action: creates tasks in Microsoft To Do with subtasks and priorities, logs bills to Finance Excel workbooks, tracks job applications, saves documents to OneDrive — all without you lifting a finger.

You get a morning briefing at 7am. Every decision is logged and auditable.



---

## Table of Contents

- [How It Works](#how-it-works)
- [5-Minute Setup](#5-minute-setup)
- [What Happens After Setup](#what-happens-after-setup)
- [Architecture](#architecture)
- [File Structure](#file-structure)
- [8 Life Dimensions](#8-life-dimensions)
- [Classification System](#classification-system)
- [Task Creation](#task-creation)
- [OneDrive Folder Structure](#onedrive-folder-structure)
- [Finance Tracker](#finance-tracker)
- [Job Tracker](#job-tracker)
- [Configuration Reference](#configuration-reference)
- [Azure Deployment — Full Guide](#azure-deployment--full-guide)
- [Updating an Existing Install](#updating-an-existing-install)
- [Testing](#testing)
- [Debug Tools](#debug-tools)
- [Privacy & Data Sovereignty](#privacy--data-sovereignty)
- [Roadmap](#roadmap)

---

## How It Works

Every email that arrives in your Outlook inbox goes through this pipeline:

1. **Labelled** `AICOS:New` — Microsoft Graph API applies a category to all incoming email
2. **Classified** — Claude AI reads the email and determines: dimension, verb, urgency, confidence, and whether you need to act
3. **Routed** — based on the verb:
   - `pay` → Finance Excel workbook (OneDrive) + Microsoft To Do task
   - `do` → To Do task with priority and due date (career emails also go to Job Tracker)
   - `meet` → Flagged in Outlook + RSVP task in To Do
   - `file` → OneDrive folder (legal/infra documents)
   - `read` → Outlook category label only (no action needed)
4. **Logged** — every routing decision recorded in Azure Table Storage
5. **Briefed** — 7am morning email summarising what needs your attention today

The system runs on 6 automatic Azure Function triggers — you never need to open the Azure portal after setup.

---

## 5-Minute Setup (Windows)

### Prerequisites

- Microsoft account (personal `@outlook.com` or Microsoft 365 work/school)
- Azure subscription ([free tier](https://azure.microsoft.com/free) works — Visual Studio Enterprise also works)
- [Anthropic API key](https://console.anthropic.com) — create one, set a $5/month spending limit
- Node.js 20+, Git, Azure CLI, Azure Functions Core Tools v4

### Install prerequisites

**Node.js 20+**

Download the LTS installer from [nodejs.org](https://nodejs.org) → run the `.msi` → or with winget:

```powershell
winget install OpenJS.NodeJS.LTS
node --version   # should print v20.x.x
```

**Git**

Download from [git-scm.com](https://git-scm.com/download/win) → run installer, or:

```powershell
winget install Git.Git
```

**Azure CLI**

```powershell
winget install Microsoft.AzureCLI
az --version
```

Or download the MSI from [aka.ms/installazurecliwindows](https://aka.ms/installazurecliwindows).

**Azure Functions Core Tools v4**

Open PowerShell as Administrator:

```powershell
npm install -g azure-functions-core-tools@4 --unsafe-perm true
func --version   # should print 4.x.x
```

**GitHub CLI + Copilot (optional but recommended)**

```powershell
winget install GitHub.cli
gh auth login
gh extension install github/gh-copilot
```

### Step 1: Clone and install

```powershell
git clone https://github.com/sukhpreetn/outlook-ai-cos.git
cd outlook-ai-cos
npm install
```

### Step 2: Log in to Azure

```powershell
az login
# Browser opens — sign in with your Microsoft account

az account show
# Verify "state": "Enabled"
```

### Step 3: Create your Azure App Registration

```powershell
# Create app registration (supports personal + work accounts)
az ad app create `
  --display-name "aicos-outlook" `
  --sign-in-audience AzureADandPersonalMicrosoftAccount

# Copy the appId from output — this is your CLIENT_ID
```

**Set token version to v2** (required for personal Microsoft accounts):

```powershell
$APP_OID = az ad app show --id "YOUR_CLIENT_ID" --query id -o tsv

az rest --method PATCH `
  --uri "https://graph.microsoft.com/v1.0/applications/$APP_OID" `
  --headers "Content-Type=application/json" `
  --body '{"api":{"requestedAccessTokenVersion":2}}'
```

**Add redirect URI for local OAuth flow:**

```powershell
az ad app update `
  --id "YOUR_CLIENT_ID" `
  --public-client-redirect-uris "http://localhost:8080/callback"
```

**Add API permissions via Azure Portal:**

Go to [portal.azure.com](https://portal.azure.com) → Microsoft Entra ID → App registrations → `aicos-outlook` → **API Permissions** → Add a permission → Microsoft Graph → Delegated permissions. Add:

```
Mail.Read          Mail.ReadWrite      Mail.Send
Calendars.Read     Tasks.ReadWrite     Files.ReadWrite.All
offline_access
```

Click **Grant admin consent** and confirm.

### Step 4: Get your OAuth2 refresh token

```bash
node scripts/getToken.mjs
```

A URL prints to your terminal. Open it in your browser → sign in with your Outlook account → the refresh token prints to your terminal automatically. Copy it.

### Step 5: Configure `local.settings.json`

```json
{
  "IsEncrypted": false,
  "Values": {
    "FUNCTIONS_WORKER_RUNTIME": "node",
    "AzureWebJobsStorage": "UseDevelopmentStorage=true",
    "AICOS_ENV": "local",
    "TENANT_ID": "your-tenant-id",
    "CLIENT_ID": "your-app-client-id",
    "CLIENT_SECRET": "",
    "USER_EMAIL": "you@outlook.com",
    "ANTHROPIC_API_KEY": "sk-ant-...",
    "STORAGE_CONNECTION_STRING": "UseDevelopmentStorage=true",
    "REFRESH_TOKEN": "M.C533_SN1.0.U.-...",
    "WEBHOOK_SECRET": "aicos-local-secret"
  }
}
```

Where to find each value:

| Setting | Where to get it |
|---------|----------------|
| `TENANT_ID` | Azure Portal → Microsoft Entra ID → Overview → Tenant ID |
| `CLIENT_ID` | Azure Portal → App registrations → aicos-outlook → Application (client) ID |
| `USER_EMAIL` | Your Outlook email address |
| `ANTHROPIC_API_KEY` | [console.anthropic.com](https://console.anthropic.com) → API Keys → Create Key |
| `REFRESH_TOKEN` | Printed by `node scripts/getToken.mjs` in Step 4 |

**Important — Anthropic spending cap:** Go to [console.anthropic.com](https://console.anthropic.com) → Settings → Limits → set a monthly budget of at least $5. The default $0 cap causes all classification to fail with HTTP 429 errors. For personal use (~50 emails/day), expect ~$0.10/month.

### Step 6: Start local development

**Terminal 1 — Start local storage emulator (Azurite):**

```powershell
npx azurite --location .\tmp\azurite
```

**Terminal 2 — Build and start functions:**

```powershell
npm run build; func start
```

You should see:

```
Functions:
  graphWebhook:         [GET,POST] http://localhost:7071/api/webhook/graph
  briefingHttp:         [POST]     http://localhost:7071/api/ops/briefing
  mailToTable:          [GET,POST] http://localhost:7071/api/ops/mailToTable
  webhookRenewalHttp:   [POST]     http://localhost:7071/api/ops/webhook/renew
  classifierTimer:      timerTrigger
  briefingTimer:        timerTrigger
  webhookRenewal:       timerTrigger
```

**Terminal 3 — Test the classifier:**

```powershell
curl -X POST "http://localhost:7071/admin/functions/classifierTimer" `
  -H "Content-Type: application/json" -d "{}"
```

---

## What Happens After Setup

Your system is now live. Here's what to expect:

| When | What happens |
|------|-------------|
| **Immediately** | Graph API starts watching your inbox for new emails |
| **Every 15 minutes** | Classifier picks up `AICOS:New` emails, classifies them, and routes them |
| **Within 1 hour** | Outlook categories appear on your emails (`AICOS:Classified`, `AICOS:Dropped`) |
| **Same day** | Microsoft To Do tasks appear in your AICOS list with priorities |
| **7am tomorrow** | First morning briefing arrives in your inbox |
| **Daily 9am UTC** | Graph webhook subscription auto-renewed (personal accounts expire every 3 days) |

**Behind the scenes (automatic, no action needed):**

- **Auto-label:** All unread emails tagged `AICOS:New` before classification
- **Idempotency:** StateEngine ensures no email is classified twice across overlapping runs
- **Token refresh:** OAuth2 refresh token silently renewed on each API call
- **Webhook renewal:** `webhookRenewal` function runs daily and renews the Graph subscription automatically

**To verify it's working:**

1. Send yourself a test email with subject: `Invoice from Test Corp - $50 due March 30`
2. Wait 15 minutes (or trigger manually — see [Debug Tools](#debug-tools))
3. Check the email — it should have categories `AICOS:Actioned` and `AICOS:Classified`
4. Check Microsoft To Do — a task should appear in the **AICOS** list: `PAY: Invoice from Test Corp - $50 due March 30`

**If something goes wrong:**

- Check the `func start` terminal — all errors printed with `[ERROR]` prefix
- Trigger classifier manually: `curl -X POST http://localhost:7071/admin/functions/classifierTimer -H "Content-Type: application/json" -d '{}'`
- Verify your `REFRESH_TOKEN` is still valid — re-run `node scripts/getToken.mjs` if needed
- Check Azure Table Storage (`AicosState` table) to see message state
- See [Debug Tools](#debug-tools) for inspection commands

---

## Architecture

```
Outlook Inbox
      │
      ▼
graphWebhook (real-time) + classifierTimer (every 15 min)
      │
      ▼
ClassifierRunner.ts   — fetches AICOS:New emails (max 25 per batch)
      │                  deduplicates via StateEngine.canProcessMessage()
      │
      ▼
Classifier.ts         — 3-tier classification engine
      │    Tier 0: Known drop patterns    (no API call — free — <1ms)
      │    Tier 1: Known domains/rules    (no API call — free — <5ms)
      │    Tier 2: Unknown senders        (Claude Haiku — ~$0.000025/email — 1-3s)
      │    Returns: cls object (dimension, verb, urgency, confidence, ...)
      │
StateEngine.ts        — AICOS:New → AICOS:Classified (category swap)
                         stores classification in Azure Table Storage
      │
      ▼
Router.ts             — 6-gate verb-driven routing engine
      ├── Gate 1: isReceipt      → FinanceTracker → Finance.xlsx (OneDrive)
      ├── Gate 2: silenceReason  → AICOS:Dropped (filed silently)
      ├── Gate 3: low confidence → AICOS:Review (unless urgency >= 4)
      ├── Gate 4: batch_key      → accumulate low-urgency emails
      ├── Gate 5: verb switch:
      │    ├── pay   → FinanceTracker + Pay task in To Do
      │    ├── do    → JobTracker (career) or To Do task
      │    ├── meet  → Flagged in Outlook + RSVP task
      │    ├── file  → Flagged for filing
      │    └── read  → Outlook category only
      └── Gate 6: hasDeadline    → To Do task with due date
      │
      ▼
MorningBriefing.ts    — 7am: Claude Sonnet generates rich HTML briefing
                         Sections: Priority Actions, Your Day, Finance Health,
                         Signals, Dimensions bar chart, Components status
                         Sends via Graph API /me/sendMail
      │
      ▼
webhookRenewal.ts     — 9am UTC daily: renews Graph subscription
                         Stores subscription ID in AicosSubscriptions table
```

---

## File Structure

```
outlook-ai-cos/
├── src/
│   ├── aicos-core/                 # System layer
│   │   ├── Config.ts               # Azure App Configuration + Key Vault reader
│   │   ├── AI.ts                   # Claude API wrapper — task-based model routing
│   │   ├── GraphClient.ts          # All Microsoft Graph operations
│   │   ├── StateEngine.ts          # Azure Table Storage state + idempotency
│   │   └── Logger.ts               # Centralised logging + diagnostic counters
│   │
│   ├── email-pipeline/             # Pipeline layer
│   │   ├── Classifier.ts           # Tier 0/1/2 email classification
│   │   ├── ClassifierRunner.ts     # Main orchestrator
│   │   ├── Router.ts               # 6-gate verb-driven routing engine
│   │   ├── FinanceTracker.ts       # Receipts/bills → Excel on OneDrive
│   │   ├── JobTracker.ts           # Career emails → Excel on OneDrive
│   │   ├── AssetTracker.ts         # Orders/subscriptions → Excel on OneDrive
│   │   └── MorningBriefing.ts      # Daily AI digest email via Claude Sonnet
│   │
│   └── functions/                  # Azure Function entry points
│       ├── timerTrigger.ts         # Every 15 min + daily briefing check
│       ├── webhookTrigger.ts       # Graph change notification handler
│       ├── webhookRenewal.ts       # Auto-renews Graph subscription daily
│       ├── briefingHttp.ts         # Manual briefing trigger (POST /ops/briefing)
│       └── mailToTable.ts          # Debug: read emails → Table Storage
│
├── scripts/
│   ├── deploy.sh                   # Full Azure infrastructure + deploy script
│   ├── getToken.mjs                # One-time OAuth2 refresh token generator
│   ├── registerWebhook.js          # Graph subscription registration
│   └── seedTestData.mjs            # Seeds inbox with test emails (all 8 dimensions)
│
├── host.json                       # Azure Functions host config
├── local.settings.json             # Local dev environment variables (git-ignored)
├── package.json
└── tsconfig.json
```

---

## 8 Life Dimensions

Every email is classified into one of 8 life dimensions:

| # | Dimension | Outlook Category | What Goes Here | Auto-Actions |
|---|-----------|-----------------|----------------|-------------|
| 000 | Priority | `AICOS:Priority` | Cross-dimension urgency flag (urgency >= 4) | Always flagged 🚩 |
| 001 | Career | `AICOS:Classified` | Job applications, recruiters, interviews, offers | JobTracker.xlsx + To Do task |
| 002 | Finance | `AICOS:Classified` | Bills, payments, bank statements, receipts | Finance.xlsx + Pay task |
| 003 | Legal | `AICOS:Classified` | Tax, compliance, government, visa, contracts | Flagged for filing |
| 004 | Knowledge | `AICOS:Classified` | Courses, newsletters, research, education | Batched or filed |
| 005 | Ventures | `AICOS:Classified` | Startups, investments, partnerships, deal flow | Flagged + high urgency |
| 006 | Life | `AICOS:Classified` | Health, family, subscriptions, personal admin | Routed by urgency |
| 007 | Commerce | `AICOS:Classified` | Shopping, orders, deliveries, promotions | Assets.xlsx |
| 009 | Infra | `AICOS:Classified` | DevOps, cloud, GitHub, system alerts | Flagged if urgent |

**Alias resolution:** The classifier handles variants automatically — "job" → career, "tax" → legal, "health" → life, "shopping" → commerce, etc.

---

## Classification System

### 3-Tier Pipeline

| Tier | Method | Model | Cost | Speed | When Used |
|------|--------|-------|------|-------|-----------|
| 0 | Known drop patterns | None | Free | < 1ms | no-reply senders, newsletters, unsubscribe links |
| 1 | Known domain map + learned rules | None | Free | < 5ms | Sender domain in known list or rule confidence >= 0.85 |
| 2 | Full AI triage | Claude Haiku 4.5 | ~$0.000025/email | 1-3 sec | Unknown senders — full classification |

**Escalation:** Tier 1 escalates to Tier 2 when: confidence < 0.85, urgency >= 4, or legal terms found in subject.

**Fallback:** If Anthropic API is unavailable, Tier 2 falls back to conservative defaults (dimension: knowledge, verb: read) and routes to `AICOS:Review`.

**Learning:** The StateEngine tracks classification patterns. When a sender reliably maps to a dimension (5+ consistent classifications), it's flagged for promotion to Tier 0/1 — reducing future AI costs.

### The `cls` Object (key fields)

```typescript
{
  messageId:       string,   // Graph message ID
  dimension:       string,   // career | finance | legal | knowledge | ventures | life | commerce | infra
  verb:            string,   // do | pay | meet | file | read
  urgency:         number,   // 1-5 (5 = requires response today)
  confidence:      number,   // 0.0-1.0
  tier:            number,   // 0 | 1 | 2
  isReceipt:       boolean,  // financial receipt/invoice?
  hasDeadline:     boolean,  // specific deadline mentioned?
  silenceReason:   string,   // why this can be auto-filed (or null)
  batchKey:        string,   // group similar low-urgency emails (or null)
  proactiveAction: string,   // draft_reply | create_task | create_event | null
  summary:         string,   // one-sentence action-oriented description
}
```

### Confidence Routing

```
confidence >= 0.85  →  Tier 1 fast path (no AI call)
confidence >= 0.60  →  Route normally
confidence <  0.60  →  AICOS:Review (human needed)
urgency    >= 4     →  Override: always route regardless of confidence
```

---

## Task Creation

When the classifier determines an email is actionable, Microsoft To Do tasks are created with rich structure:

**Example — finance (pay verb):**

```
PAY: Chase credit card bill — $247.50 due Apr 1
  List:      AICOS
  Priority:  High (urgency 4 → high importance)
  Due:       Tomorrow
  Body:      "Monthly credit card bill. Ignore = late fee + interest charges.
              From: Chase Bank <alerts@chase.com>"
```

**Example — career (do verb):**

```
[CAREER] Interview Invitation — Senior Engineer at Acme Corp
  List:      AICOS
  Priority:  High
  Body:      "We'd like to schedule an interview. From: recruiting@acme.com"
  Also:      Logged to Jobs.xlsx → Applications tab on OneDrive
```

Every task body includes the original email subject so you can always find the source email in Outlook.

---

## OneDrive Folder Structure

Created automatically on first write by each tracker module.

```
AICOS/                              (set via oneDriveFolderId config)
│
├── Finance/
│   └── Finance-2026.xlsx
│       ├── Transactions tab        (all receipts + bills logged here)
│       └── Dashboard tab           (formula-only, zero runtime cost)
│
├── Jobs/
│   └── JobSearch-2026.xlsx
│       ├── Applications tab        (state updated in-place per company)
│       └── Offers tab              (offer tracking)
│
├── Assets/
│   └── Assets-2026.xlsx
│       └── Assets tab              (orders, shipments, subscriptions)
│
└── Admin/
    └── Logs-2026.xlsx
        └── Classifier-Log tab      (diagnostic logs)
```

**Year rollover:** On the first email of January, a new workbook is created automatically (e.g. `Finance-2027.xlsx`). No manual intervention needed.

---

## Finance Tracker

### Transactions Sheet Columns

| Col | Field | Notes |
|-----|-------|-------|
| A | Date | yyyy-MM-dd |
| B | Sender | Bank or company name |
| C | Subject | Email subject |
| D | Type | receipt / bill |
| E | Amount | Numeric — extracted from email body (e.g. `$247.50`) |
| F | Notes | AI-generated one-line summary |

**Known financial domains (auto Tier 1 — no AI cost):** Chase, Bank of America, Wells Fargo, Amex, PayPal, Venmo, Stripe, Intuit, Fidelity, Vanguard, and 15+ others.

---

## Job Tracker

### Application State Machine

```
NEW_LEAD → APPLIED → INTERVIEWING → OFFER → ACCEPTED
                                          → REJECTED
                                          → WITHDRAWN
```

- Forward-only transitions — no regressions (can't go from INTERVIEWING back to APPLIED)
- Terminal states (ACCEPTED/REJECTED/WITHDRAWN) reachable from any state
- One row per company — state updated in-place, never duplicated

### Sheet Columns

| Col | Field | Notes |
|-----|-------|-------|
| A | Date | First contact yyyy-MM-dd |
| B | Status | open / interview / offer / rejected / withdrawn |
| C | Company | Extracted from sender name or domain |
| D | Subject | Email subject |
| E | Type | application_ack / interview / offer / rejection / outreach |
| F | Notes | AI-generated summary |

---

## Configuration Reference

All config lives in `local.settings.json` (local dev) or Azure Function App Settings (production). Changes take effect within 1 hour (cache TTL) or immediately after restarting `func start`.

### Core Settings

| Key | Default | What It Does |
|-----|---------|-------------|
| `TENANT_ID` | (required) | Microsoft Entra ID tenant ID |
| `CLIENT_ID` | (required) | App registration client ID |
| `USER_EMAIL` | (required) | Your Outlook email address |
| `ANTHROPIC_API_KEY` | (required) | Anthropic API key (`sk-ant-...`) |
| `REFRESH_TOKEN` | (required) | OAuth2 refresh token from `getToken.mjs` |
| `STORAGE_CONNECTION_STRING` | (required) | Azure Table Storage connection string |
| `WEBHOOK_SECRET` | `aicos-secret` | Validates incoming Graph notifications |

### Classification Behaviour

| Key | Default | What It Does |
|-----|---------|-------------|
| `aicos:batchSize` | `25` | Emails processed per timer run |
| `aicos:tier1ConfidenceFloor` | `0.85` | Min confidence for Tier 1 routing |
| `aicos:tier2ConfidenceFloor` | `0.60` | Below this → `AICOS:Review` queue |
| `aicos:draftCapPerRun` | `3` | Max AI draft replies generated per run |

### AI Models

| Key | Default | What It Does |
|-----|---------|-------------|
| `aicos:aiModelClassify` | `claude-haiku-4-5` | Fast classification (Tier 2) |
| `aicos:aiModelBrief` | `claude-sonnet-4-6` | Morning briefing generation |
| `aicos:aiModelDraft` | `claude-haiku-4-5` | Reply draft generation |

### Trigger Schedule

| Key | Default | What It Does |
|-----|---------|-------------|
| `aicos:briefingHour` | `7` | Morning briefing hour (local time, 24h) |
| `aicos:briefingTz` | `America/Los_Angeles` | IANA timezone for briefing |

---

## Azure Deployment — Full Guide

This section covers every step needed to deploy AICOS to Azure so it runs 24/7 without your computer.

### Prerequisites checklist

Before starting, confirm:

- [ ] `az login` completed and `az account show` shows `"state": "Enabled"`
- [ ] `local.settings.json` fully configured and working locally (`func start` shows all 6 functions)
- [ ] `node scripts/getToken.mjs` completed — you have a valid `REFRESH_TOKEN`
- [ ] `npm run build` succeeds with zero TypeScript errors

### Step 1: Set your deployment variables

```powershell
$RESOURCE_GROUP  = "aicos-outlook-rg"
$APP_NAME        = "aicos-outlook-yourname"       # must be globally unique
$LOCATION        = "westus2"
$STORAGE_ACCOUNT = "aicosstorageyourname"         # globally unique, lowercase only, no hyphens
$KEYVAULT_NAME   = "aicos-kv-yourname"            # globally unique
$SUB_ID          = az account show --query id -o tsv
```

### Step 2: Create resource group

```powershell
az group create --name $RESOURCE_GROUP --location $LOCATION
```

### Step 3: Create storage account

```powershell
az storage account create `
  --name $STORAGE_ACCOUNT `
  --resource-group $RESOURCE_GROUP `
  --location $LOCATION `
  --sku Standard_LRS
```

### Step 4: Create Key Vault and store secrets

```powershell
# Create Key Vault
az keyvault create `
  --name $KEYVAULT_NAME `
  --resource-group $RESOURCE_GROUP `
  --location $LOCATION

# Grant yourself write access (RBAC mode)
$USER_OID = az ad signed-in-user show --query id -o tsv

az role assignment create `
  --role "Key Vault Secrets Officer" `
  --assignee $USER_OID `
  --scope "/subscriptions/$SUB_ID/resourcegroups/$RESOURCE_GROUP/providers/Microsoft.KeyVault/vaults/$KEYVAULT_NAME"

# Wait 30 seconds for RBAC propagation
Start-Sleep -Seconds 30

# Store secrets
$STORAGE_CONN = az storage account show-connection-string `
  --name $STORAGE_ACCOUNT --resource-group $RESOURCE_GROUP --query connectionString -o tsv

az keyvault secret set --vault-name $KEYVAULT_NAME --name "anthropic-api-key"        --value "sk-ant-..."
az keyvault secret set --vault-name $KEYVAULT_NAME --name "refresh-token"             --value "M.C533_SN1..."
az keyvault secret set --vault-name $KEYVAULT_NAME --name "storage-connection-string" --value $STORAGE_CONN

Write-Host "✓ Secrets stored"
```

### Step 5: Create Function App

```powershell
az functionapp create `
  --name $APP_NAME `
  --resource-group $RESOURCE_GROUP `
  --storage-account $STORAGE_ACCOUNT `
  --consumption-plan-location $LOCATION `
  --runtime node `
  --functions-version 4 `
  --os-type Linux

# Set Node runtime explicitly (prevents "Runtime version: Error" in portal)
az functionapp config set `
  --name $APP_NAME `
  --resource-group $RESOURCE_GROUP `
  --linux-fx-version "NODE|20"

Write-Host "✓ Function App created"
```

### Step 6: Enable managed identity and grant Key Vault access

```powershell
az functionapp identity assign `
  --name $APP_NAME `
  --resource-group $RESOURCE_GROUP

$PRINCIPAL_ID = az functionapp identity show `
  --name $APP_NAME `
  --resource-group $RESOURCE_GROUP `
  --query principalId -o tsv

# Wait for identity propagation
Start-Sleep -Seconds 45

az role assignment create `
  --role "Key Vault Secrets User" `
  --assignee $PRINCIPAL_ID `
  --scope "/subscriptions/$SUB_ID/resourcegroups/$RESOURCE_GROUP/providers/Microsoft.KeyVault/vaults/$KEYVAULT_NAME"

Write-Host "✓ Managed identity configured"
```

### Step 7: Configure all app settings

```powershell
$STORAGE_CONN = az storage account show-connection-string `
  --name $STORAGE_ACCOUNT `
  --resource-group $RESOURCE_GROUP `
  --query connectionString -o tsv

$KEYVAULT_URL = "https://$KEYVAULT_NAME.vault.azure.net"

az functionapp config appsettings set `
  --name $APP_NAME `
  --resource-group $RESOURCE_GROUP `
  --settings `
    "AICOS_ENV=local" `
    "TENANT_ID=your-tenant-id" `
    "CLIENT_ID=your-client-id" `
    "USER_EMAIL=you@outlook.com" `
    "ANTHROPIC_API_KEY=sk-ant-..." `
    "REFRESH_TOKEN=M.C533_SN1..." `
    "STORAGE_CONNECTION_STRING=$STORAGE_CONN" `
    "KEY_VAULT_URL=$KEYVAULT_URL" `
    "WEBHOOK_SECRET=aicos-prod-secret" `
    "WEBSITE_NODE_DEFAULT_VERSION=~20" `
    "FUNCTIONS_EXTENSION_VERSION=~4"

Write-Host "✓ App settings configured"
```

### Step 8: Build and deploy the code

```powershell
npm run build
func azure functionapp publish $APP_NAME --typescript
```

Expected output:

```
Getting site publishing info...
Uploading package... [###########] Upload completed successfully.
Deployment completed successfully.
Syncing triggers...
```

### Step 9: Verify all 6 functions are live

```powershell
az functionapp function list `
  --name $APP_NAME `
  --resource-group $RESOURCE_GROUP `
  --query "[].{Name:name, Disabled:isDisabled}" `
  --output table
```

Expected output (all `Disabled: False`):

```
Name                  Disabled
--------------------  --------
briefingHttp          False
briefingTimer         False
classifierTimer       False
graphWebhook          False
mailToTable           False
webhookRenewal        False
webhookRenewalHttp    False
```

If the command returns `Bad Request`, wait 2 minutes for deployment propagation and retry.

### Step 10: Register Graph webhook

```powershell
$FUNC_KEY = az functionapp keys list `
  --name $APP_NAME `
  --resource-group $RESOURCE_GROUP `
  --query functionKeys.default -o tsv

# Store webhook URL in app settings
az functionapp config appsettings set `
  --name $APP_NAME `
  --resource-group $RESOURCE_GROUP `
  --settings "WEBHOOK_URL=https://$APP_NAME.azurewebsites.net/api/webhook/graph?code=$FUNC_KEY"

# Register the Graph subscription
curl -X POST "https://$APP_NAME.azurewebsites.net/api/ops/webhook/renew?code=$FUNC_KEY"
```

Expected response: `Subscription active: <subscription-id>`

### Step 11: Test the live system

```powershell
# Trigger briefing
curl --max-time 120 -X POST `
  "https://$APP_NAME.azurewebsites.net/api/ops/briefing?code=$FUNC_KEY"

# Monitor live logs (portal)
# Go to: portal.azure.com → aicos-outlook-yourname → Log stream
```

Check your Outlook inbox — a briefing email should arrive within 30-60 seconds.

### Deployment summary

| Step | What was created | How to verify |
|------|-----------------|--------------|
| 1 | Variables set | `echo $APP_NAME` |
| 2 | Resource group | `az group show -n $RESOURCE_GROUP` |
| 3 | Storage account | `az storage account show -n $STORAGE_ACCOUNT` |
| 4 | Key Vault + 3 secrets | Portal → Key Vault → Secrets (3 listed) |
| 5 | Function App + Node 20 | Portal → Overview → Runtime version: 4.x |
| 6 | Managed identity + KV access | Portal → Identity → System assigned: On |
| 7 | 11 app settings | Portal → Settings → Environment variables |
| 8 | Code deployed | Portal → Functions → 7 functions listed |
| 9 | All 7 functions enabled | `az functionapp function list` — all False |
| 10 | Graph webhook registered | `Subscription active:` response |
| 11 | Live test passed | Briefing email received in Outlook |

---

## Updating an Existing Install

```powershell
git pull origin main
npm install
npm run build
func azure functionapp publish aicos-outlook-yourname --typescript
```

All settings, stored state, and webhook subscriptions are unchanged.

**After a major update:**

```powershell
# Verify all functions are still registered
az functionapp function list `
  --name aicos-outlook-yourname `
  --resource-group aicos-outlook-rg `
  --query "[].name" --output table

# Re-register webhook if needed
$FUNC_KEY = az functionapp keys list `
  --name aicos-outlook-yourname `
  --resource-group aicos-outlook-rg `
  --query functionKeys.default -o tsv

curl -X POST `
  "https://aicos-outlook-yourname.azurewebsites.net/api/ops/webhook/renew?code=$FUNC_KEY"
```

`webhookRenewal` is self-healing — if the stored subscription ID is stale it creates a new one automatically.

---

## Testing

### Manual triggers (local)

```powershell
# Trigger classifier
curl -X POST "http://localhost:7071/admin/functions/classifierTimer" `
  -H "Content-Type: application/json" -d "{}"

# Trigger briefing
curl -X POST "http://localhost:7071/api/ops/briefing"

# Read emails to Table Storage (debug Graph connection)
curl "http://localhost:7071/api/ops/mailToTable?limit=5"

# Simulate webhook notification
curl -X POST "http://localhost:7071/api/webhook/graph" `
  -H "Content-Type: application/json" `
  -d '{"value":[{"changeType":"created","clientState":"local-dev-secret","resource":"/me/messages","subscriptionId":"test"}]}'

# Register/renew webhook
curl -X POST "http://localhost:7071/api/ops/webhook/renew"
```

### Seed test data (all 8 dimensions)

```powershell
node scripts/seedTestData.mjs
```

This sends 16 realistic test emails (2 per dimension) to your inbox so all briefing sections show real data.

### Test emails to send yourself

| Subject | Expected result |
|---------|----------------|
| `Invoice from Chase - $247.50 due June 15` | `AICOS:Classified` + Pay task + Finance.xlsx |
| `Interview Invitation - Engineer at Microsoft` | `AICOS:Classified` + flagged + RSVP task + Jobs.xlsx |
| `Your Amazon order has shipped - #1Z999AA1` | `AICOS:Classified` + Assets.xlsx |
| `DocuSign: Sign required - NDA Agreement` | `AICOS:Classified` + flagged for filing |
| `Unsubscribe from our newsletter` | `AICOS:Dropped` (Tier 0 hard drop) |
| `hello` | `AICOS:Review` (low confidence) |

### Test isolation

Tests run against `UseDevelopmentStorage=true` (Azurite). Clear state between runs:

```powershell
node -e "
const { TableClient } = require('@azure/data-tables');
const client = TableClient.fromConnectionString('UseDevelopmentStorage=true', 'AicosState');
(async () => {
  let d = 0;
  for await (const e of client.listEntities()) {
    await client.deleteEntity(e.partitionKey, e.rowKey);
    d++;
  }
  console.log('Cleared', d, 'records');
})();"
```

Production data (Azure Table Storage) is never touched during local development.

---

## Debug Tools

### Check StateEngine — what has been processed?

```powershell
node -e "
const { TableClient } = require('@azure/data-tables');
const client = TableClient.fromConnectionString('UseDevelopmentStorage=true', 'AicosState');
(async () => {
  let count = 0;
  for await (const e of client.listEntities()) {
    console.log(e.rowKey?.substring(0,20), '|', e.state, '|', e.updatedAt);
    count++;
  }
  console.log('Total:', count, 'records');
})();"
```

### Check webhook subscription status

```powershell
node -e "
const { TableClient } = require('@azure/data-tables');
const client = TableClient.fromConnectionString('UseDevelopmentStorage=true', 'AicosSubscriptions');
(async () => {
  try {
    const e = await client.getEntity('graph', 'mailInbox');
    console.log('Subscription ID:', e.subscriptionId);
    console.log('Renewed at:', e.renewedAt);
  } catch { console.log('No active subscription'); }
})();"
```

### Force reprocess all emails

```bash
# 1. Clear state
node -e "
const { TableClient } = require('@azure/data-tables');
const client = TableClient.fromConnectionString('UseDevelopmentStorage=true', 'AicosState');
(async () => {
  for await (const e of client.listEntities())
    await client.deleteEntity(e.partitionKey, e.rowKey);
  console.log('State cleared — all emails will reprocess');
})();"

# 2. Trigger classifier
curl -X POST "http://localhost:7071/admin/functions/classifierTimer" \
  -H "Content-Type: application/json" -d '{}'
```

### Force webhook renewal

```powershell
# Local
curl -X POST "http://localhost:7071/api/ops/webhook/renew"

# Production
curl -X POST "https://$APP_NAME.azurewebsites.net/api/ops/webhook/renew?code=$FUNC_KEY"
```

### View live Azure logs

```bash
# Portal: portal.azure.com → aicos-outlook-yourname → Log stream
# Or check specific function invocations:
# Portal → Functions → briefingHttp → Monitor → Invocations
```

---

## Privacy & Data Sovereignty

**Your data never leaves your Microsoft and Azure account** (except for AI classification calls to Anthropic using your own key).

| Concern | How AICOS Handles It |
|---------|---------------------|
| Email content | Read by YOUR Azure Function. Sent to Anthropic (Claude) only for Tier 2 classification — using YOUR API key |
| AI calls | Go directly from your Azure Function to Anthropic. Microsoft and AICOS developers never see your data |
| Classification results | Stored in YOUR Azure Table Storage. No external database |
| To Do tasks | Created in YOUR Microsoft To Do via Graph API. Only you can see them |
| Finance/Job data | Written to Excel workbooks in YOUR OneDrive. You own the files |
| Email bodies | Body preview (first 600 chars) sent to Claude for Tier 2 only. Not stored — only metadata (subject, sender, dimension, verb) is persisted |
| Diagnostic counters | In-memory only. Reset on each Function App restart. No telemetry sent anywhere |

**In short:** AICOS runs in your Azure account, reads your email, calls AI with your key, and writes results to your OneDrive and Microsoft To Do. The developer has zero access to your data.

### Refresh token security

Your OAuth2 refresh token gives read/write access to your Outlook mailbox.

- Never commit `local.settings.json` to git — it's in `.gitignore`
- In production it lives in Azure Key Vault — the Function App reads it via Managed Identity
- If exposed: go to [myapps.microsoft.com](https://myapps.microsoft.com) → find `aicos-outlook` → Remove. Then re-run `node scripts/getToken.mjs`
- Tokens expire after 90 days of inactivity or if your Microsoft password changes

---

## Roadmap

### Shipping Now (v1)

- ✅ 8-dimension email classification with Claude Haiku AI
- ✅ Finance tracker with receipt/bill logging to Excel on OneDrive
- ✅ Job application tracker with state machine
- ✅ Microsoft To Do task creation with priority and due dates
- ✅ Real-time Graph webhook with auto-renewal
- ✅ Rich morning briefing — Priority Actions, Your Day, Finance Health, Signals, Dimensions chart, Components status
- ✅ Idempotent StateEngine — no duplicate processing across overlapping runs
- ✅ Test data seeder — all 8 dimensions + calendar events

### Planned

- **Warm-up scan** — analyse last 200 emails on first deploy to seed Tier 1 known senders
- **LearningEngine** — auto-promote reliable senders from Tier 2 to Tier 1 over time
- **Outlook Add-in** — in-client task pane showing AICOS classification inline
- **Teams integration** — route meeting emails to Teams channels
- **Weekly digest** — Sunday summary by dimension
- **WhatsApp pipeline** — same classification core, different input source
- **Asset tracker** — purchase/subscription tracking with renewal alerts
- **Power BI dashboard** — visualise email patterns from Excel tracker data

---

## License

Personal use. Contact for commercial licensing.
