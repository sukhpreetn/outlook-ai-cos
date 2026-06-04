# AI Chief of Staff — Outlook Edition

> Port of [AIChiefOfStaff](https://github.com/...) from Gmail + Google Apps Script to **Outlook + Azure Functions + Microsoft Graph**, with **GitHub Copilot CLI** as the development companion.

---

## What This Does

An AI-powered personal operating system for your Outlook inbox:

| Original (Gmail) | This Project (Outlook) |
|---|---|
| Google Apps Script | Azure Functions (TypeScript) |
| GmailApp API | Microsoft Graph API |
| Gemini AI | Claude (Anthropic) |
| Google Sheets | Excel on OneDrive |
| Todoist | Microsoft To Do |
| PropertiesService | Azure Table Storage |
| ScriptProperties | Azure App Configuration |
| GAS triggers | Timer Trigger + Graph Webhooks |

### Pipeline (same logic, new plumbing)

```
Outlook Inbox
  → Graph Webhook / Timer Trigger
  → ClassifierRunner.ts  (orchestration)
  → Classifier.ts        (Tier 0/1/2 classification via Claude)
      → SensingEngine    (rule-based signal extraction)
      → Claude Haiku     (Tier 1/2 only)
  → Router.ts            (verb-driven routing)
      → FinanceTracker.ts    (receipts → Excel)
      → JobTracker.ts        (career → Excel)
      → AssetTracker.ts      (orders → Excel)
      → Microsoft To Do      (tasks)
  → .ts   (daily digest email via Claude Sonnet)
```

---

## Project Structure

```
outlook-ai-cos/
├── src/
│   ├── aicos-core/          # System layer (Config, AI, Logger, GraphClient, StateEngine)
│   │   ├── Config.ts        ← Azure App Configuration + Key Vault reader
│   │   ├── AI.ts            ← Claude API wrapper (task-based routing)
│   │   ├── Logger.ts        ← Centralized logging + diagnostic counters
│   │   ├── GraphClient.ts   ← All Microsoft Graph operations
│   │   └── StateEngine.ts   ← Azure Table Storage state management
│   │
│   ├── email-pipeline/      # Pipeline layer
│   │   ├── Classifier.ts         ← Tier 0/1/2 email classification
│   │   ├── ClassifierRunner.ts   ← Orchestrator (timer trigger handler)
│   │   ├── Router.ts             ← Verb-driven routing engine
│   │   ├── FinanceTracker.ts     ← Receipts/bills → Excel
│   │   ├── JobTracker.ts         ← Career emails → Excel
│   │   ├── AssetTracker.ts       ← Orders/subs → Excel
│   │   └── MorningBriefing.ts    ← Daily digest email
│   │
│   └── functions/           # Azure Function entry points
│       ├── timerTrigger.ts       ← Every 15 min + daily briefing
│       └── webhookTrigger.ts     ← Graph change notification handler
│
├── tests/
├── scripts/
│   ├── deploy.sh            ← Full Azure deploy script
│   └── registerWebhook.js   ← Graph subscription registration
├── .env.local               ← Local dev template (copy to .env)
├── package.json
└── tsconfig.json
```

---

## Prerequisites

```bash
# Azure CLI
brew install azure-cli && az login

# Azure Functions Core Tools v4
npm install -g azure-functions-core-tools@4

# GitHub CLI + Copilot extension
brew install gh
gh auth login
gh extension install github/gh-copilot

# Node.js 20+
node --version  # must be >= 20
```

---

## Azure App Registration (Entra ID)

Your function needs these **Application permissions** (not delegated):

```
Mail.Read
Mail.ReadWrite
Mail.Send
Calendars.Read
Tasks.ReadWrite
Files.ReadWrite.All
```

```bash
# Copilot CLI can help with this:
gh copilot suggest "create an Azure app registration with Mail.Read and Mail.Send application permissions using az cli"
```

---

## Local Development

```bash
# 1. Install dependencies
npm install

# 2. Set up environment
cp .env.local .env
# Edit .env with your values

# 3. Start Azurite (local Azure Storage emulator)
npx azurite --location ./tmp/azurite

# 4. Run the function app locally
npm start
# or: func start

# 5. Test the classifier manually
curl -X POST http://localhost:7071/api/webhook/graph \
  -H "Content-Type: application/json" \
  -d '{"value":[{"changeType":"created","clientState":"local-dev-secret","resource":"users/you@domain.com/messages/ABC123","subscriptionId":"test"}]}'
```

---

## Copilot CLI Workflow

GitHub Copilot CLI (`gh copilot`) is the development companion throughout this project.

### During development

```bash
# Generate boilerplate for a new pipeline module
gh copilot suggest "create a TypeScript Azure Function that reads emails from Microsoft Graph and writes to Azure Table Storage"

# Understand unfamiliar Graph API patterns
gh copilot explain "what does the Microsoft Graph subscription expirationDateTime field do and how do I renew it?"

# Get az CLI commands
gh copilot suggest "az cli command to assign Key Vault secret reader role to a managed identity"

# Debug deployment issues
gh copilot explain "why would an Azure Function with managed identity fail to authenticate to Microsoft Graph?"

# Generate test data
gh copilot suggest "TypeScript mock for a Microsoft Graph Message object with realistic email data"
```

### Alias for faster access

```bash
# Add to ~/.zshrc or ~/.bashrc
alias cs='gh copilot suggest'
alias ce='gh copilot explain'

# Then use:
cs "deploy azure function app with node 20 runtime"
ce "what is the difference between Application and Delegated permissions in Microsoft Graph"
```

---

## Deployment

```bash
# Full deploy (creates all Azure resources)
ANTHROPIC_API_KEY=sk-ant-... \
RESOURCE_GROUP=my-aicos-rg \
APP_NAME=my-aicos \
./scripts/deploy.sh

# After deploy: register Graph webhook
TENANT_ID=... CLIENT_ID=... CLIENT_SECRET=... \
USER_EMAIL=you@domain.com \
WEBHOOK_URL=https://my-aicos.azurewebsites.net/api/webhook/graph?code=... \
node scripts/registerWebhook.js
```

---

## Configuration Keys (Azure App Configuration)

| Key | Default | Description |
|-----|---------|-------------|
| `aicos:userEmail` | — | **Required.** Mailbox to monitor |
| `aicos:tenantId` | — | **Required.** Entra ID tenant |
| `aicos:clientId` | — | **Required.** App registration |
| `aicos:aiModelClassify` | `claude-haiku-20240307` | Fast classification model |
| `aicos:aiModelBrief` | `claude-sonnet-4-5` | Daily briefing model |
| `aicos:batchSize` | `20` | Emails per run |
| `aicos:briefingHour` | `7` | Hour to send briefing (local time) |
| `aicos:briefingTz` | `America/Los_Angeles` | IANA timezone |
| `aicos:draftCapPerRun` | `3` | Max AI drafts per trigger run |

### Key Vault Secrets

| Secret Name | Description |
|-------------|-------------|
| `anthropic-api-key` | Anthropic API key |
| `aad-client-secret` | Entra ID client secret |
| `storage-connection-string` | Azure Table Storage connection |

---

## Classification Model

Same as original AICOS:

| Tier | Description | Cost |
|------|-------------|------|
| **0** | Hard drop (spam/promo patterns) | Free |
| **1** | Known domain / rule match (≥ 0.85 confidence) | Free |
| **2** | Full Claude classification | ~$0.000025/email (Haiku) |

**Dimensions:** career · finance · legal · knowledge · ventures · life · commerce · infra  
**Verbs:** do · pay · meet · file · read

---

## Monitoring

```bash
# Live logs
az functionapp logs tail --name YOUR_APP --resource-group YOUR_RG

# Application Insights (if configured)
az monitor app-insights query \
  --app YOUR_APP \
  --analytics-query "traces | where message contains '[Runner]' | take 50"

# Copilot CLI for log analysis
gh copilot explain "this Azure Function error: $(az functionapp logs tail ... | tail -20)"
```

---

## Extending

### Add a new pipeline module (e.g. TravelTracker)

```bash
# Let Copilot scaffold it
gh copilot suggest "TypeScript module that extracts flight booking details from email body and appends to an Excel table using Microsoft Graph API"
```

Then:
1. Create `src/email-pipeline/TravelTracker.ts`
2. Add `recordTravelEmail()` call in `Router.ts` under the relevant dimension/verb
3. Add a workbook ID to `.env.local` and App Configuration

### Add a new classification rule

Rules are stored in Azure Table Storage under `PartitionKey = 'rule'`. Use the `saveRule()` function from `StateEngine.ts` or add via a setup script.

---

## Differences from Original AICOS

| Aspect | Original | This Port |
|--------|----------|-----------|
| Runtime | Google Apps Script (JS) | Azure Functions (TypeScript) |
| Auth | OAuth via GAS | Managed Identity + App Registration |
| Email categories | Gmail labels (`~OS:New`) | Outlook categories (`AICOS:New`) |
| State | PropertiesService | Azure Table Storage |
| Config | Google Sheets `Config` tab | Azure App Configuration |
| Secrets | Script Properties | Azure Key Vault |
| Tasks | Todoist REST API | Microsoft To Do (Graph API) |
| Trackers | Google Sheets | Excel on OneDrive (Graph API) |
| Execution limit | 6 min (GAS) | 5 min (Consumption plan, configurable) |
| AI | Gemini | Claude (Anthropic) |
