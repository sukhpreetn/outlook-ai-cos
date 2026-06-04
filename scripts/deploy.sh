#!/usr/bin/env bash
# =============================================================================
# scripts/deploy.sh
#
# Deployment script for Outlook AI Chief of Staff.
# Uses GitHub Copilot CLI (gh copilot) for AI-assisted Azure setup.
#
# USAGE:
#   ./scripts/deploy.sh [--env prod|staging] [--resource-group my-rg]
#
# PREREQUISITES:
#   - Azure CLI logged in:   az login
#   - GitHub CLI + Copilot:  gh auth login && gh extension install github/gh-copilot
#   - Node.js 20+, npm
#   - Azure Functions Core Tools v4
# =============================================================================

set -euo pipefail

# ── Config ────────────────────────────────────────────────────────────────────
RESOURCE_GROUP="${RESOURCE_GROUP:-aicos-outlook-rg}"
LOCATION="${LOCATION:-westus2}"
APP_NAME="${APP_NAME:-aicos-outlook}"
STORAGE_ACCOUNT="${STORAGE_ACCOUNT:-aicosstorage$RANDOM}"
KEYVAULT_NAME="${KEYVAULT_NAME:-aicos-kv-$RANDOM}"
APP_CONFIG_NAME="${APP_CONFIG_NAME:-aicos-appconfig}"
ENV="${DEPLOY_ENV:-prod}"

echo "================================================================"
echo "  AI Chief of Staff — Outlook Edition"
echo "  Deploying to: $RESOURCE_GROUP / $APP_NAME ($ENV)"
echo "================================================================"

# ── Step 0: Copilot CLI setup helper ─────────────────────────────────────────
copilot_suggest() {
  echo ""
  echo "  [Copilot] Tip: $1"
  echo "  Run: gh copilot suggest \"$1\""
  echo ""
}

# ── Step 1: Build ─────────────────────────────────────────────────────────────
echo "[1/8] Building TypeScript..."
npm run build
echo "      ✓ Build complete"

# ── Step 2: Resource Group ────────────────────────────────────────────────────
echo "[2/8] Ensuring resource group..."
az group create \
  --name "$RESOURCE_GROUP" \
  --location "$LOCATION" \
  --output none 2>/dev/null || true

# ── Step 3: Storage Account ───────────────────────────────────────────────────
echo "[3/8] Creating storage account..."
az storage account create \
  --name "$STORAGE_ACCOUNT" \
  --resource-group "$RESOURCE_GROUP" \
  --location "$LOCATION" \
  --sku Standard_LRS \
  --output none 2>/dev/null || true

STORAGE_CONN=$(az storage account show-connection-string \
  --name "$STORAGE_ACCOUNT" \
  --resource-group "$RESOURCE_GROUP" \
  --query connectionString -o tsv)

# ── Step 4: Key Vault ─────────────────────────────────────────────────────────
echo "[4/8] Setting up Key Vault..."
az keyvault create \
  --name "$KEYVAULT_NAME" \
  --resource-group "$RESOURCE_GROUP" \
  --location "$LOCATION" \
  --output none 2>/dev/null || true

# Store secrets — prompt if not in environment
if [[ -z "${ANTHROPIC_API_KEY:-}" ]]; then
  copilot_suggest "What is the correct az keyvault secret set command for storing an Anthropic API key?"
  read -rsp "Enter your Anthropic API key: " ANTHROPIC_API_KEY
  echo ""
fi

az keyvault secret set \
  --vault-name "$KEYVAULT_NAME" \
  --name "anthropic-api-key" \
  --value "$ANTHROPIC_API_KEY" \
  --output none

az keyvault secret set \
  --vault-name "$KEYVAULT_NAME" \
  --name "storage-connection-string" \
  --value "$STORAGE_CONN" \
  --output none

if [[ -n "${AAD_CLIENT_SECRET:-}" ]]; then
  az keyvault secret set \
    --vault-name "$KEYVAULT_NAME" \
    --name "aad-client-secret" \
    --value "$AAD_CLIENT_SECRET" \
    --output none
fi

# ── Step 5: App Configuration ─────────────────────────────────────────────────
echo "[5/8] Configuring App Configuration..."
az appconfig create \
  --name "$APP_CONFIG_NAME" \
  --resource-group "$RESOURCE_GROUP" \
  --location "$LOCATION" \
  --sku Free \
  --output none 2>/dev/null || true

APP_CONFIG_ENDPOINT=$(az appconfig show \
  --name "$APP_CONFIG_NAME" \
  --resource-group "$RESOURCE_GROUP" \
  --query endpoint -o tsv)

# Set default config values
declare -A CONFIG_VALUES=(
  ["aicos:aiModelClassify"]="claude-haiku-20240307"
  ["aicos:aiModelBrief"]="claude-sonnet-4-5"
  ["aicos:aiModelDraft"]="claude-haiku-20240307"
  ["aicos:batchSize"]="20"
  ["aicos:briefingHour"]="7"
  ["aicos:briefingTz"]="America/Los_Angeles"
  ["aicos:draftCapPerRun"]="3"
)

for key in "${!CONFIG_VALUES[@]}"; do
  az appconfig kv set \
    --name "$APP_CONFIG_NAME" \
    --key "$key" \
    --value "${CONFIG_VALUES[$key]}" \
    --yes \
    --output none
done

# ── Step 6: Function App ──────────────────────────────────────────────────────
echo "[6/8] Creating Azure Function App..."
az functionapp create \
  --name "$APP_NAME" \
  --resource-group "$RESOURCE_GROUP" \
  --storage-account "$STORAGE_ACCOUNT" \
  --consumption-plan-location "$LOCATION" \
  --runtime node \
  --runtime-version 20 \
  --functions-version 4 \
  --output none 2>/dev/null || true

# Enable managed identity
az functionapp identity assign \
  --name "$APP_NAME" \
  --resource-group "$RESOURCE_GROUP" \
  --output none

PRINCIPAL_ID=$(az functionapp identity show \
  --name "$APP_NAME" \
  --resource-group "$RESOURCE_GROUP" \
  --query principalId -o tsv)

# Grant Key Vault access
az keyvault set-policy \
  --name "$KEYVAULT_NAME" \
  --object-id "$PRINCIPAL_ID" \
  --secret-permissions get list \
  --output none

# Set app settings
KEYVAULT_URL="https://${KEYVAULT_NAME}.vault.azure.net"
az functionapp config appsettings set \
  --name "$APP_NAME" \
  --resource-group "$RESOURCE_GROUP" \
  --settings \
    "KEY_VAULT_URL=$KEYVAULT_URL" \
    "APP_CONFIG_ENDPOINT=$APP_CONFIG_ENDPOINT" \
    "WEBHOOK_SECRET=aicos-$(openssl rand -hex 8)" \
    "NODE_ENV=production" \
  --output none

# ── Step 7: Deploy ────────────────────────────────────────────────────────────
echo "[7/8] Deploying function app..."
func azure functionapp publish "$APP_NAME" --typescript
echo "      ✓ Deployed"

# ── Step 8: Register Graph Webhook ───────────────────────────────────────────
echo "[8/8] Registering Microsoft Graph webhook..."
FUNCTION_KEY=$(az functionapp keys list \
  --name "$APP_NAME" \
  --resource-group "$RESOURCE_GROUP" \
  --query functionKeys.default -o tsv 2>/dev/null || echo "")

WEBHOOK_URL="https://${APP_NAME}.azurewebsites.net/api/webhook/graph?code=${FUNCTION_KEY}"
echo "      Webhook URL: $WEBHOOK_URL"
echo ""
echo "      Register this webhook in your app registration or run:"
echo "      node scripts/registerWebhook.js"

copilot_suggest "How do I register a Microsoft Graph subscription for Outlook mail notifications using REST API?"

# ── Done ──────────────────────────────────────────────────────────────────────
echo ""
echo "================================================================"
echo "  ✓ Deployment complete!"
echo ""
echo "  Function App:  https://${APP_NAME}.azurewebsites.net"
echo "  Key Vault:     $KEYVAULT_URL"
echo "  App Config:    $APP_CONFIG_ENDPOINT"
echo ""
echo "  Next steps:"
echo "  1. Set USER_EMAIL in App Configuration (aicos:userEmail)"
echo "  2. Set TENANT_ID + CLIENT_ID in App Configuration"
echo "  3. Run: node scripts/registerWebhook.js"
echo "  4. Monitor: az functionapp logs tail --name $APP_NAME -g $RESOURCE_GROUP"
echo "================================================================"
