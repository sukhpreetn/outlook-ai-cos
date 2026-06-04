"use strict";
// =============================================================================
// src/aicos-core/Config.ts
//
// SINGLE SOURCE OF TRUTH for all settings.
// Mirrors the original AICOS Config.js pattern but reads from:
//   1. Azure App Configuration (all non-secret settings)
//   2. Azure Key Vault (secrets: API keys, client secrets)
//   3. Environment variables as local dev fallback
//
// USAGE (from any module):
//   const cfg = await getConfig();
//   const key = cfg.get('anthropicApiKey');
//
// RULE: Never hardcode values. If it can change, it belongs in Config.
// =============================================================================
Object.defineProperty(exports, "__esModule", { value: true });
exports.getConfig = getConfig;
exports.invalidateConfigCache = invalidateConfigCache;
const app_configuration_1 = require("@azure/app-configuration");
const keyvault_secrets_1 = require("@azure/keyvault-secrets");
const identity_1 = require("@azure/identity");
// ── Singleton cache ───────────────────────────────────────────────────────────
let _configCache = null;
let _cacheExpiry = 0;
const CACHE_TTL_MS = 5 * 60 * 1000; // 5 minutes
// ── Defaults (overridden by App Configuration) ────────────────────────────────
const DEFAULTS = {
    aiModelClassify: 'claude-haiku-20240307',
    aiModelBrief: 'claude-sonnet-4-5',
    aiModelDraft: 'claude-haiku-20240307',
    batchSize: 25,
    tier1ConfidenceFloor: 0.85,
    tier2ConfidenceFloor: 0.60,
    draftCapPerRun: 3,
    briefingHour: 7,
    briefingTz: 'America/Los_Angeles',
};
// ── Main export ───────────────────────────────────────────────────────────────
async function getConfig() {
    const now = Date.now();
    if (_configCache && now < _cacheExpiry)
        return _configCache;
    // Local dev: return env-var-backed config immediately
    if (process.env.AICOS_ENV === 'local') {
        _configCache = _buildLocalConfig();
        _cacheExpiry = now + CACHE_TTL_MS;
        return _configCache;
    }
    try {
        const [appCfg, secrets] = await Promise.all([
            _fetchAppConfiguration(),
            _fetchKeyVaultSecrets(),
        ]);
        _configCache = {
            ...DEFAULTS,
            ...appCfg,
            ...secrets,
        };
        _cacheExpiry = now + CACHE_TTL_MS;
        return _configCache;
    }
    catch (e) {
        throw new Error(`[Config] Failed to load configuration: ${e.message}`);
    }
}
// ── Azure App Configuration reader ───────────────────────────────────────────
async function _fetchAppConfiguration() {
    const endpoint = process.env.APP_CONFIG_ENDPOINT;
    if (!endpoint)
        return {};
    const client = new app_configuration_1.AppConfigurationClient(endpoint, new identity_1.DefaultAzureCredential());
    const result = {};
    // Map App Configuration keys → config fields
    const keyMap = {
        'aicos:aiModelClassify': 'aiModelClassify',
        'aicos:aiModelBrief': 'aiModelBrief',
        'aicos:aiModelDraft': 'aiModelDraft',
        'aicos:tenantId': 'tenantId',
        'aicos:clientId': 'clientId',
        'aicos:userEmail': 'userEmail',
        'aicos:oneDriveFolderId': 'oneDriveFolderId',
        'aicos:batchSize': 'batchSize',
        'aicos:tier1ConfidenceFloor': 'tier1ConfidenceFloor',
        'aicos:tier2ConfidenceFloor': 'tier2ConfidenceFloor',
        'aicos:draftCapPerRun': 'draftCapPerRun',
        'aicos:briefingHour': 'briefingHour',
        'aicos:briefingTz': 'briefingTz',
    };
    for (const [azKey, cfgKey] of Object.entries(keyMap)) {
        try {
            const setting = await client.getConfigurationSetting({ key: azKey });
            const raw = setting.value ?? '';
            // Coerce numerics
            result[cfgKey] = isNaN(Number(raw)) ? raw : Number(raw);
        }
        catch {
            // Key not set — fall through to defaults
        }
    }
    return result;
}
// ── Key Vault secrets reader ──────────────────────────────────────────────────
async function _fetchKeyVaultSecrets() {
    const vaultUrl = process.env.KEY_VAULT_URL;
    if (!vaultUrl)
        return {};
    const client = new keyvault_secrets_1.SecretClient(vaultUrl, new identity_1.DefaultAzureCredential());
    const [anthropicKey, clientSecret, storageConn] = await Promise.all([
        client.getSecret('anthropic-api-key').then(s => s.value ?? ''),
        client.getSecret('aad-client-secret').then(s => s.value ?? ''),
        client.getSecret('storage-connection-string').then(s => s.value ?? ''),
    ]);
    return {
        anthropicApiKey: anthropicKey,
        clientSecret: clientSecret,
        storageConnectionString: storageConn,
    };
}
// ── Local dev fallback ────────────────────────────────────────────────────────
function _buildLocalConfig() {
    return {
        ...DEFAULTS,
        aiModelClassify: process.env.AI_MODEL_CLASSIFY ?? 'claude-haiku-20240307',
        aiModelBrief: process.env.AI_MODEL_BRIEF ?? 'claude-sonnet-4-5',
        aiModelDraft: process.env.AI_MODEL_DRAFT ?? 'claude-haiku-20240307',
        tenantId: process.env.TENANT_ID ?? '',
        clientId: process.env.CLIENT_ID ?? '',
        clientSecret: process.env.CLIENT_SECRET ?? '',
        userEmail: process.env.USER_EMAIL ?? '',
        anthropicApiKey: process.env.ANTHROPIC_API_KEY ?? '',
        storageConnectionString: process.env.STORAGE_CONNECTION_STRING ?? 'UseDevelopmentStorage=true',
        oneDriveFolderId: process.env.ONEDRIVE_FOLDER_ID ?? '',
    };
}
// ── Convenience getter ────────────────────────────────────────────────────────
function invalidateConfigCache() {
    _configCache = null;
    _cacheExpiry = 0;
}
//# sourceMappingURL=Config.js.map