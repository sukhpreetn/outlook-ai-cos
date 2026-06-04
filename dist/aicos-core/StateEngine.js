"use strict";
// =============================================================================
// src/aicos-core/StateEngine.ts
//
// STATE MANAGEMENT — tracks processed messages across Azure Function runs.
// Equivalent to AICOS StateEngine.js but uses Azure Table Storage.
//
// WHY: Azure Functions are stateless. We use Table Storage as a durable
// key-value store to track which messages have been processed, preventing
// duplicate classification/routing across overlapping timer runs.
//
// TABLE SCHEMA (table: "AicosState"):
//   PartitionKey = entity type (e.g. "message", "rule", "run")
//   RowKey       = entity ID
//   state        = JSON blob with classification result, timestamps, etc.
// =============================================================================
Object.defineProperty(exports, "__esModule", { value: true });
exports.canProcessMessage = canProcessMessage;
exports.setMessageState = setMessageState;
exports.getMessageClassification = getMessageClassification;
exports.saveRule = saveRule;
exports.loadRules = loadRules;
exports.purgeOldStateRecords = purgeOldStateRecords;
const data_tables_1 = require("@azure/data-tables");
const Config_js_1 = require("./Config.js");
const Logger_js_1 = require("./Logger.js");
const TABLE_NAME = 'AicosState';
const PROCESSED_TTL_DAYS = 30; // clean up processed records after 30 days
// ── Singleton Table client ────────────────────────────────────────────────────
let _table = null;
async function getTable() {
    if (_table)
        return _table;
    const cfg = await (0, Config_js_1.getConfig)();
    _table = data_tables_1.TableClient.fromConnectionString(cfg.storageConnectionString, TABLE_NAME);
    await _table.createTable(); // no-op if exists
    return _table;
}
// ── Core state operations ─────────────────────────────────────────────────────
/**
 * Check if a message has already been processed.
 * Equivalent to canProcessItem() in AICOS StateEngine.js
 */
async function canProcessMessage(messageId) {
    try {
        const table = await getTable();
        const entity = await table.getEntity('message', _sanitizeKey(messageId));
        // Re-process only if stuck in 'review' for > 1 hour
        if (entity.state === 'review') {
            const actionedAt = entity.timestamp ? new Date(entity.timestamp) : null;
            if (actionedAt && Date.now() - actionedAt.getTime() > 60 * 60 * 1000) {
                return true; // retry
            }
        }
        return false; // already processed
    }
    catch {
        return true; // entity not found → process it
    }
}
/**
 * Mark a message as processed with a given state.
 * Equivalent to markItemProcessed() + transitionState() in AICOS.
 */
async function setMessageState(messageId, state, classification) {
    try {
        const table = await getTable();
        const now = new Date().toISOString();
        await table.upsertEntity({
            partitionKey: 'message',
            rowKey: _sanitizeKey(messageId),
            state,
            updatedAt: now,
            ...(state === 'classified' ? { classifiedAt: now } : {}),
            ...(state === 'actioned' ? { actionedAt: now } : {}),
            ...(classification ? { classification: JSON.stringify(classification) } : {}),
        }, 'Replace');
    }
    catch (e) {
        Logger_js_1.Logger.error('[StateEngine] setMessageState failed', { messageId, state, error: e.message });
    }
}
/**
 * Retrieve classification result for a message.
 */
async function getMessageClassification(messageId) {
    try {
        const table = await getTable();
        const entity = await table.getEntity('message', _sanitizeKey(messageId));
        return entity.classification ? JSON.parse(entity.classification) : null;
    }
    catch {
        return null;
    }
}
async function saveRule(rule) {
    const table = await getTable();
    await table.upsertEntity({
        partitionKey: 'rule',
        rowKey: rule.id,
        ...rule,
    }, 'Replace');
}
async function loadRules() {
    const table = await getTable();
    const rules = [];
    const entities = table.listEntities({
        queryOptions: { filter: (0, data_tables_1.odata) `PartitionKey eq 'rule'` },
    });
    for await (const entity of entities) {
        rules.push({
            id: entity.rowKey,
            name: entity.name,
            condition: entity.condition,
            dimension: entity.dimension,
            verb: entity.verb,
            confidence: Number(entity.confidence),
            hits: Number(entity.hits),
            corrections: Number(entity.corrections),
            accuracy: Number(entity.accuracy),
            lastHit: entity.lastHit,
            createdAt: entity.createdAt,
        });
    }
    return rules.sort((a, b) => b.confidence - a.confidence);
}
// ── Cleanup ───────────────────────────────────────────────────────────────────
/**
 * Delete processed message records older than PROCESSED_TTL_DAYS.
 * Run nightly from DailyDiagnostic.
 */
async function purgeOldStateRecords() {
    const table = await getTable();
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - PROCESSED_TTL_DAYS);
    let deleted = 0;
    const entities = table.listEntities({
        queryOptions: {
            filter: (0, data_tables_1.odata) `PartitionKey eq 'message' and Timestamp lt ${cutoff}`,
        },
    });
    for await (const entity of entities) {
        await table.deleteEntity('message', entity.rowKey);
        deleted++;
    }
    Logger_js_1.Logger.info(`[StateEngine] Purged ${deleted} old state records`);
    return deleted;
}
// ── Helpers ───────────────────────────────────────────────────────────────────
function _sanitizeKey(id) {
    // Azure Table Storage row keys cannot contain /, \, #, ?
    return id.replace(/[/\\#?]/g, '_');
}
//# sourceMappingURL=StateEngine.js.map