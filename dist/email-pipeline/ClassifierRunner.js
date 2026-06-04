"use strict";
// =============================================================================
// src/email-pipeline/ClassifierRunner.ts
//
// ORCHESTRATOR — bridges Classifier ↔ Router.
// Installed as an Azure Function timer trigger (every 15 minutes).
//
// FLOW PER RUN:
//   1. Auto-label: tag unread messages with "AICOS:New"
//   2. Fetch all AICOS:New messages (dynamic batch size)
//   3. Skip already-processed messages (StateEngine)
//   4. For each message: classify → state-update → route
//   5. Log batch stats
//
// EXECUTION GUARD:
//   Azure Functions have a default 5-minute timeout.
//   We check elapsed time before each message and stop at 4.5 min.
// =============================================================================
Object.defineProperty(exports, "__esModule", { value: true });
exports.runClassifierAndRouter = runClassifierAndRouter;
const Config_js_1 = require("../aicos-core/Config.js");
const Logger_js_1 = require("../aicos-core/Logger.js");
const GraphClient_js_1 = require("../aicos-core/GraphClient.js");
const StateEngine_js_1 = require("../aicos-core/StateEngine.js");
const Classifier_js_1 = require("./Classifier.js");
const Router_js_1 = require("./Router.js");
const MAX_RUN_MS = 270_000; // 4.5 minutes
// ── Dynamic batch sizing (mirrors AICOS _getDynamicBatchSize_) ────────────────
// In a fresh deploy, default to 20. After runs, the ratio of Tier 2 (AI) calls
// can be read from Table Storage to shrink the batch.
async function _getDynamicBatchSize() {
    try {
        const cfg = await (0, Config_js_1.getConfig)();
        return cfg.batchSize ?? 20;
    }
    catch {
        return 20;
    }
}
// ── Main entry point (called by Azure Function timer trigger) ─────────────────
async function runClassifierAndRouter() {
    const startTime = Date.now();
    const cfg = await (0, Config_js_1.getConfig)();
    Logger_js_1.Logger.info('================================================================');
    Logger_js_1.Logger.info(`[Runner] Timer fired at ${new Date().toISOString()}`);
    Logger_js_1.Logger.info('================================================================');
    // 0. Auto-label new messages
    const labeled = await (0, GraphClient_js_1.autoLabelNewMessages)(cfg.userEmail);
    Logger_js_1.Logger.info(`[Runner] Auto-labeled ${labeled} new messages`);
    // 1. Fetch AICOS:New messages
    const batchSize = await _getDynamicBatchSize();
    Logger_js_1.Logger.info(`[Runner] Batch size: ${batchSize}`);
    const messages = await (0, GraphClient_js_1.fetchNewMessages)(cfg.userEmail, batchSize);
    if (messages.length === 0) {
        Logger_js_1.Logger.info('[Runner] No AICOS:New messages. Done.');
        return;
    }
    Logger_js_1.Logger.info(`[Runner] Found ${messages.length} message(s) to process`);
    // 2. Process each message
    const stats = {
        processed: 0,
        dropped: 0,
        review: 0,
        errors: 0,
        skipped: 0,
        tier0: 0,
        tier1: 0,
        tier2: 0,
    };
    for (const msg of messages) {
        // Execution time guard
        if (Date.now() - startTime > MAX_RUN_MS) {
            Logger_js_1.Logger.warn(`[Runner] Time guard hit after ${stats.processed} messages. Stopping.`);
            break;
        }
        // Idempotency check
        const canProcess = await (0, StateEngine_js_1.canProcessMessage)(msg.id);
        if (!canProcess) {
            stats.skipped++;
            Logger_js_1.Logger.info(`[Runner] SKIP (already processed): ${msg.subject?.substring(0, 50)}`);
            continue;
        }
        try {
            // Classify
            const cls = await (0, Classifier_js_1.classifyMessage)(msg);
            if (cls === null) {
                // Tier 0 drop
                stats.dropped++;
                stats.tier0++;
                await (0, StateEngine_js_1.setMessageState)(msg.id, 'dropped');
                await (0, GraphClient_js_1.categorizeMessage)(cfg.userEmail, msg.id, ['AICOS:Dropped'], ['AICOS:New']);
                continue;
            }
            stats[`tier${cls.tier}`]++;
            // Low confidence → review queue
            if (cls.confidence < 0.60 && cls.urgency < 4) {
                stats.review++;
                await (0, StateEngine_js_1.setMessageState)(msg.id, 'review', cls);
                await (0, GraphClient_js_1.categorizeMessage)(cfg.userEmail, msg.id, ['AICOS:Review'], ['AICOS:New']);
                Logger_js_1.Logger.info(`[Runner] LOW CONF → Review: ${msg.subject?.substring(0, 50)}`);
                continue;
            }
            // Mark classified
            await (0, StateEngine_js_1.setMessageState)(msg.id, 'classified', cls);
            await (0, GraphClient_js_1.categorizeMessage)(cfg.userEmail, msg.id, ['AICOS:Classified'], ['AICOS:New']);
            // Route
            await (0, Router_js_1.routeClassification)(msg, cls, cfg);
            // Mark actioned
            await (0, StateEngine_js_1.setMessageState)(msg.id, 'actioned');
            await (0, GraphClient_js_1.categorizeMessage)(cfg.userEmail, msg.id, ['AICOS:Actioned'], ['AICOS:Classified']);
            stats.processed++;
            Logger_js_1.Logger.info(`[Runner] DONE id=${msg.id.substring(0, 8)} ` +
                `dim=${cls.dimension} verb=${cls.verb} tier=${cls.tier} ` +
                `urgency=${cls.urgency} conf=${cls.confidence.toFixed(2)}`);
        }
        catch (e) {
            stats.errors++;
            Logger_js_1.Logger.error(`[Runner] Error processing ${msg.id}`, e);
            await (0, StateEngine_js_1.setMessageState)(msg.id, 'review');
            await (0, GraphClient_js_1.categorizeMessage)(cfg.userEmail, msg.id, ['AICOS:Review'], ['AICOS:New']);
        }
    }
    const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
    Logger_js_1.Logger.info(`[Runner] COMPLETE in ${elapsed}s | ` +
        `processed=${stats.processed} dropped=${stats.dropped} ` +
        `review=${stats.review} errors=${stats.errors} skipped=${stats.skipped} | ` +
        `tier0=${stats.tier0} tier1=${stats.tier1} tier2=${stats.tier2}`);
    // Counters for DailyDiagnostic
    Logger_js_1.Logger.diagIncrement('DIAG_PROCESSED', stats.processed);
    Logger_js_1.Logger.diagIncrement('DIAG_DROPPED', stats.dropped);
    Logger_js_1.Logger.diagIncrement('DIAG_REVIEW', stats.review);
    Logger_js_1.Logger.diagIncrement('DIAG_ERRORS', stats.errors);
}
//# sourceMappingURL=ClassifierRunner.js.map