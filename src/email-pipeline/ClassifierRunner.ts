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

import { getConfig } from '../aicos-core/Config.js';
import { Logger } from '../aicos-core/Logger.js';
import {
  fetchNewMessages,
  autoLabelNewMessages,
  categorizeMessage,
} from '../aicos-core/GraphClient.js';
import {
  canProcessMessage,
  setMessageState,
} from '../aicos-core/StateEngine.js';
import { classifyMessage } from './Classifier.js';
import { routeClassification } from './Router.js';

const MAX_RUN_MS = 270_000; // 4.5 minutes

// ── Dynamic batch sizing (mirrors AICOS _getDynamicBatchSize_) ────────────────
// In a fresh deploy, default to 20. After runs, the ratio of Tier 2 (AI) calls
// can be read from Table Storage to shrink the batch.
async function _getDynamicBatchSize(): Promise<number> {
  try {
    const cfg = await getConfig();
    return cfg.batchSize ?? 20;
  } catch {
    return 20;
  }
}

// ── Main entry point (called by Azure Function timer trigger) ─────────────────
export async function runClassifierAndRouter(): Promise<void> {
  const startTime = Date.now();
  const cfg = await getConfig();

  Logger.info('================================================================');
  Logger.info(`[Runner] Timer fired at ${new Date().toISOString()}`);
  Logger.info('================================================================');

  // 0. Auto-label new messages
  const labeled = await autoLabelNewMessages(cfg.userEmail);
  Logger.info(`[Runner] Auto-labeled ${labeled} new messages`);

  // 1. Fetch AICOS:New messages
  const batchSize = await _getDynamicBatchSize();
  Logger.info(`[Runner] Batch size: ${batchSize}`);

  const messages = await fetchNewMessages(cfg.userEmail, batchSize);

  if (messages.length === 0) {
    Logger.info('[Runner] No AICOS:New messages. Done.');
    return;
  }

  Logger.info(`[Runner] Found ${messages.length} message(s) to process`);

  // 2. Process each message
  const stats = {
    processed: 0,
    dropped:   0,
    review:    0,
    errors:    0,
    skipped:   0,
    tier0:     0,
    tier1:     0,
    tier2:     0,
  };

  for (const msg of messages) {
    // Execution time guard
    if (Date.now() - startTime > MAX_RUN_MS) {
      Logger.warn(`[Runner] Time guard hit after ${stats.processed} messages. Stopping.`);
      break;
    }

    // Idempotency check
    const canProcess = await canProcessMessage(msg.id);
    if (!canProcess) {
      stats.skipped++;
      Logger.info(`[Runner] SKIP (already processed): ${msg.subject?.substring(0, 50)}`);
      continue;
    }

    try {
      // Classify
      const cls = await classifyMessage(msg);

      if (cls === null) {
        // Tier 0 drop
        stats.dropped++;
        stats.tier0++;
        await setMessageState(msg.id, 'dropped');
        await categorizeMessage(cfg.userEmail, msg.id, ['AICOS:Dropped'], ['AICOS:New']);
        continue;
      }

      stats[`tier${cls.tier}` as 'tier0' | 'tier1' | 'tier2']++;

      // Low confidence → review queue
      if (cls.confidence < 0.60 && cls.urgency < 4) {
        stats.review++;
        await setMessageState(msg.id, 'review', cls);
        await categorizeMessage(cfg.userEmail, msg.id, ['AICOS:Review'], ['AICOS:New']);
        Logger.info(`[Runner] LOW CONF → Review: ${msg.subject?.substring(0, 50)}`);
        continue;
      }

      // Mark classified
      await setMessageState(msg.id, 'classified', cls);
      await categorizeMessage(cfg.userEmail, msg.id, ['AICOS:Classified'], ['AICOS:New']);

      // Route
      await routeClassification(msg, cls, cfg);

      // Mark actioned
      await setMessageState(msg.id, 'actioned');
      await categorizeMessage(cfg.userEmail, msg.id, ['AICOS:Actioned'], ['AICOS:Classified']);

      stats.processed++;
      Logger.info(
        `[Runner] DONE id=${msg.id.substring(0, 8)} ` +
        `dim=${cls.dimension} verb=${cls.verb} tier=${cls.tier} ` +
        `urgency=${cls.urgency} conf=${cls.confidence.toFixed(2)}`
      );

    } catch (e) {
      stats.errors++;
      Logger.error(`[Runner] Error processing ${msg.id}`, e);
      await setMessageState(msg.id, 'review');
      await categorizeMessage(cfg.userEmail, msg.id, ['AICOS:Review'], ['AICOS:New']);
    }
  }

  const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
  Logger.info(
    `[Runner] COMPLETE in ${elapsed}s | ` +
    `processed=${stats.processed} dropped=${stats.dropped} ` +
    `review=${stats.review} errors=${stats.errors} skipped=${stats.skipped} | ` +
    `tier0=${stats.tier0} tier1=${stats.tier1} tier2=${stats.tier2}`
  );

  // Counters for DailyDiagnostic
  Logger.diagIncrement('DIAG_PROCESSED',  stats.processed);
  Logger.diagIncrement('DIAG_DROPPED',    stats.dropped);
  Logger.diagIncrement('DIAG_REVIEW',     stats.review);
  Logger.diagIncrement('DIAG_ERRORS',     stats.errors);
}
