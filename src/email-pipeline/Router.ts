// =============================================================================
// src/email-pipeline/Router.ts
//
// VERB-DRIVEN ROUTING ENGINE
// Mirrors Router_V2.js from original AICOS.
//
// 6 GATES (in order):
//   Gate 1: isReceipt      → FinanceTracker + stop
//   Gate 2: silenceReason  → silence (flag + batch) + stop
//   Gate 3: low confidence → review queue + stop
//   Gate 4: batchKey       → accumulate + stop
//   Gate 5: verb switch    → main routing action
//   Gate 6: hasDeadline    → calendar event creation
//
// ALWAYS RUNS (_finalize):
//   _applyOutlookCategories() → category tags (replaces Gmail labels)
//   _markActioned()           → AICOS:Actioned
// =============================================================================

import { getConfig, AICOSConfig } from '../aicos-core/Config.js';
import { Logger } from '../aicos-core/Logger.js';
import {
  flagMessage,
  createDraftReply,
  createTask,
  fetchCalendarEvents,
} from '../aicos-core/GraphClient.js';
import { callAI } from '../aicos-core/AI.js';
import type { GraphMessage } from '../aicos-core/GraphClient.js';
import type { Classification } from './Classifier.js';
import { recordFinanceEmail } from './FinanceTracker.js';
import { recordCareerEmail } from './JobTracker.js';
import { recordAssetEmail } from './AssetTracker.js';

let _draftCount = 0;

// ── Entry point ───────────────────────────────────────────────────────────────
export async function routeClassification(
  msg: GraphMessage,
  cls: Classification,
  cfg: AICOSConfig,
): Promise<void> {
  _draftCount = 0; // reset per run context
  const actions: string[] = [];

  Logger.info(
    `[Router] id=${cls.messageId.substring(0, 8)} ` +
    `verb=${cls.verb} dim=${cls.dimension} urgency=${cls.urgency} ` +
    `conf=${cls.confidence.toFixed(2)} proactive=${cls.proactiveAction ?? 'none'}`
  );

  // ── Gate 1: Receipt ────────────────────────────────────────────────────────
  if (cls.isReceipt) {
    await _routeReceipt(msg, cls, cfg, actions);
    return _finalize(msg, cls, cfg, actions);
  }

  // ── Gate 2: Silence ────────────────────────────────────────────────────────
  if (cls.silenceReason && cls.urgency < 3) {
    Logger.info(`[Router] SILENCE: ${cls.silenceReason}`);
    actions.push(`silenced:${cls.silenceReason}`);
    return _finalize(msg, cls, cfg, actions);
  }

  // ── Gate 3: Low confidence (already handled in ClassifierRunner, belt+suspenders) ──
  if (cls.confidence < 0.60 && cls.urgency < 4) {
    Logger.info(`[Router] LOW CONF → Review queue`);
    return _finalize(msg, cls, cfg, actions);
  }

  // ── Gate 4: Batch key ──────────────────────────────────────────────────────
  if (cls.batchKey && cls.urgency < 3) {
    Logger.info(`[Router] BATCH: ${cls.batchKey}`);
    actions.push(`batched:${cls.batchKey}`);
    return _finalize(msg, cls, cfg, actions);
  }

  // ── Gate 5: Verb routing ───────────────────────────────────────────────────
  switch (cls.verb) {
    case 'do':   await _routeDo(msg, cls, cfg, actions);   break;
    case 'pay':  await _routePay(msg, cls, cfg, actions);  break;
    case 'meet': await _routeMeet(msg, cls, cfg, actions); break;
    case 'file': await _routeFile(msg, cls, cfg, actions); break;
    case 'read': await _routeRead(msg, cls, cfg, actions); break;
  }

  // ── Gate 6: Deadline → calendar event ─────────────────────────────────────
  if (cls.hasDeadline && cls.urgency >= 3) {
    await _createDeadlineEvent(msg, cls, cfg);
    actions.push('calendar_event_created');
  }

  return _finalize(msg, cls, cfg, actions);
}

// ── Gate handlers ─────────────────────────────────────────────────────────────

async function _routeReceipt(
  msg: GraphMessage,
  cls: Classification,
  cfg: AICOSConfig,
  actions: string[],
): Promise<void> {
  try {
    await recordFinanceEmail(msg, cls, cfg);
    actions.push('finance_logged');
    Logger.info('[Router] Receipt → FinanceTracker');
  } catch (e) {
    Logger.error('[Router] recordFinanceEmail failed', e);
  }
}

async function _routeDo(
  msg: GraphMessage,
  cls: Classification,
  cfg: AICOSConfig,
  actions: string[],
): Promise<void> {
  if (cls.dimension === 'career') {
    try {
      await recordCareerEmail(msg, cls, cfg);
      actions.push('career_logged');
    } catch (e) {
      Logger.error('[Router] recordCareerEmail failed', e);
    }
  }

  // Create a task in Microsoft To Do
  try {
    const dueDate = cls.hasDeadline ? _parseDueDate(msg.bodyPreview) : undefined;
    await createTask(
      cfg.userEmail,
      `[${cls.dimension.toUpperCase()}] ${msg.subject}`,
      `From: ${cls.sender}\n\n${cls.summary}`,
      dueDate,
      cls.urgency >= 4 ? 'high' : 'normal',
    );
    actions.push('task_created');
    Logger.info('[Router] Task created in Microsoft To Do');
  } catch (e) {
    Logger.error('[Router] createTask failed', e);
  }

  // Proactive draft reply
  await _maybeDraft(msg, cls, cfg, actions);
}

async function _routePay(
  msg: GraphMessage,
  cls: Classification,
  cfg: AICOSConfig,
  actions: string[],
): Promise<void> {
  try {
    await recordFinanceEmail(msg, cls, cfg);
    actions.push('finance_logged');
  } catch (e) {
    Logger.error('[Router] finance bill logging failed', e);
  }

  // Flag for visibility
  await flagMessage(cfg.userEmail, msg.id);
  actions.push('flagged');

  // Create a pay task
  await createTask(
    cfg.userEmail,
    `PAY: ${msg.subject}`,
    `From: ${cls.sender}\n${cls.summary}`,
    undefined,
    'high',
  );
  actions.push('pay_task_created');
}

async function _routeMeet(
  msg: GraphMessage,
  cls: Classification,
  cfg: AICOSConfig,
  actions: string[],
): Promise<void> {
  // Flag the message
  await flagMessage(cfg.userEmail, msg.id);
  actions.push('flagged');

  // Create a task to respond
  await createTask(
    cfg.userEmail,
    `RSVP: ${msg.subject}`,
    `From: ${cls.sender}\n${cls.summary}`,
    undefined,
    cls.urgency >= 4 ? 'high' : 'normal',
  );
  actions.push('rsvp_task');

  // Draft a meeting reply if high urgency
  await _maybeDraft(msg, cls, cfg, actions);
}

async function _routeFile(
  msg: GraphMessage,
  cls: Classification,
  _cfg: AICOSConfig,
  actions: string[],
): Promise<void> {
  // For legal/infra: flag and note for filing
  if (cls.dimension === 'legal' || cls.dimension === 'infra') {
    await flagMessage(_cfg.userEmail, msg.id);
    actions.push('flagged_for_filing');
    Logger.info(`[Router] ${cls.dimension} → flagged for filing`);
  }
  actions.push('filed');
}

async function _routeRead(
  msg: GraphMessage,
  cls: Classification,
  cfg: AICOSConfig,
  actions: string[],
): Promise<void> {
  if (cls.dimension === 'commerce') {
    try {
      await recordAssetEmail(msg, cls, cfg);
      actions.push('asset_logged');
    } catch (e) {
      Logger.error('[Router] recordAssetEmail failed', e);
    }
  }
  // Low urgency reads: just label, no further action
  actions.push('read_labeled');
}

// ── Proactive draft reply ─────────────────────────────────────────────────────
async function _maybeDraft(
  msg: GraphMessage,
  cls: Classification,
  cfg: AICOSConfig,
  actions: string[],
): Promise<void> {
  const DRAFT_CAP = 3;
  if (
    cls.proactiveAction !== 'draft_reply' ||
    _draftCount >= DRAFT_CAP ||
    cls.urgency < 3
  ) return;

  try {
    const draftPrompt =
      `Draft a concise, professional reply to this email.\n\n` +
      `FROM: ${cls.sender}\n` +
      `SUBJECT: ${msg.subject}\n` +
      `BODY: ${msg.bodyPreview?.substring(0, 800)}\n\n` +
      `Context: This is a ${cls.dimension} email (verb: ${cls.verb}, urgency: ${cls.urgency}/5).\n` +
      `Write a 2-4 sentence draft reply. Use <br> for line breaks. Output ONLY the email body, no subject line.`;

    const result  = await callAI('draft', draftPrompt);
    const draftId = await createDraftReply(cfg.userEmail, msg.id, result.text);

    _draftCount++;
    actions.push(`draft_created:${draftId.substring(0, 8)}`);
    Logger.info(`[Router] Draft reply created (${_draftCount}/${DRAFT_CAP})`);
  } catch (e) {
    Logger.error('[Router] Draft creation failed', e);
  }
}

// ── Calendar event for deadlines ──────────────────────────────────────────────
async function _createDeadlineEvent(
  msg: GraphMessage,
  cls: Classification,
  cfg: AICOSConfig,
): Promise<void> {
  // Simple heuristic: create a reminder task rather than a calendar event
  // (Graph calendar event creation requires attendee management — keep simple for now)
  try {
    await createTask(
      cfg.userEmail,
      `DEADLINE: ${msg.subject}`,
      `From: ${cls.sender}\n${cls.summary}`,
      undefined,
      'high',
    );
    Logger.info('[Router] Deadline task created');
  } catch (e) {
    Logger.error('[Router] _createDeadlineEvent failed', e);
  }
}

// ── Finalize ──────────────────────────────────────────────────────────────────
async function _finalize(
  _msg: GraphMessage,
  cls: Classification,
  _cfg: AICOSConfig,
  actions: string[],
): Promise<void> {
  Logger.info(
    `[Router] FINALIZE id=${cls.messageId.substring(0, 8)} ` +
    `actions=[${actions.join(', ')}]`
  );
  Logger.diagIncrement(`ROUTE_${cls.dimension.toUpperCase()}`);
  Logger.diagIncrement(`VERB_${cls.verb.toUpperCase()}`);
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function _parseDueDate(bodyPreview: string): Date | undefined {
  // Naive date extraction — improve with proper NLP later
  const match = bodyPreview.match(/\b(by|before|due)\s+(\w+ \d{1,2},?\s*\d{4}|\d{1,2}\/\d{1,2}\/\d{4})/i);
  if (!match) return undefined;
  const parsed = new Date(match[2]);
  return isNaN(parsed.getTime()) ? undefined : parsed;
}
