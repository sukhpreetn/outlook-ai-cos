// =============================================================================
// src/email-pipeline/Classifier.ts
//
// EMAIL CLASSIFICATION ENGINE
// Mirrors Classifier_V2.js + SensingEngine.js from original AICOS.
//
// TIER MODEL (same as original):
//   Tier 0: Hard drop — known spam/promo patterns, no AI call
//   Tier 1: Fast path — known senders/domains, high-confidence rule match
//   Tier 2: Full AI — unknown senders, ambiguous content → Claude
//
// CLASSIFICATION OUTPUT (cls object):
//   {
//     messageId, subject, sender, dimension, verb, urgency,
//     confidence, tier, isReceipt, proactiveAction,
//     silenceReason, batchKey, hasDeadline, summary
//   }
//
// DIMENSIONS: career | finance | legal | knowledge | ventures | life | commerce | infra
// VERBS:      do | pay | meet | file | read
// =============================================================================

import { callAI } from '../aicos-core/AI.js';
import { loadRules, ClassificationRule } from '../aicos-core/StateEngine.js';
import { Logger } from '../aicos-core/Logger.js';
import type { GraphMessage } from '../aicos-core/GraphClient.js';

// ── Classification result type ────────────────────────────────────────────────
export interface Classification {
  messageId:       string;
  subject:         string;
  sender:          string;
  senderEmail:     string;
  dimension:       Dimension;
  verb:            Verb;
  urgency:         number;         // 1–5
  confidence:      number;         // 0.0–1.0
  tier:            0 | 1 | 2;
  isReceipt:       boolean;
  silenceReason:   string | null;
  batchKey:        string | null;  // for batching similar low-urgency emails
  hasDeadline:     boolean;
  proactiveAction: string | null;  // e.g. "draft_reply", "create_task"
  summary:         string;
}

type Dimension = 'career' | 'finance' | 'legal' | 'knowledge' | 'ventures' | 'life' | 'commerce' | 'infra';
type Verb      = 'do' | 'pay' | 'meet' | 'file' | 'read';

// ── Tier 0: Hard-drop patterns (no AI call) ───────────────────────────────────
const TIER0_DROP_PATTERNS = [
  /unsubscribe/i,
  /no-?reply@/i,
  /\bnewsletter\b/i,
  /\bpromotion\b/i,
  /you'?re? (invited|selected)/i,
  /\bmarketing\b/i,
  /donotreply/i,
  /noreply/i,
];

// ── Tier 1: Known domain → dimension mapping ──────────────────────────────────
const KNOWN_FINANCE_DOMAINS = new Set([
  'chase.com', 'bankofamerica.com', 'wellsfargo.com', 'citibank.com',
  'amex.com', 'paypal.com', 'venmo.com', 'stripe.com', 'square.com',
  'intuit.com', 'turbotax.com', 'fidelity.com', 'vanguard.com',
]);

const KNOWN_CAREER_DOMAINS = new Set([
  'linkedin.com', 'lever.co', 'greenhouse.io', 'workday.com',
  'icims.com', 'bamboohr.com', 'indeed.com', 'glassdoor.com',
  'smartrecruiters.com', 'jobvite.com',
]);

const KNOWN_COMMERCE_DOMAINS = new Set([
  'amazon.com', 'ebay.com', 'shopify.com', 'etsy.com', 'walmart.com',
  'target.com', 'bestbuy.com', 'apple.com', 'fedex.com', 'ups.com',
  'usps.com', 'shipbob.com',
]);

const KNOWN_LEGAL_DOMAINS = new Set([
  'docusign.com', 'hellosign.com', 'pandadoc.com', 'adobe.com',
  'irs.gov', 'state.gov', 'treasury.gov',
]);

// ── Public API ────────────────────────────────────────────────────────────────

/**
 * Classify a single email message.
 * Returns null if the message should be silently dropped (Tier 0).
 */
export async function classifyMessage(msg: GraphMessage): Promise<Classification | null> {
  const sender      = msg.from.emailAddress.name;
  const senderEmail = msg.from.emailAddress.address.toLowerCase();
  const domain      = senderEmail.split('@')[1] ?? '';
  const subject     = msg.subject ?? '';
  const body        = msg.bodyPreview ?? '';

  // ── Tier 0: Hard drop ─────────────────────────────────────────────────────
  if (_isTier0Drop(senderEmail, subject, body)) {
    Logger.info(`[Classifier] TIER0 DROP: ${subject.substring(0, 60)}`);
    Logger.diagIncrement('TIER0_DROPS');
    return null;
  }

  // ── Tier 1: Rule-based fast path ──────────────────────────────────────────
  const rules = await loadRules();
  const ruleMatch = _evalRules(rules, senderEmail, domain, subject, body);

  if (ruleMatch && ruleMatch.confidence >= 0.85) {
    Logger.info(`[Classifier] TIER1 RULE: ${ruleMatch.dimension}/${ruleMatch.verb} conf=${ruleMatch.confidence}`);
    Logger.diagIncrement('TIER1_HITS');
    return _buildFromRule(msg, sender, senderEmail, ruleMatch, 1);
  }

  // Tier 1: Known-domain fast path
  const domainResult = _evalKnownDomain(domain, subject);
  if (domainResult) {
    Logger.info(`[Classifier] TIER1 DOMAIN: ${domainResult.dimension}/${domainResult.verb}`);
    Logger.diagIncrement('TIER1_DOMAIN_HITS');
    return {
      messageId:       msg.id,
      subject,
      sender,
      senderEmail,
      tier:            1,
      confidence:      0.90,
      urgency:         2,
      isReceipt:       domainResult.isReceipt ?? false,
      silenceReason:   null,
      batchKey:        domainResult.batchKey ?? null,
      hasDeadline:     false,
      proactiveAction: null,
      summary:         `${domainResult.dimension} email from ${sender}`,
      ...domainResult,
      dimension: (domainResult.dimension ?? 'knowledge') as Dimension,
      verb:      (domainResult.verb ?? 'read') as Verb,
    };
  }

  // ── Tier 2: Full AI classification ────────────────────────────────────────
  Logger.info(`[Classifier] TIER2 AI: ${subject.substring(0, 60)}`);
  Logger.diagIncrement('TIER2_AI_CALLS');
  return await _classifyWithAI(msg, sender, senderEmail, subject, body);
}

// ── Tier 0 checks ─────────────────────────────────────────────────────────────
function _isTier0Drop(senderEmail: string, subject: string, body: string): boolean {
  for (const pattern of TIER0_DROP_PATTERNS) {
    if (pattern.test(senderEmail) || pattern.test(subject)) return true;
  }
  // Check Outlook's List-Unsubscribe equivalent (bulk mail)
  if (body.toLowerCase().includes('unsubscribe') && body.toLowerCase().includes('manage preferences')) {
    return true;
  }
  return false;
}

// ── Rule evaluation ────────────────────────────────────────────────────────────
function _evalRules(
  rules: ClassificationRule[],
  senderEmail: string,
  domain: string,
  subject: string,
  body: string,
): ClassificationRule | null {
  for (const rule of rules) {
    try {
      if (_ruleMatches(rule.condition, senderEmail, domain, subject, body)) {
        return rule;
      }
    } catch {
      // Malformed rule — skip
    }
  }
  return null;
}

function _ruleMatches(
  condition: string,
  senderEmail: string,
  domain: string,
  subject: string,
  body: string,
): boolean {
  // Simple condition DSL: "from contains 'x'" | "domain eq 'x'" | "subject contains 'x'"
  const fromMatch    = condition.match(/^from contains '(.+)'$/i);
  const domainMatch  = condition.match(/^domain eq '(.+)'$/i);
  const subjectMatch = condition.match(/^subject contains '(.+)'$/i);

  if (fromMatch)    return senderEmail.includes(fromMatch[1].toLowerCase());
  if (domainMatch)  return domain === domainMatch[1].toLowerCase();
  if (subjectMatch) return subject.toLowerCase().includes(subjectMatch[1].toLowerCase());

  return false;
}

function _buildFromRule(
  msg: GraphMessage,
  sender: string,
  senderEmail: string,
  rule: ClassificationRule,
  tier: 0 | 1 | 2,
): Classification {
  return {
    messageId:       msg.id,
    subject:         msg.subject,
    sender,
    senderEmail,
    dimension:       rule.dimension as Dimension,
    verb:            rule.verb as Verb,
    urgency:         2,
    confidence:      rule.confidence,
    tier,
    isReceipt:       false,
    silenceReason:   null,
    batchKey:        null,
    hasDeadline:     false,
    proactiveAction: null,
    summary:         `Rule match: ${rule.name}`,
  };
}

// ── Known domain fast path ────────────────────────────────────────────────────
function _evalKnownDomain(
  domain: string,
  subject: string,
): Partial<Classification> | null {
  if (KNOWN_FINANCE_DOMAINS.has(domain)) {
    const isReceipt = /receipt|order|invoice|payment|statement/i.test(subject);
    return {
      dimension: 'finance',
      verb:      isReceipt ? 'file' : 'read',
      isReceipt,
      batchKey:  isReceipt ? null : 'finance-update',
    };
  }
  if (KNOWN_CAREER_DOMAINS.has(domain)) {
    return { dimension: 'career', verb: 'do', isReceipt: false, batchKey: null };
  }
  if (KNOWN_COMMERCE_DOMAINS.has(domain)) {
    const isOrder = /order|shipped|delivery|track/i.test(subject);
    return {
      dimension: 'commerce',
      verb:      'read',
      isReceipt: isOrder,
      batchKey:  isOrder ? null : 'commerce-update',
    };
  }
  if (KNOWN_LEGAL_DOMAINS.has(domain)) {
    return { dimension: 'legal', verb: 'file', isReceipt: false, batchKey: null };
  }
  return null;
}

// ── Tier 2: AI classification ─────────────────────────────────────────────────
async function _classifyWithAI(
  msg: GraphMessage,
  sender: string,
  senderEmail: string,
  subject: string,
  body: string,
): Promise<Classification> {
  const prompt = `Classify this email and respond ONLY with a JSON object (no markdown, no preamble):

FROM: ${sender} <${senderEmail}>
SUBJECT: ${subject}
BODY PREVIEW: ${body.substring(0, 600)}

Required JSON fields:
- dimension: one of [career, finance, legal, knowledge, ventures, life, commerce, infra]
- verb: one of [do, pay, meet, file, read]
- urgency: integer 1-5 (5 = requires response today)
- confidence: float 0.0-1.0 (how sure you are)
- isReceipt: boolean (is this a financial receipt/invoice?)
- hasDeadline: boolean (does the email mention a specific deadline?)
- silenceReason: string or null (why this can be silently filed without attention)
- batchKey: string or null (group similar low-urgency emails, e.g. "newsletters")
- proactiveAction: one of [draft_reply, create_task, create_event, null]
- summary: string (one sentence, action-oriented, what the user needs to know)`;

  try {
    const result = await callAI('classify', prompt, { maxTokens: 512 });
    const cls = JSON.parse(result.text.replace(/```json\n?|```/g, "").trim()) as Partial<Classification>;

    return {
      messageId:       msg.id,
      subject,
      sender,
      senderEmail,
      dimension:       cls.dimension       ?? 'knowledge',
      verb:            cls.verb            ?? 'read',
      urgency:         cls.urgency         ?? 2,
      confidence:      cls.confidence      ?? 0.7,
      tier:            2,
      isReceipt:       cls.isReceipt       ?? false,
      silenceReason:   cls.silenceReason   ?? null,
      batchKey:        cls.batchKey        ?? null,
      hasDeadline:     cls.hasDeadline     ?? false,
      proactiveAction: cls.proactiveAction ?? null,
      summary:         cls.summary         ?? `Email from ${sender}`,
    };
  } catch (e) {
    Logger.error('[Classifier] AI parse failed', e);
    // Fallback: conservative defaults, send to review
    return {
      messageId:    msg.id,
      subject,
      sender,
      senderEmail,
      dimension:    'knowledge',
      verb:         'read',
      urgency:      2,
      confidence:   0.4,
      tier:         2,
      isReceipt:    false,
      silenceReason: null,
      batchKey:     null,
      hasDeadline:  false,
      proactiveAction: null,
      summary:      `Could not classify email from ${sender}`,
    };
  }
}
