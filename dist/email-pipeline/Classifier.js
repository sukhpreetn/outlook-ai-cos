"use strict";
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
Object.defineProperty(exports, "__esModule", { value: true });
exports.classifyMessage = classifyMessage;
const AI_js_1 = require("../aicos-core/AI.js");
const StateEngine_js_1 = require("../aicos-core/StateEngine.js");
const Logger_js_1 = require("../aicos-core/Logger.js");
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
async function classifyMessage(msg) {
    const sender = msg.from.emailAddress.name;
    const senderEmail = msg.from.emailAddress.address.toLowerCase();
    const domain = senderEmail.split('@')[1] ?? '';
    const subject = msg.subject ?? '';
    const body = msg.bodyPreview ?? '';
    // ── Tier 0: Hard drop ─────────────────────────────────────────────────────
    if (_isTier0Drop(senderEmail, subject, body)) {
        Logger_js_1.Logger.info(`[Classifier] TIER0 DROP: ${subject.substring(0, 60)}`);
        Logger_js_1.Logger.diagIncrement('TIER0_DROPS');
        return null;
    }
    // ── Tier 1: Rule-based fast path ──────────────────────────────────────────
    const rules = await (0, StateEngine_js_1.loadRules)();
    const ruleMatch = _evalRules(rules, senderEmail, domain, subject, body);
    if (ruleMatch && ruleMatch.confidence >= 0.85) {
        Logger_js_1.Logger.info(`[Classifier] TIER1 RULE: ${ruleMatch.dimension}/${ruleMatch.verb} conf=${ruleMatch.confidence}`);
        Logger_js_1.Logger.diagIncrement('TIER1_HITS');
        return _buildFromRule(msg, sender, senderEmail, ruleMatch, 1);
    }
    // Tier 1: Known-domain fast path
    const domainResult = _evalKnownDomain(domain, subject);
    if (domainResult) {
        Logger_js_1.Logger.info(`[Classifier] TIER1 DOMAIN: ${domainResult.dimension}/${domainResult.verb}`);
        Logger_js_1.Logger.diagIncrement('TIER1_DOMAIN_HITS');
        return {
            messageId: msg.id,
            subject,
            sender,
            senderEmail,
            tier: 1,
            confidence: 0.90,
            urgency: 2,
            isReceipt: domainResult.isReceipt ?? false,
            silenceReason: null,
            batchKey: domainResult.batchKey ?? null,
            hasDeadline: false,
            proactiveAction: null,
            summary: `${domainResult.dimension} email from ${sender}`,
            ...domainResult,
            dimension: (domainResult.dimension ?? 'knowledge'),
            verb: (domainResult.verb ?? 'read'),
        };
    }
    // ── Tier 2: Full AI classification ────────────────────────────────────────
    Logger_js_1.Logger.info(`[Classifier] TIER2 AI: ${subject.substring(0, 60)}`);
    Logger_js_1.Logger.diagIncrement('TIER2_AI_CALLS');
    return await _classifyWithAI(msg, sender, senderEmail, subject, body);
}
// ── Tier 0 checks ─────────────────────────────────────────────────────────────
function _isTier0Drop(senderEmail, subject, body) {
    for (const pattern of TIER0_DROP_PATTERNS) {
        if (pattern.test(senderEmail) || pattern.test(subject))
            return true;
    }
    // Check Outlook's List-Unsubscribe equivalent (bulk mail)
    if (body.toLowerCase().includes('unsubscribe') && body.toLowerCase().includes('manage preferences')) {
        return true;
    }
    return false;
}
// ── Rule evaluation ────────────────────────────────────────────────────────────
function _evalRules(rules, senderEmail, domain, subject, body) {
    for (const rule of rules) {
        try {
            if (_ruleMatches(rule.condition, senderEmail, domain, subject, body)) {
                return rule;
            }
        }
        catch {
            // Malformed rule — skip
        }
    }
    return null;
}
function _ruleMatches(condition, senderEmail, domain, subject, body) {
    // Simple condition DSL: "from contains 'x'" | "domain eq 'x'" | "subject contains 'x'"
    const fromMatch = condition.match(/^from contains '(.+)'$/i);
    const domainMatch = condition.match(/^domain eq '(.+)'$/i);
    const subjectMatch = condition.match(/^subject contains '(.+)'$/i);
    if (fromMatch)
        return senderEmail.includes(fromMatch[1].toLowerCase());
    if (domainMatch)
        return domain === domainMatch[1].toLowerCase();
    if (subjectMatch)
        return subject.toLowerCase().includes(subjectMatch[1].toLowerCase());
    return false;
}
function _buildFromRule(msg, sender, senderEmail, rule, tier) {
    return {
        messageId: msg.id,
        subject: msg.subject,
        sender,
        senderEmail,
        dimension: rule.dimension,
        verb: rule.verb,
        urgency: 2,
        confidence: rule.confidence,
        tier,
        isReceipt: false,
        silenceReason: null,
        batchKey: null,
        hasDeadline: false,
        proactiveAction: null,
        summary: `Rule match: ${rule.name}`,
    };
}
// ── Known domain fast path ────────────────────────────────────────────────────
function _evalKnownDomain(domain, subject) {
    if (KNOWN_FINANCE_DOMAINS.has(domain)) {
        const isReceipt = /receipt|order|invoice|payment|statement/i.test(subject);
        return {
            dimension: 'finance',
            verb: isReceipt ? 'file' : 'read',
            isReceipt,
            batchKey: isReceipt ? null : 'finance-update',
        };
    }
    if (KNOWN_CAREER_DOMAINS.has(domain)) {
        return { dimension: 'career', verb: 'do', isReceipt: false, batchKey: null };
    }
    if (KNOWN_COMMERCE_DOMAINS.has(domain)) {
        const isOrder = /order|shipped|delivery|track/i.test(subject);
        return {
            dimension: 'commerce',
            verb: 'read',
            isReceipt: isOrder,
            batchKey: isOrder ? null : 'commerce-update',
        };
    }
    if (KNOWN_LEGAL_DOMAINS.has(domain)) {
        return { dimension: 'legal', verb: 'file', isReceipt: false, batchKey: null };
    }
    return null;
}
// ── Tier 2: AI classification ─────────────────────────────────────────────────
async function _classifyWithAI(msg, sender, senderEmail, subject, body) {
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
        const result = await (0, AI_js_1.callAI)('classify', prompt, { maxTokens: 512 });
        const cls = JSON.parse(result.text.replace(/```json\n?|```/g, "").trim());
        return {
            messageId: msg.id,
            subject,
            sender,
            senderEmail,
            dimension: cls.dimension ?? 'knowledge',
            verb: cls.verb ?? 'read',
            urgency: cls.urgency ?? 2,
            confidence: cls.confidence ?? 0.7,
            tier: 2,
            isReceipt: cls.isReceipt ?? false,
            silenceReason: cls.silenceReason ?? null,
            batchKey: cls.batchKey ?? null,
            hasDeadline: cls.hasDeadline ?? false,
            proactiveAction: cls.proactiveAction ?? null,
            summary: cls.summary ?? `Email from ${sender}`,
        };
    }
    catch (e) {
        Logger_js_1.Logger.error('[Classifier] AI parse failed', e);
        // Fallback: conservative defaults, send to review
        return {
            messageId: msg.id,
            subject,
            sender,
            senderEmail,
            dimension: 'knowledge',
            verb: 'read',
            urgency: 2,
            confidence: 0.4,
            tier: 2,
            isReceipt: false,
            silenceReason: null,
            batchKey: null,
            hasDeadline: false,
            proactiveAction: null,
            summary: `Could not classify email from ${sender}`,
        };
    }
}
//# sourceMappingURL=Classifier.js.map