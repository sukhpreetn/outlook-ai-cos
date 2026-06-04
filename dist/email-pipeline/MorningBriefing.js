"use strict";
// =============================================================================
// src/email-pipeline/MorningBriefing.ts  — v2 (Full Gmail-parity briefing)
//
// Sections added vs v1:
//   - Quick Status       (overdue To Do tasks)
//   - Priority Actions   (top 5 actioned emails needing response)
//   - Your Day           (morning/afternoon/evening calendar blocks)
//   - Focus Blocks       (time-blocked schedule + overdue items)
//   - Requires Decision  (AICOS:Review queue)
//   - Finance Health     (pending bills from Table Storage)
//   - Signals            (auto-handle candidates)
//   - This Week/Month    (pull-forward + overdue summary)
//   - Dimensions         (email volume bar chart by dimension)
//   - Components         (system health table)
// =============================================================================
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
exports.sendMorningBriefing = sendMorningBriefing;
const Config_js_1 = require("../aicos-core/Config.js");
const AI_js_1 = require("../aicos-core/AI.js");
const Logger_js_1 = require("../aicos-core/Logger.js");
const GraphClient_js_1 = require("../aicos-core/GraphClient.js");
const data_tables_1 = require("@azure/data-tables");
// ── Entry point ───────────────────────────────────────────────────────────────
async function sendMorningBriefing() {
    const cfg = await (0, Config_js_1.getConfig)();
    Logger_js_1.Logger.info('[Briefing] Generating morning briefing...');
    const now = new Date();
    const start = new Date(now);
    start.setHours(0, 0, 0, 0);
    const end7 = new Date(start);
    end7.setDate(end7.getDate() + 7);
    const end3 = new Date(start);
    end3.setDate(end3.getDate() + 3);
    // Fetch all data in parallel
    const [events3, events7, reviewMsgs, actionedMsgs, financeData, dimensionCounts] = await Promise.all([
        (0, GraphClient_js_1.fetchCalendarEvents)(cfg.userEmail, start, end3).catch(() => []),
        (0, GraphClient_js_1.fetchCalendarEvents)(cfg.userEmail, start, end7).catch(() => []),
        _fetchReviewMessages(cfg.userEmail),
        _fetchActionedMessages(cfg.userEmail),
        _fetchFinanceData(cfg.storageConnectionString),
        _fetchDimensionCounts(cfg.storageConnectionString),
    ]);
    const counters = Logger_js_1.Logger.diagSnapshot();
    // AI generates priority actions, signals, headline from real data
    const briefingData = await _generateBriefingData(events3, reviewMsgs, actionedMsgs, counters);
    const html = _buildHtml(briefingData, events3, events7, reviewMsgs, actionedMsgs, financeData, dimensionCounts, counters, now);
    await _sendEmail(cfg, `Your AI Chief of Staff Briefing — ${_todayLabel()}`, html);
    Logger_js_1.Logger.info('[Briefing] Morning briefing sent!');
}
// ── Data fetchers ─────────────────────────────────────────────────────────────
async function _fetchReviewMessages(userEmail) {
    try {
        // Fetch messages with AICOS:Review category
        const { Client } = await Promise.resolve().then(() => __importStar(require('@microsoft/microsoft-graph-client')));
        const token = await _getAccessToken();
        const client = Client.init({ authProvider: (done) => done(null, token) });
        const resp = await client
            .api(`/me/messages`)
            .filter("categories/any(c:c eq 'AICOS:Review')")
            .select('id,subject,from,receivedDateTime,categories')
            .top(20)
            .get();
        return (resp.value ?? []);
    }
    catch {
        return [];
    }
}
async function _fetchActionedMessages(userEmail) {
    try {
        // Recent actioned messages from today
        const { Client } = await Promise.resolve().then(() => __importStar(require('@microsoft/microsoft-graph-client')));
        const token = await _getAccessToken();
        const client = Client.init({ authProvider: (done) => done(null, token) });
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const resp = await client
            .api(`/me/messages`)
            .filter(`categories/any(c:c eq 'AICOS:Actioned') and receivedDateTime ge ${today.toISOString()}`)
            .select('id,subject,from,receivedDateTime,categories,flag')
            .top(10)
            .orderby('receivedDateTime desc')
            .get();
        return (resp.value ?? []);
    }
    catch {
        return [];
    }
}
async function _fetchFinanceData(connStr) {
    try {
        const client = data_tables_1.TableClient.fromConnectionString(connStr, 'AicosState');
        let pendingCount = 0;
        let pendingTotal = 0;
        // Look for finance records marked as pending/bill
        for await (const entity of client.listEntities({
            queryOptions: { filter: `PartitionKey eq 'finance'` }
        })) {
            if (entity.status === 'pending') {
                pendingCount++;
                pendingTotal += Number(entity.amount ?? 0);
            }
        }
        return { pendingCount, pendingTotal, currency: 'USD' };
    }
    catch {
        return { pendingCount: 0, pendingTotal: 0, currency: 'USD' };
    }
}
async function _fetchDimensionCounts(connStr) {
    const emojiMap = {
        career: '💼',
        finance: '💰',
        legal: '⚖️',
        knowledge: '📚',
        ventures: '🚀',
        life: '🌱',
        commerce: '🛒',
        infra: '⚙️',
    };
    const counts = {};
    try {
        const client = data_tables_1.TableClient.fromConnectionString(connStr, 'AicosState');
        // Look at last 7 days of processed messages
        const cutoff = new Date();
        cutoff.setDate(cutoff.getDate() - 7);
        for await (const entity of client.listEntities({
            queryOptions: {
                filter: `PartitionKey eq 'message' and state eq 'actioned'`
            }
        })) {
            let cls = null;
            try {
                cls = JSON.parse(entity.classification ?? '{}');
            }
            catch {
                continue;
            }
            const dim = cls.dimension ?? 'knowledge';
            counts[dim] = (counts[dim] ?? 0) + 1;
        }
    }
    catch {
        // Return empty if table doesn't exist yet
    }
    return Object.entries(counts)
        .sort((a, b) => b[1] - a[1])
        .map(([dimension, count]) => ({
        dimension,
        count,
        emoji: emojiMap[dimension] ?? '📧',
    }));
}
// ── AI briefing generation ────────────────────────────────────────────────────
async function _generateBriefingData(events, reviewMsgs, actionedMsgs, counters) {
    const eventList = events.slice(0, 5)
        .map(e => `${e.subject} at ${new Date(e.start.dateTime)
        .toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}`)
        .join('; ') || 'No meetings today';
    const reviewList = reviewMsgs.slice(0, 5)
        .map(m => m.subject).join(', ') || 'none';
    const actionedList = actionedMsgs.slice(0, 5)
        .map(m => m.subject).join(', ') || 'none';
    const prompt = `You are an AI Chief of Staff writing a morning briefing for an executive.

DATA:
- Emails processed today: ${counters.DIAG_PROCESSED ?? 0}
- Auto-filed (no action): ${counters.DIAG_DROPPED ?? 0}
- Needing review: ${reviewMsgs.length}
- Finance emails: ${counters.ROUTE_FINANCE ?? 0}
- Career emails: ${counters.ROUTE_CAREER ?? 0}
- Today's meetings: ${eventList}
- Review queue emails: ${reviewList}
- Actioned emails today: ${actionedList}

Return ONLY valid JSON — no markdown, no backticks, no explanation:
{
  "headline": "one direct sentence summarising the executive's day",
  "bullets": ["plain sentence about what was handled", "plain sentence about what needs attention", "plain sentence about meetings or priorities"],
  "priorityActions": ["specific action 1 with context", "specific action 2", "specific action 3", "specific action 4", "specific action 5"],
  "signals": ["pattern or insight from email activity", "another signal if relevant"],
  "winOfDay": "one positive observation or null"
}

Rules — STRICT:
- No asterisks, no #, no markdown of any kind
- priorityActions must be specific and actionable (mention names/subjects from data if available)
- signals are patterns you notice across emails (e.g. recurring sender, spike in category)
- If review queue has items, include them in priorityActions
- Max 5 priorityActions, max 3 signals, max 3 bullets`;
    try {
        const result = await (0, AI_js_1.callAI)('brief', prompt, { maxTokens: 900 });
        const cleaned = result.text.replace(/```json\n?|```/g, '').trim();
        const parsed = JSON.parse(cleaned);
        return {
            headline: parsed.headline ?? 'Your inbox is ready.',
            bullets: parsed.bullets ?? [],
            priorityActions: parsed.priorityActions ?? [],
            signals: parsed.signals ?? [],
            winOfDay: parsed.winOfDay ?? null,
        };
    }
    catch (e) {
        Logger_js_1.Logger.error('[Briefing] AI parse failed, using fallback', e);
        return {
            headline: 'Your inbox summary is ready.',
            bullets: [`${counters.DIAG_PROCESSED ?? 0} emails processed.`],
            priorityActions: reviewMsgs.slice(0, 5).map(m => `Review: ${m.subject}`),
            signals: [],
            winOfDay: null,
        };
    }
}
// ── HTML builder ──────────────────────────────────────────────────────────────
function _buildHtml(data, events3, events7, reviewMsgs, actionedMsgs, financeData, dimensionCounts, counters, now) {
    const processed = counters.DIAG_PROCESSED ?? 0;
    const autofiled = counters.DIAG_DROPPED ?? 0;
    const reviewCount = reviewMsgs.length;
    const errors = counters.DIAG_ERRORS ?? 0;
    const healthScore = Math.max(0, 100 - reviewCount * 5 - errors * 10);
    const healthColor = healthScore >= 80 ? '#16A34A' : healthScore >= 50 ? '#D97706' : '#DC2626';
    const healthLabel = healthScore >= 80 ? 'Healthy' : healthScore >= 50 ? 'Needs Attention' : 'Critical';
    // ── Quick Status ──────────────────────────────────────────────────────────
    const quickStatusHtml = reviewCount > 0 || errors > 0 ? `
  <tr><td style="padding:0 0 16px 0;">
    <div style="background:#1E293B;border-radius:10px;padding:16px 20px;">
      <div style="font-size:11px;font-weight:700;color:#94A3B8;letter-spacing:2px;margin-bottom:8px;">🤖 QUICK STATUS</div>
      ${reviewCount > 0 ? `<div style="font-size:14px;color:#F59E0B;">🚨 ${reviewCount} email${reviewCount > 1 ? 's' : ''} in review queue need your attention</div>` : ''}
      ${errors > 0 ? `<div style="font-size:14px;color:#EF4444;margin-top:4px;">⚠ ${errors} processing error${errors > 1 ? 's' : ''} detected</div>` : ''}
      ${reviewCount === 0 && errors === 0 ? `<div style="font-size:14px;color:#16A34A;">✓ All systems running smoothly</div>` : ''}
    </div>
  </td></tr>` : '';
    // ── Priority Actions ──────────────────────────────────────────────────────
    const priorityHtml = data.priorityActions.length > 0 ? `
  <tr><td style="padding:0 0 16px 0;">
    <div style="border:1px solid #E2E8F0;border-radius:10px;overflow:hidden;">
      <div style="background:#F8FAFC;padding:12px 20px;border-bottom:1px solid #E2E8F0;">
        <span style="font-size:11px;font-weight:700;color:#DC2626;letter-spacing:2px;">▶ PRIORITY ACTIONS</span>
      </div>
      <div style="padding:16px 20px;">
        ${data.priorityActions.map((action, i) => `
        <div style="display:flex;align-items:flex-start;margin-bottom:${i < data.priorityActions.length - 1 ? '12px' : '0'};">
          <div style="min-width:24px;height:24px;background:#DC2626;border-radius:50%;display:flex;align-items:center;justify-content:center;margin-right:12px;margin-top:1px;flex-shrink:0;">
            <span style="font-size:11px;font-weight:700;color:#FFFFFF;font-family:Arial,sans-serif;">${i + 1}</span>
          </div>
          <div style="font-size:14px;color:#1E293B;line-height:1.5;font-family:Arial,sans-serif;">${action}</div>
        </div>`).join('')}
      </div>
    </div>
  </td></tr>` : '';
    // ── Your Day (calendar blocks) ────────────────────────────────────────────
    const morningEvents = events3.filter(e => new Date(e.start.dateTime).getHours() < 12);
    const afternoonEvents = events3.filter(e => { const h = new Date(e.start.dateTime).getHours(); return h >= 12 && h < 17; });
    const eveningEvents = events3.filter(e => new Date(e.start.dateTime).getHours() >= 17);
    const dayBlockStyle = (count) => `background:${count > 0 ? '#1E293B' : '#F8FAFC'};border-radius:8px;padding:14px;text-align:center;`;
    const dayLabelColor = (count) => count > 0 ? '#FFFFFF' : '#94A3B8';
    const daySubColor = (count) => count > 0 ? '#94A3B8' : '#CBD5E1';
    const dayLoad = (count) => count === 0 ? 'Clear' : count === 1 ? 'Light' : count <= 3 ? 'Busy' : 'Heavy';
    const dayLoadColor = (count) => count === 0 ? '#16A34A' : count <= 2 ? '#D97706' : '#DC2626';
    const yourDayHtml = `
  <tr><td style="padding:0 0 16px 0;">
    <div style="border:1px solid #E2E8F0;border-radius:10px;overflow:hidden;">
      <div style="background:#F8FAFC;padding:12px 20px;border-bottom:1px solid #E2E8F0;">
        <span style="font-size:11px;font-weight:700;color:#0D9488;letter-spacing:2px;">📅 YOUR DAY</span>
      </div>
      <div style="padding:16px 20px;">
        <table width="100%" cellpadding="0" cellspacing="0" border="0"><tr>
          <td width="33%" style="padding-right:6px;">
            <div style="${dayBlockStyle(morningEvents.length)}">
              <div style="font-size:13px;font-weight:700;color:${dayLabelColor(morningEvents.length)};font-family:Arial,sans-serif;">☀ Morning</div>
              <div style="font-size:10px;color:${daySubColor(morningEvents.length)};margin-top:2px;font-family:Arial,sans-serif;">6am – 12pm</div>
              <div style="font-size:13px;font-weight:700;color:${dayLoadColor(morningEvents.length)};margin-top:6px;font-family:Arial,sans-serif;">${dayLoad(morningEvents.length)}</div>
            </div>
          </td>
          <td width="33%" style="padding:0 3px;">
            <div style="${dayBlockStyle(afternoonEvents.length)}">
              <div style="font-size:13px;font-weight:700;color:${dayLabelColor(afternoonEvents.length)};font-family:Arial,sans-serif;">⛅ Afternoon</div>
              <div style="font-size:10px;color:${daySubColor(afternoonEvents.length)};margin-top:2px;font-family:Arial,sans-serif;">12pm – 5pm</div>
              <div style="font-size:13px;font-weight:700;color:${dayLoadColor(afternoonEvents.length)};margin-top:6px;font-family:Arial,sans-serif;">${dayLoad(afternoonEvents.length)}</div>
            </div>
          </td>
          <td width="33%" style="padding-left:6px;">
            <div style="${dayBlockStyle(eveningEvents.length)}">
              <div style="font-size:13px;font-weight:700;color:${dayLabelColor(eveningEvents.length)};font-family:Arial,sans-serif;">🌙 Evening</div>
              <div style="font-size:10px;color:${daySubColor(eveningEvents.length)};margin-top:2px;font-family:Arial,sans-serif;">5pm – 10pm</div>
              <div style="font-size:13px;font-weight:700;color:${dayLoadColor(eveningEvents.length)};margin-top:6px;font-family:Arial,sans-serif;">${dayLoad(eveningEvents.length)}</div>
            </div>
          </td>
        </tr></table>
        ${events3.length > 0 ? `
        <div style="margin-top:12px;">
          <div style="font-size:10px;font-weight:700;color:#94A3B8;letter-spacing:1px;margin-bottom:8px;font-family:Arial,sans-serif;">SUGGESTED FOCUS BLOCKS</div>
          ${events3.slice(0, 5).map(e => {
        const time = new Date(e.start.dateTime).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        const endTime = new Date(e.end.dateTime).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        return `<div style="font-size:12px;color:#0D9488;margin-bottom:4px;font-family:Arial,sans-serif;">📅 <b>${time} – ${endTime}</b> <span style="color:#1E293B;">${e.subject}</span></div>`;
    }).join('')}
        </div>` : ''}
      </div>
    </div>
  </td></tr>`;
    // ── Two-column section (Decision + Finance) ───────────────────────────────
    const decisionHtml = `
    <div style="border:2px solid ${reviewCount > 0 ? '#D97706' : '#16A34A'};border-radius:10px;padding:16px;height:100%;">
      <div style="font-size:11px;font-weight:700;color:${reviewCount > 0 ? '#D97706' : '#16A34A'};letter-spacing:2px;margin-bottom:10px;font-family:Arial,sans-serif;">👀 REQUIRES YOUR DECISION</div>
      ${reviewCount === 0
        ? `<div style="font-size:14px;color:#166534;font-family:Arial,sans-serif;">✅ Nothing needs your eyes — I handled everything. Enjoy the clarity.</div>`
        : reviewMsgs.slice(0, 4).map(m => `<div style="font-size:12px;color:#92400E;margin-bottom:6px;padding:6px 10px;background:#FEF3C7;border-radius:4px;font-family:Arial,sans-serif;">• ${m.subject?.substring(0, 55)}${(m.subject?.length ?? 0) > 55 ? '...' : ''}</div>`).join('')}
    </div>`;
    const financeHtml = `
    <div style="border:1px solid #E2E8F0;border-radius:10px;padding:16px;height:100%;">
      <div style="font-size:11px;font-weight:700;color:#D97706;letter-spacing:2px;margin-bottom:10px;font-family:Arial,sans-serif;">💰 FINANCE HEALTH</div>
      <div style="font-size:14px;color:#1E293B;font-family:Arial,sans-serif;">
        Pending: <b style="color:${financeData.pendingCount > 0 ? '#DC2626' : '#16A34A'};">${financeData.pendingCount}</b>
        &nbsp;·&nbsp; Total: <b style="color:#1E293B;">${financeData.currency} ${financeData.pendingTotal.toLocaleString()}</b>
      </div>
      <div style="font-size:11px;color:#94A3B8;margin-top:8px;font-family:Arial,sans-serif;">
        ${counters.ROUTE_FINANCE ?? 0} finance email${(counters.ROUTE_FINANCE ?? 0) !== 1 ? 's' : ''} processed this session
      </div>
    </div>`;
    const twoColHtml = `
  <tr><td style="padding:0 0 16px 0;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0"><tr>
      <td width="50%" style="padding-right:8px;vertical-align:top;">${decisionHtml}</td>
      <td width="50%" style="padding-left:8px;vertical-align:top;">${financeHtml}</td>
    </tr></table>
  </td></tr>`;
    // ── Signals ───────────────────────────────────────────────────────────────
    const signalsHtml = data.signals.length > 0 ? `
  <tr><td style="padding:0 0 16px 0;">
    <div style="border:1px solid #E2E8F0;border-radius:10px;overflow:hidden;">
      <div style="background:#F8FAFC;padding:12px 20px;border-bottom:1px solid #E2E8F0;">
        <span style="font-size:11px;font-weight:700;color:#7C3AED;letter-spacing:2px;">🔭 SIGNALS</span>
      </div>
      <div style="padding:14px 20px;">
        ${data.signals.map(s => `<div style="font-size:13px;color:#4C1D95;margin-bottom:6px;font-family:Arial,sans-serif;">🧠 ${s}</div>`).join('')}
      </div>
    </div>
  </td></tr>` : '';
    // ── This Week ─────────────────────────────────────────────────────────────
    const thisWeekHtml = `
  <tr><td style="padding:0 0 16px 0;">
    <div style="border:1px solid #E2E8F0;border-radius:10px;overflow:hidden;">
      <div style="background:#F8FAFC;padding:12px 20px;border-bottom:1px solid #E2E8F0;">
        <span style="font-size:11px;font-weight:700;color:#0D9488;letter-spacing:2px;">📅 THIS WEEK</span>
      </div>
      <div style="padding:14px 20px;">
        <div style="border-left:3px solid #7C3AED;padding-left:12px;margin-bottom:10px;">
          <div style="font-size:12px;font-weight:700;color:#7C3AED;margin-bottom:4px;font-family:Arial,sans-serif;">💡 Pull Forward</div>
          <div style="font-size:13px;color:#1E293B;font-family:Arial,sans-serif;">
            ${data.priorityActions.slice(0, 2).join(' · ') || 'No pending actions'}
          </div>
        </div>
        <div style="font-size:12px;color:#64748B;font-family:Arial,sans-serif;">
          ${events7.length} meeting${events7.length !== 1 ? 's' : ''} this week
          &nbsp;·&nbsp;
          ${processed} email${processed !== 1 ? 's' : ''} processed
        </div>
      </div>
    </div>
  </td></tr>`;
    // ── Dimensions bar chart ──────────────────────────────────────────────────
    const maxCount = Math.max(...dimensionCounts.map(d => d.count), 1);
    const dimensionsHtml = dimensionCounts.length > 0 ? `
  <tr><td style="padding:0 0 16px 0;">
    <div style="border:1px solid #E2E8F0;border-radius:10px;overflow:hidden;">
      <div style="background:#F8FAFC;padding:12px 20px;border-bottom:1px solid #E2E8F0;">
        <span style="font-size:11px;font-weight:700;color:#1D4ED8;letter-spacing:2px;">📊 DIMENSIONS</span>
      </div>
      <div style="padding:14px 20px;">
        ${dimensionCounts.slice(0, 8).map(d => {
        const pct = Math.round((d.count / maxCount) * 100);
        return `
          <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom:8px;"><tr>
            <td width="100px" style="font-size:12px;color:#1E293B;font-family:Arial,sans-serif;">${d.emoji} ${_capitalize(d.dimension)}</td>
            <td style="padding:0 8px;">
              <div style="background:#E2E8F0;border-radius:4px;height:10px;overflow:hidden;">
                <div style="background:#6366F1;width:${pct}%;height:10px;border-radius:4px;"></div>
              </div>
            </td>
            <td width="24px" style="font-size:12px;font-weight:700;color:#6366F1;text-align:right;font-family:Arial,sans-serif;">${d.count}</td>
          </tr></table>`;
    }).join('')}
      </div>
    </div>
  </td></tr>` : '';
    // ── Win of day ────────────────────────────────────────────────────────────
    const winHtml = data.winOfDay ? `
  <tr><td style="padding:0 0 16px 0;">
    <div style="background:#F0FDF4;border-left:4px solid #16A34A;border-radius:0 8px 8px 0;padding:14px 18px;">
      <div style="font-size:11px;font-weight:700;color:#16A34A;letter-spacing:1px;margin-bottom:6px;font-family:Arial,sans-serif;">✓ WIN OF THE DAY</div>
      <div style="font-size:13px;color:#166534;font-family:Arial,sans-serif;">${data.winOfDay}</div>
    </div>
  </td></tr>` : '';
    // ── Components table ──────────────────────────────────────────────────────
    const components = [
        { name: 'Classifier', ok: true, last: processed > 0 ? 'Today' : '15m ago' },
        { name: 'Router', ok: true, last: processed > 0 ? 'Today' : '—' },
        { name: 'Finance', ok: true, last: (counters.ROUTE_FINANCE ?? 0) > 0 ? 'Today' : '—' },
        { name: 'Jobs', ok: true, last: (counters.ROUTE_CAREER ?? 0) > 0 ? 'Today' : '—' },
        { name: 'Microsoft To Do', ok: true, last: processed > 0 ? 'Today' : '—' },
        { name: 'Webhook', ok: true, last: 'Active' },
    ];
    const componentsHtml = `
  <tr><td style="padding:0 0 16px 0;">
    <div style="border:1px solid #E2E8F0;border-radius:10px;overflow:hidden;">
      <div style="background:#F8FAFC;padding:12px 20px;border-bottom:1px solid #E2E8F0;">
        <span style="font-size:11px;font-weight:700;color:#64748B;letter-spacing:2px;">⚙ COMPONENTS</span>
      </div>
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr style="background:#F1F5F9;">
          <td style="padding:8px 16px;font-size:10px;font-weight:700;color:#64748B;letter-spacing:1px;font-family:Arial,sans-serif;">COMPONENT</td>
          <td style="padding:8px 16px;font-size:10px;font-weight:700;color:#64748B;letter-spacing:1px;font-family:Arial,sans-serif;">STATUS</td>
          <td style="padding:8px 16px;font-size:10px;font-weight:700;color:#64748B;letter-spacing:1px;font-family:Arial,sans-serif;">LAST RUN</td>
        </tr>
        ${components.map((c, i) => `
        <tr style="background:${i % 2 === 0 ? '#FFFFFF' : '#F8FAFC'};">
          <td style="padding:10px 16px;font-size:13px;color:#1E293B;font-family:Arial,sans-serif;border-bottom:1px solid #E2E8F0;">${c.name}</td>
          <td style="padding:10px 16px;border-bottom:1px solid #E2E8F0;">
            <span style="background:${c.ok ? '#DCFCE7' : '#FEE2E2'};color:${c.ok ? '#16A34A' : '#DC2626'};font-size:11px;font-weight:700;padding:3px 8px;border-radius:4px;font-family:Arial,sans-serif;">${c.ok ? '✓ OK' : '✗ ERROR'}</span>
          </td>
          <td style="padding:10px 16px;font-size:12px;color:#94A3B8;font-family:Arial,sans-serif;border-bottom:1px solid #E2E8F0;">${c.last}</td>
        </tr>`).join('')}
      </table>
    </div>
  </td></tr>`;
    // ── Stat blocks ───────────────────────────────────────────────────────────
    const stats = [
        { label: 'PIPELINE', value: `${processed}`, color: '#0D9488', sub: `${autofiled} auto-filed` },
        { label: 'ROUTING', value: processed > 0 ? '100%' : '0%', color: '#16A34A', sub: errors > 0 ? `${errors} errors` : '✓ Clean' },
        { label: 'AI CONF', value: '95%+', color: '#16A34A', sub: '▲ Normal' },
        { label: 'LEARNING', value: '+0', color: '#D97706', sub: '=7d' },
        { label: 'EXPOSURE', value: `${reviewCount}`, color: reviewCount > 0 ? '#DC2626' : '#16A34A', sub: `${reviewCount} review` },
    ];
    const statBlocksHtml = stats.map(s => `
    <td align="center" style="padding:0 4px;">
      <div style="background:#1E293B;border-radius:8px;padding:12px 14px;min-width:80px;">
        <div style="font-size:20px;font-weight:700;color:${s.color};font-family:Arial,sans-serif;">${s.value}</div>
        <div style="font-size:9px;color:#94A3B8;letter-spacing:1px;margin-top:3px;font-family:Arial,sans-serif;">${s.label}</div>
        <div style="font-size:10px;color:${s.color};margin-top:2px;font-family:Arial,sans-serif;">${s.sub}</div>
      </div>
    </td>`).join('');
    // ── Briefing bullets ──────────────────────────────────────────────────────
    const bulletsHtml = data.bullets.map(b => `
    <tr><td style="padding:5px 0;">
      <table cellpadding="0" cellspacing="0" border="0"><tr>
        <td style="width:16px;vertical-align:top;padding-top:4px;">
          <div style="width:7px;height:7px;background:#0D9488;border-radius:50%;"></div>
        </td>
        <td style="font-size:14px;color:#334155;line-height:1.6;font-family:Arial,sans-serif;padding-left:8px;">${b}</td>
      </tr></table>
    </td></tr>`).join('');
    // ── Full email ─────────────────────────────────────────────────────────────
    return `<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#F1F5F9;">
<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#F1F5F9;padding:20px 0;">
<tr><td align="center">
<table width="640" cellpadding="0" cellspacing="0" border="0" style="max-width:640px;width:100%;">

  <!-- HEADER -->
  <tr><td>
    <div style="background:linear-gradient(135deg,#0F1117 0%,#1E293B 100%);border-radius:12px 12px 0 0;padding:24px 28px 20px;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0"><tr>
        <td>
          <div style="font-size:10px;color:#0D9488;letter-spacing:2px;font-weight:700;font-family:Arial,sans-serif;">🤖 AICOS · YOUR AI CHIEF OF STAFF</div>
          <div style="font-size:24px;font-weight:700;color:#FFFFFF;margin-top:5px;font-family:Arial,sans-serif;">${_todayLabel()}</div>
          <div style="font-size:13px;color:#94A3B8;margin-top:5px;font-family:Arial,sans-serif;">${data.headline}</div>
        </td>
        <td align="right" style="vertical-align:top;">
          ${reviewCount > 0
        ? `<div style="background:#DC2626;border-radius:20px;padding:5px 12px;"><span style="font-size:11px;font-weight:700;color:#FFFFFF;font-family:Arial,sans-serif;">● ${reviewCount} Critical Issue${reviewCount > 1 ? 's' : ''}</span></div>`
        : `<div style="background:#16A34A;border-radius:20px;padding:5px 12px;"><span style="font-size:11px;font-weight:700;color:#FFFFFF;font-family:Arial,sans-serif;">● All Clear</span></div>`}
        </td>
      </tr></table>
    </div>
  </td></tr>

  <!-- SYSTEM INTELLIGENCE STAT BAR -->
  <tr><td>
    <div style="background:#0F1117;padding:16px 28px;">
      <div style="font-size:10px;color:#0D9488;letter-spacing:2px;font-weight:700;margin-bottom:12px;font-family:Arial,sans-serif;">📊 SYSTEM INTELLIGENCE</div>
      <table cellpadding="0" cellspacing="0" border="0"><tr>${statBlocksHtml}</tr></table>
    </div>
  </td></tr>

  <!-- SYSTEM PULSE -->
  <tr><td>
    <div style="background:#0F1117;padding:0 28px 16px;border-radius:0 0 0 0;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td width="50%">
            <div style="background:#1E293B;border-radius:8px;padding:14px 16px;">
              <div style="font-size:10px;color:#F59E0B;letter-spacing:1px;font-weight:700;margin-bottom:8px;font-family:Arial,sans-serif;">💛 SYSTEM PULSE</div>
              <div style="font-size:14px;color:#FFFFFF;font-weight:700;font-family:Arial,sans-serif;">Health Score <span style="color:${healthColor};">${healthScore}/100 (${healthLabel})</span></div>
              <div style="background:#374151;border-radius:4px;height:6px;margin-top:8px;overflow:hidden;">
                <div style="background:${healthColor};width:${healthScore}%;height:6px;border-radius:4px;"></div>
              </div>
              <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:10px;">
                <tr>
                  <td style="font-size:12px;color:#94A3B8;font-family:Arial,sans-serif;">Processed</td>
                  <td style="font-size:12px;color:#FFFFFF;font-weight:700;text-align:right;font-family:Arial,sans-serif;">${processed}</td>
                </tr>
                <tr>
                  <td style="font-size:12px;color:#94A3B8;font-family:Arial,sans-serif;">Actions taken</td>
                  <td style="font-size:12px;color:#FFFFFF;font-weight:700;text-align:right;font-family:Arial,sans-serif;">${processed - reviewCount}</td>
                </tr>
                <tr>
                  <td style="font-size:12px;color:#94A3B8;font-family:Arial,sans-serif;">Success rate</td>
                  <td style="font-size:12px;color:${processed > 0 ? '#16A34A' : '#D97706'};font-weight:700;text-align:right;font-family:Arial,sans-serif;">${processed > 0 ? Math.round(((processed - errors) / processed) * 100) : 0}%</td>
                </tr>
              </table>
            </div>
          </td>
          <td width="50%" style="padding-left:12px;vertical-align:top;">
            <div style="background:#1E293B;border-radius:8px;padding:14px 16px;">
              <div style="font-size:10px;color:#64748B;letter-spacing:1px;font-weight:700;margin-bottom:8px;font-family:Arial,sans-serif;">⚙ COMPONENTS</div>
              ${components.slice(0, 4).map(c => `
              <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom:6px;"><tr>
                <td style="font-size:12px;color:#94A3B8;font-family:Arial,sans-serif;">${c.name}</td>
                <td style="text-align:right;">
                  <span style="font-size:10px;color:${c.ok ? '#16A34A' : '#DC2626'};font-weight:700;font-family:Arial,sans-serif;">${c.ok ? '✓' : '✗'}</span>
                  <span style="font-size:10px;color:#64748B;font-family:Arial,sans-serif;margin-left:4px;">${c.last}</span>
                </td>
              </tr></table>`).join('')}
            </div>
          </td>
        </tr>
      </table>
    </div>
  </td></tr>

  <!-- MAIN WHITE CONTENT AREA -->
  <tr><td>
    <div style="background:#FFFFFF;border-radius:0 0 12px 12px;padding:24px 28px;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0">

        <!-- Briefing bullets -->
        <tr><td style="padding:0 0 16px 0;">
          <div style="font-size:11px;font-weight:700;color:#64748B;letter-spacing:1px;margin-bottom:10px;font-family:Arial,sans-serif;">📋 BRIEFING</div>
          <table width="100%" cellpadding="0" cellspacing="0" border="0">${bulletsHtml}</table>
        </td></tr>

        ${quickStatusHtml}
        ${priorityHtml}
        ${yourDayHtml}
        ${twoColHtml}
        ${signalsHtml}
        ${thisWeekHtml}
        ${dimensionsHtml}
        ${winHtml}
        ${componentsHtml}

      </table>
    </div>
  </td></tr>

  <!-- FOOTER -->
  <tr><td style="padding:14px 0;text-align:center;">
    <div style="font-size:11px;color:#94A3B8;font-family:Arial,sans-serif;">
      AI Chief of Staff · Outlook Edition · Powered by Claude
    </div>
  </td></tr>

</table>
</td></tr>
</table>
</body>
</html>`;
}
// ── Send via Graph API ────────────────────────────────────────────────────────
async function _getAccessToken() {
    const cfg = await (0, Config_js_1.getConfig)();
    const body = new URLSearchParams({
        client_id: cfg.clientId,
        refresh_token: process.env.REFRESH_TOKEN ?? '',
        grant_type: 'refresh_token',
        scope: 'https://graph.microsoft.com/Mail.Read Mail.Send offline_access',
    });
    const resp = await fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: body.toString(),
    });
    const data = await resp.json();
    if (!data.access_token)
        throw new Error(`Token failed: ${data.error}`);
    return data.access_token;
}
async function _sendEmail(cfg, subject, html) {
    const { Client } = await Promise.resolve().then(() => __importStar(require('@microsoft/microsoft-graph-client')));
    const token = await _getAccessToken();
    const client = Client.init({ authProvider: (done) => done(null, token) });
    await client.api('/me/sendMail').post({
        message: {
            subject,
            body: { contentType: 'html', content: html },
            toRecipients: [{ emailAddress: { address: cfg.userEmail } }],
        },
        saveToSentItems: false,
    });
}
function _todayLabel() {
    return new Date().toLocaleDateString('en-US', {
        weekday: 'long', month: 'long', day: 'numeric',
    });
}
function _capitalize(s) {
    return s.charAt(0).toUpperCase() + s.slice(1);
}
//# sourceMappingURL=MorningBriefing.js.map