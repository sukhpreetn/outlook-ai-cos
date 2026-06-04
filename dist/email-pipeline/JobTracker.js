"use strict";
// =============================================================================
// src/email-pipeline/JobTracker.ts
// Records career emails to Excel. Mirrors JobTrackerV2.js.
// =============================================================================
Object.defineProperty(exports, "__esModule", { value: true });
exports.recordCareerEmail = recordCareerEmail;
const GraphClient_js_1 = require("../aicos-core/GraphClient.js");
const Logger_js_1 = require("../aicos-core/Logger.js");
// Columns: Date | Company | Subject | Type | Status | Notes
async function recordCareerEmail(msg, cls, cfg) {
    if (!cfg.oneDriveFolderId)
        return;
    const type = _detectCareerType(msg.subject, msg.bodyPreview);
    const company = _extractCompany(cls.sender, cls.senderEmail);
    const row = [
        new Date(msg.receivedDateTime).toISOString().split('T')[0],
        company,
        msg.subject,
        type,
        'open',
        cls.summary,
    ];
    const workbookId = process.env.CAREER_WORKBOOK_ID ?? cfg.oneDriveFolderId;
    await (0, GraphClient_js_1.appendExcelRow)(cfg.userEmail, workbookId, 'Applications', row);
    Logger_js_1.Logger.info(`[JobTracker] Logged: ${company} — ${type}`);
}
function _detectCareerType(subject, body) {
    const text = `${subject} ${body}`.toLowerCase();
    if (/interview|schedule|zoom|meet/i.test(text))
        return 'interview';
    if (/offer|offer letter|compensation/i.test(text))
        return 'offer';
    if (/application.*received|thank.*applying/i.test(text))
        return 'application_ack';
    if (/rejection|unfortunately|moved forward/i.test(text))
        return 'rejection';
    if (/recruiter|opportunity|role.*open/i.test(text))
        return 'outreach';
    return 'career_misc';
}
function _extractCompany(sender, email) {
    // Try to get company from sender name
    const parts = sender.split(/\s+at\s+|\s+-\s+|,\s*/i);
    if (parts.length > 1)
        return parts[parts.length - 1].trim();
    // Fall back to domain
    const domain = email.split('@')[1] ?? '';
    return domain.split('.')[0] ?? sender;
}
//# sourceMappingURL=JobTracker.js.map