"use strict";
// =============================================================================
// src/email-pipeline/AssetTracker.ts
// Records commerce / asset emails to Excel. Mirrors AssetTrackerV2.js.
// =============================================================================
Object.defineProperty(exports, "__esModule", { value: true });
exports.recordAssetEmail = recordAssetEmail;
const GraphClient_js_1 = require("../aicos-core/GraphClient.js");
const Logger_js_1 = require("../aicos-core/Logger.js");
// Columns: Date | Sender | Subject | SubType | TrackingNo | Notes
const ASSET_SUB_TYPES = ['order', 'delivery', 'warranty', 'activation', 'subscription', 'renewal', 'refund', 'recall'];
async function recordAssetEmail(msg, cls, cfg) {
    if (!cfg.oneDriveFolderId)
        return;
    const subType = _detectSubType(msg.subject, msg.bodyPreview);
    const trackingNo = _extractTracking(msg.bodyPreview);
    const row = [
        new Date(msg.receivedDateTime).toISOString().split('T')[0],
        cls.sender,
        msg.subject,
        subType,
        trackingNo ?? '',
        cls.summary,
    ];
    const workbookId = process.env.ASSET_WORKBOOK_ID ?? cfg.oneDriveFolderId;
    await (0, GraphClient_js_1.appendExcelRow)(cfg.userEmail, workbookId, 'Assets', row);
    Logger_js_1.Logger.info(`[AssetTracker] Logged: ${subType} — ${msg.subject}`);
}
function _detectSubType(subject, body) {
    const text = `${subject} ${body}`.toLowerCase();
    for (const t of ASSET_SUB_TYPES) {
        if (text.includes(t))
            return t;
    }
    return 'commerce_misc';
}
function _extractTracking(body) {
    // Common tracking number patterns
    const patterns = [
        /tracking\s*(number|#|no\.?)?\s*:?\s*([A-Z0-9]{10,30})/i,
        /\b(1Z[A-Z0-9]{16})\b/, // UPS
        /\b(\d{22})\b/, // USPS
        /\b([A-Z]{2}\d{9}[A-Z]{2})\b/, // USPS international
    ];
    for (const p of patterns) {
        const m = body.match(p);
        if (m)
            return m[m.length - 1];
    }
    return null;
}
//# sourceMappingURL=AssetTracker.js.map