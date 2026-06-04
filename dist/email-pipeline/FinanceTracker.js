"use strict";
// =============================================================================
// src/email-pipeline/FinanceTracker.ts
//
// Records finance emails (receipts, bills, payments) to an Excel workbook
// on OneDrive. Equivalent to FinanceTrackerV2.js in original AICOS.
// =============================================================================
Object.defineProperty(exports, "__esModule", { value: true });
exports.recordFinanceEmail = recordFinanceEmail;
const GraphClient_js_1 = require("../aicos-core/GraphClient.js");
const Logger_js_1 = require("../aicos-core/Logger.js");
// Excel table columns (must match the workbook template)
// A: Date | B: Sender | C: Subject | D: Type | E: Amount | F: Notes
async function recordFinanceEmail(msg, cls, cfg) {
    if (!cfg.oneDriveFolderId) {
        Logger_js_1.Logger.warn('[FinanceTracker] oneDriveFolderId not configured — skipping');
        return;
    }
    const amount = _extractAmount(msg.bodyPreview);
    const row = [
        new Date(msg.receivedDateTime).toISOString().split('T')[0],
        cls.sender,
        msg.subject,
        cls.isReceipt ? 'receipt' : 'bill',
        amount ?? '',
        cls.summary,
    ];
    // TODO: replace with your actual Finance workbook ID from OneDrive
    const workbookId = process.env.FINANCE_WORKBOOK_ID ?? cfg.oneDriveFolderId;
    await (0, GraphClient_js_1.appendExcelRow)(cfg.userEmail, workbookId, 'Transactions', row);
    Logger_js_1.Logger.info(`[FinanceTracker] Logged: ${msg.subject}`);
}
function _extractAmount(body) {
    const match = body.match(/\$\s*([\d,]+\.?\d{0,2})/);
    if (!match)
        return null;
    return parseFloat(match[1].replace(',', ''));
}
//# sourceMappingURL=FinanceTracker.js.map