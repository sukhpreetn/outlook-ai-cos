// =============================================================================
// src/email-pipeline/FinanceTracker.ts
//
// Records finance emails (receipts, bills, payments) to an Excel workbook
// on OneDrive. Equivalent to FinanceTrackerV2.js in original AICOS.
// =============================================================================

import { getConfig } from '../aicos-core/Config.js';
import { appendExcelRow } from '../aicos-core/GraphClient.js';
import { Logger } from '../aicos-core/Logger.js';
import type { GraphMessage } from '../aicos-core/GraphClient.js';
import type { Classification } from './Classifier.js';
import type { AICOSConfig } from '../aicos-core/Config.js';

// Excel table columns (must match the workbook template)
// A: Date | B: Sender | C: Subject | D: Type | E: Amount | F: Notes

export async function recordFinanceEmail(
  msg: GraphMessage,
  cls: Classification,
  cfg: AICOSConfig,
): Promise<void> {
  if (!cfg.oneDriveFolderId) {
    Logger.warn('[FinanceTracker] oneDriveFolderId not configured — skipping');
    return;
  }

  const amount = _extractAmount(msg.bodyPreview);
  const row: (string | number | boolean)[] = [
    new Date(msg.receivedDateTime).toISOString().split('T')[0],
    cls.sender,
    msg.subject,
    cls.isReceipt ? 'receipt' : 'bill',
    amount ?? '',
    cls.summary,
  ];

  // TODO: replace with your actual Finance workbook ID from OneDrive
  const workbookId = process.env.FINANCE_WORKBOOK_ID ?? cfg.oneDriveFolderId;

  await appendExcelRow(cfg.userEmail, workbookId, 'Transactions', row);
  Logger.info(`[FinanceTracker] Logged: ${msg.subject}`);
}

function _extractAmount(body: string): number | null {
  const match = body.match(/\$\s*([\d,]+\.?\d{0,2})/);
  if (!match) return null;
  return parseFloat(match[1].replace(',', ''));
}
