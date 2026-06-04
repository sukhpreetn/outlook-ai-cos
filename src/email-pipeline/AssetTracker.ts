// =============================================================================
// src/email-pipeline/AssetTracker.ts
// Records commerce / asset emails to Excel. Mirrors AssetTrackerV2.js.
// =============================================================================

import { appendExcelRow } from '../aicos-core/GraphClient.js';
import { Logger } from '../aicos-core/Logger.js';
import type { GraphMessage } from '../aicos-core/GraphClient.js';
import type { Classification } from './Classifier.js';
import type { AICOSConfig } from '../aicos-core/Config.js';

// Columns: Date | Sender | Subject | SubType | TrackingNo | Notes
const ASSET_SUB_TYPES = ['order', 'delivery', 'warranty', 'activation', 'subscription', 'renewal', 'refund', 'recall'];

export async function recordAssetEmail(
  msg: GraphMessage,
  cls: Classification,
  cfg: AICOSConfig,
): Promise<void> {
  if (!cfg.oneDriveFolderId) return;

  const subType    = _detectSubType(msg.subject, msg.bodyPreview);
  const trackingNo = _extractTracking(msg.bodyPreview);

  const row: (string | number | boolean)[] = [
    new Date(msg.receivedDateTime).toISOString().split('T')[0],
    cls.sender,
    msg.subject,
    subType,
    trackingNo ?? '',
    cls.summary,
  ];

  const workbookId = process.env.ASSET_WORKBOOK_ID ?? cfg.oneDriveFolderId;
  await appendExcelRow(cfg.userEmail, workbookId, 'Assets', row);
  Logger.info(`[AssetTracker] Logged: ${subType} — ${msg.subject}`);
}

function _detectSubType(subject: string, body: string): string {
  const text = `${subject} ${body}`.toLowerCase();
  for (const t of ASSET_SUB_TYPES) {
    if (text.includes(t)) return t;
  }
  return 'commerce_misc';
}

function _extractTracking(body: string): string | null {
  // Common tracking number patterns
  const patterns = [
    /tracking\s*(number|#|no\.?)?\s*:?\s*([A-Z0-9]{10,30})/i,
    /\b(1Z[A-Z0-9]{16})\b/,        // UPS
    /\b(\d{22})\b/,                 // USPS
    /\b([A-Z]{2}\d{9}[A-Z]{2})\b/, // USPS international
  ];
  for (const p of patterns) {
    const m = body.match(p);
    if (m) return m[m.length - 1];
  }
  return null;
}
