// =============================================================================
// src/email-pipeline/JobTracker.ts
// Records career emails to Excel. Mirrors JobTrackerV2.js.
// =============================================================================

import { appendExcelRow } from '../aicos-core/GraphClient.js';
import { Logger } from '../aicos-core/Logger.js';
import type { GraphMessage } from '../aicos-core/GraphClient.js';
import type { Classification } from './Classifier.js';
import type { AICOSConfig } from '../aicos-core/Config.js';

// Columns: Date | Company | Subject | Type | Status | Notes
export async function recordCareerEmail(
  msg: GraphMessage,
  cls: Classification,
  cfg: AICOSConfig,
): Promise<void> {
  if (!cfg.oneDriveFolderId) return;

  const type = _detectCareerType(msg.subject, msg.bodyPreview);
  const company = _extractCompany(cls.sender, cls.senderEmail);

  const row: (string | number | boolean)[] = [
    new Date(msg.receivedDateTime).toISOString().split('T')[0],
    company,
    msg.subject,
    type,
    'open',
    cls.summary,
  ];

  const workbookId = process.env.CAREER_WORKBOOK_ID ?? cfg.oneDriveFolderId;
  await appendExcelRow(cfg.userEmail, workbookId, 'Applications', row);
  Logger.info(`[JobTracker] Logged: ${company} — ${type}`);
}

function _detectCareerType(subject: string, body: string): string {
  const text = `${subject} ${body}`.toLowerCase();
  if (/interview|schedule|zoom|meet/i.test(text))      return 'interview';
  if (/offer|offer letter|compensation/i.test(text))    return 'offer';
  if (/application.*received|thank.*applying/i.test(text)) return 'application_ack';
  if (/rejection|unfortunately|moved forward/i.test(text)) return 'rejection';
  if (/recruiter|opportunity|role.*open/i.test(text))   return 'outreach';
  return 'career_misc';
}

function _extractCompany(sender: string, email: string): string {
  // Try to get company from sender name
  const parts = sender.split(/\s+at\s+|\s+-\s+|,\s*/i);
  if (parts.length > 1) return parts[parts.length - 1].trim();
  // Fall back to domain
  const domain = email.split('@')[1] ?? '';
  return domain.split('.')[0] ?? sender;
}
