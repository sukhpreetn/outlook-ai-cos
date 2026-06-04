// =============================================================================
// src/aicos-core/StateEngine.ts
//
// STATE MANAGEMENT — tracks processed messages across Azure Function runs.
// Equivalent to AICOS StateEngine.js but uses Azure Table Storage.
//
// WHY: Azure Functions are stateless. We use Table Storage as a durable
// key-value store to track which messages have been processed, preventing
// duplicate classification/routing across overlapping timer runs.
//
// TABLE SCHEMA (table: "AicosState"):
//   PartitionKey = entity type (e.g. "message", "rule", "run")
//   RowKey       = entity ID
//   state        = JSON blob with classification result, timestamps, etc.
// =============================================================================

import { TableClient, TableEntity, odata } from '@azure/data-tables';
import { getConfig } from './Config.js';
import { Logger } from './Logger.js';

const TABLE_NAME = 'AicosState';
const PROCESSED_TTL_DAYS = 30; // clean up processed records after 30 days

// State enum mirrors AICOS OS label progression
export type MessageState =
  | 'new'
  | 'classified'
  | 'actioned'
  | 'review'
  | 'dropped';

export interface MessageStateRecord {
  messageId:    string;
  state:        MessageState;
  classifiedAt?: string;
  actionedAt?:  string;
  classification?: string; // JSON blob of cls object
}

// ── Singleton Table client ────────────────────────────────────────────────────
let _table: TableClient | null = null;

async function getTable(): Promise<TableClient> {
  if (_table) return _table;
  const cfg = await getConfig();
  _table = TableClient.fromConnectionString(cfg.storageConnectionString, TABLE_NAME);
  await _table.createTable(); // no-op if exists
  return _table;
}

// ── Core state operations ─────────────────────────────────────────────────────

/**
 * Check if a message has already been processed.
 * Equivalent to canProcessItem() in AICOS StateEngine.js
 */
export async function canProcessMessage(messageId: string): Promise<boolean> {
  try {
    const table = await getTable();
    const entity = await table.getEntity<TableEntity & { state: MessageState }>(
      'message',
      _sanitizeKey(messageId),
    );
    // Re-process only if stuck in 'review' for > 1 hour
    if (entity.state === 'review') {
      const actionedAt = entity.timestamp ? new Date(entity.timestamp) : null;
      if (actionedAt && Date.now() - actionedAt.getTime() > 60 * 60 * 1000) {
        return true; // retry
      }
    }
    return false; // already processed
  } catch {
    return true; // entity not found → process it
  }
}

/**
 * Mark a message as processed with a given state.
 * Equivalent to markItemProcessed() + transitionState() in AICOS.
 */
export async function setMessageState(
  messageId: string,
  state: MessageState,
  classification?: object,
): Promise<void> {
  try {
    const table = await getTable();
    const now = new Date().toISOString();

    await table.upsertEntity<TableEntity>({
      partitionKey: 'message',
      rowKey:       _sanitizeKey(messageId),
      state,
      updatedAt:    now,
      ...(state === 'classified' ? { classifiedAt: now } : {}),
      ...(state === 'actioned'   ? { actionedAt:   now } : {}),
      ...(classification ? { classification: JSON.stringify(classification) } : {}),
    }, 'Replace');
  } catch (e) {
    Logger.error('[StateEngine] setMessageState failed', { messageId, state, error: (e as Error).message });
  }
}

/**
 * Retrieve classification result for a message.
 */
export async function getMessageClassification(messageId: string): Promise<object | null> {
  try {
    const table = await getTable();
    const entity = await table.getEntity<TableEntity & { classification?: string }>(
      'message',
      _sanitizeKey(messageId),
    );
    return entity.classification ? JSON.parse(entity.classification) : null;
  } catch {
    return null;
  }
}

// ── Rule storage (for LearningEngine) ────────────────────────────────────────

export interface ClassificationRule {
  id:           string;
  name:         string;
  condition:    string; // e.g. "from contains '@amazon.com'"
  dimension:    string;
  verb:         string;
  confidence:   number;
  hits:         number;
  corrections:  number;
  accuracy:     number;
  lastHit:      string;
  createdAt:    string;
}

export async function saveRule(rule: ClassificationRule): Promise<void> {
  const table = await getTable();
  await table.upsertEntity<TableEntity>({
    partitionKey: 'rule',
    rowKey:       rule.id,
    ...rule,
  }, 'Replace');
}

export async function loadRules(): Promise<ClassificationRule[]> {
  const table = await getTable();
  const rules: ClassificationRule[] = [];

  const entities = table.listEntities<TableEntity & ClassificationRule>({
    queryOptions: { filter: odata`PartitionKey eq 'rule'` },
  });

  for await (const entity of entities) {
    rules.push({
      id:          entity.rowKey as string,
      name:        entity.name,
      condition:   entity.condition,
      dimension:   entity.dimension,
      verb:        entity.verb,
      confidence:  Number(entity.confidence),
      hits:        Number(entity.hits),
      corrections: Number(entity.corrections),
      accuracy:    Number(entity.accuracy),
      lastHit:     entity.lastHit,
      createdAt:   entity.createdAt,
    });
  }

  return rules.sort((a, b) => b.confidence - a.confidence);
}

// ── Cleanup ───────────────────────────────────────────────────────────────────

/**
 * Delete processed message records older than PROCESSED_TTL_DAYS.
 * Run nightly from DailyDiagnostic.
 */
export async function purgeOldStateRecords(): Promise<number> {
  const table = await getTable();
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - PROCESSED_TTL_DAYS);

  let deleted = 0;
  const entities = table.listEntities<TableEntity>({
    queryOptions: {
      filter: odata`PartitionKey eq 'message' and Timestamp lt ${cutoff}`,
    },
  });

  for await (const entity of entities) {
    await table.deleteEntity('message', entity.rowKey as string);
    deleted++;
  }

  Logger.info(`[StateEngine] Purged ${deleted} old state records`);
  return deleted;
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function _sanitizeKey(id: string): string {
  // Azure Table Storage row keys cannot contain /, \, #, ?
  return id.replace(/[/\\#?]/g, '_');
}
