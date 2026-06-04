export type MessageState = 'new' | 'classified' | 'actioned' | 'review' | 'dropped';
export interface MessageStateRecord {
    messageId: string;
    state: MessageState;
    classifiedAt?: string;
    actionedAt?: string;
    classification?: string;
}
/**
 * Check if a message has already been processed.
 * Equivalent to canProcessItem() in AICOS StateEngine.js
 */
export declare function canProcessMessage(messageId: string): Promise<boolean>;
/**
 * Mark a message as processed with a given state.
 * Equivalent to markItemProcessed() + transitionState() in AICOS.
 */
export declare function setMessageState(messageId: string, state: MessageState, classification?: object): Promise<void>;
/**
 * Retrieve classification result for a message.
 */
export declare function getMessageClassification(messageId: string): Promise<object | null>;
export interface ClassificationRule {
    id: string;
    name: string;
    condition: string;
    dimension: string;
    verb: string;
    confidence: number;
    hits: number;
    corrections: number;
    accuracy: number;
    lastHit: string;
    createdAt: string;
}
export declare function saveRule(rule: ClassificationRule): Promise<void>;
export declare function loadRules(): Promise<ClassificationRule[]>;
/**
 * Delete processed message records older than PROCESSED_TTL_DAYS.
 * Run nightly from DailyDiagnostic.
 */
export declare function purgeOldStateRecords(): Promise<number>;
//# sourceMappingURL=StateEngine.d.ts.map