import type { GraphMessage } from '../aicos-core/GraphClient.js';
export interface Classification {
    messageId: string;
    subject: string;
    sender: string;
    senderEmail: string;
    dimension: Dimension;
    verb: Verb;
    urgency: number;
    confidence: number;
    tier: 0 | 1 | 2;
    isReceipt: boolean;
    silenceReason: string | null;
    batchKey: string | null;
    hasDeadline: boolean;
    proactiveAction: string | null;
    summary: string;
}
type Dimension = 'career' | 'finance' | 'legal' | 'knowledge' | 'ventures' | 'life' | 'commerce' | 'infra';
type Verb = 'do' | 'pay' | 'meet' | 'file' | 'read';
/**
 * Classify a single email message.
 * Returns null if the message should be silently dropped (Tier 0).
 */
export declare function classifyMessage(msg: GraphMessage): Promise<Classification | null>;
export {};
//# sourceMappingURL=Classifier.d.ts.map