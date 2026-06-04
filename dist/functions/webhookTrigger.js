"use strict";
// =============================================================================
// src/functions/webhookTrigger.ts
//
// AZURE FUNCTION — HTTP Trigger (Graph Webhook)
// Receives Microsoft Graph change notifications for new Outlook messages.
// Faster than the timer trigger (near-real-time for urgent emails).
//
// Graph requires webhook validation on subscription creation.
// =============================================================================
Object.defineProperty(exports, "__esModule", { value: true });
const functions_1 = require("@azure/functions");
const ClassifierRunner_js_1 = require("../email-pipeline/ClassifierRunner.js");
const Logger_js_1 = require("../aicos-core/Logger.js");
functions_1.app.http('graphWebhook', {
    methods: ['GET', 'POST'],
    authLevel: 'function',
    route: 'webhook/graph',
    handler: async (req, context) => {
        // ── Validation handshake (Graph subscription creation) ─────────────────
        const validationToken = req.query.get('validationToken');
        if (validationToken) {
            context.log('[webhook] Validation handshake received');
            return {
                status: 200,
                headers: { 'Content-Type': 'text/plain' },
                body: validationToken,
            };
        }
        // ── Change notification ────────────────────────────────────────────────
        let body;
        try {
            body = await req.json();
        }
        catch {
            return { status: 400, body: 'Invalid JSON' };
        }
        // Validate client state (CSRF protection)
        const expectedState = process.env.WEBHOOK_SECRET ?? 'aicos-secret';
        const invalidNotif = body.value?.find(n => n.clientState !== expectedState);
        if (invalidNotif) {
            Logger_js_1.Logger.warn('[webhook] Invalid clientState — possible spoofed notification');
            return { status: 401, body: 'Invalid clientState' };
        }
        // Acknowledge immediately (Graph requires < 3s response)
        // Process asynchronously
        context.log(`[webhook] Received ${body.value?.length ?? 0} notification(s)`);
        // Fire-and-forget: run classifier on the new message
        setImmediate(async () => {
            try {
                await (0, ClassifierRunner_js_1.runClassifierAndRouter)();
            }
            catch (e) {
                Logger_js_1.Logger.error('[webhook] ClassifierRunner failed after notification', e);
            }
        });
        return { status: 202 };
    },
});
//# sourceMappingURL=webhookTrigger.js.map