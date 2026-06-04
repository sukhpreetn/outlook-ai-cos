// =============================================================================
// src/functions/webhookTrigger.ts
//
// AZURE FUNCTION — HTTP Trigger (Graph Webhook)
// Receives Microsoft Graph change notifications for new Outlook messages.
// Faster than the timer trigger (near-real-time for urgent emails).
//
// Graph requires webhook validation on subscription creation.
// =============================================================================

import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { runClassifierAndRouter } from '../email-pipeline/ClassifierRunner.js';
import { Logger } from '../aicos-core/Logger.js';

interface GraphNotification {
  value: Array<{
    changeType:       string;
    clientState:      string;
    resource:         string;
    resourceData?:    { id?: string };
    subscriptionId:   string;
  }>;
}

app.http('graphWebhook', {
  methods:    ['GET', 'POST'],
  authLevel:  'function',
  route:      'webhook/graph',
  handler: async (req: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> => {

    // ── Validation handshake (Graph subscription creation) ─────────────────
    const validationToken = req.query.get('validationToken');
    if (validationToken) {
      context.log('[webhook] Validation handshake received');
      return {
        status:  200,
        headers: { 'Content-Type': 'text/plain' },
        body:    validationToken,
      };
    }

    // ── Change notification ────────────────────────────────────────────────
    let body: GraphNotification;
    try {
      body = await req.json() as GraphNotification;
    } catch {
      return { status: 400, body: 'Invalid JSON' };
    }

    // Validate client state (CSRF protection)
    const expectedState = process.env.WEBHOOK_SECRET ?? 'aicos-secret';
    const invalidNotif = body.value?.find(n => n.clientState !== expectedState);
    if (invalidNotif) {
      Logger.warn('[webhook] Invalid clientState — possible spoofed notification');
      return { status: 401, body: 'Invalid clientState' };
    }

    // Acknowledge immediately (Graph requires < 3s response)
    // Process asynchronously
    context.log(`[webhook] Received ${body.value?.length ?? 0} notification(s)`);

    // Fire-and-forget: run classifier on the new message
    setImmediate(async () => {
      try {
        await runClassifierAndRouter();
      } catch (e) {
        Logger.error('[webhook] ClassifierRunner failed after notification', e);
      }
    });

    return { status: 202 };
  },
});
