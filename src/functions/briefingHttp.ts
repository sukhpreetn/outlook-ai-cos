import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { sendMorningBriefing } from '../email-pipeline/MorningBriefing.js';
import { Logger } from '../aicos-core/Logger.js';

app.http('briefingHttp', {
  methods: ['POST'],
  authLevel: 'function',
  route: 'ops/briefing',
  handler: async (req: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> => {
    context.log('[briefingHttp] Manual briefing triggered');
    try {
      await sendMorningBriefing();
      return { status: 200, body: 'Briefing sent! Check your Outlook inbox.' };
    } catch (e) {
      Logger.error('[briefingHttp] failed', e);
      return { status: 500, body: `Error: ${(e as Error).message}` };
    }
  },
});
