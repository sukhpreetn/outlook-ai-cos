// =============================================================================
// src/functions/timerTrigger.ts
//
// AZURE FUNCTION — Timer Trigger (every 15 minutes)
// Runs ClassifierRunner on schedule.
// Also triggers MorningBriefing at configured hour.
// =============================================================================

import { app, InvocationContext, Timer } from '@azure/functions';
import { runClassifierAndRouter } from '../email-pipeline/ClassifierRunner.js';
import { sendMorningBriefing } from '../email-pipeline/MorningBriefing.js';
import { getConfig } from '../aicos-core/Config.js';
import { Logger } from '../aicos-core/Logger.js';

// Register the 15-minute classifier timer
app.timer('classifierTimer', {
  schedule: '0 */15 * * * *', // every 15 minutes
  handler: async (timer: Timer, context: InvocationContext) => {
    context.log('[timerTrigger] Classifier timer fired');
    try {
      await runClassifierAndRouter();
    } catch (e) {
      Logger.error('[timerTrigger] Classifier run failed', e);
      throw e; // re-throw so Azure marks the invocation as failed
    }
  },
});

// Register the daily briefing timer (runs hourly, checks if it's briefing hour)
app.timer('briefingTimer', {
  schedule: '0 0 * * * *', // every hour on the hour
  handler: async (timer: Timer, context: InvocationContext) => {
    try {
      const cfg = await getConfig();
      const now = new Date();
      const hourInTz = new Intl.DateTimeFormat('en-US', {
        hour: 'numeric',
        hour12: false,
        timeZone: cfg.briefingTz,
      }).format(now);

      if (parseInt(hourInTz, 10) === cfg.briefingHour) {
        context.log('[timerTrigger] Briefing time! Sending morning briefing...');
        await sendMorningBriefing();
      }
    } catch (e) {
      Logger.error('[timerTrigger] Briefing failed', e);
    }
  },
});
