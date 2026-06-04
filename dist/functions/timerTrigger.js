"use strict";
// =============================================================================
// src/functions/timerTrigger.ts
//
// AZURE FUNCTION — Timer Trigger (every 15 minutes)
// Runs ClassifierRunner on schedule.
// Also triggers MorningBriefing at configured hour.
// =============================================================================
Object.defineProperty(exports, "__esModule", { value: true });
const functions_1 = require("@azure/functions");
const ClassifierRunner_js_1 = require("../email-pipeline/ClassifierRunner.js");
const MorningBriefing_js_1 = require("../email-pipeline/MorningBriefing.js");
const Config_js_1 = require("../aicos-core/Config.js");
const Logger_js_1 = require("../aicos-core/Logger.js");
// Register the 15-minute classifier timer
functions_1.app.timer('classifierTimer', {
    schedule: '0 */15 * * * *', // every 15 minutes
    handler: async (timer, context) => {
        context.log('[timerTrigger] Classifier timer fired');
        try {
            await (0, ClassifierRunner_js_1.runClassifierAndRouter)();
        }
        catch (e) {
            Logger_js_1.Logger.error('[timerTrigger] Classifier run failed', e);
            throw e; // re-throw so Azure marks the invocation as failed
        }
    },
});
// Register the daily briefing timer (runs hourly, checks if it's briefing hour)
functions_1.app.timer('briefingTimer', {
    schedule: '0 0 * * * *', // every hour on the hour
    handler: async (timer, context) => {
        try {
            const cfg = await (0, Config_js_1.getConfig)();
            const now = new Date();
            const hourInTz = new Intl.DateTimeFormat('en-US', {
                hour: 'numeric',
                hour12: false,
                timeZone: cfg.briefingTz,
            }).format(now);
            if (parseInt(hourInTz, 10) === cfg.briefingHour) {
                context.log('[timerTrigger] Briefing time! Sending morning briefing...');
                await (0, MorningBriefing_js_1.sendMorningBriefing)();
            }
        }
        catch (e) {
            Logger_js_1.Logger.error('[timerTrigger] Briefing failed', e);
        }
    },
});
//# sourceMappingURL=timerTrigger.js.map