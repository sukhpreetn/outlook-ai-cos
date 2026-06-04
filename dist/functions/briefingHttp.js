"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const functions_1 = require("@azure/functions");
const MorningBriefing_js_1 = require("../email-pipeline/MorningBriefing.js");
const Logger_js_1 = require("../aicos-core/Logger.js");
functions_1.app.http('briefingHttp', {
    methods: ['POST'],
    authLevel: 'function',
    route: 'ops/briefing',
    handler: async (req, context) => {
        context.log('[briefingHttp] Manual briefing triggered');
        try {
            await (0, MorningBriefing_js_1.sendMorningBriefing)();
            return { status: 200, body: 'Briefing sent! Check your Outlook inbox.' };
        }
        catch (e) {
            Logger_js_1.Logger.error('[briefingHttp] failed', e);
            return { status: 500, body: `Error: ${e.message}` };
        }
    },
});
//# sourceMappingURL=briefingHttp.js.map