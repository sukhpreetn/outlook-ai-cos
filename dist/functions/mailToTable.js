"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const functions_1 = require("@azure/functions");
const GraphClient_js_1 = require("../aicos-core/GraphClient.js");
const Config_js_1 = require("../aicos-core/Config.js");
const Logger_js_1 = require("../aicos-core/Logger.js");
const data_tables_1 = require("@azure/data-tables");
const DEFAULT_TABLE_NAME = 'Emails';
functions_1.app.http('mailToTable', {
    methods: ['POST', 'GET'],
    authLevel: 'function',
    route: 'ops/mailToTable',
    handler: async (req, context) => {
        context.log('[mailToTable] start');
        try {
            const cfg = await (0, Config_js_1.getConfig)();
            const userEmail = cfg.userEmail;
            if (!userEmail)
                return { status: 500, body: 'No userEmail configured' };
            const tableName = (req.query.get('table') ?? DEFAULT_TABLE_NAME);
            const limit = Number(req.query.get('limit') ?? cfg.batchSize ?? 25);
            const messages = await (0, GraphClient_js_1.fetchNewMessages)(userEmail, limit);
            if ((messages?.length ?? 0) === 0)
                return { status: 200, body: 'No new messages' };
            const conn = cfg.storageConnectionString;
            if (!conn)
                return { status: 500, body: 'No storage connection string' };
            const serviceClient = data_tables_1.TableServiceClient.fromConnectionString(conn);
            try {
                await serviceClient.createTable(tableName);
                context.log(`[mailToTable] Created table ${tableName}`);
            }
            catch (e) {
                if (e.statusCode !== 409) {
                    Logger_js_1.Logger.error('[mailToTable] createTable failed', e);
                    throw e;
                }
            }
            const tableClient = data_tables_1.TableClient.fromConnectionString(conn, tableName);
            let written = 0;
            for (const msg of messages) {
                const entity = {
                    partitionKey: userEmail,
                    rowKey: msg.id,
                    subject: msg.subject ?? '',
                    bodyPreview: msg.bodyPreview ?? '',
                    fromName: msg.from?.emailAddress?.name ?? '',
                    fromAddress: msg.from?.emailAddress?.address ?? '',
                    receivedDateTime: msg.receivedDateTime ?? '',
                    isRead: String(msg.isRead ?? false),
                    categories: (msg.categories ?? []).join(','),
                    hasAttachments: String(msg.hasAttachments ?? false),
                    conversationId: msg.conversationId ?? '',
                    internetMessageId: msg.internetMessageId ?? ''
                };
                try {
                    await tableClient.upsertEntity(entity, 'Merge');
                    written++;
                }
                catch (e) {
                    Logger_js_1.Logger.error('[mailToTable] upsertEntity failed ' + msg.id, e);
                }
            }
            context.log(`[mailToTable] Wrote ${written}/${messages.length} entities`);
            return { status: 200, body: `Wrote ${written} entities to ${tableName}` };
        }
        catch (e) {
            Logger_js_1.Logger.error('[mailToTable] failed', e);
            return { status: 500, body: 'Internal error' };
        }
    }
});
//# sourceMappingURL=mailToTable.js.map