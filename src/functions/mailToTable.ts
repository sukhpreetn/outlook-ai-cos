import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { fetchNewMessages, GraphMessage } from '../aicos-core/GraphClient.js';
import { getConfig } from '../aicos-core/Config.js';
import { Logger } from '../aicos-core/Logger.js';
import { TableServiceClient, TableClient } from '@azure/data-tables';

const DEFAULT_TABLE_NAME = 'Emails';

app.http('mailToTable', {
  methods: ['POST','GET'],
  authLevel: 'function',
  route: 'ops/mailToTable',
  handler: async (req: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> => {
    context.log('[mailToTable] start');
    try {
      const cfg = await getConfig();
      const userEmail = cfg.userEmail;
      if (!userEmail) return { status: 500, body: 'No userEmail configured' };

      const tableName = (req.query.get('table') ?? DEFAULT_TABLE_NAME) as string;
      const limit = Number(req.query.get('limit') ?? cfg.batchSize ?? 25);

      const messages: GraphMessage[] = await fetchNewMessages(userEmail, limit);
      if ((messages?.length ?? 0) === 0) return { status: 200, body: 'No new messages' };

      const conn = cfg.storageConnectionString;
      if (!conn) return { status: 500, body: 'No storage connection string' };

      const serviceClient = TableServiceClient.fromConnectionString(conn);
      try {
        await serviceClient.createTable(tableName);
        context.log(`[mailToTable] Created table ${tableName}`);
      } catch (e: any) {
        if (e.statusCode !== 409) { Logger.error('[mailToTable] createTable failed', e); throw e; }
      }

      const tableClient: TableClient = TableClient.fromConnectionString(conn, tableName);
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
        } as { partitionKey: string; rowKey: string; [key: string]: unknown };

        try {
          await tableClient.upsertEntity(entity, 'Merge');
          written++;
        } catch (e) {
          Logger.error('[mailToTable] upsertEntity failed ' + msg.id, e);
        }
      }

      context.log(`[mailToTable] Wrote ${written}/${messages.length} entities`);
      return { status: 200, body: `Wrote ${written} entities to ${tableName}` };
    } catch (e) {
      Logger.error('[mailToTable] failed', e);
      return { status: 500, body: 'Internal error' };
    }
  }
});
