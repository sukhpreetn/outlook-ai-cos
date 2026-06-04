import { Client } from '@microsoft/microsoft-graph-client';
import { getConfig } from './Config.js';
import { Logger } from './Logger.js';

export interface GraphMessage {
  id:            string;
  subject:       string;
  bodyPreview:   string;
  body:          { content: string; contentType: string };
  from:          { emailAddress: { name: string; address: string } };
  receivedDateTime: string;
  isRead:        boolean;
  categories:    string[];
  flag:          { flagStatus: string };
  conversationId: string;
  internetMessageId: string;
  hasAttachments: boolean;
}

export interface GraphEvent {
  id:            string;
  subject:       string;
  start:         { dateTime: string; timeZone: string };
  end:           { dateTime: string; timeZone: string };
  attendees:     Array<{ emailAddress: { name: string; address: string } }>;
  bodyPreview:   string;
  isOnlineMeeting: boolean;
  onlineMeetingUrl: string | null;
}

let _accessToken: string | null = null;
let _tokenExpiry = 0;

async function getAccessToken(): Promise<string> {
  if (_accessToken && Date.now() < _tokenExpiry - 60000) return _accessToken;

  const cfg = await getConfig();
  const refreshToken = process.env.REFRESH_TOKEN ?? '';

  if (!refreshToken) throw new Error('[GraphClient] REFRESH_TOKEN is not set in local.settings.json');

  const body = new URLSearchParams({
    client_id:     cfg.clientId,
    refresh_token: refreshToken,
    grant_type:    'refresh_token',
    scope:         'offline_access Mail.Read Mail.ReadWrite Mail.Send Calendars.Read Tasks.ReadWrite Files.ReadWrite.All',
  });

  const resp = await fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
    method:  'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body:    body.toString(),
  });

  const data = await resp.json() as {
    access_token?:  string;
    expires_in?:    number;
    refresh_token?: string;
    error?:         string;
    error_description?: string;
  };

  if (!data.access_token) {
    throw new Error(`[GraphClient] Token refresh failed: ${data.error} — ${data.error_description}`);
  }

  _accessToken = data.access_token;
  _tokenExpiry = Date.now() + (data.expires_in ?? 3600) * 1000;

  if (data.refresh_token && data.refresh_token !== refreshToken) {
    Logger.warn('[GraphClient] New refresh token issued — update REFRESH_TOKEN in local.settings.json:');
    Logger.warn(data.refresh_token);
  }

  return _accessToken;
}

async function getClient(): Promise<Client> {
  const token = await getAccessToken();
  return Client.init({
    authProvider: (done) => done(null, token),
  });
}

export async function fetchNewMessages(
  _userEmail: string,
  limit = 25,
): Promise<GraphMessage[]> {
  const client = await getClient();
  try {
    const response = await client
      .api('/me/messages')
      .filter('isRead eq false')
      .select('id,subject,bodyPreview,body,from,receivedDateTime,isRead,categories,flag,conversationId,internetMessageId,hasAttachments')
      .top(limit)
      .orderby('receivedDateTime desc')
      .get();
    return (response.value ?? []) as GraphMessage[];
  } catch (e) {
    Logger.error('[GraphClient] fetchNewMessages failed', e);
    return [];
  }
}

export async function autoLabelNewMessages(_userEmail: string): Promise<number> {
  Logger.info('[GraphClient] autoLabelNewMessages: skipped for personal account');
  return 0;
}

export async function categorizeMessage(
  _userEmail: string,
  messageId: string,
  addCategories: string[],
  _removeCategories: string[] = [],
): Promise<void> {
  const client = await getClient();
  try {
    const msg = await client.api(`/me/messages/${messageId}`).select('categories').get() as { categories: string[] };
    const current = new Set(msg.categories ?? []);
    addCategories.forEach(c => current.add(c));
    await client.api(`/me/messages/${messageId}`).patch({ categories: Array.from(current) });
  } catch (e) {
    Logger.error('[GraphClient] categorizeMessage failed', e);
  }
}

export async function flagMessage(_userEmail: string, messageId: string): Promise<void> {
  const client = await getClient();
  await client.api(`/me/messages/${messageId}`).patch({ flag: { flagStatus: 'flagged' } });
}

export async function createDraftReply(
  _userEmail: string,
  messageId: string,
  body: string,
): Promise<string> {
  const client = await getClient();
  const draft = await client.api(`/me/messages/${messageId}/createReply`).post({});
  await client.api(`/me/messages/${draft.id}`).patch({
    body: { contentType: 'html', content: body },
  });
  return draft.id as string;
}

export async function fetchCalendarEvents(
  _userEmail: string,
  startDate: Date,
  endDate: Date,
): Promise<GraphEvent[]> {
  const client = await getClient();
  const response = await client
    .api('/me/calendarView')
    .query({ startDateTime: startDate.toISOString(), endDateTime: endDate.toISOString() })
    .select('id,subject,start,end,attendees,bodyPreview,isOnlineMeeting,onlineMeetingUrl')
    .orderby('start/dateTime')
    .get();
  return (response.value ?? []) as GraphEvent[];
}

export async function createTask(
  _userEmail: string,
  title: string,
  body: string,
  dueDate?: Date,
  importance: 'low' | 'normal' | 'high' = 'normal',
): Promise<string> {
  const client = await getClient();
  const lists = await client.api('/me/todo/lists').get();
  const existing = (lists.value as Array<{ id: string; displayName: string }>)
    .find(l => l.displayName === 'AICOS');
  let listId: string;
  if (existing) {
    listId = existing.id;
  } else {
    const created = await client.api('/me/todo/lists').post({ displayName: 'AICOS' });
    listId = created.id as string;
  }
  const task = await client.api(`/me/todo/lists/${listId}/tasks`).post({
    title,
    body:       { contentType: 'text', content: body },
    importance,
    ...(dueDate ? {
      dueDateTime: {
        dateTime: dueDate.toISOString().split('T')[0] + 'T00:00:00',
        timeZone: 'UTC',
      },
    } : {}),
  });
  return task.id as string;
}

export async function appendExcelRow(
  _userEmail: string,
  workbookId: string,
  sheetName: string,
  values: (string | number | boolean)[],
): Promise<void> {
  const client = await getClient();
  await client
    .api(`/me/drive/items/${workbookId}/workbook/worksheets/${sheetName}/tables/Table1/rows/add`)
    .post({ values: [values] });
}

export async function getExcelRows(
  _userEmail: string,
  workbookId: string,
  sheetName: string,
): Promise<(string | number | boolean)[][]> {
  const client = await getClient();
  const response = await client
    .api(`/me/drive/items/${workbookId}/workbook/worksheets/${sheetName}/tables/Table1/rows`)
    .get();
  return (response.value as Array<{ values: (string | number | boolean)[][] }>).map(r => r.values[0]);
}

export async function createMailSubscription(
  _userEmail: string,
  notificationUrl: string,
): Promise<string> {
  const client = await getClient();
  const expiry = new Date();
  expiry.setDate(expiry.getDate() + 3);
  const sub = await client.api('/subscriptions').post({
    changeType:         'created',
    notificationUrl,
    resource:           '/me/mailFolders/inbox/messages',
    expirationDateTime: expiry.toISOString(),
    clientState:        process.env.WEBHOOK_SECRET ?? 'aicos-secret',
  });
  return sub.id as string;
}
