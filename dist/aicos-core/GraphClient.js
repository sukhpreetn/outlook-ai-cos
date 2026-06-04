"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.fetchNewMessages = fetchNewMessages;
exports.autoLabelNewMessages = autoLabelNewMessages;
exports.categorizeMessage = categorizeMessage;
exports.flagMessage = flagMessage;
exports.createDraftReply = createDraftReply;
exports.fetchCalendarEvents = fetchCalendarEvents;
exports.createTask = createTask;
exports.appendExcelRow = appendExcelRow;
exports.getExcelRows = getExcelRows;
exports.createMailSubscription = createMailSubscription;
const microsoft_graph_client_1 = require("@microsoft/microsoft-graph-client");
const Config_js_1 = require("./Config.js");
const Logger_js_1 = require("./Logger.js");
let _accessToken = null;
let _tokenExpiry = 0;
async function getAccessToken() {
    if (_accessToken && Date.now() < _tokenExpiry - 60000)
        return _accessToken;
    const cfg = await (0, Config_js_1.getConfig)();
    const refreshToken = process.env.REFRESH_TOKEN ?? '';
    if (!refreshToken)
        throw new Error('[GraphClient] REFRESH_TOKEN is not set in local.settings.json');
    const body = new URLSearchParams({
        client_id: cfg.clientId,
        refresh_token: refreshToken,
        grant_type: 'refresh_token',
        scope: 'offline_access Mail.Read Mail.ReadWrite Mail.Send Calendars.Read Tasks.ReadWrite Files.ReadWrite.All',
    });
    const resp = await fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: body.toString(),
    });
    const data = await resp.json();
    if (!data.access_token) {
        throw new Error(`[GraphClient] Token refresh failed: ${data.error} — ${data.error_description}`);
    }
    _accessToken = data.access_token;
    _tokenExpiry = Date.now() + (data.expires_in ?? 3600) * 1000;
    if (data.refresh_token && data.refresh_token !== refreshToken) {
        Logger_js_1.Logger.warn('[GraphClient] New refresh token issued — update REFRESH_TOKEN in local.settings.json:');
        Logger_js_1.Logger.warn(data.refresh_token);
    }
    return _accessToken;
}
async function getClient() {
    const token = await getAccessToken();
    return microsoft_graph_client_1.Client.init({
        authProvider: (done) => done(null, token),
    });
}
async function fetchNewMessages(_userEmail, limit = 25) {
    const client = await getClient();
    try {
        const response = await client
            .api('/me/messages')
            .filter('isRead eq false')
            .select('id,subject,bodyPreview,body,from,receivedDateTime,isRead,categories,flag,conversationId,internetMessageId,hasAttachments')
            .top(limit)
            .orderby('receivedDateTime desc')
            .get();
        return (response.value ?? []);
    }
    catch (e) {
        Logger_js_1.Logger.error('[GraphClient] fetchNewMessages failed', e);
        return [];
    }
}
async function autoLabelNewMessages(_userEmail) {
    Logger_js_1.Logger.info('[GraphClient] autoLabelNewMessages: skipped for personal account');
    return 0;
}
async function categorizeMessage(_userEmail, messageId, addCategories, _removeCategories = []) {
    const client = await getClient();
    try {
        const msg = await client.api(`/me/messages/${messageId}`).select('categories').get();
        const current = new Set(msg.categories ?? []);
        addCategories.forEach(c => current.add(c));
        await client.api(`/me/messages/${messageId}`).patch({ categories: Array.from(current) });
    }
    catch (e) {
        Logger_js_1.Logger.error('[GraphClient] categorizeMessage failed', e);
    }
}
async function flagMessage(_userEmail, messageId) {
    const client = await getClient();
    await client.api(`/me/messages/${messageId}`).patch({ flag: { flagStatus: 'flagged' } });
}
async function createDraftReply(_userEmail, messageId, body) {
    const client = await getClient();
    const draft = await client.api(`/me/messages/${messageId}/createReply`).post({});
    await client.api(`/me/messages/${draft.id}`).patch({
        body: { contentType: 'html', content: body },
    });
    return draft.id;
}
async function fetchCalendarEvents(_userEmail, startDate, endDate) {
    const client = await getClient();
    const response = await client
        .api('/me/calendarView')
        .query({ startDateTime: startDate.toISOString(), endDateTime: endDate.toISOString() })
        .select('id,subject,start,end,attendees,bodyPreview,isOnlineMeeting,onlineMeetingUrl')
        .orderby('start/dateTime')
        .get();
    return (response.value ?? []);
}
async function createTask(_userEmail, title, body, dueDate, importance = 'normal') {
    const client = await getClient();
    const lists = await client.api('/me/todo/lists').get();
    const existing = lists.value
        .find(l => l.displayName === 'AICOS');
    let listId;
    if (existing) {
        listId = existing.id;
    }
    else {
        const created = await client.api('/me/todo/lists').post({ displayName: 'AICOS' });
        listId = created.id;
    }
    const task = await client.api(`/me/todo/lists/${listId}/tasks`).post({
        title,
        body: { contentType: 'text', content: body },
        importance,
        ...(dueDate ? {
            dueDateTime: {
                dateTime: dueDate.toISOString().split('T')[0] + 'T00:00:00',
                timeZone: 'UTC',
            },
        } : {}),
    });
    return task.id;
}
async function appendExcelRow(_userEmail, workbookId, sheetName, values) {
    const client = await getClient();
    await client
        .api(`/me/drive/items/${workbookId}/workbook/worksheets/${sheetName}/tables/Table1/rows/add`)
        .post({ values: [values] });
}
async function getExcelRows(_userEmail, workbookId, sheetName) {
    const client = await getClient();
    const response = await client
        .api(`/me/drive/items/${workbookId}/workbook/worksheets/${sheetName}/tables/Table1/rows`)
        .get();
    return response.value.map(r => r.values[0]);
}
async function createMailSubscription(_userEmail, notificationUrl) {
    const client = await getClient();
    const expiry = new Date();
    expiry.setDate(expiry.getDate() + 3);
    const sub = await client.api('/subscriptions').post({
        changeType: 'created',
        notificationUrl,
        resource: '/me/mailFolders/inbox/messages',
        expirationDateTime: expiry.toISOString(),
        clientState: process.env.WEBHOOK_SECRET ?? 'aicos-secret',
    });
    return sub.id;
}
//# sourceMappingURL=GraphClient.js.map