#!/usr/bin/env node
// scripts/registerWebhook.js
// Registers a Microsoft Graph change notification subscription for new Outlook mail.
// Run after deploy: node scripts/registerWebhook.js

import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from
  '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js';

const {
  TENANT_ID,
  CLIENT_ID,
  CLIENT_SECRET,
  USER_EMAIL,
  WEBHOOK_URL,
  WEBHOOK_SECRET = 'aicos-secret',
} = process.env;

if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET || !USER_EMAIL || !WEBHOOK_URL) {
  console.error('Missing required env vars: TENANT_ID, CLIENT_ID, CLIENT_SECRET, USER_EMAIL, WEBHOOK_URL');
  process.exit(1);
}

const credential = new ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
const authProvider = new TokenCredentialAuthenticationProvider(credential, {
  scopes: ['https://graph.microsoft.com/.default'],
});
const client = Client.initWithMiddleware({ authProvider });

const expiry = new Date();
expiry.setDate(expiry.getDate() + 3);

try {
  const sub = await client.api('/subscriptions').post({
    changeType: 'created',
    notificationUrl: WEBHOOK_URL,
    resource: `/users/${USER_EMAIL}/mailFolders/inbox/messages`,
    expirationDateTime: expiry.toISOString(),
    clientState: WEBHOOK_SECRET,
  });

  console.log('✓ Subscription registered:');
  console.log(`  ID:      ${sub.id}`);
  console.log(`  Expires: ${sub.expirationDateTime}`);
  console.log('');
  console.log('Remember to renew this subscription before it expires (3-day max).');
} catch (e) {
  console.error('Failed to register subscription:', e.message);
  process.exit(1);
}
