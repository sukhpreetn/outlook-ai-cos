// =============================================================================
// scripts/seedTestData.mjs
// Seeds Outlook inbox with realistic test emails across all 8 AICOS dimensions
// USAGE: node scripts/seedTestData.mjs
// =============================================================================

import { readFileSync } from 'fs';

const settings = JSON.parse(readFileSync('./local.settings.json', 'utf8'));
const CLIENT_ID     = settings.Values.CLIENT_ID;
const REFRESH_TOKEN = settings.Values.REFRESH_TOKEN;
const USER_EMAIL    = settings.Values.USER_EMAIL;

if (!CLIENT_ID || !REFRESH_TOKEN || !USER_EMAIL) {
  console.error('❌ Missing CLIENT_ID, REFRESH_TOKEN, or USER_EMAIL in local.settings.json');
  process.exit(1);
}

// ── Get access token ──────────────────────────────────────────────────────────
async function getToken() {
  const body = new URLSearchParams({
    client_id:     CLIENT_ID,
    refresh_token: REFRESH_TOKEN,
    grant_type:    'refresh_token',
    scope:         'https://graph.microsoft.com/Mail.Send Mail.ReadWrite offline_access',
  });
  const resp = await fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body.toString(),
  });
  const data = await resp.json();
  if (!data.access_token) {
    console.error('❌ Token failed:', data.error_description);
    process.exit(1);
  }
  console.log('✓ Got access token');
  return data.access_token;
}

// ── Send email ────────────────────────────────────────────────────────────────
async function sendEmail(token, subject, body, importance = 'normal') {
  const resp = await fetch('https://graph.microsoft.com/v1.0/me/sendMail', {
    method: 'POST',
    headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      message: {
        subject,
        importance,
        body: { contentType: 'html', content: body },
        toRecipients: [{ emailAddress: { address: USER_EMAIL } }],
      },
      saveToSentItems: false,
    }),
  });
  if (resp.status === 202) {
    console.log(`  ✓ Sent: ${subject}`);
  } else {
    const err = await resp.text();
    console.log(`  ✗ Failed: ${subject} — ${err.substring(0, 120)}`);
  }
  await new Promise(r => setTimeout(r, 800));
}

// ── Email definitions ─────────────────────────────────────────────────────────
const emails = [

  // ── FINANCE (3 emails) ──────────────────────────────────────────────────────
  {
    dim: 'finance', importance: 'high',
    subject: 'Chase Credit Card Statement — Payment Due June 15',
    body: '<p>Your Chase Sapphire statement balance is <strong>$2,847.50</strong> due June 15, 2026. Minimum payment: $35.00. Log in to chase.com to pay.</p>',
  },
  {
    dim: 'finance', importance: 'high',
    subject: 'HOA Payment Reminder — June Dues $385 Due June 10',
    body: '<p>Dear Homeowner,</p><p>This is a reminder that your <strong>June HOA dues of $385.00</strong> are due by June 10, 2026.</p><p>Late fee of $50 applies after June 15. Please pay via the resident portal at myhoa.com or mail a check to: Everett HOA, PO Box 1234, Everett WA 98201.</p><p>Thank you,<br>Everett HOA Management</p>',
  },
  {
    dim: 'finance', importance: 'normal',
    subject: 'PayPal Receipt — $149.99 Adobe Creative Cloud Annual Plan',
    body: '<p>You sent <strong>$149.99 USD</strong> to Adobe Inc. Transaction ID: 7RK49285XD123456T. Date: June 3, 2026. Your annual Creative Cloud subscription has been renewed.</p>',
  },

  // ── CAREER (1 email — no interview invitation) ───────────────────────────────
  {
    dim: 'career', importance: 'normal',
    subject: 'Application Received — Staff Engineer at Stripe',
    body: '<p>Thanks for applying to the <strong>Staff Engineer, Infrastructure Platform</strong> role at Stripe. Application ID: APP-2026-089234. Our team will review within 5-7 business days.</p><p>The Stripe Recruiting Team</p>',
  },

  // ── LEGAL (2 emails) ────────────────────────────────────────────────────────
  {
    dim: 'legal', importance: 'high',
    subject: 'DocuSign: Sign Required — NDA with Acme Ventures (Due June 7)',
    body: '<p>Acme Ventures sent you a <strong>Mutual Non-Disclosure Agreement</strong> to sign. Please sign by June 7, 2026. Sent by: John Davis, john.davis@acmeventures.com</p>',
  },
  {
    dim: 'legal', importance: 'high',
    subject: 'Rainier Title — Escrow Refund Follow-Up: $1,240 Pending Release',
    body: '<p>Dear Sukhpreet,</p><p>We are following up on your escrow refund of <strong>$1,240.00</strong> from the closing of your property at 4521 Meridian Ave N.</p><p>Your refund was processed on May 15, 2026 and should have been received by now. Please confirm receipt or contact us if you have not received the check.</p><p>If not received, we can reissue via ACH. Please reply with your preferred bank details.</p><p>Best regards,<br>Amanda Torres<br>Escrow Officer, Rainier Title Company<br>(425) 555-0187</p>',
  },

  // ── KNOWLEDGE (3 emails) ────────────────────────────────────────────────────
  {
    dim: 'knowledge', importance: 'normal',
    subject: 'You are enrolled — Agentic AI Development with Azure + Claude (Starts June 10)',
    body: '<p>Hi Sukhpreet,</p><p>Your enrollment is confirmed for <strong>Agentic AI Development: Building Autonomous Systems with Azure Functions + Claude AI</strong>.</p><ul><li>Start date: June 10, 2026</li><li>Format: Live sessions + async content</li><li>Session 1: Introduction to Agentic Patterns (June 10, 7pm PST)</li><li>Session 2: Azure Functions + Graph API integration (June 17, 7pm PST)</li><li>Session 3: Claude AI for classification and reasoning (June 24, 7pm PST)</li></ul><p>Join link will be sent 1 hour before each session. Add to your calendar now.</p><p>The AI Engineering Academy</p>',
  },
  {
    dim: 'knowledge', importance: 'normal',
    subject: 'Microsoft Learn — Your Azure AI Engineer Learning Path: 3 Modules Due This Week',
    body: '<p>Hi Sukhpreet,</p><p>You have <strong>3 modules due this week</strong> in your Azure AI Engineer learning path:</p><ul><li>Module 5: Natural Language Processing with Azure Cognitive Services (Due June 6)</li><li>Module 6: Computer Vision solutions with Azure (Due June 7)</li><li>Module 7: Implementing knowledge mining with Azure Cognitive Search (Due June 8)</li></ul><p>Each module takes approximately 45-60 minutes. Complete at learn.microsoft.com</p>',
  },
  {
    dim: 'knowledge', importance: 'normal',
    subject: 'Weekly Digest — AI, Azure, TypeScript Highlights This Week',
    body: '<ul><li><strong>Claude 4 released</strong> — extended thinking + 200k context</li><li><strong>Azure Functions v5</strong> — what is new in the latest release</li><li><strong>TypeScript 5.6</strong> — new features and breaking changes</li><li><strong>Microsoft Graph API</strong> — new Mail.ReadWrite.Shared permission</li><li><strong>GitHub Copilot Workspace</strong> — now generally available</li></ul><p>Read time: ~15 minutes. The Dev Weekly Team</p>',
  },

  // ── LIFE (4 emails) ─────────────────────────────────────────────────────────
  {
    dim: 'life', importance: 'high',
    subject: 'Reminder — Dentist Appointment June 6 at 10:00 AM — Dr. Patel Family Dental',
    body: '<p>Dear Sukhpreet,</p><p>This is a reminder for your upcoming dental appointment:</p><p><strong>Provider:</strong> Dr. Priya Patel, DDS<br><strong>Type:</strong> Bi-Annual Cleaning + X-rays<br><strong>Date:</strong> Friday, June 6, 2026<br><strong>Time:</strong> 10:00 AM PST<br><strong>Location:</strong> Patel Family Dental, 8901 Evergreen Way, Everett WA 98208</p><p>Please arrive 10 minutes early. Bring your insurance card. Call (425) 555-0234 to reschedule (24h notice required).</p>',
  },
  {
    dim: 'life', importance: 'high',
    subject: 'Silver Firs Elementary — 6th Grade PTM Scheduled: June 12 at 4:00 PM',
    body: '<p>Dear Parent/Guardian,</p><p>This is a confirmation of your Parent-Teacher Meeting for <strong>6th Grade</strong>:</p><p><strong>Date:</strong> Thursday, June 12, 2026<br><strong>Time:</strong> 4:00 PM – 4:30 PM<br><strong>Teacher:</strong> Ms. Jennifer Walsh<br><strong>Location:</strong> Room 214, Silver Firs Elementary</p><p>Topics to discuss: Academic progress, state test results, summer reading recommendations, and 7th grade transition planning.</p><p>Please bring any questions you have. If you need to reschedule, contact the office at (425) 555-0100 by June 10.</p><p>Silver Firs Elementary School</p>',
  },
  {
    dim: 'life', importance: 'high',
    subject: 'June Rent Follow-Up — Unit 4B: Payment Not Received as of June 3',
    body: '<p>Hi,</p><p>This is a friendly follow-up regarding the <strong>June rent for Unit 4B</strong>.</p><p>As of today, June 3, 2026, we have not received your June rent payment of <strong>$1,850.00</strong> which was due June 1, 2026.</p><p>Please submit payment via Zelle (landlord@email.com) or check by June 5 to avoid the $75 late fee per the lease agreement.</p><p>If you have already sent payment, please disregard this message and reply with confirmation.</p><p>Thank you,<br>Sukhpreet Nagi<br>Property Owner</p>',
  },
  {
    dim: 'life', importance: 'normal',
    subject: 'Best Brains Summer Camp 2026 — Registration Deadline June 15 (Save $300)',
    body: '<p>STEM + Math Enrichment camp. Dates: July 7 – Aug 15, 2026. Everett Community Center. <strong>Early bird: $1,200</strong> (save $300). Deadline: June 15. Register at bestbrains.com/summer2026</p>',
  },

  // ── VENTURES (2 emails) ─────────────────────────────────────────────────────
  {
    dim: 'ventures', importance: 'high',
    subject: 'Investment Opportunity — NeuralOps AI Seed Round $2.5M at $10M Valuation',
    body: '<p>Hi Sukhpreet,</p><p>Exciting seed round opportunity. Company: <strong>NeuralOps Inc.</strong> (AI Infrastructure/MLOps). Raising $2.5M at $10M valuation. Min check: $25,000. Strong traction: 12 enterprise pilots, $180k ARR. Deck attached.</p><p>— Michael Torres, Sequoia Scout</p>',
  },
  {
    dim: 'ventures', importance: 'normal',
    subject: 'Partnership Proposal — Integrate AICOS with TalentFlow HR Platform (20% Rev Share)',
    body: '<p>Impressed by AICOS on GitHub. We have 500+ enterprise customers, $5M ARR. Proposal: <strong>20% revenue share</strong> partnership. Available for a 30-min call? — Priya Sharma, CEO TalentFlow AI</p>',
  },

  // ── COMMERCE (2 emails) ─────────────────────────────────────────────────────
  {
    dim: 'commerce', importance: 'normal',
    subject: 'Amazon Order Shipped — Logitech MX Master 3S #1Z999AA10123456784',
    body: '<p>Your order has shipped. Item: <strong>Logitech MX Master 3S Wireless Mouse</strong>. Tracking: 1Z999AA10123456784 (UPS). Estimated delivery: June 5, 2026.</p>',
  },
  {
    dim: 'commerce', importance: 'normal',
    subject: 'Netflix Subscription Renewed — $22.99/month Premium Plan',
    body: '<p>Your Netflix Premium plan has been renewed. Amount: <strong>$22.99</strong>. Next billing: July 3, 2026. Manage at netflix.com/account</p>',
  },

  // ── INFRA (2 emails) ────────────────────────────────────────────────────────
  {
    dim: 'infra', importance: 'high',
    subject: 'GitHub: New OAuth App Authorized — VS Code GitHub Copilot',
    body: '<p>A new OAuth application was authorized on your GitHub account. App: <strong>VS Code GitHub Copilot</strong>. Permissions: Read repositories, Read user profile. If you did not authorize this, revoke at github.com/settings/applications</p>',
  },
  {
    dim: 'infra', importance: 'normal',
    subject: 'Azure Alert — Migrate aicos-outlook-nagi to Flex Consumption by Sept 2028',
    body: '<p>Action required: Your Function App <strong>aicos-outlook-nagi</strong> runs on Linux Consumption (Y1) which reaches EOL September 30, 2028. Migrate to Flex Consumption for better performance.</p>',
  },

];

// ── Main ──────────────────────────────────────────────────────────────────────
async function main() {
  console.log('\n=== AICOS Test Data Seeder ===');
  console.log(`Seeding inbox for: ${USER_EMAIL}\n`);

  const token = await getToken();

  const dims = [...new Set(emails.map(e => e.dim))];
  for (const dim of dims) {
    console.log(`\n📧 ${dim.toUpperCase()}:`);
    for (const email of emails.filter(e => e.dim === dim)) {
      await sendEmail(token, email.subject, email.body, email.importance ?? 'normal');
    }
  }

  console.log('\n=== Done! ===');
  console.log(`✓ ${emails.length} emails sent across 8 dimensions`);
  console.log('\nDimension breakdown:');
  dims.forEach(d => console.log(`  ${d}: ${emails.filter(e=>e.dim===d).length} emails`));
  console.log('\nNext steps:');
  console.log('  1. Trigger classifier:');
  console.log('     curl -X POST http://localhost:7071/admin/functions/classifierTimer \\');
  console.log('       -H "Content-Type: application/json" -d "{}"');
  console.log('  2. Wait 2 min, then trigger briefing:');
  console.log('     curl -X POST http://localhost:7071/api/ops/briefing');
  console.log('\nNote: Add calendar events manually in Outlook for the Your Day section.');
  console.log('  Suggested events to add:');
  console.log('  - Dentist Appointment (June 6, 10am)');
  console.log('  - 6th Grade PTM (June 12, 4pm)');
  console.log('  - AI Training Session 1 (June 10, 7pm)');
  console.log('  - Study Focus Hour (evening, daily)');
}

main().catch(e => { console.error('Fatal:', e); process.exit(1); });
