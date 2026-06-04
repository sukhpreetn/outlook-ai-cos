import { createServer } from 'http';

const CLIENT_ID    = 'e3bc9bf6-3a9a-4f05-a109-b92d71716c54';
const REDIRECT_URI = 'http://localhost:8080/callback';
const SCOPES       = 'offline_access Mail.Read Mail.ReadWrite Mail.Send Calendars.Read Calendars.ReadWrite Tasks.ReadWrite Files.ReadWrite.All';

const authUrl = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize' +
  '?client_id='     + CLIENT_ID +
  '&response_type=code' +
  '&redirect_uri='  + encodeURIComponent(REDIRECT_URI) +
  '&scope='         + encodeURIComponent(SCOPES) +
  '&response_mode=query' +
  '&prompt=consent';

console.log('\n1. Open this URL in your browser:\n');
console.log(authUrl);
console.log('\n2. Sign in with sukhpreetn@outlook.com');
console.log('3. IMPORTANT: Click Accept on ALL permission checkboxes');
console.log('4. Token will appear here automatically\n');

const server = createServer(async (req, res) => {
  const url  = new URL(req.url, 'http://localhost:8080');
  const code = url.searchParams.get('code');
  if (!code) { res.end('No code found'); return; }

  const body = new URLSearchParams({
    client_id:    CLIENT_ID,
    code,
    redirect_uri: REDIRECT_URI,
    grant_type:   'authorization_code',
    scope:        SCOPES,
  });

  const resp = await fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
    method:  'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body:    body.toString(),
  });

  const data = await resp.json();

  if (data.refresh_token) {
    console.log('\n✓ Add this to local.settings.json:\n');
    console.log('"REFRESH_TOKEN": "' + data.refresh_token + '"\n');
    console.log('Granted scopes:', data.scope);
    res.end('<h2>Success! Check your terminal for the REFRESH_TOKEN.</h2>');
  } else {
    console.log('\n✗ Error:', JSON.stringify(data, null, 2));
    res.end('Error - check terminal');
  }
  server.close();
});

server.listen(8080, () => {
  console.log('Waiting on http://localhost:8080/callback ...\n');
});
