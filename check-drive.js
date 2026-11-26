// Diagnostic: shows which service account the key belongs to and what it can see of the folder.
// Prints no secret values.
const crypto = require('crypto');
(async () => {
  const key = JSON.parse(process.env.GDRIVE_SERVICE_ACCOUNT_KEY || '{}');
  const folder = (process.env.GDRIVE_FOLDER_ID || '');
  const email = String(key.client_email || '');
  console.log('Service account:', email.replace('@', ' at '));
  console.log('Folder id: length', folder.length, 'starts', folder.slice(0, 4), 'ends', folder.slice(-4), 'trimmed same:', folder.trim() === folder);
  const now = Math.floor(Date.now() / 1000);
  const b64 = (o) => Buffer.from(JSON.stringify(o)).toString('base64url');
  const unsigned = b64({alg: 'RS256', typ: 'JWT'}) + '.' + b64({iss: email, scope: 'https://www.googleapis.com/auth/drive.readonly', aud: 'https://oauth2.googleapis.com/token', iat: now, exp: now + 3600});
  const sig = crypto.createSign('RSA-SHA256').update(unsigned).sign(key.private_key).toString('base64url');
  const tr = await fetch('https://oauth2.googleapis.com/token', {method: 'POST', headers: {'content-type': 'application/x-www-form-urlencoded'}, body: new URLSearchParams({grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: unsigned + '.' + sig})});
  console.log('Token:', tr.status);
  const token = (await tr.json()).access_token;
  const H = {authorization: 'Bearer ' + token};
  const g = await fetch(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(folder.trim())}?fields=id,name,mimeType&supportsAllDrives=true`, {headers: H});
  const gj = await g.json();
  console.log('Folder lookup:', g.status, gj.name ? 'name=' + gj.name : JSON.stringify(gj.error && gj.error.message));
  const all = await fetch('https://www.googleapis.com/drive/v3/files?pageSize=20&fields=files(name,mimeType)&supportsAllDrives=true&includeItemsFromAllDrives=true', {headers: H});
  const aj = await all.json();
  console.log('Everything this account can see:', all.status, (aj.files || []).map((f) => f.name).join(' | ') || '(nothing)');
})().catch((e) => { console.error('Check failed:', e.message); process.exit(1); });
