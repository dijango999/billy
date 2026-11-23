/* Billy — sends the claim email from your own Gmail / Google Workspace address (e.g. shaheer@choppershoot.com).
 *
 * Sign-in: Google Identity Services "token model". Google shows its own sign-in / consent window; Billy gets a
 * short-lived (about 1 hour) token that can ONLY send mail (scope gmail.send) — it cannot read your inbox.
 * No password or secret is stored. The OAuth Client ID is public by design; it is set once in Settings.
 *
 * The message is built here (MIME: text + HTML body, Excel and bills PDF attached) and sent with the Gmail
 * API media upload, which accepts messages up to 35 MB.
 */
'use strict';

const GMAIL = (() => {
  const SCOPE = 'https://www.googleapis.com/auth/gmail.send';
  const MAX = 24 * 1024 * 1024;       // stay under Gmail's 25 MB attachment limit
  let gis = null, token = null, client = null, clientFor = '';

  function load() {
    if (window.google?.accounts?.oauth2) return Promise.resolve();
    return gis ||= new Promise((res, rej) => {
      const s = document.createElement('script'); s.src = 'https://accounts.google.com/gsi/client'; s.async = true;
      s.onload = () => window.google?.accounts?.oauth2 ? res() : rej(new Error('Google sign-in did not load.'));
      s.onerror = () => { gis = null; rej(new Error('Could not reach Google. Check your internet connection.')); };
      document.head.appendChild(s);
    });
  }
  const valid = () => token && token.exp > Date.now() + 60000;

  /**
   * Gets a send-only token. Must be called straight from a tap (Google opens a window).
   * Call preload() earlier (e.g. when the screen opens) so the Google script is ready by then.
   */
  function authorize(clientId, hint) {
    if (valid()) return Promise.resolve(token.value);
    if (!window.google?.accounts?.oauth2) return Promise.reject(new Error('Google sign-in is still loading — tap again in a second.'));
    return new Promise((res, rej) => {
      if (!client || clientFor !== clientId) {
        clientFor = clientId;
        client = google.accounts.oauth2.initTokenClient({ client_id: clientId, scope: SCOPE, callback: () => {} });
      }
      client.callback = r => {
        if (r.error || !r.access_token) return rej(new Error(r.error === 'access_denied' ? 'Gmail permission was not given.' : (r.error_description || 'Google sign-in did not complete.')));
        if (!google.accounts.oauth2.hasGrantedAllScopes(r, SCOPE)) return rej(new Error('Please tick “Send email on your behalf” so Billy can send the claim.'));
        token = { value: r.access_token, exp: Date.now() + (+r.expires_in || 3600) * 1000 };
        res(token.value);
      };
      client.error_callback = e => rej(new Error(e?.type === 'popup_closed' ? 'Google window was closed.' : e?.type === 'popup_failed_to_open' ? 'Allow pop-ups for Billy, then try again.' : 'Google sign-in did not complete.'));
      client.requestAccessToken({ hint: hint || undefined, prompt: '' });
    });
  }

  /* ---------- building the email ---------- */
  const enc = s => btoa(unescape(encodeURIComponent(s)));
  const wrap = b64 => b64.replace(/.{1,76}/g, '$&\r\n');
  const header = s => /^[\x20-\x7E]*$/.test(s) ? s : `=?UTF-8?B?${enc(s)}?=`;
  async function fileB64(f) {
    const a = new Uint8Array(await f.arrayBuffer()); let s = '';
    for (let i = 0; i < a.length; i += 0x8000) s += String.fromCharCode.apply(null, a.subarray(i, i + 0x8000));
    return btoa(s);
  }
  async function mime({ from, to, cc, subject, text, html, files }) {
    const b1 = 'billy_mixed_' + Math.random().toString(36).slice(2), b2 = 'billy_alt_' + Math.random().toString(36).slice(2);
    let m = '';
    m += `To: ${to}\r\n`;
    if (cc) m += `Cc: ${cc}\r\n`;
    m += `Subject: ${header(subject)}\r\nMIME-Version: 1.0\r\nContent-Type: multipart/mixed; boundary="${b1}"\r\n\r\n`;
    m += `--${b1}\r\nContent-Type: multipart/alternative; boundary="${b2}"\r\n\r\n`;
    m += `--${b2}\r\nContent-Type: text/plain; charset=UTF-8\r\nContent-Transfer-Encoding: base64\r\n\r\n${wrap(enc(text))}\r\n`;
    m += `--${b2}\r\nContent-Type: text/html; charset=UTF-8\r\nContent-Transfer-Encoding: base64\r\n\r\n${wrap(enc(html))}\r\n--${b2}--\r\n`;
    for (const f of files) {
      const name = header(f.name);
      m += `--${b1}\r\nContent-Type: ${f.type || 'application/octet-stream'}; name="${name}"\r\nContent-Disposition: attachment; filename="${name}"\r\nContent-Transfer-Encoding: base64\r\n\r\n${wrap(await fileB64(f))}\r\n`;
    }
    return m + `--${b1}--\r\n`;
  }

  /** Sends the message. Returns the Gmail message id.
   * Small messages (most claims) go as JSON to the standard endpoint; large ones use the documented upload
   * endpoint on www.googleapis.com (gmail.googleapis.com/upload is not reachable from Safari). */
  async function send(tok, msg) {
    const raw = await mime(msg);
    const b64url = s => { const b = new TextEncoder().encode(s); let x = ''; for (let i = 0; i < b.length; i += 0x8000) x += String.fromCharCode.apply(null, b.subarray(i, i + 0x8000)); return btoa(x).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, ''); };
    const small = raw.length < 4.5 * 1024 * 1024;
    let r;
    try {
      r = small
        ? await fetch('https://gmail.googleapis.com/gmail/v1/users/me/messages/send', { method: 'POST', headers: { Authorization: 'Bearer ' + tok, 'Content-Type': 'application/json' }, body: JSON.stringify({ raw: b64url(raw) }) })
        : await fetch('https://www.googleapis.com/upload/gmail/v1/users/me/messages/send?uploadType=media', { method: 'POST', headers: { Authorization: 'Bearer ' + tok, 'Content-Type': 'message/rfc822' }, body: raw });
    } catch (e) { throw new Error('Could not reach Gmail (' + (e.message || 'network') + '). Check the internet and try again.'); }
    const j = await r.json().catch(() => ({}));
    if (r.status === 401) { token = null; throw new Error('Google sign-in expired — tap Send again.'); }
    if (r.status === 403 && /insufficient|scope/i.test(j.error?.message || '')) { token = null; throw new Error('Gmail permission missing — tap Send again and tick “Send email on your behalf”.'); }
    if (!r.ok) throw new Error('Gmail: ' + (j.error?.message || `could not send the email (${r.status}).`));
    return j.id;
  }
  const tooBig = files => files.reduce((a, f) => a + f.size, 0) * 1.37 > MAX;

  return { load, authorize, send, tooBig, valid, signOut() { if (token) try { google.accounts.oauth2.revoke(token.value); } catch {} token = null; } };
})();
