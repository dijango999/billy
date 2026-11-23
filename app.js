/* Billy — office expenses, bills and the monthly claim, kept in your own OneDrive.
 *
 * Data lives in a OneDrive folder (default: Documents › 02 ACCOUNTS › OFFICE › DOC › EXPENCE › Billy):
 *   billy.json      every expense, project, monthly item and claim
 *   Bills/YYYY-MM/  bills and receipts (photos are saved as JPG, PDFs as they are)
 *   Claims/         one folder per claim: the Excel sheet and one PDF with all the bills
 *   Backups/        one copy of billy.json per day, taken before the first change of the day
 *
 * Every change is applied to the newest copy from OneDrive and saved with an eTag check, so a
 * phone and a laptop editing at the same time never overwrite each other. Sign-in is the same
 * Microsoft sign-in as Reno App and My Gold (PKCE, no secrets in the code).
 */
'use strict';

const CFG = {
  clientId: 'c010a16c-4e83-493b-b778-86ad96b766e9',
  authority: 'https://login.microsoftonline.com/consumers/oauth2/v2.0',
  scopes: 'Files.ReadWrite User.Read offline_access',
  graph: 'https://graph.microsoft.com/v1.0',
  dataFile: 'billy.json',
  pointer: 'Apps/Billy/location.json',
  version: '1.1.0',
};
const CANON = 'https://dijango999.github.io/billy/';
const ON_SITE = location.hostname === 'dijango999.github.io';
if (ON_SITE && location.origin + location.pathname !== CANON) location.replace(CANON + location.search);
const REDIRECT = ON_SITE ? CANON : location.origin + location.pathname.replace(/index\.html$/, '');
const PAST_FOLDER = 'Documents/02 ACCOUNTS/OFFICE/DOC/EXPENCE';
const DEFAULT_PATH = PAST_FOLDER + '/Billy';

const CATS = [
  ['Music & licensing', '🎵'], ['Air tickets', '✈️'], ['Hotel & stay', '🏨'], ['Taxi & transport', '🚕'], ['Fuel', '⛽'],
  ['Parking & tolls', '🅿️'], ['Food & meals', '🍽️'], ['Equipment rental', '🎥'], ['Equipment & gear', '🛠️'],
  ['Software & subscriptions', '💻'], ['Stock footage & assets', '🎞️'], ['Props & set', '🎬'], ['Crew & freelancers', '👥'],
  ['Permits & fees', '📄'], ['Courier & printing', '📦'], ['Phone & internet', '📱'], ['Home rent', '🏠'], ['School fees', '🎒'],
  ['Shoot day allowance', '📅'], ['Other', '🧾'],
].map(([name, icon]) => ({ name, icon }));
const MONTHLY_CATS = ['Phone & internet', 'Home rent', 'School fees', 'Software & subscriptions', 'Shoot day allowance'];
const CURRENCIES = ['AED', 'USD', 'EUR', 'GBP', 'INR', 'SAR', 'OMR', 'QAR', 'KWD', 'BHD', 'EGP', 'TRY', 'PKR', 'LKR', 'THB', 'SGD', 'JPY', 'CNY', 'AUD', 'CAD', 'CHF'];
const PROJECT_COLORS = ['#E4572E', '#2563EB', '#1C7B62', '#7A5AA6', '#C98A1A', '#C0447A', '#2E8A9E', '#5B6577'];

/* ---------------- helpers ---------------- */
const $ = (s, el = document) => el.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const ls = {
  get(k) { try { return localStorage.getItem(k); } catch { return null; } },
  set(k, v) { try { localStorage.setItem(k, v); } catch {} },
  del(k) { try { localStorage.removeItem(k); } catch {} },
};
const lsJSON = (k, d = null) => { try { return JSON.parse(ls.get(k) || 'null') ?? d; } catch { return d; } };
const uuid = () => (crypto.randomUUID ? crypto.randomUUID() : 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => {
  const r = Math.random() * 16 | 0; return (c === 'x' ? r : (r & 3 | 8)).toString(16); }));
const isoNow = () => new Date().toISOString().replace(/\.\d{3}Z$/, 'Z');
const pad2 = n => String(n).padStart(2, '0');
const ymd = d => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
const todayStr = () => ymd(new Date());
const parseDay = s => { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s || ''); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null; };
const shortDate = s => { const d = parseDay(s); return d ? d.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' }) : '–'; };
const dayHead = s => { const d = parseDay(s); if (!d) return '–'; const t = todayStr(); if (s === t) return 'Today'; return d.toLocaleDateString('en-GB', { weekday: 'short', day: 'numeric', month: 'short', year: d.getFullYear() === new Date().getFullYear() ? undefined : 'numeric' }); };
const timeStr = d => d ? new Date(d).toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' }) : '';
const mk = s => (s || '').slice(0, 7);
const thisMonth = () => mk(todayStr());
const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const monthLabel = k => { const [y, m] = (k || '').split('-'); return m ? `${MONTH_NAMES[+m - 1]} ${y}` : ''; };
const monthShort = k => { const [y, m] = (k || '').split('-'); return m ? `${MONTH_NAMES[+m - 1].slice(0, 3)} ${y}` : ''; };
const prevMonth = k => { const [y, m] = k.split('-').map(Number); return m === 1 ? `${y - 1}-12` : `${y}-${pad2(m - 1)}`; };
const cmpText = (a, b) => String(a).localeCompare(String(b), undefined, { sensitivity: 'base', numeric: true });
const num = v => { const n = parseFloat(String(v ?? '').replace(/[,\s]/g, '').replace(/^[A-Z]{3}/i, '')); return isFinite(n) ? n : NaN; };
const round2 = n => Math.round(n * 100) / 100;
const base = () => S.data?.settings?.currency || 'AED';
function money(n, cur = base(), dp = 2) {
  if (n == null || !isFinite(n)) return '–';
  return (n < 0 ? '-' : '') + cur + ' ' + Math.abs(n).toLocaleString('en-US', { minimumFractionDigits: dp, maximumFractionDigits: dp });
}
const moneyS = (n, cur = base()) => money(n, cur, Math.abs(n) >= 10000 || Number.isInteger(n) ? 0 : 2);
const _tu = new Map();
function tu(d) {
  if (!d || !d.startsWith('data:')) return d || '';
  let u = _tu.get(d); if (u) return u;
  try { const [h, b] = d.split(','), bin = atob(b), a = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) a[i] = bin.charCodeAt(i);
    u = URL.createObjectURL(new Blob([a], { type: /:(.*?);/.exec(h)?.[1] || 'image/jpeg' })); } catch { u = d; }
  _tu.set(d, u); return u;
}
const initials = n => String(n).split(/\s+/).filter(Boolean).slice(0, 2).map(w => w[0]).join('').toUpperCase() || '?';
const plural = (n, w) => `${n} ${w}${n === 1 ? '' : 's'}`;

/* ---------------- state ---------------- */
const S = {
  data: null, etag: null, account: ls.get('by.account'), demo: ls.get('by.demo') === '1',
  phase: 'loading',                 // loading | welcome | needsignin | folder | setup | ready
  tab: ls.get('by.tab') || 'home',
  f: { month: '', q: '', status: 'all' },
  sync: { state: 'idle', at: null, msg: '' },
  blobs: {},                        // claim id → { xlsx, pdf } kept in memory for sharing
};
let tokens = lsJSON('by.tokens');

/* ---------------- Microsoft sign-in (PKCE, no secrets) ---------------- */
class AuthError extends Error {}
class HttpError extends Error { constructor(status, msg) { super(msg); this.status = status; } }
const b64url = buf => btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
const randStr = n => { const a = new Uint8Array(n); crypto.getRandomValues(a); return b64url(a).slice(0, n); };

async function signIn(chooseAccount, silent) {
  const verifier = randStr(64);
  const challenge = b64url(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)));
  const state = randStr(24);
  ls.set('by.pkce', JSON.stringify({ verifier, state }));
  const p = new URLSearchParams({
    client_id: CFG.clientId, response_type: 'code', redirect_uri: REDIRECT, scope: CFG.scopes,
    code_challenge: challenge, code_challenge_method: 'S256', state, response_mode: 'query',
  });
  const hint = S.account || ls.get('mg.account');
  if (hint && !chooseAccount) p.set('login_hint', hint); else p.set('prompt', 'select_account');
  if (silent) { p.set('prompt', 'none'); ls.set('by.silentAt', String(Date.now())); }
  location.href = `${CFG.authority}/authorize?${p}`;
}
async function handleRedirect() {
  const q = new URLSearchParams(location.search);
  if (!q.has('code') && !q.has('error')) return;
  history.replaceState(null, '', REDIRECT);
  if (q.get('error')) {
    const quiet = ['access_denied', 'login_required', 'interaction_required', 'consent_required'].includes(q.get('error'));
    if (!quiet) toast(q.get('error_description') || 'Sign-in did not complete.');
    return;
  }
  const saved = lsJSON('by.pkce'); ls.del('by.pkce');
  if (!saved || saved.state !== q.get('state')) { toast('That sign-in expired. Please try again.'); return; }
  await tokenRequest({ grant_type: 'authorization_code', code: q.get('code'), redirect_uri: REDIRECT, code_verifier: saved.verifier });
  try {
    const me = await (await g('/me?$select=userPrincipalName,displayName,mail')).json();
    S.account = me.userPrincipalName || me.mail || me.displayName || 'Microsoft account';
    ls.set('by.account', S.account);
    if (me.displayName) ls.set('by.displayName', me.displayName);
  } catch {}
}
async function tokenRequest(fields) {
  const body = new URLSearchParams({ client_id: CFG.clientId, scope: CFG.scopes, ...fields });
  let r; try { r = await fetch(`${CFG.authority}/token`, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body }); }
  catch { throw new HttpError(0, 'No internet connection.'); }
  const j = await r.json().catch(() => ({}));
  if (!r.ok || !j.access_token) throw new AuthError(j.error_description || 'Microsoft sign-in failed.');
  tokens = { access: j.access_token, refresh: j.refresh_token || tokens?.refresh, exp: Date.now() + ((j.expires_in || 3600) - 120) * 1000 };
  ls.set('by.tokens', JSON.stringify(tokens));
}
let refreshing = null;
async function accessToken() {
  if (tokens?.access && tokens.exp > Date.now()) return tokens.access;
  if (!tokens?.refresh) throw new AuthError('Please sign in.');
  refreshing ||= tokenRequest({ grant_type: 'refresh_token', refresh_token: tokens.refresh }).finally(() => { refreshing = null; });
  await refreshing; return tokens.access;
}
function trySilent() {
  const last = +(ls.get('by.silentAt') || 0);
  if (!S.account || Date.now() - last < 120000 || !navigator.onLine) return false;
  S.phase = 'loading'; render(); signIn(false, true); return true;
}
function signOut() {
  for (const k of ['by.account', 'by.tokens', 'by.cache', 'by.etag', 'by.loc', 'by.demo', 'by.demoData']) ls.del(k);
  tokens = null; LOC = null; S.account = null; S.demo = false; S.data = null; S.etag = null;
  S.phase = 'welcome'; closeSheet(); render();
}

/* ---------------- OneDrive (Microsoft Graph) ---------------- */
async function g(path, opts = {}) {
  const url = path.startsWith('http') ? path : CFG.graph + path;
  for (let i = 0; i < 2; i++) {
    const t = await accessToken();
    let r; try { r = await fetch(url, { ...opts, headers: { ...(opts.headers || {}), Authorization: 'Bearer ' + t } }); }
    catch { throw new HttpError(0, 'No internet connection.'); }
    if (r.status === 401 && i === 0) { tokens.exp = 0; continue; }
    if (!r.ok) { let m = ''; try { m = (await r.json()).error?.message; } catch {} throw new HttpError(r.status, m || `OneDrive error ${r.status}`); }
    return r;
  }
}
let LOC = lsJSON('by.loc');
const encRel = rel => rel.split('/').filter(Boolean).map(encodeURIComponent).join('/');
const item = rel => { if (!LOC) throw new Error('Choose where to keep your data first.'); return `/drives/${LOC.driveId}/items/${LOC.id}:/${encRel(rel)}`; };
const locFull = (l = LOC) => l ? 'OneDrive › ' + (l.path ? l.path.split('/').join(' › ') : 'My files') : 'not chosen';
const locShort = (l = LOC) => l ? (l.path ? l.path.split('/').slice(-2).join(' › ') : 'My files') : '';
function folderPath(meta) {
  const parent = decodeURIComponent((meta.parentReference?.path || '').replace(/^\/drive(s\/[^/]+)?\/root:?\/?/, ''));
  return meta.root ? '' : [parent, meta.name].filter(Boolean).join('/');
}
function saveLoc(l) {
  const changed = !LOC || LOC.id !== l.id;
  LOC = { driveId: l.driveId, id: l.id, path: l.path || '' }; ls.set('by.loc', JSON.stringify(LOC));
  if (changed) { ls.del('by.cache'); ls.del('by.etag'); S.data = null; S.etag = null; }
}
async function readPointer() {
  try {
    const m = await (await g(`/me/drive/root:/${encRel(CFG.pointer)}`)).json();
    const p = JSON.parse(await (await fetch(m['@microsoft.graph.downloadUrl'], { cache: 'no-store' })).text());
    return p && p.id && p.driveId ? p : null;
  } catch (e) { if (e.status === 404 || e instanceof SyntaxError) return null; throw e; }
}
async function writePointer(l) {
  await g(`/me/drive/root:/${encRel(CFG.pointer)}:/content`, { method: 'PUT', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ driveId: l.driveId, id: l.id, path: l.path, savedAt: isoNow(), by: S.account || '' }, null, 2) });
}
/** Finds the data folder: the saved pointer, then EXPENCE › Billy (made automatically if EXPENCE exists). */
async function resolveLocation() {
  const p = await readPointer();
  if (p) {
    try { const m = await (await g(`/drives/${p.driveId}/items/${p.id}?$select=id,name,root,parentReference`)).json();
      saveLoc({ driveId: p.driveId, id: p.id, path: folderPath(m) }); return true; }
    catch (e) { if (e.status !== 404) throw e; }
  }
  if (LOC) { await writePointer(LOC).catch(() => {}); return true; }
  const use = async f => { const l = { driveId: f.parentReference.driveId, id: f.id, path: DEFAULT_PATH }; saveLoc(l); await writePointer(l).catch(() => {}); return true; };
  try { return await use(await (await g(`/me/drive/root:/${encRel(DEFAULT_PATH)}?$select=id,name,parentReference`)).json()); }
  catch (e) { if (e.status !== 404) throw e; }
  try {
    const parent = await (await g(`/me/drive/root:/${encRel(PAST_FOLDER)}?$select=id,parentReference`)).json();
    const f = await (await g(`/drives/${parent.parentReference.driveId}/items/${parent.id}/children`, { method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: 'Billy', folder: {}, '@microsoft.graph.conflictBehavior': 'fail' }) })).json();
    return await use(f);
  } catch (e) { if (e.status !== 404 && e.status !== 409) throw e; }
  return false;
}
async function loadRemote() {
  let meta;
  try { meta = await (await g(item(CFG.dataFile))).json(); }
  catch (e) { if (e.status === 404) return null; throw e; }
  const r = await fetch(meta['@microsoft.graph.downloadUrl'], { cache: 'no-store' });
  if (!r.ok) throw new HttpError(r.status, 'Could not download your data.');
  const text = await r.text();
  let data; try { data = JSON.parse(text); } catch { throw new Error(`${CFG.dataFile} in OneDrive could not be read. Nothing was changed.`); }
  return { data: normalize(data), etag: meta.eTag, text };
}
function normalize(d) {
  d.projects ||= []; d.expenses ||= []; d.claims ||= []; d.monthly ||= []; d.learned ||= [];
  if (!d.categories?.length) d.categories = CATS.map(c => ({ ...c }));
  for (const c of CATS) if (!d.categories.some(x => x.name === c.name)) d.categories.push({ ...c });
  d.settings = { name: '', accountant: 'Adnan Ahmad', email: 'accounts@choppershoot.com', company: 'Choppershoot', currency: 'AED', dayRate: 50, fromEmail: '', cc: '', googleClientId: '', ...(d.settings || {}) };
  for (const x of d.expenses) { x.bills ||= []; x.currency ||= d.settings.currency; }
  return d;
}
const stableJSON = v => JSON.stringify(v, null, 2);
async function putData(data, etag, createOnly) {
  const headers = { 'Content-Type': 'application/json' };
  if (createOnly) headers['If-None-Match'] = '*'; else if (etag) headers['If-Match'] = etag;
  const r = await g(item(CFG.dataFile) + ':/content', { method: 'PUT', headers, body: stableJSON(data) });
  return (await r.json()).eTag;
}
async function dailyBackup(text) {
  const day = todayStr();
  if (ls.get('by.backup') === day) return;
  try {
    await g(item(`Backups/billy-${day}.json`) + ':/content?@microsoft.graph.conflictBehavior=fail',
      { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: text });
  } catch (e) { if (e.status !== 409) { console.warn('backup', e); return; } }
  ls.set('by.backup', day);
}

let saving = Promise.resolve();
/** Applies `change` to the newest data and saves it. Retries if another device saved first. */
function mutate(change) {
  const run = async () => {
    if (S.demo) {
      const d = normalize(JSON.parse(JSON.stringify(S.data))); change(d); d.savedAt = isoNow();
      ls.set('by.demoData', JSON.stringify(d)); applyData(d, null); return true;
    }
    if (!navigator.onLine) throw new Error('You are offline. Connect to the internet to save changes.');
    setSync('busy', 'Saving…');
    for (let attempt = 0; attempt < 5; attempt++) {
      let d, etag;
      if (attempt === 0 && S.base && S.etag) { d = JSON.parse(JSON.stringify(S.base)); etag = S.etag; }
      else {
        const remote = await loadRemote();
        if (!remote) throw new Error('Your data file is missing in OneDrive. Nothing was saved.');
        await dailyBackup(remote.text); d = remote.data; etag = remote.etag;
      }
      if (attempt === 0 && ls.get('by.backup') !== todayStr()) { try { await dailyBackup(stableJSON(d)); } catch {} }
      change(d); d.savedAt = isoNow();
      try {
        const e2 = await putData(d, etag);
        if (S.fast > 0) { S.base = JSON.parse(JSON.stringify(d)); S.etag = e2; ls.set('by.etag', e2 || ''); ls.set('by.cache', JSON.stringify(d)); }
        else applyData(d, e2);
        setSync('ok'); return true;
      }
      catch (e) { if (e.status === 412 || e.status === 409) { S.base = null; continue; } throw e; }
    }
    throw new Error('OneDrive is busy. Please try again.');
  };
  const p = saving.then(run, run);
  saving = p.catch(() => {});
  return p.catch(e => { setSync('bad', e.message); toast(e.message); return false; });
}
/** Instant version for small changes: shows the result at once, saves in the background (undone if the save fails). */
function mutateFast(change) {
  if (S.demo || !S.data) return mutate(change);
  change(S.data); S.data.savedAt = isoNow();
  try { ls.set('by.cache', JSON.stringify(S.data)); } catch {}
  render();
  S.fast = (S.fast || 0) + 1;
  mutate(change).then(ok => { S.fast--; if (!ok) pull(true); else if (S.fast === 0) { S.data = JSON.parse(JSON.stringify(S.base || S.data)); render(); } });
  return Promise.resolve(true);
}
function applyData(d, etag) {
  S.data = d; S.etag = etag; S.base = S.demo ? null : JSON.parse(JSON.stringify(d));
  if (!S.demo) { ls.set('by.cache', JSON.stringify(d)); ls.set('by.etag', etag || ''); }
  S.memo = null;
  if (S.phase !== 'ready') S.phase = 'ready';
  render();
}
async function pull(force) {
  if (S.demo || !tokens) return;
  try {
    if (!force && S.etag) {
      const meta = await (await g(item(CFG.dataFile) + '?$select=eTag')).json();
      if (meta.eTag === S.etag) { setSync('ok'); return; }
    }
    setSync('busy', 'Syncing…');
    const remote = await loadRemote();
    if (!remote) { S.phase = 'setup'; render(); return; }
    applyData(remote.data, remote.etag); setSync('ok');
    ensureMonthly();
  } catch (e) {
    if (e instanceof AuthError) { if (trySilent()) return; S.phase = 'needsignin'; render(); return; }
    setSync('bad', e.status === 0 ? 'Offline — showing the last synced copy' : e.message);
  }
}
function setSync(state, msg = '') {
  S.sync = { state, msg, at: state === 'ok' ? new Date() : S.sync.at };
  const el = $('#syncline'); if (el) el.innerHTML = syncLine();
}
function syncLine() {
  if (S.demo) return `<span class="dot" style="background:var(--info)"></span>Sample data on this device only — nothing is saved to OneDrive`;
  const s = S.sync;
  if (s.state === 'busy') return `<span class="dot busy"></span>${esc(s.msg)}`;
  if (s.state === 'bad') return `<span class="dot bad"></span>${esc(s.msg)}`;
  return `<span class="dot"></span>Saved in OneDrive › ${esc(locShort())}${s.at ? ' · ' + timeStr(s.at) : ''}`;
}

/* ---------------- bills: upload, view, remove ---------------- */
async function loadBitmap(file) {
  try { return await createImageBitmap(file, { imageOrientation: 'from-image' }); }
  catch {
    const url = URL.createObjectURL(file);
    try { const img = new Image(); img.src = url; await img.decode(); return img; } finally { setTimeout(() => URL.revokeObjectURL(url), 2000); }
  }
}
async function resizeImage(file, max, q) {
  const bmp = await loadBitmap(file);
  const w = bmp.width || bmp.naturalWidth, h = bmp.height || bmp.naturalHeight;
  const sc = Math.min(1, max / Math.max(w, h));
  const c = document.createElement('canvas'); c.width = Math.round(w * sc); c.height = Math.round(h * sc);
  c.getContext('2d').drawImage(bmp, 0, 0, c.width, c.height);
  return { blob: await new Promise(r => c.toBlob(r, 'image/jpeg', q)), dataUrl: c.toDataURL('image/jpeg', q) };
}
/** Small preview (kept in billy.json) so lists show the bill straight away, even offline. */
async function makeThumb(file) {
  const bmp = await loadBitmap(file);
  const w = bmp.width || bmp.naturalWidth, h = bmp.height || bmp.naturalHeight, T = 240, sc = T / Math.max(w, h);
  const c = document.createElement('canvas'); c.width = Math.round(w * sc); c.height = Math.round(h * sc);
  c.getContext('2d').drawImage(bmp, 0, 0, c.width, c.height);
  return c.toDataURL('image/jpeg', 0.7);
}
async function uploadBlob(rel, blob) {
  if (blob.size <= 4 * 1024 * 1024) {
    await g(item(rel) + ':/content?@microsoft.graph.conflictBehavior=replace',
      { method: 'PUT', headers: { 'Content-Type': blob.type || 'application/octet-stream' }, body: blob });
    return;
  }
  const s = await (await g(item(rel) + ':/createUploadSession', { method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'replace' } }) })).json();
  const chunk = 320 * 1024 * 16;
  for (let start = 0; start < blob.size; start += chunk) {
    const end = Math.min(start + chunk, blob.size);
    const r = await fetch(s.uploadUrl, { method: 'PUT', headers: { 'Content-Range': `bytes ${start}-${end - 1}/${blob.size}` }, body: blob.slice(start, end) });
    if (!r.ok) throw new HttpError(r.status, 'Upload failed. Please try again.');
  }
}
const cleanName = s => String(s || '').replace(/[\\/:*?"<>|#%{}~&]+/g, ' ').replace(/\s+/g, ' ').trim();
const isPdfFile = f => /pdf/i.test(f.type) || /\.pdf$/i.test(f.name || '');
/** Bill or receipt → Bills/YYYY-MM/ in OneDrive. Photos become JPG (max 2400px). */
async function saveBill(file, x) {
  let blob = file, ext = (/\.(\w{1,5})$/.exec(file.name || '')?.[1] || (file.type.split('/')[1] || 'bin')).toLowerCase(), thumb = '';
  if (/^image\//.test(file.type) || /\.(heic|heif|jpe?g|png|webp)$/i.test(file.name || '')) {
    try { blob = (await resizeImage(file, 2400, 0.85)).blob; ext = 'jpg'; thumb = await makeThumb(file); } catch {}
  }
  const code = Math.random().toString(16).slice(2, 6).toUpperCase();
  const name = [x.date, cleanName(projectText(x)), cleanName(x.reason).slice(0, 50), x.amount ? String(x.amount) : '', code].filter(Boolean).join(' ').slice(0, 150) + '.' + ext;
  const file_ = `${mk(x.date) || thisMonth()}/${name}`;
  if (!S.demo) await uploadBlob('Bills/' + file_, blob);
  return { id: uuid(), file: file_, name: file.name || name, ext, thumb, addedAt: isoNow() };
}
async function downloadUrl(rel) {
  const m = await (await g(item(rel))).json();
  const u = m['@microsoft.graph.downloadUrl'];
  if (!u) throw new Error('OneDrive did not return the file.');
  return u;
}
async function fileBlob(rel) {
  try { const r = await fetch(await downloadUrl(rel), { cache: 'no-store' }); if (r.ok) return await r.blob(); } catch (e) { if (e instanceof HttpError && e.status !== 0) throw e; }
  return await (await g(item(rel) + ':/content')).blob();
}
async function billBlob(b) {
  if (S.demo) { const r = await fetch(b.thumb); return r.blob(); }
  return fileBlob('Bills/' + b.file);
}
async function moveToRemoved(rel) {
  if (S.demo) return;
  try {
    try { await g(`/drives/${LOC.driveId}/items/${LOC.id}/children`, { method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: 'Removed', folder: {}, '@microsoft.graph.conflictBehavior': 'fail' }) }); }
    catch (e) { if (e.status !== 409) throw e; }
    const removed = await (await g(item('Removed') + '?$select=id')).json();
    const nm = rel.split('/').pop();
    await g(item(rel), { method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ parentReference: { id: removed.id }, name: `${todayStr()} ${nm}` }) });
  } catch (e) { console.warn('remove file', e); }
}
const PDFJS = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js', PDFJS_W = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
async function renderPdf(blob, box, url) {
  try {
    if (!window.pdfjsLib) await new Promise((ok, no) => { const sc = document.createElement('script'); sc.src = PDFJS; sc.onload = ok; sc.onerror = () => no(new Error('pdf')); document.head.appendChild(sc); });
    if (!pdfjsLib.GlobalWorkerOptions.workerSrc) {
      try { pdfjsLib.GlobalWorkerOptions.workerSrc = URL.createObjectURL(new Blob([await (await fetch(PDFJS_W)).text()], { type: 'text/javascript' })); }
      catch { pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS_W; }
    }
    const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(await blob.arrayBuffer()) }).promise;
    box.textContent = '';
    const w = Math.min(box.clientWidth - 20, 900), dpr = Math.min(window.devicePixelRatio || 1, 3);
    for (let n = 1; n <= pdf.numPages; n++) {
      const page = await pdf.getPage(n), vp0 = page.getViewport({ scale: 1 }), sc = (w * dpr) / vp0.width, vp = page.getViewport({ scale: sc });
      const c = document.createElement('canvas'); c.width = Math.round(vp.width); c.height = Math.round(vp.height);
      c.style.cssText = `width:${w}px;height:auto;max-width:100%;display:block;margin:0 auto 12px;background:#fff;border-radius:6px`;
      box.appendChild(c); await page.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
    }
  } catch (e) {
    box.innerHTML = `<iframe title="Document" src="${url}" style="width:100%;height:100%;border:0;background:#fff"></iframe>`;
  }
}
/** Shows an image or PDF full screen. */
function viewBlob(blob, name) {
  const pdf = blob.type === 'application/pdf' || /\.pdf$/i.test(name || '');
  const url = URL.createObjectURL(blob);
  const v = document.createElement('div'); v.className = 'viewer';
  const closeBtn = '<button aria-label="Close">✕</button>';
  const link = `<a href="${url}" target="_blank" rel="noopener" ${name ? `download="${esc(name)}"` : ''} style="position:absolute;left:14px;top:calc(var(--safet) + 12px);padding:10px 14px;border-radius:20px;background:rgba(255,255,255,.2);color:#fff;font-weight:600;text-decoration:none;z-index:2">${pdf ? 'Open / save' : 'Save'}</a>`;
  v.innerHTML = !pdf ? `<img alt="" src="${url}">${link}${closeBtn}` : `<div class="pdfpages" style="position:absolute;inset:0;overflow:auto;-webkit-overflow-scrolling:touch;padding:calc(var(--safet) + 64px) 10px 24px;text-align:center;color:#fff">Loading…</div>${link}${closeBtn}`;
  $('button', v).addEventListener('click', () => { v.remove(); setTimeout(() => URL.revokeObjectURL(url), 1000); });
  if (!pdf) v.addEventListener('click', e => { if (e.target.tagName === 'IMG') { v.remove(); } });
  document.body.appendChild(v);
  if (pdf) renderPdf(blob, $('.pdfpages', v), url);
}
async function openBill(b) {
  toast('Opening bill…');
  try {
    let blob = await billBlob(b);
    const ext = (b.ext || b.file.split('.').pop() || '').toLowerCase();
    blob = new Blob([blob], { type: ext === 'pdf' ? 'application/pdf' : 'image/' + (ext === 'jpg' ? 'jpeg' : ext) });
    document.querySelector('.toast')?.remove();
    viewBlob(blob, b.name);
  } catch (e) { toast(e.message); }
}

/* ---------------- exchange rates (free, no key) ---------------- */
async function fetchJSON(url, ms = 9000) {
  const c = new AbortController(); const t = setTimeout(() => c.abort(), ms);
  try { const r = await fetch(url, { signal: c.signal, cache: 'no-store' }); if (!r.ok) throw new Error('HTTP ' + r.status); return await r.json(); }
  finally { clearTimeout(t); }
}
/** 1 unit of `cur` in the base currency on a day (fawazahmed0 currency-api). */
async function fxRate(cur, day) {
  const b = base().toLowerCase(), c = cur.toLowerCase();
  if (c === b) return 1;
  if (b === 'aed' && c === 'usd') return 3.6725;
  const key = `${c}-${b}-${day}`, cache = lsJSON('by.fx', {});
  if (cache[key]) return cache[key];
  const d = !day || day >= todayStr() ? 'latest' : day;
  const one = async u => { const j = await fetchJSON(u, 7000); const v = j[c]?.[b]; if (v) return v; throw 0; };
  try {
    const v = await Promise.any([`https://cdn.jsdelivr.net/npm/@fawazahmed0/currency-api@${d}/v1/currencies/${c}.min.json`, `https://${d}.currency-api.pages.dev/v1/currencies/${c}.min.json`].map(one));
    const r = Math.round(v * 1e6) / 1e6; cache[key] = r; ls.set('by.fx', JSON.stringify(cache)); return r;
  } catch { return null; }
}

/* ---------------- lookups and totals ---------------- */
const st = () => S.data?.settings || {};
const projects = () => S.data?.projects || [];
const projectById = id => projects().find(p => p.id === id);
const claimById = id => (S.data?.claims || []).find(c => c.id === id);
const catIcon = name => (S.data?.categories || CATS).find(c => c.name === name)?.icon || '🧾';
const projectText = x => projectById(x.projectId)?.name || x.label || x.category || '';
const projColor = p => PROJECT_COLORS[Math.max(0, projects().indexOf(p)) % PROJECT_COLORS.length];
/** confirm (monthly item not confirmed) | open (not claimed) | claimed | sent | paid */
function status(x) {
  if (x.pending) return 'confirm';
  if (!x.claimId) return 'open';
  const c = claimById(x.claimId);
  return !c ? 'open' : c.paidAt ? 'paid' : c.sentAt ? 'sent' : 'claimed';
}
const STATUS = {
  confirm: ['To confirm', 'warn'], open: ['To claim', 'acc'], claimed: ['In a claim', 'info'], sent: ['Sent', 'info'], paid: ['Paid', 'good'],
};
const live = () => (S.data?.expenses || []).filter(x => !x.imported);
const openItems = () => live().filter(x => status(x) === 'open');
const sum = list => round2(list.reduce((a, x) => a + (+x.amount || 0), 0));
const byDate = (a, b) => (b.date || '').localeCompare(a.date || '') || (b.createdAt || '').localeCompare(a.createdAt || '');
function recentProjects() {
  const last = new Map();
  for (const x of S.data.expenses) if (x.projectId && (!last.has(x.projectId) || last.get(x.projectId) < x.date)) last.set(x.projectId, x.date);
  return projects().filter(p => !p.archived).sort((a, b) => (last.get(b.id) || b.createdAt || '').localeCompare(last.get(a.id) || a.createdAt || ''));
}
function memory() {
  if (S.memo) return S.memo;
  S.memo = SCAN.buildMemory(S.data.expenses.filter(x => !x.pending).map(x => ({ vendor: x.vendor, category: x.category, projectName: projectById(x.projectId)?.name || '', reason: x.reason, date: x.date })));
  return S.memo;
}
const aiKey = () => ls.get('by.aiKey') || ls.get('mg.aiKey') || '';

/** Makes this month's copy of every monthly item (once), waiting for you to confirm the amount. */
function ensureMonthly() {
  if (!S.data) return;
  const m = thisMonth(), need = S.data.monthly.filter(t => t.active !== false && (t.startMonth || '') <= m &&
    !S.data.expenses.some(x => x.monthlyId === t.id && mk(x.date) === m));
  if (!need.length) return;
  mutate(d => {
    for (const t of d.monthly) {
      if (t.active === false || (t.startMonth || '') > m || d.expenses.some(x => x.monthlyId === t.id && mk(x.date) === m)) continue;
      d.expenses.push({ id: uuid(), date: `${m}-01`, projectId: t.projectId || null, label: t.label || '', category: t.category, reason: t.reason, vendor: t.vendor || '',
        amount: t.amount, currency: d.settings.currency, bills: [], claimId: null, monthlyId: t.id, pending: true, createdAt: isoNow() });
    }
  });
}

/* ---------------- UI basics ---------------- */
let toastTimer;
function toast(msg) {
  document.querySelector('.toast')?.remove();
  const t = document.createElement('div'); t.className = 'toast'; t.textContent = msg; document.body.appendChild(t);
  clearTimeout(toastTimer); toastTimer = setTimeout(() => t.remove(), 3600);
}
function progress(msg) {
  let p = $('.progress');
  if (msg == null) { p?.remove(); return; }
  if (!p) { p = document.createElement('div'); p.className = 'progress'; p.innerHTML = '<div><span class="spin"></span><p style="margin:12px 0 0"></p></div>'; document.body.appendChild(p); }
  $('p', p).textContent = msg;
}
function applyTheme() {
  const t = ls.get('by.theme') || 'auto';
  if (t === 'auto') document.documentElement.removeAttribute('data-theme'); else document.documentElement.dataset.theme = t;
}
const isDark = () => { const t = ls.get('by.theme') || 'auto'; return t === 'dark' || (t === 'auto' && matchMedia('(prefers-color-scheme: dark)').matches); };
matchMedia('(prefers-color-scheme: dark)').addEventListener?.('change', () => { if (S.phase === 'ready') render(); });
const SV = p => `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">${p}</svg>`;
const ICON = {
  home: SV('<path d="M3 10.5 12 3l9 7.5V20a1 1 0 0 1-1 1h-5v-6h-6v6H4a1 1 0 0 1-1-1z"/>'),
  expenses: SV('<path d="M6 3h12v18l-3-2-3 2-3-2-3 2z"/><path d="M9 8h6M9 12h6M9 16h3"/>'),
  projects: SV('<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>'),
  claims: SV('<path d="m22 2-7 20-4-9-9-4z"/><path d="M22 2 11 13"/>'),
  gear: SV('<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"/>'),
};
const TABS = [['home', 'Home'], ['expenses', 'Expenses'], ['projects', 'Projects'], ['claims', 'Claims']];

let _rq = 0, _last = '';
function render() { if (_rq) return; _rq = requestAnimationFrame(() => { _rq = 0; renderNow(); }); }
function renderNow() {
  const app = $('#app'); if (S.phase !== 'ready') _last = '';
  const sy = window.scrollY;
  if (S.phase === 'loading') { app.innerHTML = `<div class="welcome"><img src="icons/icon-512.png" alt=""><h1>Billy</h1><p><span class="spin"></span></p></div>`; return; }
  if (S.phase === 'welcome' || S.phase === 'needsignin') { app.innerHTML = welcomeView(); return; }
  if (S.phase === 'folder') { app.innerHTML = folderView(); folderBrowser($('#fb')); return; }
  if (S.phase === 'setup') { app.innerHTML = setupView(); return; }
  const title = { home: 'Billy', expenses: 'Expenses', projects: 'Projects', claims: 'Claims', settings: 'Settings' }[S.tab] || 'Billy';
  const body = { expenses: expensesView, projects: projectsView, claims: claimsView, settings: settingsView }[S.tab]?.() ?? homeView();
  const html = `
    <header class="top"><div class="topbar"><h1>${title}</h1>
      <button class="iconbtn" data-a="theme-toggle" title="Light / dark">${isDark() ? '☀️' : '🌙'}</button>
      <button class="iconbtn" data-a="settings" title="Settings" style="${S.tab === 'settings' ? 'color:var(--acc);border-color:var(--acc)' : ''}"><span style="width:19px;height:19px;display:block">${ICON.gear}</span></button>
    </div><div class="sync" id="syncline">${syncLine()}</div></header>
    <main>${body}</main>
    ${S.tab === 'settings' || S.tab === 'claims' ? '' : `<button class="fab" data-a="add"><span style="font-size:19px;line-height:1">＋</span> Add expense</button>`}
    <nav class="tabs">${TABS.map(([k, l]) => `<button data-tab="${k}" class="${S.tab === k ? 'on' : ''}">${ICON[k]}<span>${l}</span></button>`).join('')}</nav>`;
  if (html !== _last || !app.firstElementChild) { app.innerHTML = html; _last = html; window.scrollTo(0, sy); }
}

/* ---------------- welcome / folder / setup ---------------- */
function welcomeView() {
  const again = S.phase === 'needsignin';
  return `<div class="welcome">
    <img src="icons/icon-512.png" alt="">
    <h1>Billy</h1>
    <p>${again ? 'Your Microsoft sign-in has expired. Sign in again to continue — your data is safe in OneDrive.' :
      'Snap a bill and Billy works out what it was for. Every expense, by project, ready as a monthly claim for accounts.'}</p>
    <div class="feat"><div>📸 Reads bills &amp; tickets</div><div>📁 By project</div><div>🔁 Monthly items pre-filled</div><div>📤 Excel + bills PDF for accounts</div></div>
    <div class="acts">
      <button class="btn" data-a="signin">Sign in with Microsoft</button>
      ${again ? `<button class="btn sec" data-a="signin-other">Use a different account</button>` : `<button class="btn sec" data-a="demo">Look around with sample data</button>`}
    </div>
    <p class="small" style="opacity:.7;margin-top:18px">Your data stays in your OneDrive. Nothing is stored anywhere else.</p>
  </div>`;
}
function folderView() {
  return `<header class="top"><div class="topbar"><h1>Where to keep it</h1></div></header>
    <main><div class="note">Choose (or create) a folder in your OneDrive — for example <b>Documents › 02 ACCOUNTS › OFFICE › DOC › EXPENCE › Billy</b>. Your expenses, bills and claims will all be kept there.</div>
    <div id="fb"></div></main>`;
}
function setupView() {
  const su = (S.setup ||= { name: ls.get('by.displayName') || '', accountant: 'Adnan Ahmad', email: 'accounts@choppershoot.com', company: 'Choppershoot', fromEmail: 'shaheer@choppershoot.com', projects: '' });
  return `<header class="top"><div class="topbar"><h1>Welcome to Billy</h1></div></header>
    <main>
      <div class="card"><h3>You</h3>
        <label class="field"><span>Your name, as on the claim sheet</span><input data-su="name" value="${esc(su.name)}" placeholder="e.g. Shaheer V"></label>
        <div class="small faint" style="margin-top:6px">The sheet will be called “${esc(su.name || 'Your name')} ${monthLabel(thisMonth())} Expenses”.</div>
      </div>
      <div class="card"><h3>Accounts</h3>
        <label class="field"><span>Accountant</span><input data-su="accountant" value="${esc(su.accountant)}"></label>
        <label class="field"><span>Send claims to</span><input data-su="email" type="email" value="${esc(su.email)}"></label>
        <label class="field"><span>Send from (your Gmail)</span><input data-su="fromEmail" type="email" value="${esc(su.fromEmail)}"></label>
        <label class="field"><span>Company</span><input data-su="company" value="${esc(su.company)}"></label>
      </div>
      <div class="card"><h3>Projects you're working on <span class="grow"></span><span class="faint" style="text-transform:none;letter-spacing:0;font-weight:500">optional</span></h3>
        <label class="field"><span>One per line</span><textarea data-su="projects" rows="4" placeholder="DHRE Palm Jebel Ali&#10;Emaar brand film">${esc(su.projects)}</textarea></label>
        <div class="small faint" style="margin-top:6px">Next, Billy can read your past expense sheets in OneDrive to learn the rest.</div>
      </div>
      <button class="btn" data-a="setup-done" style="margin-top:18px">Start</button>
      <div class="small faint" style="text-align:center;margin-top:12px">Billy ${CFG.version} · data in ${esc(locFull())}</div>
    </main>`;
}

/* ---------------- Home ---------------- */
function expenseRow(x, opts = {}) {
  const s = status(x), [sl, sc] = STATUS[s], p = projectById(x.projectId);
  const thumb = x.bills.find(b => b.thumb)?.thumb;
  const sub = [p?.name || x.label, x.category].filter(Boolean).join(' · ');
  return `<button class="row" data-a="open-exp" data-id="${x.id}">
    ${thumb ? `<span class="thumb" style="width:42px;height:42px;border-radius:10px;background-image:url('${tu(thumb)}')"></span>` : `<span class="ic cat">${catIcon(x.category)}</span>`}
    <span class="main"><span class="t">${esc(x.reason || x.vendor || x.category)}</span><span class="s">${esc(sub)}${opts.date ? ' · ' + shortDate(x.date) : ''}</span></span>
    <span class="r"><span class="a amt">${moneyS(+x.amount || 0)}</span><span class="b" style="display:flex;gap:4px;justify-content:flex-end;margin-top:2px">${x.bills.length ? '' : s === 'open' || s === 'confirm' ? '<span class="chip bad">No bill</span>' : ''}${s === 'open' ? '' : `<span class="chip ${sc}">${sl}</span>`}</span></span>
  </button>`;
}
function homeView() {
  const d = S.data, m = thisMonth(), open = openItems(), noBill = open.filter(x => !x.bills.length).length;
  const pending = live().filter(x => x.pending).sort(byDate);
  const monthSpent = sum(live().filter(x => mk(x.date) === m));
  const waiting = sum(live().filter(x => ['claimed', 'sent'].includes(status(x))));
  const yr = String(new Date().getFullYear());
  const paidYear = sum(live().filter(x => status(x) === 'paid' && (claimById(x.claimId)?.paidAt || '').startsWith(yr)));
  const byCat = {}; for (const x of live().filter(x => mk(x.date) === m)) byCat[x.category] = (byCat[x.category] || 0) + (+x.amount || 0);
  const cats = Object.entries(byCat).sort((a, b) => b[1] - a[1]), catMax = cats[0]?.[1] || 1;
  const learnedAny = d.learned.length || d.expenses.some(x => x.imported);
  return `
    <div class="hero">
      <div class="lbl">Ready to claim</div>
      <div class="big num">${money(sum(open))}</div>
      <div class="sub">${open.length ? `${plural(open.length, 'expense')}${noBill ? ` · <b>${noBill} without a bill</b>` : ' · all bills attached'}` : 'Nothing waiting — add an expense or snap a bill.'}${pending.length ? ` · ${pending.length} monthly to confirm` : ''}</div>
      <div class="stats">
        <div class="stat"><div class="k">${MONTH_NAMES[+m.slice(5) - 1]}</div><div class="v num">${moneyS(monthSpent)}</div></div>
        <div class="stat"><div class="k">Awaiting pay</div><div class="v num">${moneyS(waiting)}</div></div>
        <div class="stat"><div class="k">Paid ${yr}</div><div class="v num">${moneyS(paidYear)}</div></div>
      </div>
      ${open.length ? `<button class="btn cta" data-a="send-open">📧 Send ${esc(monthLabel(periodOf(open).month).split(' ')[0])} expenses to ${esc(firstName())}</button>
      <button class="linkbtn" data-a="new-claim" style="color:#FFC2A8;display:block;margin:8px auto 0">Choose which ones first</button>` : ''}
    </div>
    ${pending.length ? `<div class="card"><h3>🔁 Monthly items · ${monthLabel(m)}</h3>
      <div class="small muted" style="margin:-4px 0 8px">Filled in with last month's amounts. Tap one to change the amount or attach the bill.</div>
      <div class="list" style="margin-top:0">${pending.map(x => `<div class="row" style="padding:10px 12px">
        <span class="ic cat">${catIcon(x.category)}</span>
        <button class="main" data-a="open-exp" data-id="${x.id}" style="border:0;background:none;text-align:left;padding:0"><span class="t">${esc(x.label || x.reason)}</span><span class="s">${esc(x.reason)}</span></button>
        <span class="amt">${moneyS(+x.amount || 0)}</span>
        <button class="chip good" data-a="confirm" data-id="${x.id}" style="border:0;padding:7px 11px;font-size:13px">✓</button></div>`).join('')}</div>
      ${pending.length > 1 ? `<button class="btn sec" data-a="confirm-all" style="margin-top:10px">✓ Confirm all ${pending.length} · ${moneyS(sum(pending))}</button>` : ''}</div>` : ''}
    ${!learnedAny && !S.demo && ls.get('by.learnSkip') !== '1' ? `<div class="card" style="border-color:var(--acc)"><h3>🧠 Learn from your past sheets</h3>
      <div class="small muted">Billy can read the monthly expense sheets you sent before (in ${esc(PAST_FOLDER.split('/').slice(-2).join(' › '))}) to learn your projects, categories and the items that repeat every month.</div>
      <div class="actions"><button class="btn" data-a="learn">Read past sheets</button><button class="btn sec" data-a="learn-skip">Not now</button></div></div>` : ''}
    ${cats.length ? `<div class="card"><h3>${monthLabel(m)} by category</h3>
      ${cats.slice(0, 6).map(([c, v]) => `<div style="margin-top:10px"><div style="display:flex;justify-content:space-between;font-size:14px"><span>${catIcon(c)} ${esc(c)}</span><b class="num">${moneyS(v)}</b></div><div class="pbar"><i style="width:${Math.max(3, v / catMax * 100)}%"></i></div></div>`).join('')}</div>` : ''}
    <div class="card" style="padding:0;overflow:hidden"><h3 style="padding:14px 14px 0">Latest<span class="grow"></span><button class="linkbtn" data-tab="expenses" style="text-transform:none;letter-spacing:0">See all</button></h3>
      <div class="list" style="border:0;border-radius:0;margin-top:4px">${live().filter(x => !x.pending).sort(byDate).slice(0, 7).map(x => expenseRow(x, { date: true })).join('') || '<div class="empty">No expenses yet.<br>Tap <b>Add expense</b> and snap a bill.</div>'}</div></div>`;
}

/* ---------------- Expenses ---------------- */
function expensesView() {
  const all = live(), months = [...new Set([thisMonth(), ...all.map(x => mk(x.date))])].filter(Boolean).sort().reverse();
  const f = S.f; if (f.month && !months.includes(f.month)) f.month = '';
  const q = f.q.trim().toLowerCase();
  let list = all.filter(x => (!f.month || mk(x.date) === f.month) &&
    (f.status === 'all' || (f.status === 'open' ? ['open', 'confirm'].includes(status(x)) : f.status === 'claimed' ? ['claimed', 'sent'].includes(status(x)) : status(x) === 'paid')) &&
    (!q || [x.reason, x.vendor, x.label, x.category, projectById(x.projectId)?.name, String(x.amount)].join(' ').toLowerCase().includes(q))).sort(byDate);
  const groups = []; for (const x of list) { const k = x.date; let gp = groups.find(z => z.k === k); if (!gp) groups.push(gp = { k, items: [] }); gp.items.push(x); }
  return `
    <div class="months">${[['', 'All'], ...months.map(k => [k, monthShort(k)])].map(([k, l]) => `<button data-a="f-month" data-k="${k}" class="${f.month === k ? 'on' : ''}">${l}</button>`).join('')}</div>
    <div class="filters">
      <div class="search"><span class="faint">⌕</span><input id="q" type="search" placeholder="Search reason, vendor, project, amount" value="${esc(f.q)}" autocomplete="off"></div>
      <div class="seg">${[['all', 'All'], ['open', 'To claim'], ['claimed', 'Claimed'], ['paid', 'Paid']].map(([k, l]) => `<button data-a="f-status" data-k="${k}" class="${f.status === k ? 'on' : ''}">${l}</button>`).join('')}</div>
    </div>
    <div style="display:flex;justify-content:space-between;margin:14px 2px 0;font-size:13.5px"><span class="muted">${plural(list.length, 'expense')}</span><b class="num">${money(sum(list))}</b></div>
    ${groups.length ? `<div class="list">${groups.map(gp => `<div class="gh"><span>${dayHead(gp.k)}</span><span class="num">${moneyS(sum(gp.items))}</span></div>${gp.items.map(x => expenseRow(x)).join('')}`).join('')}</div>`
      : `<div class="empty">No expenses here.</div>`}`;
}

/* ---------------- Projects ---------------- */
function projectsView() {
  const showArch = S.showArch;
  const ps = projects().filter(p => showArch || !p.archived);
  const stats = p => { const l = live().filter(x => x.projectId === p.id); return { n: l.length, total: sum(l), open: sum(l.filter(x => status(x) === 'open')), last: l.map(x => x.date).sort().pop() }; };
  const rows = ps.map(p => ({ p, s: stats(p) })).sort((a, b) => (b.s.last || b.p.createdAt || '').localeCompare(a.s.last || a.p.createdAt || ''));
  const loose = live().filter(x => !x.projectId);
  return `
    <div class="list">${rows.map(({ p, s }) => `<button class="row" data-a="open-proj" data-id="${p.id}">
      <span class="av" style="--c:${projColor(p)};width:38px;height:38px;font-size:14px">${esc(initials(p.name))}</span>
      <span class="main"><span class="t">${esc(p.name)}${p.archived ? ' <span class="chip">Archived</span>' : ''}</span><span class="s">${s.n ? `${plural(s.n, 'expense')} · last ${shortDate(s.last)}` : 'No expenses yet'}</span></span>
      <span class="r"><span class="a amt">${moneyS(s.total)}</span>${s.open ? `<span class="b"><span class="chip acc">${moneyS(s.open)} to claim</span></span>` : ''}</span></button>`).join('') || '<div class="empty">No projects yet.</div>'}
      ${loose.length ? `<button class="row" data-a="open-proj" data-id="">
        <span class="ic">🗂️</span><span class="main"><span class="t">No project</span><span class="s">Monthly items, office and personal allowances</span></span>
        <span class="r"><span class="a amt">${moneyS(sum(loose))}</span></span></button>` : ''}</div>
    <div class="actions"><button class="btn sec" data-a="new-proj">＋ New project</button><button class="btn sec" data-a="toggle-arch">${showArch ? 'Hide archived' : 'Show archived'}</button></div>`;
}
function openProject(id) {
  const p = projectById(id), list = live().filter(x => id ? x.projectId === id : !x.projectId).sort(byDate);
  const byCat = {}; for (const x of list) byCat[x.category] = (byCat[x.category] || 0) + (+x.amount || 0);
  openSheet(p ? p.name : 'No project', `
    <div class="split" style="margin-top:12px">
      <div class="box"><div class="k">Total</div><div class="v num">${money(sum(list))}</div></div>
      <div class="box"><div class="k">To claim</div><div class="v num">${money(sum(list.filter(x => status(x) === 'open')))}</div></div>
    </div>
    ${Object.keys(byCat).length ? `<div class="legend" style="margin-top:12px">${Object.entries(byCat).sort((a, b) => b[1] - a[1]).map(([c, v]) => `<span>${catIcon(c)} ${esc(c)} <b>${moneyS(v)}</b></span>`).join('')}</div>` : ''}
    <div class="list">${list.map(x => expenseRow(x, { date: true })).join('') || '<div class="empty">No expenses yet.</div>'}</div>
    ${p ? `<div class="actions">
      <button class="btn sec" data-a="proj-excel" data-id="${p.id}">⬇︎ Excel of project</button>
      <button class="btn sec" data-a="proj-rename" data-id="${p.id}">Rename</button>
      <button class="btn sec" data-a="proj-archive" data-id="${p.id}">${p.archived ? 'Unarchive' : 'Archive'}</button>
      <button class="btn danger" data-a="proj-delete" data-id="${p.id}">Delete</button></div>` : ''}`);
}

/* ---------------- Claims ---------------- */
function claimStatus(c) { return c.paidAt ? ['Paid ' + shortDate(c.paidAt), 'good', 3] : c.sentAt ? ['Sent ' + shortDate(c.sentAt), 'info', 2] : ['Not sent yet', 'warn', 1]; }
function claimsView() {
  const open = openItems(), pending = live().filter(x => x.pending);
  const cs = (S.data.claims || []).filter(c => !c.imported).sort((a, b) => (b.createdAt || '').localeCompare(a.createdAt || ''));
  const past = (S.data.claims || []).filter(c => c.imported).sort((a, b) => (b.month || '').localeCompare(a.month || ''));
  return `
    <div class="card" style="border-color:var(--acc)"><h3>Ready to claim</h3>
      <div class="claimcard" style="margin:0;border:0;padding:0;background:none"><div class="top"><span class="sum">${money(sum(open))}</span><span class="muted small">${plural(open.length, 'expense')}</span></div></div>
      ${pending.length ? `<div class="note warn">${plural(pending.length, 'monthly item')} still to confirm — they're left out until you confirm them.</div>` : ''}
      <button class="btn" data-a="send-open" style="margin-top:12px" ${open.length ? '' : 'disabled'}>📧 Send to ${esc(firstName())}</button>
      <button class="btn sec" data-a="new-claim" style="margin-top:8px" ${open.length ? '' : 'disabled'}>Choose which ones first</button>
    </div>
    ${cs.map(c => { const [l, k, n] = claimStatus(c); return `<button class="claimcard" data-a="open-claim" data-id="${c.id}">
      <div class="top"><span class="ttl">${esc(c.title)}</span><span class="chip ${k}">${l}</span></div>
      <div class="top" style="margin-top:6px"><span class="sum">${money(c.total)}</span><span class="muted small">${plural(c.expenseIds.length, 'expense')}</span></div>
      <div class="steps"><span class="on"></span><span class="${n >= 2 ? 'on' : ''}"></span><span class="${n >= 3 ? 'on' : ''}"></span></div>
      ${c.stale && !c.sentAt ? '<div class="small" style="color:var(--warn);margin-top:6px">Changed since the files were made — open to update them.</div>' : ''}
    </button>`; }).join('') || (open.length ? '' : '<div class="empty">No claims yet.</div>')}
    ${past.length ? `<div class="card" style="padding:0;overflow:hidden"><h3 style="padding:14px 14px 0">Past sheets Billy learned from</h3><div class="list" style="border:0;border-radius:0;margin-top:4px">
      ${past.map(c => `<button class="row" data-a="open-claim" data-id="${c.id}"><span class="ic">📗</span><span class="main"><span class="t">${esc(c.title)}</span><span class="s">${plural(c.expenseIds.length, 'item')}</span></span><span class="amt">${moneyS(c.total)}</span></button>`).join('')}</div></div>` : ''}`;
}

/* ---------------- Settings ---------------- */
function settingsView() {
  const s = st(), mon = S.data.monthly.slice().sort((a, b) => cmpText(a.label || a.reason, b.label || b.reason));
  const key = aiKey(), keyFrom = ls.get('by.aiKey') ? 'Billy' : ls.get('mg.aiKey') ? 'My Gold' : '';
  return `
    <div class="card"><h3>You and accounts</h3>
      <label class="field"><span>Your name on the sheet</span><input data-set="name" value="${esc(s.name)}" placeholder="e.g. Shaheer V"></label>
      <div class="two"><label class="field"><span>Accountant</span><input data-set="accountant" value="${esc(s.accountant)}"></label>
      <label class="field"><span>Company</span><input data-set="company" value="${esc(s.company)}"></label></div>
      <label class="field"><span>Send claims to</span><input data-set="email" type="email" value="${esc(s.email)}"></label>
      <div class="two"><label class="field"><span>Shoot day rate (${esc(base())})</span><input data-set="dayRate" inputmode="decimal" value="${esc(s.dayRate)}"></label>
      <label class="field"><span>Currency</span><select data-set="currency">${CURRENCIES.map(c => `<option ${c === s.currency ? 'selected' : ''}>${c}</option>`).join('')}</select></label></div>
    </div>
    <div class="card"><h3>📧 Email to accounts</h3>
      <div class="two"><label class="field"><span>Send from (your Gmail)</span><input data-set="fromEmail" type="email" value="${esc(s.fromEmail)}" placeholder="shaheer@choppershoot.com"></label>
      <label class="field"><span>Cc (optional)</span><input data-set="cc" type="email" value="${esc(s.cc)}" placeholder="e.g. yourself"></label></div>
      <label class="field"><span>Google Client ID ${s.googleClientId ? '· ✓ set' : ''}</span><input data-set="googleClientId" value="${esc(s.googleClientId)}" placeholder="….apps.googleusercontent.com" autocapitalize="off" spellcheck="false"></label>
      <button class="linkbtn" data-a="gmail-help" style="margin-top:6px">How to get the Client ID (once, 5 minutes)</button>
    </div>
    <div class="card" style="padding:0;overflow:hidden"><h3 style="padding:14px 14px 0">🔁 Monthly items<span class="grow"></span><button class="linkbtn" data-a="monthly-new" style="text-transform:none;letter-spacing:0">＋ Add</button></h3>
      <div class="small muted" style="padding:0 14px">Billy adds these at the start of every month with the last amount, for you to confirm.</div>
      <div class="list" style="border:0;border-radius:0;margin-top:8px">${mon.map(t => `<button class="row" data-a="monthly-edit" data-id="${t.id}" style="${t.active === false ? 'opacity:.5' : ''}">
        <span class="ic cat">${catIcon(t.category)}</span><span class="main"><span class="t">${esc(t.label || t.reason)}</span><span class="s">${esc(t.reason)}${t.active === false ? ' · paused' : ''}</span></span><span class="amt">${moneyS(+t.amount || 0)}</span></button>`).join('') || '<div class="empty" style="padding:16px">None yet. Phone bill, rent, school fees and subscriptions go here.</div>'}</div></div>
    <div class="card"><h3>🧠 Smart bill reading</h3>
      <div class="small muted">Billy always reads bills on your phone. For sharper results — and better guesses of the project and reason — add a free Google Gemini key (aistudio.google.com). ${keyFrom === 'My Gold' ? '<b>Using the key you saved in My Gold.</b>' : ''} The key stays on this device.</div>
      <div class="field"><span>AI key ${keyFrom === 'Billy' ? '· ✓ saved' : ''}</span><input id="aiKey" type="password" autocomplete="off" autocapitalize="off" placeholder="${key ? '•••••••• (saved)' : 'Paste Gemini key'}"></div>
      <div style="display:flex;gap:8px;margin-top:8px"><button class="btn sec" data-a="ai-save">Save key</button>${keyFrom === 'Billy' ? '<button class="btn sec" data-a="ai-off">Remove</button>' : ''}</div>
      <button class="btn sec" data-a="learn" style="margin-top:12px">📗 Learn from past expense sheets</button>
      ${S.data.learned.length ? `<div class="small faint" style="margin-top:6px">Learned from ${plural(S.data.learned.length, 'sheet')}.</div>` : ''}
    </div>
    <div class="card"><h3>Categories</h3>
      <div class="pills">${S.data.categories.map(c => `<button data-a="cat-edit" data-k="${esc(c.name)}">${c.icon} ${esc(c.name)}</button>`).join('')}<button data-a="cat-new">＋ New</button></div></div>
    <div class="card"><h3>Look</h3><div class="seg">${[['auto', 'Automatic'], ['light', 'Light'], ['dark', 'Dark']].map(([k, l]) => `<button data-a="theme" data-k="${k}" class="${(ls.get('by.theme') || 'auto') === k ? 'on' : ''}">${l}</button>`).join('')}</div></div>
    <div class="card"><h3>Data</h3>
      <dl class="kv"><dt>Signed in as</dt><dd>${esc(S.demo ? 'Sample data' : S.account || '–')}</dd><dt>Folder</dt><dd style="font-weight:500">${esc(locFull())}</dd><dt>Expenses</dt><dd>${S.data.expenses.length}</dd></dl>
      <div class="actions"><button class="btn sec" data-a="export-all">⬇︎ Excel of everything</button>${S.demo ? '' : '<button class="btn sec" data-a="change-folder">Change folder</button>'}</div>
      <button class="btn danger" data-a="signout" style="margin-top:8px">${S.demo ? 'Leave sample data' : 'Sign out'}</button>
    </div>
    <div class="small faint" style="text-align:center;margin:16px 0">Billy ${CFG.version}</div>`;
}

/* ---------------- sheets ---------------- */
let sheetEl = null, scrimEl = null, sheetOnClose = null;
function openSheet(title, html, { onClose, right = '' } = {}) {
  closeSheet(true);
  scrimEl = document.createElement('div'); scrimEl.className = 'scrim';
  sheetEl = document.createElement('div'); sheetEl.className = 'sheet';
  sheetEl.innerHTML = `<div class="sh"><button class="linkbtn" data-a="close" style="font-weight:600">Close</button><h2>${esc(title)}</h2>${right}</div><div class="sb">${html}</div>`;
  document.body.append(scrimEl, sheetEl);
  scrimEl.addEventListener('click', () => closeSheet());
  // stopPropagation: the sheet redraws itself, so the page-level listeners must not see the same event again
  sheetEl.addEventListener('click', e => { e.stopPropagation(); onClick(e); });
  sheetEl.addEventListener('change', e => { e.stopPropagation(); onChange(e); });
  sheetEl.addEventListener('input', e => { e.stopPropagation(); onInput(e); });
  requestAnimationFrame(() => { scrimEl.classList.add('on'); sheetEl.classList.add('on'); });
  sheetOnClose = onClose; document.body.style.overflow = 'hidden';
  return sheetEl;
}
function closeSheet(instant) {
  if (!sheetEl) return;
  const s = sheetEl, c = scrimEl; sheetEl = scrimEl = null; document.body.style.overflow = '';
  if (instant) { s.remove(); c.remove(); } else { s.classList.remove('on'); c.classList.remove('on'); setTimeout(() => { s.remove(); c.remove(); }, 260); }
  const f = sheetOnClose; sheetOnClose = null; f?.();
}
const sheetBody = () => sheetEl?.querySelector('.sb');
function redrawSheet(html) { const sb = sheetBody(); if (!sb) return; const y = sb.scrollTop; sb.innerHTML = html; sb.scrollTop = y; }

/* ---- expense detail ---- */
function openExpense(id) {
  const x = S.data.expenses.find(e => e.id === id); if (!x) return;
  const s = status(x), [sl, sc] = STATUS[s], c = claimById(x.claimId), locked = s === 'sent' || s === 'paid';
  openSheet(x.reason || x.category, `
    <div style="text-align:center;margin-top:14px">
      <div style="font-size:34px">${catIcon(x.category)}</div>
      <div class="serif num" style="font-size:34px;font-weight:700;letter-spacing:-.02em">${money(+x.amount || 0)}</div>
      ${x.currency && x.currency !== base() && x.origAmount ? `<div class="small muted">${money(x.origAmount, x.currency)} at ${x.fx}</div>` : ''}
      <div style="margin-top:6px"><span class="chip ${sc}">${sl}</span> ${x.bills.length ? '' : '<span class="chip bad">No bill</span>'}</div>
    </div>
    <div class="card"><dl class="kv">
      <dt>Date</dt><dd>${shortDate(x.date)}</dd>
      <dt>Project</dt><dd>${esc(projectById(x.projectId)?.name || x.label || '—')}</dd>
      <dt>Category</dt><dd>${esc(x.category)}</dd>
      ${x.vendor ? `<dt>Vendor</dt><dd>${esc(x.vendor)}</dd>` : ''}
      ${x.days ? `<dt>Days</dt><dd>${x.days} × ${moneyS(+x.rate || 0)}</dd>` : ''}
      ${c ? `<dt>Claim</dt><dd>${esc(c.title)}</dd>` : ''}
      ${x.monthlyId ? `<dt>Repeats</dt><dd>Every month</dd>` : ''}
    </dl>${x.reason ? `<div style="margin-top:12px;font-size:15px">${esc(x.reason)}</div>` : ''}</div>
    <div class="card"><h3>Bills</h3><div class="bills">
      ${x.bills.map((b, i) => `<button class="g" data-a="view-bill" data-id="${x.id}" data-i="${i}" style="${b.thumb ? `background-image:url('${tu(b.thumb)}')` : ''}">${b.thumb ? '' : `<span><span class="pdf">📄</span>${esc((b.ext || 'PDF').toUpperCase())}</span>`}</button>`).join('')}
      ${locked ? '' : `<label class="g add">＋ Add bill<input type="file" accept="image/*,application/pdf,.pdf,.heic" multiple hidden data-addbill="${x.id}"></label>`}
    </div></div>
    ${x.pending ? `<button class="btn" data-a="confirm" data-id="${x.id}" style="margin-top:14px">✓ Confirm ${moneyS(+x.amount || 0)}</button>` : ''}
    <div class="actions">
      ${locked ? '' : `<button class="btn sec" data-a="edit-exp" data-id="${x.id}">Edit</button>`}
      <button class="btn sec" data-a="dup-exp" data-id="${x.id}">Duplicate</button>
      ${locked ? '' : `<button class="btn danger" data-a="del-exp" data-id="${x.id}">Delete</button>`}
    </div>
    ${locked ? `<div class="small muted" style="margin-top:10px">This expense was sent to accounts, so it can't be changed. Undo “Sent” on the claim to edit it.</div>` : ''}`);
}

/* ---- add / edit expense ---- */
let F = null;   // the form being edited
function newForm(from = {}) {
  const recent = recentProjects()[0];
  return { id: null, date: todayStr(), projectId: from.projectId ?? (S.f.lastProject ?? recent?.id ?? null), label: '', category: '', reason: '', vendor: '',
    amount: '', currency: base(), origAmount: '', fx: null, days: '', rate: st().dayRate || 50, bills: [], add: [], removed: [], touched: new Set(), scan: null, repeat: false, ...from };
}
function openForm(id, preset = {}) {
  const x = id ? S.data.expenses.find(e => e.id === id) : null;
  F = x ? newForm({ ...JSON.parse(JSON.stringify(x)), amount: x.currency !== base() ? '' : String(x.amount ?? ''), origAmount: x.origAmount ? String(x.origAmount) : '', days: x.days ? String(x.days) : '', rate: x.rate || st().dayRate || 50, touched: new Set(['all']), add: [], removed: [] })
        : newForm(preset);
  if (x && x.currency !== base()) F.amount = String(x.origAmount || '');
  if (x && x.currency === base()) F.amount = String(x.amount ?? '');
  openSheet(x ? 'Edit expense' : 'New expense', formHtml(), { right: `<button class="linkbtn" data-a="save-exp">Save</button>` });
  if (!x && preset.autoPick) setTimeout(() => $('[data-addfile]', sheetEl)?.click(), 250);
}
function formHtml() {
  const cats = S.data.categories, ps = recentProjects(), foreign = F.currency !== base();
  const curP = F.projectId && projectById(F.projectId);
  const perDay = F.category === 'Shoot day allowance';
  const conv = foreign && F.fx && num(F.amount) > 0 ? round2(num(F.amount) * F.fx) : null;
  return `
    <div class="bills" style="margin-top:12px">
      ${F.bills.map((b, i) => `<div class="g" style="${b.thumb ? `background-image:url('${tu(b.thumb)}')` : ''}">${b.thumb ? '' : `<span><span class="pdf">📄</span>${esc((b.ext || 'pdf').toUpperCase())}</span>`}<button class="x" data-a="rm-bill" data-i="${i}">✕</button></div>`).join('')}
      ${F.add.map((b, i) => `<div class="g" style="${b.thumb ? `background-image:url('${b.thumb}')` : ''}">${b.thumb ? '' : `<span><span class="pdf">📄</span>${esc(b.name.slice(0, 14))}</span>`}<button class="x" data-a="rm-add" data-i="${i}">✕</button></div>`).join('')}
      <label class="g add">${F.bills.length + F.add.length ? '＋ Another' : '📸 Snap or<br>add bill'}<input type="file" accept="image/*,application/pdf,.pdf,.heic" multiple hidden data-addfile></label>
    </div>
    <div id="scanBox">${scanBoxHtml()}</div>
    ${perDay ? `<div class="two" style="margin-top:12px"><label class="field" style="margin:0"><span>Shoot days</span><input name="days" inputmode="numeric" value="${esc(F.days)}" placeholder="0"></label>
      <label class="field" style="margin:0"><span>Rate per day</span><input name="rate" inputmode="decimal" value="${esc(F.rate)}"></label></div>` : ''}
    <div class="bigamt"><input name="amount" inputmode="decimal" placeholder="0.00" value="${esc(F.amount)}" autocomplete="off">
      <select name="currency">${CURRENCIES.map(c => `<option ${c === F.currency ? 'selected' : ''}>${c}</option>`).join('')}</select></div>
    ${foreign ? `<div class="fxline">1 ${esc(F.currency)} = <input name="fx" inputmode="decimal" value="${esc(F.fx ?? '')}" placeholder="rate"> ${esc(base())} ${conv != null ? `→ <b>${money(conv)}</b>` : F.fx == null ? '<span class="spin" style="width:14px;height:14px"></span>' : ''}</div>` : ''}
    <div class="two">
      <label class="field"><span>Date</span><input name="date" type="date" value="${esc(F.date)}"></label>
      <label class="field"><span>Project</span><select name="projectId">
        <option value="">— No project —</option>
        ${ps.map(p => `<option value="${p.id}" ${p.id === F.projectId ? 'selected' : ''}>${esc(p.name)}</option>`).join('')}
        ${curP && !ps.includes(curP) ? `<option value="${curP.id}" selected>${esc(curP.name)}</option>` : ''}
        <option value="__new">＋ New project…</option></select></label>
    </div>
    ${F.projectId ? '' : `<label class="field"><span>Name in the PROJECT column <span class="faint" style="font-weight:500">(optional, e.g. DU, School)</span></span><input name="label" value="${esc(F.label)}" placeholder="Leave empty to use the category"></label>`}
    <div class="field"><span>Category</span><div class="pills">${cats.map(c => `<button type="button" data-a="pick-cat" data-k="${esc(c.name)}" class="${F.category === c.name ? 'on' : ''}">${c.icon} ${esc(c.name)}</button>`).join('')}</div></div>
    <label class="field"><span>Description / reason</span><textarea name="reason" rows="2" placeholder="e.g. Music licence for the brand film">${esc(F.reason)}</textarea></label>
    <label class="field"><span>Vendor / shop <span class="faint" style="font-weight:500">(optional)</span></span><input name="vendor" value="${esc(F.vendor)}" placeholder="e.g. Artlist, Careem, ENOC"></label>
    ${F.monthlyId ? '' : `<label style="display:flex;gap:10px;align-items:center;margin-top:14px;font-weight:600"><input type="checkbox" name="repeat" ${F.repeat ? 'checked' : ''} style="width:20px;height:20px;accent-color:var(--acc)"> Repeat every month <span class="faint small" style="font-weight:500">(phone, rent, school…)</span></label>`}
    <button class="btn" data-a="save-exp" style="margin-top:18px">${F.id ? 'Save changes' : 'Save expense'}</button>
    ${F.id ? '' : `<button class="btn sec" data-a="save-exp" data-more="1" style="margin-top:8px">Save and add another</button>`}`;
}
function scanBoxHtml() {
  const s = F.scan; if (!s) return '';
  if (s.state === 'busy') return `<div class="note info" style="display:flex;gap:10px;align-items:center"><span class="spin" style="width:18px;height:18px;flex:none"></span><span>${esc(s.msg)}</span></div>`;
  if (s.state === 'bad') return `<div class="note warn">${esc(s.msg)}</div>`;
  const r = s.result || {};
  return `<div class="note" style="display:flex;gap:10px;align-items:flex-start"><span>✨</span><span style="flex:1">Filled in from the ${esc((r.docType || 'bill').toLowerCase())}${r.ai ? ' (AI)' : ''}${r.learned ? ' and your past expenses' : ''}${r.projectOnBill ? '; project name found on it' : ''}. Please check it.
    ${r.amount == null ? '<br><b>Couldn’t find the total — type it in.</b>' : ''}</span><button class="linkbtn" data-a="scan-undo" style="padding:0">Undo</button></div>`;
}
function readForm() {
  if (!sheetEl || !F) return;
  for (const el of sheetEl.querySelectorAll('[name]')) {
    if (el.type === 'file') continue;
    if (el.name === 'repeat') { F.repeat = el.checked; continue; }
    F[el.name] = el.value;
  }
  if (F.fx !== null && F.fx !== '' && typeof F.fx === 'string') F.fx = num(F.fx) || null;
}
/** Redraws the form from F. Call readForm() first if the screen may hold newer values. */
function redrawForm() { redrawSheet(formHtml()); }
async function updateFx() {
  readForm();
  if (F.currency === base()) { F.fx = null; redrawForm(); return; }
  F.fx = null; redrawForm();
  const want = F.currency, r = await fxRate(want, F.date);
  if (F.currency !== want) return;
  F.fx = r ?? ''; if (r == null) toast('Could not get the exchange rate — type it in.');
  redrawForm();
}
async function addFiles(files) {
  readForm();
  const firstScan = !F.add.length && !F.bills.length;
  for (const file of files) {
    const e = { id: uuid(), file, name: file.name || 'bill', isPdf: isPdfFile(file), thumb: '' };
    if (!e.isPdf) { try { e.thumb = await makeThumb(file); } catch {} }
    F.add.push(e);
  }
  redrawForm();
  if (firstScan || !num(F.amount)) scan(files[0]);
}
async function scan(file) {
  F.scan = { state: 'busy', msg: aiKey() ? 'Reading the bill with AI…' : 'Reading the bill…' }; paintScan();
  const before = { ...F };
  try {
    const r = await SCAN.read(file, {
      categories: S.data.categories.map(c => c.name), projects: recentProjects().map(p => p.name), memory: memory(), aiKey: aiKey(),
      onProgress: p => { if (F?.scan?.state === 'busy') { F.scan.msg = `Reading the bill… ${p}%`; paintScan(); } },
    });
    if (!F || !sheetEl) return;
    readForm();
    const free = k => !F.touched.has(k) && !F.touched.has('all');
    if (r.amount != null && free('amount')) F.amount = String(r.amount);
    if (r.currency && CURRENCIES.includes(r.currency) && free('currency')) F.currency = r.currency;
    if (r.date && free('date')) F.date = r.date;
    if (r.vendor && free('vendor')) F.vendor = r.vendor;
    if (r.category && free('category')) F.category = r.category;
    if (r.reason && free('reason')) F.reason = r.reason;
    if (r.project && free('projectId')) { const p = projects().find(p => p.name === r.project); if (p) F.projectId = p.id; }
    F.scan = { state: 'ok', result: r, before };
    redrawForm();
    if (F.currency !== base()) updateFx();
  } catch (e) {
    if (!F) return;
    F.scan = { state: 'bad', msg: e.message || 'Could not read this bill — fill it in yourself.' }; paintScan();
  }
}
function paintScan() { const b = $('#scanBox', sheetEl); if (b) b.innerHTML = scanBoxHtml(); }

async function saveForm(btn, more) {
  readForm();
  const foreign = F.currency !== base();
  let amount = num(F.amount);
  if (F.category === 'Shoot day allowance' && num(F.days) > 0 && !(amount > 0)) amount = num(F.days) * (num(F.rate) || 0);
  if (!(amount > 0)) return toast('Enter the amount.');
  if (!F.category) return toast('Choose a category.');
  if (foreign && !(num(F.fx) > 0)) return toast(`Enter the ${F.currency} → ${base()} rate.`);
  const orig = foreign ? round2(amount) : null, fx = foreign ? num(F.fx) : null, inBase = foreign ? round2(amount * fx) : round2(amount);
  const old = F.id ? S.data.expenses.find(e => e.id === F.id) : null;
  btn.disabled = true; const label = btn.innerHTML; btn.innerHTML = '<span class="spin" style="width:16px;height:16px"></span>';
  try {
    const rec = { date: F.date || todayStr(), projectId: F.projectId || null, label: F.projectId ? '' : F.label.trim(), category: F.category, reason: F.reason.trim(), vendor: F.vendor.trim(),
      amount: inBase, currency: F.currency, origAmount: orig, fx, days: num(F.days) > 0 ? num(F.days) : null, rate: num(F.days) > 0 ? num(F.rate) : null };
    const uploaded = [];
    for (let i = 0; i < F.add.length; i++) {
      if (F.add.length > 1) btn.innerHTML = `Uploading ${i + 1}/${F.add.length}…`;
      uploaded.push(await saveBill(F.add[i].file, rec));
    }
    if (S.demo) uploaded.forEach((b, i) => { if (!b.thumb && F.add[i].thumb) b.thumb = F.add[i].thumb; });
    const id = F.id || uuid(), removed = F.removed.slice(), repeat = F.repeat && !F.monthlyId, keepBills = F.bills;
    const ok = await mutate(d => {
      let x = d.expenses.find(e => e.id === id);
      if (!x) { x = { id, bills: [], claimId: null, createdAt: isoNow() }; d.expenses.push(x); }
      Object.assign(x, rec, { updatedAt: isoNow() });
      x.bills = [...keepBills.filter(b => (x.bills.length ? x.bills : keepBills).some(o => o.id === b.id)), ...uploaded];
      if (x.pending) delete x.pending;
      if (x.claimId) { const c = d.claims.find(c => c.id === x.claimId); if (c && !c.sentAt) { c.stale = true; c.total = round2(c.expenseIds.reduce((a, eid) => a + (+d.expenses.find(e => e.id === eid)?.amount || 0), 0)); } }
      if (repeat) {
        const t = { id: uuid(), label: rec.label || projectById(rec.projectId)?.name || '', projectId: rec.projectId, category: rec.category, reason: rec.reason, vendor: rec.vendor, amount: rec.amount, active: true, startMonth: mk(rec.date) };
        d.monthly.push(t); x.monthlyId = t.id;
      }
      if (x.monthlyId) { const t = d.monthly.find(t => t.id === x.monthlyId); if (t && mk(x.date) >= (t.lastMonth || '')) { t.amount = x.amount; t.lastMonth = mk(x.date); } }
    });
    if (!ok) { btn.disabled = false; btn.innerHTML = label; return; }
    for (const b of removed) moveToRemoved('Bills/' + b.file);
    S.f.lastProject = rec.projectId;
    toast(old ? 'Saved ✓' : `Saved ${money(inBase)} ✓`);
    if (more) { closeSheet(true); openForm(null, { projectId: rec.projectId, date: rec.date }); }
    else closeSheet();
  } catch (e) { toast(e.message); btn.disabled = false; btn.innerHTML = label; }
}

/* ---- monthly item editor ---- */
function openMonthly(id) {
  const t = id ? S.data.monthly.find(m => m.id === id) : { label: '', reason: '', category: 'Phone & internet', amount: '', active: true, projectId: null };
  if (!t) return;
  openSheet(id ? 'Monthly item' : 'New monthly item', `
    <label class="field"><span>Name in the PROJECT column</span><input id="mLabel" value="${esc(t.label)}" placeholder="e.g. DU, Home Rent, School"></label>
    <label class="field"><span>Description</span><input id="mReason" value="${esc(t.reason)}" placeholder="e.g. Mobile bill"></label>
    <div class="two"><label class="field"><span>Usual amount (${esc(base())})</span><input id="mAmount" inputmode="decimal" value="${esc(t.amount)}"></label>
    <label class="field"><span>Category</span><select id="mCat">${S.data.categories.map(c => `<option ${c.name === t.category ? 'selected' : ''}>${esc(c.name)}</option>`).join('')}</select></label></div>
    <label style="display:flex;gap:10px;align-items:center;margin-top:14px;font-weight:600"><input type="checkbox" id="mActive" ${t.active !== false ? 'checked' : ''} style="width:20px;height:20px;accent-color:var(--acc)"> Add it every month</label>
    <button class="btn" data-a="monthly-save" data-id="${id || ''}" style="margin-top:18px">Save</button>
    ${id ? `<button class="btn danger" data-a="monthly-del" data-id="${id}" style="margin-top:8px">Remove</button>` : ''}`);
}

/* ---------------- claims: make, build files, send ---------------- */
function openNewClaim() {
  const open = openItems().sort((a, b) => (a.date || '').localeCompare(b.date || ''));
  if (!open.length) return toast('Nothing to claim yet.');
  const months = [...new Set(open.map(x => mk(x.date)))].sort();
  S.nc = { months, month: 'all', sel: new Set(open.map(x => x.id)) };
  openSheet('New claim', newClaimHtml());
}
function newClaimHtml() {
  const open = openItems().sort((a, b) => (a.date || '').localeCompare(b.date || '')), nc = S.nc;
  const sel = open.filter(x => nc.sel.has(x.id)), noBill = sel.filter(x => !x.bills.length).length, pending = live().filter(x => x.pending).length;
  const period = periodOf(sel);
  return `
    <div class="months" style="margin-top:12px">${[['all', 'Everything'], ...nc.months.map(k => [k, 'Up to ' + monthShort(k)])].map(([k, l]) => `<button data-a="nc-month" data-k="${k}" class="${nc.month === k ? 'on' : ''}">${l}</button>`).join('')}</div>
    <div class="card" style="text-align:center"><div class="small muted">${esc(claimTitle(period))}</div>
      <div class="serif num" style="font-size:32px;font-weight:700;margin-top:4px">${money(sum(sel))}</div><div class="small muted">${plural(sel.length, 'expense')}</div></div>
    ${noBill ? `<div class="note warn">${plural(noBill, 'expense')} without a bill — ${noBill === 1 ? 'it' : 'they'} will be marked “No bill”.</div>` : ''}
    ${pending ? `<div class="note warn">${plural(pending, 'monthly item')} not confirmed yet, so not included. Confirm them on Home first if they belong in this claim.</div>` : ''}
    <div class="list">${open.map(x => `<button class="row ${nc.sel.has(x.id) ? 'picked' : ''}" data-a="nc-toggle" data-id="${x.id}">
      <span class="tick ${nc.sel.has(x.id) ? 'on' : ''}">${nc.sel.has(x.id) ? '✓' : ''}</span>
      <span class="main"><span class="t">${esc(x.reason || x.category)}</span><span class="s">${shortDate(x.date)} · ${esc(projectText(x))}${x.bills.length ? '' : ' · <b style="color:var(--bad)">no bill</b>'}</span></span>
      <span class="amt">${moneyS(+x.amount || 0)}</span></button>`).join('')}</div>
    <button class="btn" data-a="make-claim" style="margin-top:16px" ${sel.length ? '' : 'disabled'}>Make claim · Excel + bills PDF</button>`;
}
function periodOf(list) {
  const ms = [...new Set(list.map(x => mk(x.date)))].sort();
  if (!ms.length) return { month: thisMonth(), label: monthLabel(thisMonth()) };
  const last = ms[ms.length - 1];
  // Monthly items and late bills from the month before still belong to "this month's" sheet.
  const label = ms.length === 1 || (ms.length === 2 && ms[0] === prevMonth(last)) ? monthLabel(last) : `${monthShort(ms[0])} – ${monthShort(last)}`;
  return { month: last, label };
}
const claimTitle = p => `${st().name || 'My'} ${p.label} Expenses`.replace(/\s+/g, ' ').trim();

async function makeClaim(btn, quiet) {
  const ids = [...S.nc.sel], list = S.data.expenses.filter(x => ids.includes(x.id));
  if (!list.length) return null;
  const p = periodOf(list), id = uuid();
  const n = S.data.claims.filter(c => !c.imported && c.month === p.month).length;
  const title = claimTitle(p) + (n ? ` (${n + 1})` : '');
  if (btn) btn.disabled = true;
  const ok = await mutate(d => {
    d.claims.push({ id, title, month: p.month, period: p.label, expenseIds: ids, total: sum(list), createdAt: isoNow(), folder: `${p.month} ${cleanName(title)}` });
    for (const x of d.expenses) if (ids.includes(x.id)) x.claimId = id;
  });
  if (!ok) { if (btn) btn.disabled = false; return null; }
  if (quiet) return (await buildClaimFiles(id)) ? id : null;
  closeSheet(true);
  const built = await buildClaimFiles(id);
  openClaim(id);
  if (built) toast('Claim ready ✓');
  return id;
}
/** Makes the Excel sheet and the bills PDF and saves both in Claims/<folder>/ in OneDrive. */
async function buildClaimFiles(id) {
  const c = claimById(id); if (!c) return false;
  const list = c.expenseIds.map(eid => S.data.expenses.find(x => x.id === eid)).filter(Boolean)
    .sort((a, b) => (a.date || '').localeCompare(b.date || '') || (a.createdAt || '').localeCompare(b.createdAt || ''));
  try {
    progress('Collecting bills…');
    const entries = [];
    let i = 0;
    for (const x of list) {
      i++;
      const files = [];
      for (const b of x.bills) {
        progress(`Collecting bills… ${i}/${list.length}`);
        try { files.push({ blob: await billBlob(b), ext: (b.ext || b.file.split('.').pop()).toLowerCase(), name: b.name }); } catch (e) { console.warn(e); files.push({ blob: new Blob([]), ext: 'bad', name: b.name }); }
      }
      entries.push({ no: i, x, amount: +x.amount || 0, dateText: shortDate(x.date), project: projectText(x), category: x.category, reason: x.reason || x.category, amountText: money(+x.amount || 0), files });
    }
    progress('Making the bills PDF…');
    const info = { title: c.title, name: st().name, accountant: st().accountant, email: st().email, company: st().company, preparedOn: shortDate(todayStr()), currency: base() };
    const pdf = await CLAIM.makeBillsPdf(entries, info, (k, n) => progress(`Making the bills PDF… ${k}/${n}`));
    progress('Making the Excel sheet…');
    const xlsx = await CLAIM.makeExcel(entries.map(e => ({
      no: e.no, project: e.project, date: e.x.date, reason: e.reason, category: e.category, amount: +e.x.amount || 0,
      original: e.x.currency && e.x.currency !== base() && e.x.origAmount ? `${money(e.x.origAmount, e.x.currency)} @ ${e.x.fx}` : '',
      billPages: pdf.pages.get(e.no) || '',
    })), info);
    const files = { xlsx: `${cleanName(c.title)}.xlsx`, pdf: `${cleanName(c.title)} - Bills.pdf` };
    if (!S.demo) {
      progress('Saving to OneDrive…');
      await uploadBlob(`Claims/${c.folder}/${files.xlsx}`, xlsx);
      await uploadBlob(`Claims/${c.folder}/${files.pdf}`, pdf.blob);
    }
    S.blobs[id] = { xlsx: new File([xlsx], files.xlsx, { type: CLAIM.XLSX_TYPE }), pdf: new File([pdf.blob], files.pdf, { type: 'application/pdf' }) };
    await mutate(d => { const k = d.claims.find(z => z.id === id); if (k) { k.files = files; k.builtAt = isoNow(); k.stale = false; k.total = sum(list); k.noBill = entries.filter(e => !e.files.length).length; } });
    progress(null); return true;
  } catch (e) { progress(null); toast('Could not make the claim files: ' + e.message); return false; }
}
async function claimFiles(id) {
  if (S.blobs[id]) return S.blobs[id];
  const c = claimById(id); if (!c?.files || S.demo) return null;
  const [x, p] = await Promise.all([fileBlob(`Claims/${c.folder}/${c.files.xlsx}`), fileBlob(`Claims/${c.folder}/${c.files.pdf}`)]);
  return (S.blobs[id] = { xlsx: new File([x], c.files.xlsx, { type: CLAIM.XLSX_TYPE }), pdf: new File([p], c.files.pdf, { type: 'application/pdf' }) });
}
function emailText(c) {
  const list = c.expenseIds.map(id => S.data.expenses.find(x => x.id === id)).filter(Boolean);
  const byProj = {}; for (const x of list) { const k = projectText(x); byProj[k] = (byProj[k] || 0) + (+x.amount || 0); }
  const first = (st().accountant || '').split(' ')[0] || 'Sir';
  const lines = Object.entries(byProj).sort((a, b) => b[1] - a[1]).map(([k, v]) => `• ${k} — ${money(v)}`).join('\n');
  const noBill = list.filter(x => !x.bills.length).length;
  return {
    subject: `${c.title} — ${money(c.total)}`,
    body: `Dear ${first},\n\nPlease find attached my expenses for ${c.period}: ${plural(list.length, 'item')}, total ${money(c.total)}.\n\n${lines}\n\nThe Excel sheet lists every expense. The bills PDF has the bills in the same order — the BILL column shows the page.${noBill ? `\n${plural(noBill, 'item')} without a bill.` : ''}\n\nThank you,\n${st().name || ''}`,
  };
}

/* ---------------- one-click email to accounts (Gmail) ---------------- */
const firstName = () => (st().accountant || '').split(' ')[0] || 'accounts';
function emailHtml(c) {
  const list = c.expenseIds.map(id => S.data.expenses.find(x => x.id === id)).filter(Boolean).sort((a, b) => (a.date || '').localeCompare(b.date || ''));
  const td = 'padding:6px 10px;border-bottom:1px solid #e5e3dd;font-size:13px;vertical-align:top';
  const rows = list.map((x, i) => `<tr><td style="${td}">${i + 1}</td><td style="${td}">${esc(projectText(x))}</td><td style="${td};white-space:nowrap">${esc(shortDate(x.date))}</td><td style="${td}">${esc(x.reason || x.category)}</td><td style="${td};text-align:right;white-space:nowrap">${esc(money(+x.amount || 0))}</td><td style="${td};color:${x.bills.length ? '#1C7B62' : '#C53B3B'}">${x.bills.length ? 'Attached' : 'No bill'}</td></tr>`).join('');
  const th = 'padding:7px 10px;background:#16181D;color:#fff;font-size:12px;text-align:left';
  const t = emailText(c);
  const intro = t.body.split('\n\n')[1] || '';
  return `<div style="font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif;color:#16181D;font-size:14px;line-height:1.5">
    <p>Dear ${esc(firstName())},</p><p>${esc(intro)}</p>
    <table style="border-collapse:collapse;margin:12px 0;min-width:520px"><tr><th style="${th}">No.</th><th style="${th}">Project</th><th style="${th}">Date</th><th style="${th}">Description</th><th style="${th};text-align:right">Amount</th><th style="${th}">Bill</th></tr>${rows}
    <tr><td colspan="4" style="padding:8px 10px;font-weight:700;text-align:right;border-top:2px solid #E4572E">TOTAL</td><td style="padding:8px 10px;font-weight:700;text-align:right;white-space:nowrap;border-top:2px solid #E4572E">${esc(money(c.total))}</td><td style="border-top:2px solid #E4572E"></td></tr></table>
    <p>The Excel sheet and the bills PDF are attached. In the PDF the bills are in the same order — the BILL column of the sheet shows the page.</p>
    <p>Thank you,<br>${esc(st().name || '')}</p></div>`;
}
/** Opens the "send to accounts" check screen. claimId = existing claim, or null to claim everything that is waiting. */
function openSend(claimId) {
  const c = claimId ? claimById(claimId) : null;
  const list = c ? c.expenseIds.map(id => S.data.expenses.find(x => x.id === id)).filter(Boolean) : openItems();
  if (!list.length) return toast('Nothing to send yet.');
  const pending = c ? [] : live().filter(x => x.pending);
  const p = c ? { label: c.period } : periodOf(list), title = c ? c.title : claimTitle(p), noBill = list.filter(x => !x.bills.length).length;
  const ready = !!st().googleClientId || S.demo;
  S.sendFor = { claimId, ids: list.map(x => x.id) };
  GMAIL.load().catch(() => {});
  openSheet('Send to ' + firstName(), `
    <div class="card" style="margin-top:12px"><dl class="kv">
      <dt>To</dt><dd>${esc(st().accountant)} &lt;${esc(st().email)}&gt;</dd>
      <dt>From</dt><dd>${esc(st().fromEmail || 'your Gmail')}</dd>
      ${st().cc ? `<dt>Cc</dt><dd>${esc(st().cc)}</dd>` : ''}
      <dt>Subject</dt><dd style="font-weight:500">${esc(title)} — ${esc(money(sum(list)))}</dd>
    </dl></div>
    <div class="card" style="text-align:center"><div class="serif num" style="font-size:32px;font-weight:700">${money(sum(list))}</div>
      <div class="small muted">${plural(list.length, 'expense')} · ${esc(p.label)}</div>
      <div class="small" style="margin-top:8px">📗 Excel sheet &nbsp; 📕 Bills PDF &nbsp; 📝 Table in the email</div></div>
    ${noBill ? `<div class="note warn">${plural(noBill, 'expense')} without a bill — marked “No bill” in the sheet.</div>` : ''}
    ${pending.length ? `<div class="note warn">${plural(pending.length, 'monthly item')} not confirmed yet, so not included. <button class="linkbtn" data-a="send-confirm-all" style="padding:0;color:inherit;text-decoration:underline">Confirm and include them</button></div>` : ''}
    ${ready ? `<button class="btn" data-a="send-now" style="margin-top:14px">📧 Send now</button>
      <div class="small faint" style="text-align:center;margin-top:6px">${S.demo ? 'Sample data — nothing is really sent.' : 'Sent from your Gmail. Google may ask you to allow “send email” the first time.'}</div>`
    : `<div class="note">One-time setup needed so Billy can send from your Gmail. <button class="linkbtn" data-a="gmail-help" style="padding:0">How to set it up</button></div>`}
    <button class="btn sec" data-a="send-other" style="margin-top:8px">Send another way (Mail / share)</button>`);
}
async function sendNow(btn) {
  const sf = S.sendFor; if (!sf) return;
  let tok = null;
  btn.disabled = true; btn.innerHTML = '<span class="spin" style="width:16px;height:16px"></span> Sending…';
  try {
    if (!S.demo) tok = await GMAIL.authorize(st().googleClientId, st().fromEmail);   // first, while the tap still counts
    let id = sf.claimId;
    if (!id) {
      S.nc = { sel: new Set(sf.ids) };
      id = await makeClaim(null, true);
      if (!id) throw new Error('Could not make the claim.');
    } else if (!S.blobs[id] || claimById(id).stale) {
      if (claimById(id).stale || !claimById(id).files) { if (!(await buildClaimFiles(id))) throw new Error('Could not make the files.'); }
      else { progress('Getting the files…'); await claimFiles(id); progress(null); }
    }
    const c = claimById(id), f = S.blobs[id], t = emailText(c);
    let files = [f.xlsx, f.pdf], html = emailHtml(c), text = t.body;
    if (GMAIL.tooBig(files) && !S.demo) {     // bills too large for an attachment: send a OneDrive link instead
      progress('Bills PDF is large — adding a OneDrive link…');
      const link = await claimLink(c);
      files = [f.xlsx];
      text = text.replace('The Excel sheet', `Bills PDF (too large to attach): ${link}\n\nThe Excel sheet`);
      html = html.replace('<p>The Excel sheet and the bills PDF are attached.', `<p>The Excel sheet is attached. The bills PDF is too large to attach — <a href="${link}">open it in OneDrive</a>.`);
    }
    progress('Sending to ' + firstName() + '…');
    let msgId = 'demo';
    if (!S.demo) msgId = await GMAIL.send(tok, { from: st().fromEmail, to: `${st().accountant} <${st().email}>`, cc: st().cc, subject: t.subject, text, html, files });
    await mutate(d => { const k = d.claims.find(z => z.id === id); if (k) { k.sentAt = todayStr(); k.sentVia = 'gmail'; k.messageId = msgId; k.sentTo = st().email; } });
    progress(null); closeSheet(true); openClaim(id);
    toast(S.demo ? 'Sample — not really sent ✓' : `Sent to ${firstName()} ✓`);
  } catch (e) { progress(null); toast(e.message); btn.disabled = false; btn.innerHTML = '📧 Send now'; }
}
async function claimLink(c) {
  const meta = await (await g(item(`Claims/${c.folder}/${c.files.pdf}`) + '?$select=id')).json();
  const l = (await (await g(`/drives/${LOC.driveId}/items/${meta.id}/createLink`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ type: 'view', scope: 'anonymous' }) })).json()).link?.webUrl;
  if (!l) throw new Error('OneDrive did not give a link.');
  return l;
}
function openGmailHelp() {
  openSheet('Send from your Gmail', `
    <div class="note" style="margin-top:12px">Done once. It lets Billy send the claim from <b>${esc(st().fromEmail || 'your work Gmail')}</b>. Billy can only <b>send</b> — it can't read your mail.</div>
    <ol style="padding-left:20px;line-height:1.6;font-size:14.5px">
      <li>Open <b>console.cloud.google.com</b> signed in as ${esc(st().fromEmail || 'your work account')} and create a project called <b>Billy</b>.</li>
      <li><b>APIs &amp; Services → Library</b> → turn on <b>Gmail API</b>.</li>
      <li><b>OAuth consent screen</b> → user type <b>Internal</b> → app name Billy → save.</li>
      <li><b>Credentials → Create credentials → OAuth client ID → Web application</b>. Under <b>Authorised JavaScript origins</b> add <code>https://dijango999.github.io</code>. Create.</li>
      <li>Copy the <b>Client ID</b> (ends with .apps.googleusercontent.com) and paste it in Billy → Settings → Email to accounts.</li>
    </ol>
    <button class="btn" data-a="close-to-settings">Go to Settings</button>`);
}

function openClaim(id) {
  const c = claimById(id); if (!c) return;
  const list = c.expenseIds.map(eid => S.data.expenses.find(x => x.id === eid)).filter(Boolean).sort((a, b) => (a.date || '').localeCompare(b.date || ''));
  if (c.imported) {
    openSheet(c.title, `<div class="note">Read from your past sheet <b>${esc(c.source || '')}</b>. Billy uses it to learn — it isn't counted as money waiting.</div>
      <div class="list">${list.map(x => expenseRow(x, { date: true })).join('')}</div>
      <button class="btn danger" data-a="forget-sheet" data-id="${c.id}" style="margin-top:14px">Forget this sheet</button>`);
    return;
  }
  const [l, k, n] = claimStatus(c), ready = !!S.blobs[id];
  const canShare = !!navigator.canShare;
  openSheet(c.title, `
    <div style="text-align:center;margin-top:12px"><div class="serif num" style="font-size:34px;font-weight:700">${money(c.total)}</div>
      <div class="small muted">${plural(c.expenseIds.length, 'expense')} · ${esc(c.period)}</div>
      <div style="margin-top:6px"><span class="chip ${k}">${l}</span></div></div>
    <div class="steps"><span class="on"></span><span class="${n >= 2 ? 'on' : ''}"></span><span class="${n >= 3 ? 'on' : ''}"></span></div>
    <div style="display:flex;justify-content:space-between;font-size:11.5px;color:var(--muted);margin-top:4px;font-weight:600"><span>Made</span><span>Sent to ${esc(st().accountant?.split(' ')[0] || 'accounts')}</span><span>Paid</span></div>
    ${c.stale && !c.sentAt ? `<div class="note warn">Some expenses changed after the files were made. <button class="linkbtn" data-a="rebuild" data-id="${id}" style="padding:0;color:inherit;text-decoration:underline">Update the files</button></div>` : ''}
    ${!c.files ? `<div class="note warn">The files haven't been made yet. <button class="linkbtn" data-a="rebuild" data-id="${id}" style="padding:0;color:inherit;text-decoration:underline">Make them now</button></div>` : `
    <div class="card"><h3>Files ${S.demo ? '' : `<span class="grow"></span><span class="faint" style="text-transform:none;letter-spacing:0;font-weight:500">in OneDrive › Claims</span>`}</h3>
      <div class="list" style="margin-top:0">
        <button class="row" data-a="claim-file" data-id="${id}" data-k="xlsx"><span class="ic">📗</span><span class="main"><span class="t">${esc(c.files.xlsx)}</span><span class="s">Excel · same layout as your monthly sheet</span></span><span class="faint">›</span></button>
        <button class="row" data-a="claim-file" data-id="${id}" data-k="pdf"><span class="ic">📕</span><span class="main"><span class="t">${esc(c.files.pdf)}</span><span class="s">All bills in order${c.noBill ? ` · ${c.noBill} without bill` : ''}</span></span><span class="faint">›</span></button>
      </div></div>
    ${c.sentVia === 'gmail' ? `<div class="note" style="background:var(--goodSoft);color:var(--good)">📧 Emailed to ${esc(c.sentTo || st().email)} on ${shortDate(c.sentAt)}.</div>` : ''}
    <button class="btn ${c.sentAt ? 'sec' : ''}" data-a="send-claim" data-id="${id}" style="margin-top:14px">📧 ${c.sentAt ? 'Send again' : 'Send to ' + esc(firstName())}</button>
    <div class="actions">
      <button class="btn sec" data-a="share-claim" data-id="${id}" ${ready || !canShare ? '' : 'disabled'}>📤 Share / Mail app</button>
      ${S.demo ? '' : `<button class="btn sec" data-a="email-link" data-id="${id}">✉︎ Email a OneDrive link</button>`}
      <button class="btn sec" data-a="copy-email" data-id="${id}">Copy email text</button>
    </div>`}
    <div class="actions">
      ${c.sentAt ? `<button class="btn sec" data-a="unsent" data-id="${id}">Undo “Sent”</button>` : `<button class="btn sec" data-a="sent" data-id="${id}">Mark as sent</button>`}
      ${c.paidAt ? `<button class="btn sec" data-a="unpaid" data-id="${id}">Undo “Paid”</button>` : `<button class="btn sec" data-a="paid" data-id="${id}" ${c.sentAt ? '' : 'disabled'}>Mark as paid</button>`}
    </div>
    <div class="card" style="padding:0;overflow:hidden"><h3 style="padding:14px 14px 0">In this claim</h3><div class="list" style="border:0;border-radius:0;margin-top:4px">${list.map(x => expenseRow(x, { date: true })).join('')}</div></div>
    ${c.sentAt ? '' : `<div class="actions"><button class="btn sec" data-a="rebuild" data-id="${id}">Remake files</button><button class="btn danger" data-a="cancel-claim" data-id="${id}">Cancel claim</button></div>`}`);
  if (c.files && !ready) claimFiles(id).then(() => { if (sheetEl && $('[data-a="share-claim"]', sheetEl)?.dataset.id === id) { const b = $('[data-a="share-claim"]', sheetEl); b.disabled = false; b.innerHTML = `📤 Send to ${esc(st().accountant?.split(' ')[0] || 'accounts')}`; } }).catch(e => toast(e.message));
}
function downloadFile(file) {
  const a = document.createElement('a'); a.href = URL.createObjectURL(file); a.download = file.name; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 8000);
}
async function shareClaim(id) {
  const c = claimById(id), f = S.blobs[id], m = emailText(c);
  const mail = `mailto:${encodeURIComponent(st().email || '')}?subject=${encodeURIComponent(m.subject)}&body=${encodeURIComponent(m.body)}`;
  if (f && navigator.canShare?.({ files: [f.xlsx, f.pdf] })) {
    try {
      try { await navigator.clipboard?.writeText(st().email || ''); } catch {}
      await navigator.share({ files: [f.xlsx, f.pdf], title: m.subject, text: m.body });
      if (!c.sentAt && confirm(`Did you send it to ${st().accountant || 'accounts'}? Mark the claim as sent?`)) await mutate(d => { const k = d.claims.find(z => z.id === id); if (k) k.sentAt = todayStr(); });
      if (sheetEl) { closeSheet(true); openClaim(id); }
    } catch (e) { if (e.name !== 'AbortError') toast(e.message); }
    return;
  }
  // No file sharing here (e.g. Chrome on a Mac): save both files and open the email to attach them.
  const files = f || await claimFiles(id).catch(e => { toast(e.message); return null; });
  if (!files) return;
  downloadFile(files.xlsx); setTimeout(() => downloadFile(files.pdf), 400);
  setTimeout(() => { location.href = mail; }, 900);
  toast('Both files are in Downloads — attach them to the email that opens.');
}
async function emailLink(id) {
  const c = claimById(id), m = emailText(c);
  toast('Making a OneDrive link…');
  try {
    const meta = await (await g(item(`Claims/${c.folder}`) + '?$select=id')).json();
    const link = (await (await g(`/drives/${LOC.driveId}/items/${meta.id}/createLink`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ type: 'view', scope: 'anonymous' }) })).json()).link?.webUrl;
    if (!link) throw new Error('OneDrive did not give a link.');
    const body = m.body.replace('Please find attached my expenses', 'Here are my expenses').replace('\n\nThe Excel sheet', `\n\nFiles (Excel sheet and bills PDF):\n${link}\n\nThe Excel sheet`);
    location.href = `mailto:${encodeURIComponent(st().email || '')}?subject=${encodeURIComponent(m.subject)}&body=${encodeURIComponent(body)}`;
  } catch (e) { toast(e.message); }
}

/* ---------------- learning from past sheets ---------------- */
async function findPastSheets() {
  const out = [];
  async function walk(path, depth) {
    let next = `/me/drive/root:/${encRel(path)}:/children?$select=id,name,file,folder,size,lastModifiedDateTime,parentReference&$top=200`;
    while (next) {
      const j = await (await g(next)).json();
      for (const f of j.value || []) {
        if (f.folder && depth < 2 && !/^(billy|removed|backups)$/i.test(f.name)) await walk(path + '/' + f.name, depth + 1);
        else if (f.file && /\.(xlsx|xls|xlsm|csv)$/i.test(f.name) && !/^~\$/.test(f.name)) out.push({ id: f.id, driveId: f.parentReference?.driveId, name: f.name, path: path.split('/').slice(4).join(' › '), at: f.lastModifiedDateTime, size: f.size });
      }
      next = j['@odata.nextLink'];
    }
  }
  await walk(PAST_FOLDER, 0);
  return out.sort((a, b) => (b.at || '').localeCompare(a.at || ''));
}
async function openLearn() {
  if (S.demo) return toast('Sign in to read your real past sheets.');
  S.lr = { files: null, sel: new Set(), err: '' };
  openSheet('Learn from past sheets', learnHtml());
  try { S.lr.files = await findPastSheets(); const done = new Set(S.data.learned.map(l => l.id)); S.lr.sel = new Set(S.lr.files.filter(f => !done.has(f.id)).map(f => f.id)); }
  catch (e) { S.lr.err = e.status === 404 ? 'The EXPENCE folder was not found in your OneDrive.' : e.message; S.lr.files = []; }
  if (sheetEl) redrawSheet(learnHtml());
}
function learnHtml() {
  const lr = S.lr, done = new Set(S.data.learned.map(l => l.id));
  return `<div class="note" style="margin-top:12px">Billy reads each sheet, keeps the rows as history (not as money waiting), creates your projects and finds the items that repeat every month. Nothing in your old files is changed.</div>
    ${lr.files == null ? '<div class="empty"><span class="spin"></span><br>Looking in OneDrive › … › EXPENCE</div>' : `
    ${lr.err ? `<div class="note warn">${esc(lr.err)}</div>` : ''}
    <div class="list">${lr.files.map(f => `<button class="row ${lr.sel.has(f.id) ? 'picked' : ''}" data-a="lr-toggle" data-id="${f.id}">
      <span class="tick ${lr.sel.has(f.id) ? 'on' : ''}">${lr.sel.has(f.id) ? '✓' : ''}</span>
      <span class="main"><span class="t">${esc(f.name)}</span><span class="s">${esc(f.path || 'EXPENCE')} · ${shortDate((f.at || '').slice(0, 10))}${done.has(f.id) ? ' · already learned' : ''}</span></span></button>`).join('') || '<div class="empty">No Excel sheets found there.</div>'}</div>`}
    <label class="btn sec" style="margin-top:10px">📎 Choose sheets from this device<input type="file" accept=".xlsx,.xls,.xlsm,.csv" multiple hidden data-learnfiles></label>
    <button class="btn" data-a="lr-go" style="margin-top:10px" ${lr.files && lr.sel.size ? '' : 'disabled'}>Learn from ${plural(lr.sel.size || 0, 'sheet')}</button>`;
}
const fp = x => `${x.date}|${round2(+x.amount)}|${String(x.reason || '').toLowerCase().replace(/\s+/g, ' ').trim()}`;
/** sources: [{ id, name, buffer }] → imports every row as history and suggests monthly items. */
async function learnFrom(sources) {
  const found = [];
  for (const [i, s] of sources.entries()) {
    progress(`Reading ${i + 1}/${sources.length}: ${s.name}`);
    try { for (const sh of await LEARN.parse(s.buffer, s.name)) found.push({ src: s, sh }); } catch (e) { console.warn(s.name, e); }
  }
  progress(null);
  if (!found.length) return toast('No expense rows were found in those sheets.');
  const known = new Set(S.data.expenses.map(fp)), cats = S.data.categories.map(c => c.name);
  const newProjects = new Map(), claims = [], expenses = [], seen = new Set();
  for (const { src, sh } of found) {
    const cid = uuid(), ids = [];
    for (const r of sh.rows) {
      const g1 = SCAN.parseText(`${r.project}\n${r.reason}`);
      let category = cats.find(c => c.toLowerCase() === (r.category || '').toLowerCase()) || g1.category || 'Other';
      if (!cats.includes(category)) category = 'Other';
      const isLabel = !!g1.category || !r.project || MONTHLY_CATS.includes(category);
      let projectId = null, label = '';
      if (isLabel) label = r.project;
      else {
        const key = r.project.toLowerCase();
        let p = projects().find(p => p.name.toLowerCase() === key) || newProjects.get(key);
        if (!p) { p = { id: uuid(), name: r.project, createdAt: isoNow(), learned: true }; newProjects.set(key, p); }
        projectId = p.id;
      }
      const x = { id: uuid(), date: r.date || '', dateText: r.dateText || '', projectId, label, category, reason: r.reason || r.project, vendor: g1.vendor && g1.category ? g1.vendor : '', amount: r.amount,
        currency: base(), bills: [], claimId: cid, imported: true, createdAt: isoNow() };
      const k = fp(x); if (known.has(k) || seen.has(k)) continue; seen.add(k);
      expenses.push(x); ids.push(x.id);
    }
    if (ids.length) {
      const month = sh.month ? `${sh.month.y}-${pad2(sh.month.m)}` : mk(expenses.find(x => x.id === ids[0])?.date);
      claims.push({ id: cid, imported: true, title: sh.title || src.name.replace(/\.\w+$/, ''), month, period: monthLabel(month), source: src.name, expenseIds: ids,
        total: round2(expenses.filter(x => ids.includes(x.id)).reduce((a, x) => a + x.amount, 0)), createdAt: isoNow(), sentAt: `${month}-28`, paidAt: `${month}-28` });
    }
  }
  const learned = sources.map(s => ({ id: s.id || s.name, name: s.name, at: isoNow() }));
  // Monthly candidates: recurring-type categories, or the same label/reason in 2+ months.
  const groups = new Map();
  for (const x of expenses) {
    const pname = projectById(x.projectId)?.name || [...newProjects.values()].find(p => p.id === x.projectId)?.name || '';
    const key = `${(x.label || pname).toLowerCase()}|${x.reason.toLowerCase()}`;
    const gp = groups.get(key) || { x, months: new Set(), last: x }; gp.months.add(mk(x.date)); if ((x.date || '') >= (gp.last.date || '')) gp.last = x; groups.set(key, gp);
  }
  const have = new Set(S.data.monthly.map(t => `${(t.label || '').toLowerCase()}|${(t.reason || '').toLowerCase()}`));
  const suggest = [...groups.entries()].filter(([k, gp]) => !have.has(k) && (MONTHLY_CATS.includes(gp.x.category) || gp.months.size >= 2)).map(([, gp]) => gp.last);
  const ok = await mutate(d => {
    for (const p of newProjects.values()) if (!d.projects.some(q => q.name.toLowerCase() === p.name.toLowerCase())) d.projects.push(p);
    d.expenses.push(...expenses); d.claims.push(...claims);
    for (const l of learned) if (!d.learned.some(z => z.id === l.id)) d.learned.push(l);
  });
  if (!ok) return;
  closeSheet(true);
  S.sugg = { list: suggest, sel: new Set(suggest.filter(x => x.category !== 'Shoot day allowance').map(x => x.id)), n: expenses.length, sheets: claims.length, np: newProjects.size };
  openSheet('What Billy learned', learnedHtml());
}
function learnedHtml() {
  const sg = S.sugg, { n, sheets, np } = sg;
  return `<div style="text-align:center;margin-top:14px"><div style="font-size:40px">🧠</div>
      <div class="serif" style="font-size:22px;font-weight:700">${plural(n, 'expense')} from ${plural(sheets, 'sheet')}</div>
      <div class="small muted">${np ? `${plural(np, 'new project')} · ` : ''}saved as history, so they don't count as money waiting.</div></div>
    ${sg.list.length ? `<div class="card"><h3>Repeat these every month?</h3>
      <div class="small muted" style="margin-bottom:8px">Ticked items are added at the start of each month with the last amount — you just confirm or attach the bill.</div>
      <div class="list" style="margin-top:0">${sg.list.map(x => `<button class="row ${sg.sel.has(x.id) ? 'picked' : ''}" data-a="sg-toggle" data-id="${x.id}">
        <span class="tick ${sg.sel.has(x.id) ? 'on' : ''}">${sg.sel.has(x.id) ? '✓' : ''}</span><span class="ic cat">${catIcon(x.category)}</span>
        <span class="main"><span class="t">${esc(x.label || x.reason)}</span><span class="s">${esc(x.reason)}</span></span><span class="amt">${moneyS(x.amount)}</span></button>`).join('')}</div>
      <button class="btn" data-a="sg-save" style="margin-top:12px">Repeat ${plural(sg.sel.size, 'item')} monthly</button></div>` : ''}
    <button class="btn sec" data-a="close" style="margin-top:12px">Done</button>`;
}

/* ---------------- Excel of a project / everything ---------------- */
async function exportExcel(list, title) {
  try {
    progress('Making the Excel sheet…');
    const rows = list.slice().sort((a, b) => (a.date || '').localeCompare(b.date || '')).map((x, i) => ({
      no: i + 1, project: projectText(x), date: x.date, reason: x.reason || x.category, category: x.category, amount: +x.amount || 0,
      original: x.currency && x.currency !== base() && x.origAmount ? `${money(x.origAmount, x.currency)} @ ${x.fx}` : '',
      billPages: x.bills.length ? `${x.bills.length} in OneDrive` : '',
    }));
    const blob = await CLAIM.makeExcel(rows, { title, name: st().name, preparedOn: shortDate(todayStr()), currency: base() });
    progress(null); downloadFile(new File([blob], cleanName(title) + '.xlsx', { type: CLAIM.XLSX_TYPE }));
  } catch (e) { progress(null); toast(e.message); }
}

/* ---------------- folder browser ---------------- */
function folderBrowser(box) {
  const fs = { stack: [{ name: 'My files', driveId: null, id: null, path: '' }], items: [], loading: true, err: '' };
  const here = () => fs.stack[fs.stack.length - 1];
  async function load() {
    fs.loading = true; fs.err = ''; draw();
    try {
      const h = here();
      let next = (h.id ? `/drives/${h.driveId}/items/${h.id}/children` : '/me/drive/root/children') + '?$select=id,name,folder,file,parentReference,root&$top=200', all = [];
      while (next) { const j = await (await g(next)).json(); all = all.concat(j.value || []); next = j['@odata.nextLink']; }
      if (!h.id) { try { const r = await (await g('/me/drive/root?$select=id,parentReference')).json(); h.id = r.id; h.driveId = r.parentReference?.driveId; } catch {} }
      fs.items = all.filter(x => x.folder || x.name === CFG.dataFile).sort((a, b) => (!!b.folder - !!a.folder) || cmpText(a.name, b.name));
      if (!h.driveId && all[0]) h.driveId = all[0].parentReference?.driveId;
    } catch (e) { fs.err = e.message; }
    fs.loading = false; draw();
  }
  function draw() {
    const hasData = fs.items.some(x => x.name === CFG.dataFile), atRoot = fs.stack.length === 1;
    box.innerHTML = `<div class="list">
      <div class="gh" style="text-transform:none;letter-spacing:0;align-items:center;gap:8px">${atRoot ? '' : `<button class="linkbtn" data-b="up" style="padding:0">‹ Back</button>`}
        <span style="flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">📁 ${esc(fs.stack.map(x => x.name).join(' › '))}</span></div>
      ${fs.loading ? '<div class="empty"><span class="spin"></span></div>' : fs.err ? `<div class="empty" style="color:var(--bad)">${esc(fs.err)}</div>` :
        (fs.items.map((x, i) => x.folder
          ? `<button class="row" data-b="open" data-i="${i}"><span class="ic">📁</span><span class="main"><span class="t">${esc(x.name)}</span><span class="s">${x.folder.childCount} item${x.folder.childCount === 1 ? '' : 's'}</span></span><span class="faint">›</span></button>`
          : `<div class="row"><span class="ic">✅</span><span class="main"><span class="t">Billy data found here</span><span class="s">${CFG.dataFile}</span></span></div>`).join('') || '<div class="empty">No folders here.</div>')}
    </div>
    <div style="display:flex;gap:8px;margin-top:10px"><button class="btn sec" data-b="new" style="flex:1">＋ New folder</button>
      <button class="btn" data-b="use" style="flex:1.4" ${atRoot || fs.loading ? 'disabled' : ''}>${hasData ? 'Use this data' : 'Use this folder'}</button></div>`;
  }
  box.addEventListener('click', async e => {
    const b = e.target.closest('[data-b]'); if (!b || fs.loading) return;
    if (b.dataset.b === 'up') { fs.stack.pop(); load(); }
    else if (b.dataset.b === 'open') { const x = fs.items[+b.dataset.i]; fs.stack.push({ name: x.name, id: x.id, driveId: x.parentReference?.driveId || here().driveId, path: [here().path, x.name].filter(Boolean).join('/') }); load(); }
    else if (b.dataset.b === 'new') {
      const name = (prompt('Name for the new folder:', 'Billy') || '').trim().replace(/[\\/:*?"<>|]/g, ' '); if (!name) return;
      try {
        const h = here(), url = h.id ? `/drives/${h.driveId}/items/${h.id}/children` : '/me/drive/root/children';
        const f = await (await g(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ name, folder: {}, '@microsoft.graph.conflictBehavior': 'rename' }) })).json();
        fs.stack.push({ name: f.name, id: f.id, driveId: f.parentReference?.driveId || h.driveId, path: [h.path, f.name].filter(Boolean).join('/') }); load();
      } catch (err) { toast(err.message); }
    } else if (b.dataset.b === 'use') {
      const h = here(); b.disabled = true;
      try { saveLoc({ driveId: h.driveId, id: h.id, path: h.path }); await writePointer(LOC).catch(() => {}); closeSheet(); S.phase = 'loading'; render(); await pull(true); }
      catch (err) { toast(err.message); b.disabled = false; }
    }
  });
  load();
}

/* ---------------- sample data ---------------- */
function fakeBill(vendor, amount, color = '#E4572E') {
  const c = document.createElement('canvas'); c.width = 180; c.height = 240;
  const x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, 180, 240);
  x.fillStyle = color; x.fillRect(0, 0, 180, 34); x.fillStyle = '#fff'; x.font = 'bold 15px sans-serif'; x.fillText(vendor.slice(0, 18), 12, 23);
  x.fillStyle = '#bbb'; for (let i = 0; i < 7; i++) x.fillRect(12, 52 + i * 18, 90 + (i * 37) % 60, 6);
  x.fillStyle = '#222'; x.font = 'bold 16px sans-serif'; x.fillText('TOTAL', 12, 205); x.fillText(String(amount), 100, 205);
  return c.toDataURL('image/jpeg', 0.8);
}
function sampleData() {
  const m = thisMonth(), pm = prevMonth(m), P = (name) => ({ id: uuid(), name, createdAt: isoNow() });
  const dh = P('DHRE Palm Jebel Ali'), em = P('Emaar Beachfront film'), sr = P('Choppershoot showreel');
  const monthly = [
    ['DU', 'Mobile bill', 'Phone & internet', 250, 'du'], ['Home Internet', 'Etisalat', 'Phone & internet', 250, 'Etisalat'], ['Home Rent', 'Home Rent', 'Home rent', 3826, ''],
    ['School', 'Daughter School Grade 4', 'School fees', 1785, ''], ['School', 'Son School KG2', 'School fees', 1510, ''], ['Black Magic Cloud Subscription', 'Cloud Storage for Project', 'Software & subscriptions', 19.86, 'Blackmagic Cloud'],
  ].map(([label, reason, category, amount, vendor]) => ({ id: uuid(), label, reason, category, amount, vendor, active: true, startMonth: pm, lastMonth: pm }));
  const E = (date, p, category, reason, amount, vendor, bill = true, extra = {}) => ({ id: uuid(), date, projectId: p?.id || null, label: '', category, reason, vendor, amount, currency: 'AED',
    bills: bill ? [{ id: uuid(), file: 'demo.jpg', name: 'bill.jpg', ext: 'jpg', thumb: fakeBill(vendor || category, amount), addedAt: isoNow() }] : [], claimId: null, createdAt: isoNow(), ...extra });
  const oldClaim = uuid();
  const ex = [
    ...monthly.map(t => E(`${pm}-01`, null, t.category, t.reason, t.amount, t.vendor, t.category !== 'School fees', { label: t.label, monthlyId: t.id, claimId: oldClaim })),
    E(`${pm}-09`, dh, 'Air tickets', 'Flight DXB-COK for drone crew', 1845, 'Emirates', true, { claimId: oldClaim }),
    E(`${pm}-14`, dh, 'Music & licensing', 'Music licence for the launch film', 734.49, 'Artlist', true, { claimId: oldClaim, currency: 'USD', origAmount: 199.99, fx: 3.6725 }),
    E(`${m}-03`, em, 'Equipment rental', 'Extra FX3 body for 2 days', 650, 'Camera House', true),
    E(`${m}-04`, em, 'Taxi & transport', 'Taxi to location and back', 86.5, 'Careem'),
    E(`${m}-06`, sr, 'Music & licensing', 'Music for showreel', 734.49, 'Musicbed', true, { currency: 'USD', origAmount: 199.99, fx: 3.6725 }),
    E(`${m}-08`, dh, 'Fuel', 'Fuel for shoot day at Palm Jebel Ali', 128.4, 'ENOC'),
    E(`${m}-08`, dh, 'Food & meals', 'Lunch for crew (5)', 212, 'Al Mallah', false),
    E(`${m}-10`, dh, 'Parking & tolls', 'Parking at site', 40, 'Parkin'),
    E(`${m}-11`, null, 'Shoot day allowance', '6 Days of Shoot', 300, '', false, { label: 'Shoot Days', days: 6, rate: 50 }),
  ];
  const oldIds = ex.filter(x => x.claimId === oldClaim).map(x => x.id);
  return normalize({
    settings: { name: 'Shaheer V', accountant: 'Adnan Ahmad', email: 'accounts@choppershoot.com', company: 'Choppershoot', currency: 'AED', dayRate: 50, fromEmail: 'shaheer@choppershoot.com' },
    projects: [dh, em, sr], monthly, expenses: ex,
    claims: [{ id: oldClaim, title: `Shaheer V ${monthLabel(pm)} Expenses`, month: pm, period: monthLabel(pm), expenseIds: oldIds, total: sum(ex.filter(x => oldIds.includes(x.id))),
      createdAt: `${m}-02T09:00:00Z`, sentAt: `${m}-02`, paidAt: null, folder: `${pm} demo`, files: null }],
    createdAt: isoNow(),
  });
}

/* ---------------- events ---------------- */
async function onClick(e) {
  const t = e.target;
  const tab = t.closest('[data-tab]');
  if (tab && !t.closest('.sheet')) { S.tab = tab.dataset.tab; ls.set('by.tab', S.tab); window.scrollTo(0, 0); render(); return; }
  if (tab && t.closest('.sheet')) { closeSheet(); S.tab = tab.dataset.tab; render(); return; }
  const a = t.closest('[data-a]'); if (!a) return;
  const id = a.dataset.id, k = a.dataset.k;
  switch (a.dataset.a) {
    case 'signin': signIn(false); break;
    case 'signin-other': signIn(true); break;
    case 'demo': S.demo = true; ls.set('by.demo', '1'); applyData(sampleData(), null); ls.set('by.demoData', JSON.stringify(S.data)); ensureMonthly(); break;
    case 'setup-done': await finishSetup(a); break;
    case 'close': closeSheet(); break;
    case 'settings': S.tab = S.tab === 'settings' ? 'home' : 'settings'; window.scrollTo(0, 0); render(); break;
    case 'theme-toggle': ls.set('by.theme', isDark() ? 'light' : 'dark'); applyTheme(); render(); break;
    case 'theme': ls.set('by.theme', k); applyTheme(); render(); break;
    case 'add': openForm(null, { autoPick: false }); break;
    case 'open-exp': openExpense(id); break;
    case 'edit-exp': closeSheet(true); openForm(id); break;
    case 'dup-exp': { const x = S.data.expenses.find(e => e.id === id); closeSheet(true); openForm(null, { projectId: x.projectId, label: x.label, category: x.category, reason: x.reason, vendor: x.vendor, amount: x.currency === base() ? String(x.amount) : String(x.origAmount || ''), currency: x.currency, fx: x.fx, date: todayStr() }); break; }
    case 'del-exp': {
      const x = S.data.expenses.find(e => e.id === id); if (!x || !confirm(`Delete “${x.reason || x.category}” (${money(+x.amount || 0)})? Its bills are moved to a Removed folder in OneDrive.`)) break;
      if (await mutate(d => {
        d.expenses = d.expenses.filter(e => e.id !== id);
        for (const c of d.claims) if (c.expenseIds.includes(id)) { c.expenseIds = c.expenseIds.filter(z => z !== id); c.stale = true; c.total = round2(c.expenseIds.reduce((s, eid) => s + (+d.expenses.find(e => e.id === eid)?.amount || 0), 0)); }
      })) { for (const b of x.bills) moveToRemoved('Bills/' + b.file); closeSheet(); toast('Deleted'); }
      break;
    }
    case 'confirm': await mutateFast(d => { const x = d.expenses.find(e => e.id === id); if (x) delete x.pending; }); if (sheetEl) closeSheet(); toast('Confirmed ✓'); break;
    case 'confirm-all': await mutateFast(d => { for (const x of d.expenses) delete x.pending; }); toast('All confirmed ✓'); break;
    case 'view-bill': { const x = S.data.expenses.find(e => e.id === id); if (x) openBill(x.bills[+a.dataset.i]); break; }
    case 'pick-cat': readForm(); F.category = k; F.touched.add('category'); redrawForm(); break;
    case 'rm-bill': readForm(); F.removed.push(F.bills.splice(+a.dataset.i, 1)[0]); redrawForm(); break;
    case 'rm-add': readForm(); F.add.splice(+a.dataset.i, 1); redrawForm(); break;
    case 'scan-undo': { const b = F.scan?.before; if (b) { for (const key of ['amount', 'currency', 'date', 'vendor', 'category', 'reason', 'projectId', 'fx']) F[key] = b[key]; } F.scan = null; redrawSheet(formHtml()); break; }
    case 'save-exp': await saveForm(a, a.dataset.more === '1'); break;
    case 'f-month': S.f.month = k; render(); break;
    case 'f-status': S.f.status = k; render(); break;
    case 'open-proj': openProject(id || ''); break;
    case 'new-proj': { const n = (prompt('Project name:') || '').trim(); if (n) await mutateFast(d => { d.projects.push({ id: uuid(), name: n, createdAt: isoNow() }); }); break; }
    case 'toggle-arch': S.showArch = !S.showArch; render(); break;
    case 'proj-rename': { const p = projectById(id); const n = (prompt('New name:', p.name) || '').trim(); if (n) { await mutateFast(d => { const q = d.projects.find(q => q.id === id); if (q) q.name = n; }); closeSheet(true); openProject(id); } break; }
    case 'proj-archive': await mutateFast(d => { const q = d.projects.find(q => q.id === id); if (q) q.archived = !q.archived; }); closeSheet(); break;
    case 'proj-delete': {
      const used = S.data.expenses.filter(x => x.projectId === id).length;
      if (!confirm(used ? `Delete this project? Its ${plural(used, 'expense')} will stay, without a project.` : 'Delete this project?')) break;
      const name = projectById(id)?.name || '';
      await mutate(d => { d.projects = d.projects.filter(q => q.id !== id); for (const x of d.expenses) if (x.projectId === id) { x.projectId = null; x.label ||= name; } }); closeSheet(); break;
    }
    case 'proj-excel': { const p = projectById(id); exportExcel(live().filter(x => x.projectId === id), `${p.name} — expenses ${todayStr()}`); break; }
    case 'export-all': exportExcel(S.data.expenses, `${st().name || 'Billy'} all expenses ${todayStr()}`); break;
    case 'new-claim': openNewClaim(); break;
    case 'nc-month': {
      S.nc.month = k;
      S.nc.sel = new Set(openItems().filter(x => k === 'all' || mk(x.date) <= k).map(x => x.id));
      redrawSheet(newClaimHtml()); break;
    }
    case 'nc-toggle': S.nc.sel.has(id) ? S.nc.sel.delete(id) : S.nc.sel.add(id); redrawSheet(newClaimHtml()); break;
    case 'make-claim': await makeClaim(a); break;
    case 'open-claim': openClaim(id); break;
    case 'rebuild': closeSheet(true); if (await buildClaimFiles(id)) toast('Files updated ✓'); openClaim(id); break;
    case 'claim-file': { const f = S.blobs[id] || await claimFiles(id).catch(err => { toast(err.message); return null; }); if (f) { k === 'pdf' ? viewBlob(f.pdf, f.pdf.name) : downloadFile(f.xlsx); } break; }
    case 'share-claim': await shareClaim(id); break;
    case 'send-open': openSend(null); break;
    case 'send-claim': openSend(id); break;
    case 'send-now': await sendNow(a); break;
    case 'send-confirm-all': await mutateFast(d => { for (const x of d.expenses) delete x.pending; }); closeSheet(true); openSend(null); break;
    case 'send-other': { const sf = S.sendFor; closeSheet(true); if (sf.claimId) { await shareClaim(sf.claimId); } else { S.nc = { sel: new Set(sf.ids) }; const cid = await makeClaim(null, true); if (cid) openClaim(cid); } break; }
    case 'gmail-help': openGmailHelp(); break;
    case 'close-to-settings': closeSheet(); S.tab = 'settings'; render(); break;
    case 'email-link': await emailLink(id); break;
    case 'copy-email': { const m = emailText(claimById(id)); try { await navigator.clipboard.writeText(`To: ${st().email}\nSubject: ${m.subject}\n\n${m.body}`); toast('Email text copied ✓'); } catch { toast('Could not copy'); } break; }
    case 'sent': await mutateFast(d => { const c = d.claims.find(z => z.id === id); if (c) c.sentAt = todayStr(); }); closeSheet(true); openClaim(id); break;
    case 'unsent': await mutateFast(d => { const c = d.claims.find(z => z.id === id); if (c) { c.sentAt = null; c.paidAt = null; } }); closeSheet(true); openClaim(id); break;
    case 'paid': await mutateFast(d => { const c = d.claims.find(z => z.id === id); if (c) c.paidAt = todayStr(); }); closeSheet(true); openClaim(id); toast('Marked as paid ✓'); break;
    case 'unpaid': await mutateFast(d => { const c = d.claims.find(z => z.id === id); if (c) c.paidAt = null; }); closeSheet(true); openClaim(id); break;
    case 'cancel-claim': {
      if (!confirm('Cancel this claim? Its expenses go back to “To claim”. The files stay in OneDrive.')) break;
      await mutate(d => { d.claims = d.claims.filter(c => c.id !== id); for (const x of d.expenses) if (x.claimId === id) x.claimId = null; });
      delete S.blobs[id]; closeSheet(); toast('Claim cancelled'); break;
    }
    case 'forget-sheet': {
      if (!confirm('Forget this sheet? Its rows are removed from Billy (your Excel file is not touched).')) break;
      await mutate(d => { const c = d.claims.find(z => z.id === id); d.expenses = d.expenses.filter(x => x.claimId !== id); d.claims = d.claims.filter(z => z.id !== id); if (c) d.learned = d.learned.filter(l => l.name !== c.source); });
      closeSheet(); break;
    }
    case 'learn': openLearn(); break;
    case 'learn-skip': ls.set('by.learnSkip', '1'); render(); break;
    case 'lr-toggle': S.lr.sel.has(id) ? S.lr.sel.delete(id) : S.lr.sel.add(id); redrawSheet(learnHtml()); break;
    case 'lr-go': {
      const files = S.lr.files.filter(f => S.lr.sel.has(f.id)); a.disabled = true;
      try {
        const sources = [];
        for (const [i, f] of files.entries()) { progress(`Downloading ${i + 1}/${files.length}`); const blob = await (await g(`/drives/${f.driveId}/items/${f.id}/content`)).blob(); sources.push({ id: f.id, name: f.name, buffer: new Uint8Array(await blob.arrayBuffer()) }); }
        progress(null); await learnFrom(sources);
      } catch (err) { progress(null); toast(err.message); a.disabled = false; }
      break;
    }
    case 'sg-toggle': S.sugg.sel.has(id) ? S.sugg.sel.delete(id) : S.sugg.sel.add(id); redrawSheet(learnedHtml()); break;
    case 'sg-save': {
      const pick = S.sugg.list.filter(x => S.sugg.sel.has(x.id)), m = thisMonth();
      await mutate(d => { for (const x of pick) d.monthly.push({ id: uuid(), label: x.label || projectById(x.projectId)?.name || '', projectId: x.projectId, category: x.category, reason: x.reason, vendor: x.vendor || '', amount: x.amount, active: true, startMonth: m, lastMonth: mk(x.date) }); });
      closeSheet(); ensureMonthly(); toast(`${plural(pick.length, 'monthly item')} set up ✓`); break;
    }
    case 'monthly-new': openMonthly(null); break;
    case 'monthly-edit': openMonthly(id); break;
    case 'monthly-save': {
      const v = { label: $('#mLabel', sheetEl).value.trim(), reason: $('#mReason', sheetEl).value.trim(), amount: num($('#mAmount', sheetEl).value) || 0, category: $('#mCat', sheetEl).value, active: $('#mActive', sheetEl).checked };
      if (!v.label && !v.reason) return toast('Give it a name.');
      await mutate(d => { const t = id && d.monthly.find(m => m.id === id); if (t) Object.assign(t, v); else d.monthly.push({ id: uuid(), ...v, startMonth: thisMonth() }); });
      closeSheet(); ensureMonthly(); break;
    }
    case 'monthly-del': if (confirm('Stop adding this item every month? Past months stay.')) { await mutate(d => { d.monthly = d.monthly.filter(m => m.id !== id); d.expenses = d.expenses.filter(x => !(x.monthlyId === id && x.pending)); }); closeSheet(); } break;
    case 'cat-new': { const n = (prompt('New category name:') || '').trim(); if (!n) break; const icon = (prompt('An emoji for it (optional):', '🧾') || '🧾').trim().slice(0, 4); await mutateFast(d => { if (!d.categories.some(c => c.name.toLowerCase() === n.toLowerCase())) d.categories.push({ name: n, icon }); }); break; }
    case 'cat-edit': { const icon = prompt(`Emoji for “${k}”:`, catIcon(k)); if (icon) await mutateFast(d => { const c = d.categories.find(c => c.name === k); if (c) c.icon = icon.trim().slice(0, 4); }); break; }
    case 'ai-save': { const v = ($('#aiKey')?.value || '').trim(); if (!/^[\w\-.]{25,}$/.test(v)) { toast('That doesn’t look like an API key'); break; } ls.set('by.aiKey', v); toast('AI bill reading is on'); render(); break; }
    case 'ai-off': ls.del('by.aiKey'); toast('AI key removed'); render(); break;
    case 'change-folder': S.phase = 'folder'; render(); break;
    case 'signout': if (confirm(S.demo ? 'Leave the sample data?' : 'Sign out on this device? Your data stays in OneDrive.')) signOut(); break;
  }
}
async function finishSetup(btn) {
  const su = S.setup;
  if (!su.name.trim()) return toast('Add your name — it goes on the claim sheet.');
  btn.disabled = true; btn.innerHTML = '<span class="spin" style="width:16px;height:16px"></span>';
  const d = normalize({
    settings: { name: su.name.trim(), accountant: su.accountant.trim(), email: su.email.trim(), company: su.company.trim(), fromEmail: su.fromEmail.trim() },
    projects: su.projects.split('\n').map(s => s.trim()).filter(Boolean).map(name => ({ id: uuid(), name, createdAt: isoNow() })),
    createdAt: isoNow(),
  });
  try { const etag = await putData(d, null, true); applyData(d, etag); setSync('ok'); }
  catch (e) { if (e.status === 412) { await pull(true); return; } toast(e.message); btn.disabled = false; btn.textContent = 'Start'; }
}
async function onChange(e) {
  const t = e.target;
  if (t.matches('[data-addfile]') && t.files?.length) { await addFiles([...t.files]); return; }
  if (t.matches('[data-learnfiles]') && t.files?.length) {
    const sources = []; for (const f of t.files) sources.push({ id: 'local:' + f.name, name: f.name, buffer: new Uint8Array(await f.arrayBuffer()) });
    await learnFrom(sources); return;
  }
  if (t.dataset.addbill && t.files?.length) {
    const x = S.data.expenses.find(e => e.id === t.dataset.addbill); if (!x) return;
    toast('Uploading…');
    try {
      const added = []; for (const f of t.files) added.push(await saveBill(f, x));
      if (S.demo) for (const [i, b] of added.entries()) if (!b.thumb && !isPdfFile(t.files[i])) b.thumb = await makeThumb(t.files[i]);
      if (await mutate(d => { const y = d.expenses.find(e => e.id === x.id); if (y) { y.bills.push(...added); y.updatedAt = isoNow(); const c = y.claimId && d.claims.find(c => c.id === y.claimId); if (c && !c.sentAt) c.stale = true; } })) { closeSheet(true); openExpense(x.id); toast('Bill added ✓'); }
    } catch (err) { toast(err.message); }
    return;
  }
  if (t.dataset.set) {
    const key = t.dataset.set, v = key === 'dayRate' ? num(t.value) || 0 : t.value.trim();
    await mutateFast(d => { d.settings[key] = v; }); return;
  }
  if (F && sheetEl && t.closest('.sheet') && t.name) {
    F.touched.add(t.name);
    if (t.name === 'projectId' && t.value === '__new') {
      const n = (prompt('New project name:') || '').trim();
      readForm();
      if (!n) { F.projectId = null; redrawForm(); return; }
      const p = { id: uuid(), name: n, createdAt: isoNow() };
      F.projectId = p.id;
      await mutateFast(d => { d.projects.push(p); });
      redrawForm(); return;
    }
    if (t.name === 'currency' || (t.name === 'date' && F.currency !== base())) { await updateFx(); return; }
    if (t.name === 'projectId' || t.name === 'fx') { readForm(); redrawForm(); return; }
    if (t.name === 'days' || t.name === 'rate') { readForm(); if (num(F.days) > 0) { F.amount = String(round2(num(F.days) * (num(F.rate) || 0))); redrawForm(); } return; }
  }
}
let qTimer;
function onInput(e) {
  const t = e.target;
  if (t.dataset?.su != null) { S.setup[t.dataset.su] = t.value; return; }
  if (F && sheetEl && t.closest?.('.sheet') && t.name) F.touched.add(t.name);
  if (t.id === 'q') { clearTimeout(qTimer); const v = t.value; qTimer = setTimeout(() => { S.f.q = v; const pos = t.selectionStart; render(); const q = $('#q'); if (q) { q.focus(); q.setSelectionRange(pos, pos); } }, 200); }
}
document.addEventListener('click', e => { if (!e.target.closest('.sheet')) onClick(e); });
document.addEventListener('change', e => { if (!e.target.closest('.sheet')) onChange(e); });
document.addEventListener('input', e => { if (!e.target.closest('.sheet')) onInput(e); });

/* ---------------- updates and start ---------------- */
async function checkForUpdate() {
  if (location.protocol !== 'https:') return;
  try {
    const v = (await fetchJSON('version.json?t=' + Date.now(), 6000)).version;
    if (v && v !== CFG.version && ls.get('by.updTried') !== v) {
      ls.set('by.updTried', v);
      try { const rs = await navigator.serviceWorker?.getRegistrations?.(); for (const r of rs || []) await r.update(); } catch {}
      try { for (const k of await caches.keys()) await caches.delete(k); } catch {}
      location.reload();
    }
  } catch {}
}
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') checkForUpdate(); });

async function boot() {
  applyTheme();
  if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js', { updateViaCache: 'none' }).catch(() => {});
  checkForUpdate();
  if (S.demo) {
    const d = lsJSON('by.demoData'); if (d) { applyData(normalize(d), null); ensureMonthly(); return; }
    S.demo = false; ls.del('by.demo');
  }
  try { await handleRedirect(); } catch (e) { toast(e.message); }
  if (!tokens) { S.phase = 'welcome'; render(); return; }
  const cached = lsJSON('by.cache');
  if (cached && LOC) { S.etag = ls.get('by.etag') || null; applyData(normalize(cached), S.etag); }
  else { S.phase = 'loading'; render(); }
  try {
    if (!(await resolveLocation())) { S.phase = 'folder'; render(); return; }
    await pull(!cached);
    if (!S.data && S.phase !== 'setup' && S.phase !== 'needsignin') S.phase = 'setup';
    render();
  } catch (e) {
    if (e instanceof AuthError) { if (trySilent()) return; S.phase = cached ? 'ready' : 'needsignin'; render(); return; }
    if (cached) setSync('bad', e.status === 0 ? 'Offline — showing the last synced copy' : e.message);
    else { S.phase = 'welcome'; render(); toast(e.message); }
  }
}
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible' && S.phase === 'ready') pull(); });
window.addEventListener('scroll', () => $('header.top')?.classList.toggle('scrolled', window.scrollY > 4), { passive: true });
boot();
