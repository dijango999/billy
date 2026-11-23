/* Billy — reads your shoot days from Team Sigma (sigma.choppershoot.com › My days).
 *
 * Two ways in, both read on the device:
 *   • the Excel file from My days › Export Excel
 *   • the My days page copied as text (select all, copy, paste)
 * Each shoot day becomes { date, project, client, role, note, type, wrapped }. Billy then adds one
 * "Shoot day allowance" line for the month: number of shoot days × your day rate.
 */
'use strict';

const SIGMA = (() => {
  const pad = n => String(n).padStart(2, '0');
  const DATE = /^(\d{1,2})\/(\d{1,2})\/(\d{4})\b(.*)$/;
  const title = s => /[a-z]/.test(s) ? s : s.toLowerCase().replace(/(^|[\s(\-–/&])([a-z])/g, (m, a, b) => a + b.toUpperCase()).replace(/\b(X|I3|Pfl|Det|Adnoc|Nmdc|Dhre|Ad)\b/gi, w => w.toUpperCase());

  /** Text of the My days page → shoot days. */
  function fromText(text) {
    const L = String(text || '').split(/\r?\n/).map(s => s.trim()).filter(Boolean);
    const days = []; let proj = '', client = '', role = '', wrapped = null;
    for (let i = 0; i < L.length; i++) {
      const l = L[i];
      if (/^(wrapped|not wrapped yet)$/i.test(l)) { proj = title(L[i - 1] || ''); wrapped = /^wrapped$/i.test(l); client = ''; role = ''; continue; }
      if (/^client$/i.test(l)) { client = L[i + 1] || ''; i++; continue; }
      if (/^role$/i.test(l)) { role = L[i + 1] || ''; i++; continue; }
      const m = DATE.exec(l);
      if (m && proj) {
        const type = (L[i + 1] && !DATE.exec(L[i + 1]) && !/^(client|role)$/i.test(L[i + 1]) && /day|travel|recce|prep|edit/i.test(L[i + 1])) ? L[i + 1] : '';
        const note = (/\(([^)]*)\)\s*$/.exec(m[4]) || [])[1] || '';
        days.push({ date: `${m[3]}-${pad(m[2])}-${pad(m[1])}`, project: proj, client, role, time: m[4].replace(/\([^)]*\)\s*$/, '').trim(), note, type: type || 'Shoot Day', wrapped });
        if (type) i++;
      }
    }
    return dedupe(days);
  }

  /** Rows of the exported sheet (array of arrays) → shoot days. Works with a header row or with a text-like layout. */
  function fromRows(rows) {
    const str = v => v == null ? '' : v instanceof Date ? `${pad(v.getDate())}/${pad(v.getMonth() + 1)}/${v.getFullYear()}` : String(v).trim();
    let h = -1, c = {};
    for (let i = 0; i < Math.min(rows.length, 20) && h < 0; i++) {
      const cc = {};
      (rows[i] || []).forEach((v, j) => {
        const t = str(v).toLowerCase();
        if (cc.date == null && /^(shoot\s*)?date|^day$/.test(t)) cc.date = j;
        else if (cc.project == null && /project|job|event|shoot name|title/.test(t)) cc.project = j;
        else if (cc.client == null && /client/.test(t)) cc.client = j;
        else if (cc.role == null && /role|position/.test(t)) cc.role = j;
        else if (cc.type == null && /type|day type|category/.test(t)) cc.type = j;
        else if (cc.status == null && /status|wrap/.test(t)) cc.status = j;
        else if (cc.note == null && /note|detail|description|time/.test(t)) cc.note = j;
      });
      if (cc.date != null && cc.project != null) { h = i; c = cc; }
    }
    if (h < 0) return fromText(rows.map(r => (r || []).map(str).filter(Boolean).join('\n')).join('\n'));
    const days = [];
    for (const r of rows.slice(h + 1)) {
      let d = r[c.date], date = '';
      if (d instanceof Date) date = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
      else if (typeof d === 'number' && d > 30000) { const x = new Date(Math.round((d - 25569) * 864e5)); date = `${x.getUTCFullYear()}-${pad(x.getUTCMonth() + 1)}-${pad(x.getUTCDate())}`; }
      else { const m = /(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(str(d)) || []; if (m[0]) date = `${m[3]}-${pad(m[2])}-${pad(m[1])}`; else { const n = /(\d{4})-(\d{2})-(\d{2})/.exec(str(d)); if (n) date = n[0]; } }
      if (!date) continue;
      const type = c.type != null ? str(r[c.type]) : 'Shoot Day';
      days.push({ date, project: title(str(r[c.project])), client: c.client != null ? str(r[c.client]) : '', role: c.role != null ? str(r[c.role]) : '', note: c.note != null ? str(r[c.note]) : '', type: type || 'Shoot Day',
        wrapped: c.status != null ? !/not/i.test(str(r[c.status])) : null });
    }
    return dedupe(days);
  }
  function dedupe(days) {
    const seen = new Set();
    return days.filter(d => { const k = d.date + '|' + d.project.toLowerCase(); if (seen.has(k)) return false; seen.add(k); return true; }).sort((a, b) => a.date.localeCompare(b.date));
  }
  /** Only real shoot days count for the allowance (not leave, travel or office days unless they say "shoot"). */
  const counts = d => /shoot|event|filming/i.test(d.type || 'shoot');
  return { fromText, fromRows, counts };
})();
if (typeof module !== 'undefined') module.exports = SIGMA;
