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

  const MON = 'janfebmaraprmayjunjulaugsepoctnovdec';
  /** Any common way a date is written in Sigma or Excel → YYYY-MM-DD (day first, as in the UAE). */
  function anyDate(v) {
    if (v instanceof Date) return `${v.getFullYear()}-${pad(v.getMonth() + 1)}-${pad(v.getDate())}`;
    if (typeof v === 'number' && v > 30000 && v < 60000) { const x = new Date(Math.round((v - 25569) * 864e5)); return `${x.getUTCFullYear()}-${pad(x.getUTCMonth() + 1)}-${pad(x.getUTCDate())}`; }
    const t = String(v ?? '');
    let m = /\b(\d{4})-(\d{1,2})-(\d{1,2})\b/.exec(t); if (m) return `${m[1]}-${pad(m[2])}-${pad(m[3])}`;
    m = /\b(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4}|\d{2})\b/.exec(t); if (m) return `${m[3].length === 2 ? '20' + m[3] : m[3]}-${pad(m[2])}-${pad(m[1])}`;
    m = /\b(\d{1,2})(?:st|nd|rd|th)?[\s\-]+(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[\s\-,]+(\d{4})\b/i.exec(t); if (m) return `${m[3]}-${pad(MON.indexOf(m[2].toLowerCase()) / 3 + 1)}-${pad(m[1])}`;
    m = /\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s+(\d{1,2}),?\s+(\d{4})\b/i.exec(t); if (m) return `${m[3]}-${pad(MON.indexOf(m[1].toLowerCase()) / 3 + 1)}-${pad(m[2])}`;
    return '';
  }
  /** Text copied from Excel (tab-separated cells) or from the My days page. */
  function fromPaste(text) {
    const t = String(text || '');
    if (/\t/.test(t)) {
      const rows = t.split(/\r?\n/).filter(l => l.trim()).map(l => l.split('\t').map(c => c.replace(/^"|"$/g, '').trim()));
      const d = fromRows(rows); if (d.length) return d;
    }
    return fromText(t);
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
    if (h < 0) {
      // No header row (e.g. only the data cells were copied): the date column is the one full of dates,
      // the project is the longest text in the row that isn't a date, time or role.
      const days = [];
      for (const r of rows) {
        const cells = (r || []).map(str); const di = cells.findIndex(x => anyDate(x));
        if (di < 0) continue;
        const txt = cells.filter((x, j) => j !== di && x && !anyDate(x) && !/^\d{1,2}:\d{2}/.test(x) && !/operator|camera|editor|wrapped|shoot day$/i.test(x) && /[a-z]{3}/i.test(x)).sort((a, b) => b.length - a.length);
        if (!txt.length) continue;
        days.push({ date: anyDate(cells[di]), project: title(txt[0]), client: txt[1] || '', role: '', note: '', type: cells.find(x => /day/i.test(x) && !anyDate(x)) || 'Shoot Day', wrapped: null });
      }
      if (days.length) return dedupe(days);
      return fromText(rows.map(r => (r || []).map(str).filter(Boolean).join('\n')).join('\n'));
    }
    const days = [];
    for (const r of rows.slice(h + 1)) {
      const date = anyDate(r[c.date]);
      if (!date || !str(r[c.project])) continue;
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
  return { fromText, fromRows, fromPaste, anyDate, counts };
})();
if (typeof module !== 'undefined') module.exports = SIGMA;
