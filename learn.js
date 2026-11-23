/* Billy — reads the monthly expense sheets you sent before (Excel or CSV) so Billy can learn
 * your projects, categories, usual amounts and the items that come back every month.
 *
 * Works with the usual layout ("Shaheer V April 2025 Expenses", then PROJECT | DATE | DISCRIPTION |
 * AMOUNT and a TOTAL row) and with most other layouts: it looks for a header row that has an
 * amount column plus a project / description column, wherever it is on the sheet.
 */
'use strict';

const LEARN = (() => {
  const SHEETJS = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
  let lib = typeof XLSX !== 'undefined' ? XLSX : null;
  async function load() {
    if (lib) return lib;
    if (typeof require === 'function' && typeof window === 'undefined') return (lib = require('xlsx'));
    await new Promise((res, rej) => { const s = document.createElement('script'); s.src = SHEETJS; s.onload = res; s.onerror = () => rej(new Error('Could not load the Excel reader. Check your internet connection.')); document.head.appendChild(s); });
    return (lib = window.XLSX);
  }
  const MONTHS = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
  const monthOf = w => { const t = String(w || '').toLowerCase().trim(); if (t.length < 3) return 0; const i = MONTHS.findIndex(m => m.startsWith(t.slice(0, 3))); return i + 1; };
  const pad = n => String(n).padStart(2, '0');
  const ymd = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const dayOf = d => ymd(new Date(d.getTime() + 6 * 3600e3));   // SheetJS dates can land a few seconds before midnight
  const str = v => v == null ? '' : v instanceof Date ? dayOf(v) : String(v).replace(/\s+/g, ' ').trim();
  const money = v => { if (typeof v === 'number') return v; const n = parseFloat(String(v ?? '').replace(/[^\d.\-]/g, '')); return isFinite(n) ? n : NaN; };

  /** "April 2025" anywhere in a text → { y, m } */
  function findMonth(text) {
    const m = /\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[\s_\-.,]*(20\d{2})\b/i.exec(text || '');
    if (m) return { y: +m[2], m: monthOf(m[1]) };
    const n = /\b(20\d{2})[\s_\-.](\d{1,2})\b/.exec(text || ''); if (n && +n[2] >= 1 && +n[2] <= 12) return { y: +n[1], m: +n[2] };
    return null;
  }

  const ROLE = {
    project: /^(project|projects|job|client|item|head|account|particular|expense\s*head)s?$/i,
    date: /^(date|day|month|bill\s*date|txn\s*date)$/i,
    reason: /^(d[ie]s?c?r?i?p?tion|discription|description|details?|reason|purpose|narration|remarks?|particulars?|notes?|for)$/i,
    amount: /^(amount|amt|total|price|cost|value|aed|amount\s*\(?aed\)?|amount\s*aed)$/i,
    category: /^(category|type|cat)$/i,
  };

  /** Returns [{ sheet, title, month:{y,m}, rows:[{ project, date, dateText, reason, amount, category }] }] */
  /** A month written without a year ("DEC", "Sept") → month number. */
  function monthWord(text) {
    const m = /\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\b/i.exec(text || '');
    return m ? monthOf(m[1]) : 0;
  }
  /** fileDate: when the file was last saved (YYYY-MM-DD…), used to work out the year when the sheet only says "DEC". */
  async function parse(buffer, fileName = '', fileDate = '') {
    const X = await load();
    const wb = X.read(buffer, { type: 'array', cellDates: true });
    const out = [];
    for (const name of wb.SheetNames) {
      const aoa = X.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null, blankrows: false });
      // header row: has an amount column and a project or description column
      let h = -1, cols = null;
      for (let i = 0; i < Math.min(aoa.length, 30) && h < 0; i++) {
        const c = {};
        (aoa[i] || []).forEach((v, j) => { const t = str(v).replace(/[:.]$/, ''); for (const [k, re] of Object.entries(ROLE)) if (c[k] == null && re.test(t)) c[k] = j; });
        if (c.amount != null && (c.project != null || c.reason != null)) { h = i; cols = c; }
      }
      if (h < 0) continue;
      const before = aoa.slice(0, h).flat().map(str).filter(Boolean).join(' ');
      const title = aoa.slice(0, h).flat().map(str).find(s => /[a-z]{3,}/i.test(s)) || '';
      let month = findMonth(before) || findMonth(fileName) || findMonth(name);
      if (!month) {
        const mw = monthWord(before) || monthWord(fileName) || monthWord(name);
        // the year: from real dates in the sheet, else from when the file was saved (a December sheet saved in January is last year's)
        const ys = {}; for (let i = h + 1; i < aoa.length; i++) for (const v of aoa[i] || []) if (v instanceof Date) { const y = v.getFullYear(); if (y > 2000) ys[y] = (ys[y] || 0) + 1; }
        let y = +Object.entries(ys).sort((a, b) => b[1] - a[1])[0]?.[0] || 0;
        if (!y && mw && /^\d{4}-\d{2}/.test(fileDate)) { const fy = +fileDate.slice(0, 4), fm = +fileDate.slice(5, 7); y = mw > fm + 1 ? fy - 1 : fy; }
        if (mw && y) month = { y, m: mw };
      }
      const rows = [];
      for (let i = h + 1; i < aoa.length; i++) {
        const r = aoa[i] || [];
        const cells = r.map(str);
        if (cells.some(s => /^(grand\s*)?total$/i.test(s))) break;
        const amount = money(r[cols.amount]);
        if (!(amount > 0)) continue;
        const project = cols.project != null ? str(r[cols.project]) : '';
        const reason = cols.reason != null ? str(r[cols.reason]) : '';
        if (!project && !reason) continue;
        let date = '', dateText = '';
        const dv = cols.date != null ? r[cols.date] : null;
        if (dv instanceof Date) date = dayOf(dv);
        else if (typeof dv === 'number' && dv > 30000 && dv < 60000) date = ymd(new Date(Math.round((dv - 25569) * 864e5)));
        else if (dv != null && str(dv)) {
          const t = str(dv), mm = monthOf(t), dm = /^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})$/.exec(t);
          if (dm) date = `${dm[3].length === 2 ? '20' + dm[3] : dm[3]}-${pad(dm[2])}-${pad(dm[1])}`;
          else if (mm && month) { const y = mm < month.m - 6 ? month.y + 1 : mm > month.m + 6 ? month.y - 1 : month.y; date = `${y}-${pad(mm)}-01`; dateText = MONTHS[mm - 1][0].toUpperCase() + MONTHS[mm - 1].slice(1); }
          else dateText = t;
        }
        if (!date && month) date = `${month.y}-${pad(month.m)}-01`;
        rows.push({ project, date, dateText, reason, amount: Math.round(amount * 100) / 100, category: cols.category != null ? str(r[cols.category]) : '' });
      }
      if (rows.length) out.push({ sheet: name, title, month, rows });
    }
    return out;
  }
  return { parse, findMonth, load };
})();
if (typeof module !== 'undefined') module.exports = LEARN;
