/* Billy — builds the monthly claim for the accounts team.
 *
 *   makeExcel()  → Expense Claim .xlsx: every expense in date order with a total, plus a
 *                  "Summary" sheet with totals by project and by category (live Excel formulas).
 *   makeBillsPdf() → one PDF with every bill in the same order. Each page carries a header line
 *                  (No., date, project, amount) so the accountant can match it to the Excel row.
 *
 * Both libraries load from a CDN only when a claim is made (ExcelJS, pdf-lib).
 */
'use strict';

const CLAIM = (() => {
  const CDN = {
    excel: 'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js',
    pdf: 'https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/dist/pdf-lib.min.js',
  };
  const loaded = {};
  function loadScript(src) {
    return loaded[src] ||= new Promise((res, rej) => {
      const s = document.createElement('script'); s.src = src; s.async = true;
      s.onload = res; s.onerror = () => { delete loaded[src]; rej(new Error('Could not load the claim builder. Check your internet connection.')); };
      document.head.appendChild(s);
    });
  }
  const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const day = s => { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s || ''); return m ? new Date(Date.UTC(+m[1], +m[2] - 1, +m[3])) : null; };
  const nice = s => { const d = day(s); return d ? d.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric', timeZone: 'UTC' }) : ''; };
  const money2 = n => (+n || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  /**
   * Same layout as the monthly sheet the accounts team already gets:
   *   "<Name> <Month Year> Expenses"  then  No. | PROJECT | DATE | DESCRIPTION | CATEGORY | AMOUNT | BILL  and a TOTAL.
   * rows: [{ no, project, date, dateText, reason, category, amount, original, billPages }]
   * info: { title, name, accountant, email, preparedOn, currency }
   */
  async function makeExcel(rows, info) {
    await loadScript(CDN.excel);
    const wb = new ExcelJS.Workbook();
    wb.creator = info.name || 'Billy'; wb.created = new Date();
    const C = info.currency || 'AED';
    const ws = wb.addWorksheet('Expenses', {
      views: [{ state: 'frozen', ySplit: 4 }],
      pageSetup: { orientation: 'landscape', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: .4, right: .4, top: .5, bottom: .5, header: .2, footer: .2 } },
    });
    const cols = [['No.', 6], ['PROJECT', 34], ['DATE', 14], ['DESCRIPTION', 58], ['CATEGORY', 22], [`AMOUNT (${C})`, 15], ['BILL', 13]];
    cols.forEach(([, w], i) => { ws.getColumn(i + 1).width = w; });
    const ink = { argb: 'FF16181D' }, acc = { argb: 'FFE4572E' }, soft = { argb: 'FFF6F5F1' }, line = { style: 'thin', color: { argb: 'FFD9D6CE' } };
    ws.mergeCells('A1:G1');
    ws.getCell('A1').value = info.title;
    ws.getCell('A1').font = { bold: true, size: 16, color: ink };
    ws.getRow(1).height = 26;
    ws.mergeCells('A2:G2');
    ws.getCell('A2').value = [info.accountant || info.email ? `To: ${[info.accountant, info.email].filter(Boolean).join(' — ')}` : '', `Prepared on ${info.preparedOn}`, `${rows.length} item${rows.length === 1 ? '' : 's'}`].filter(Boolean).join('   ·   ');
    ws.getCell('A2').font = { color: { argb: 'FF676B73' } };
    const H = 4, hr = ws.getRow(H);
    cols.forEach(([t], i) => {
      const c = hr.getCell(i + 1); c.value = t;
      c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: ink };
      c.alignment = { vertical: 'middle', horizontal: i === 5 ? 'right' : 'left' };
    });
    hr.height = 22;
    rows.forEach((x, i) => {
      const r = ws.getRow(H + 1 + i);
      const desc = x.original ? `${x.reason}  (${x.original})` : x.reason;
      r.values = [x.no, x.project, x.date ? day(x.date) : (x.dateText || ''), desc, x.category, +x.amount || 0, x.billPages || 'No bill'];
      if (x.date) r.getCell(3).numFmt = 'dd mmm yyyy';
      r.getCell(6).numFmt = '#,##0.00';
      for (let c = 1; c <= 7; c++) {
        const cell = r.getCell(c); cell.border = { bottom: line };
        cell.alignment = { vertical: 'top', wrapText: c === 4 || c === 2 };
        if (i % 2) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: soft };
      }
      if (!x.billPages) r.getCell(7).font = { color: { argb: 'FFC53B3B' }, italic: true };
    });
    const first = H + 1, last = H + rows.length, T = last + 2;
    const tr = ws.getRow(T);
    tr.getCell(2).value = 'TOTAL';
    tr.getCell(6).value = rows.length ? { formula: `SUM(F${first}:F${last})`, result: rows.reduce((a, x) => a + (+x.amount || 0), 0) } : 0;
    tr.getCell(6).numFmt = `"${C} "#,##0.00`;
    for (let c = 1; c <= 7; c++) { const cell = tr.getCell(c); cell.font = { bold: true, size: 12 }; cell.border = { top: { style: 'medium', color: acc }, bottom: { style: 'medium', color: acc } }; }
    tr.height = 22;
    if (rows.length) ws.autoFilter = { from: { row: H, column: 1 }, to: { row: last, column: 7 } };

    // Summary — formulas point back to the Expenses sheet, so edits there update these totals.
    const sm = wb.addWorksheet('Summary', { pageSetup: { orientation: 'portrait', paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
    sm.getColumn(1).width = 36; sm.getColumn(2).width = 10; sm.getColumn(3).width = 18;
    sm.getCell('A1').value = info.title; sm.getCell('A1').font = { bold: true, size: 15 };
    let r = 3;
    const block = (title, colLetter, key) => {
      const names = [...new Set(rows.map(x => x[key]))].sort((a, b) => String(a).localeCompare(String(b)));
      const h = sm.getRow(r);
      [title, 'Items', `Amount (${C})`].forEach((t, i) => { const c = h.getCell(i + 1); c.value = t; c.font = { bold: true, color: { argb: 'FFFFFFFF' } }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: ink }; if (i) c.alignment = { horizontal: 'right' }; });
      const start = ++r;
      for (const n of names) {
        const row = sm.getRow(r), exp = rows.filter(x => x[key] === n);
        row.getCell(1).value = n;
        row.getCell(2).value = { formula: `COUNTIF(Expenses!${colLetter}${first}:${colLetter}${last},A${r})`, result: exp.length };
        row.getCell(3).value = { formula: `SUMIF(Expenses!${colLetter}${first}:${colLetter}${last},A${r},Expenses!F${first}:F${last})`, result: exp.reduce((a, x) => a + (+x.amount || 0), 0) };
        row.getCell(3).numFmt = '#,##0.00';
        for (let c = 1; c <= 3; c++) row.getCell(c).border = { bottom: line };
        r++;
      }
      const t = sm.getRow(r);
      t.getCell(1).value = 'Total';
      t.getCell(2).value = { formula: `SUM(B${start}:B${r - 1})`, result: rows.length };
      t.getCell(3).value = { formula: `SUM(C${start}:C${r - 1})`, result: rows.reduce((a, x) => a + (+x.amount || 0), 0) };
      t.getCell(3).numFmt = '#,##0.00'; t.font = { bold: true };
      for (let c = 1; c <= 3; c++) t.getCell(c).border = { top: { style: 'medium', color: acc } };
      r += 3;
    };
    block('Project', 'B', 'project');
    block('Category', 'E', 'category');
    return new Blob([await wb.xlsx.writeBuffer()], { type: XLSX_TYPE });
  }

  /** Characters the built-in PDF font cannot draw are replaced so the header never fails. */
  const safe = s => String(s ?? '').replace(/[‘’]/g, "'").replace(/[“”]/g, '"').replace(/[–—]/g, '-').replace(/…/g, '...')
    .replace(/[^\x20-\x7E\xA0-\xFF]/g, '?');

  function fitText(font, text, size, max) {
    let t = safe(text);
    if (font.widthOfTextAtSize(t, size) <= max) return t;
    while (t.length > 1 && font.widthOfTextAtSize(t + '...', size) > max) t = t.slice(0, -1);
    return t + '...';
  }

  /**
   * entries: [{ no, dateText, project, category, reason, amountText, files: [{ blob, ext, name }] }]
   * Returns { blob, pages: Map(no → "Page 3" | "Pages 3-4") }. The cover lists every expense
   * with its page number; expenses without a bill are marked "No bill".
   */
  async function makeBillsPdfCore(entries, info, extraCover, onStep) {
    const { PDFDocument, StandardFonts, rgb } = PDFLib;
    // Simple path: cover pages first (pre-sized), then bills.
    const doc = await PDFDocument.create();
    doc.setTitle(safe(info.title + ' - Bills')); doc.setCreator('Billy');
    const font = await doc.embedFont(StandardFonts.Helvetica), bold = await doc.embedFont(StandardFonts.HelveticaBold);
    const A4 = [595.28, 841.89], M = 36, HEAD = 46, inkC = rgb(.086, .094, .114), mutedC = rgb(.4, .42, .45), accC = rgb(.894, .341, .18);
    const covers = [];
    for (let i = 0; i <= extraCover; i++) covers.push(doc.addPage(A4));
    const pagesMap = new Map(), rowsOut = [];
    let n = 0;
    for (const e of entries) {
      onStep?.(++n, entries.length);
      if (!e.files.length) { rowsOut.push([e, 'No bill']); continue; }
      const start = doc.getPageCount() + 1;
      for (const f of e.files) {
        try {
          if (f.ext === 'pdf') {
            const src = await PDFDocument.load(await f.blob.arrayBuffer(), { ignoreEncryption: true });
            for (const p of await doc.copyPages(src, src.getPageIndices())) {
              const pg = doc.addPage(p); const { width: w, height: h } = pg.getSize(); const sc = (h - HEAD - 6) / h;
              pg.scaleContent(sc, sc); pg.translateContent((w - w * sc) / 2, 0); headerOn(pg, e);
            }
          } else {
            const b = await f.blob.arrayBuffer(); const img = f.ext === 'png' ? await doc.embedPng(b) : await doc.embedJpg(b);
            const pg = doc.addPage(A4); const sc = Math.min((A4[0] - 2 * M) / img.width, (A4[1] - HEAD - 2 * M + 10) / img.height, 2);
            pg.drawImage(img, { x: (A4[0] - img.width * sc) / 2, y: A4[1] - HEAD - 14 - img.height * sc, width: img.width * sc, height: img.height * sc }); headerOn(pg, e);
          }
        } catch (err) { const pg = doc.addPage(A4); headerOn(pg, e); pg.drawText('This bill could not be added. It is in the Bills folder in OneDrive.', { x: M, y: A4[1] / 2, size: 11, font, color: mutedC }); }
      }
      const end = doc.getPageCount(), lbl = end > start ? `Pages ${start}-${end}` : `Page ${start}`;
      pagesMap.set(e.no, lbl); rowsOut.push([e, lbl]);
    }
    function headerOn(page, e) {
      const w = page.getWidth(), h = page.getHeight();
      page.drawRectangle({ x: 0, y: h - HEAD, width: w, height: HEAD, color: rgb(1, 1, 1) });
      page.drawRectangle({ x: 0, y: h - HEAD, width: 5, height: HEAD, color: accC });
      page.drawText(fitText(bold, `No. ${e.no}  -  ${e.dateText}  -  ${e.project}`, 11, w - 2 * M - 120), { x: M, y: h - 20, size: 11, font: bold, color: inkC });
      const amt = safe(e.amountText); page.drawText(amt, { x: w - M - bold.widthOfTextAtSize(amt, 11), y: h - 20, size: 11, font: bold, color: inkC });
      page.drawText(fitText(font, `${e.category}  -  ${e.reason}`, 9, w - 2 * M), { x: M, y: h - 35, size: 9, font, color: mutedC });
    }
    let ci = 0, page = covers[0], y = A4[1] - M - 10;
    page.drawText('BILLS', { x: M, y, size: 22, font: bold, color: inkC }); y -= 22;
    page.drawText(safe(info.title), { x: M, y, size: 12, font, color: mutedC }); y -= 30;
    const colX = [M, M + 34, M + 104, A4[0] - M - 150, A4[0] - M - 60];
    const th = () => { ['No.', 'Date', 'Project / reason', 'Amount', 'Bill'].forEach((t, i) => page.drawText(t, { x: colX[i], y, size: 9, font: bold, color: mutedC })); y -= 20; };
    th();
    for (const [e, lbl] of rowsOut) {
      if (y < M + 30 && ci < covers.length - 1) { page = covers[++ci]; y = A4[1] - M - 10; th(); }
      page.drawText(String(e.no), { x: colX[0], y, size: 9.5, font, color: inkC });
      page.drawText(safe(e.dateText), { x: colX[1], y, size: 9.5, font, color: inkC });
      page.drawText(fitText(font, `${e.project} - ${e.reason}`, 9.5, colX[3] - colX[2] - 10), { x: colX[2], y, size: 9.5, font, color: inkC });
      page.drawText(fitText(font, e.amountText, 9.5, 85), { x: colX[3], y, size: 9.5, font, color: inkC });
      page.drawText(safe(lbl), { x: colX[4], y, size: 9.5, font: lbl === 'No bill' ? font : bold, color: lbl === 'No bill' ? accC : inkC });
      y -= 16;
    }
    const total = entries.reduce((a, e) => a + (+e.amount || 0), 0);
    if (total) {
      page.drawLine({ start: { x: M, y: y + 8 }, end: { x: A4[0] - M, y: y + 8 }, thickness: .6, color: mutedC });
      page.drawText('TOTAL', { x: colX[2], y: y - 6, size: 10.5, font: bold, color: inkC });
      page.drawText(safe(`${info.currency || 'AED'} ${money2(total)}`), { x: colX[3], y: y - 6, size: 10.5, font: bold, color: inkC });
    }
    return { blob: new Blob([await doc.save()], { type: 'application/pdf' }), pages: pagesMap };
  }

  return { makeExcel, makeBillsPdf: (entries, info, onStep) => {
    // Cover holds ~44 rows on the first page and ~48 on each extra page.
    const extra = entries.length <= 44 ? 0 : Math.ceil((entries.length - 44) / 48);
    return (async () => { await loadScript(CDN.pdf); return makeBillsPdfCore(entries, info, extra, onStep); })();
  }, nice, money2, XLSX_TYPE };
})();
