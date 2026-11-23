/* Billy — reads a bill, receipt, ticket or invoice and works out what it is for.
 *
 * Two ways to read:
 *   • On the phone (default, free, works without any key): PDF text is read directly; photos and
 *     scanned PDFs go through text recognition (Tesseract.js). Simple rules then find the vendor,
 *     date, total and currency.
 *   • With AI (optional — a free Google Gemini key, or an Anthropic key, saved in Settings): the
 *     bill is sent to the AI with your project list, categories and past habits, and it answers
 *     with every field plus the category, project and a one-line reason.
 *
 * Either way the guess is then improved with what you did before (MEMORY): the same vendor usually
 * goes to the same category and project, and a project name printed on the bill wins.
 */
'use strict';

const SCAN = (() => {
  const CDN = {
    tesseract: 'https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js',
    pdf: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js',
    pdfWorker: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js',
  };
  const loaded = {};
  function loadScript(src) {
    return loaded[src] ||= new Promise((res, rej) => {
      const s = document.createElement('script'); s.src = src; s.async = true;
      s.onload = res; s.onerror = () => { delete loaded[src]; rej(new Error('Could not load the bill reader. Check your internet connection.')); };
      document.head.appendChild(s);
    });
  }
  const isPdf = f => /pdf/i.test(f.type) || /\.pdf$/i.test(f.name || '');

  /* ---------- reading text on the device ---------- */
  async function bitmap(file) {
    try { return await createImageBitmap(file, { imageOrientation: 'from-image' }); }
    catch { const u = URL.createObjectURL(file); const i = new Image(); i.src = u; await i.decode(); return i; }
  }
  function prepare(src, w0, h0) {
    const sc = Math.min(3, 2000 / w0);
    const c = document.createElement('canvas'); c.width = Math.round(w0 * sc); c.height = Math.round(h0 * sc);
    const x = c.getContext('2d', { willReadFrequently: true });
    x.drawImage(src, 0, 0, c.width, c.height);
    const d = x.getImageData(0, 0, c.width, c.height), p = d.data;
    let lo = 255, hi = 0;
    for (let i = 0; i < p.length; i += 16) { const v = (p[i] * 299 + p[i + 1] * 587 + p[i + 2] * 114) / 1000; if (v < lo) lo = v; if (v > hi) hi = v; }
    const span = Math.max(40, hi - lo);
    for (let i = 0; i < p.length; i += 4) {
      const v = (p[i] * 299 + p[i + 1] * 587 + p[i + 2] * 114) / 1000, g = Math.max(0, Math.min(255, (v - lo) * 255 / span));
      p[i] = p[i + 1] = p[i + 2] = g;
    }
    x.putImageData(d, 0, 0);
    return c;
  }
  let worker = null;
  async function ocr(canvas, onProgress) {
    await loadScript(CDN.tesseract);
    if (!worker) worker = await Tesseract.createWorker('eng', 1, { logger: m => { if (m.status === 'recognizing text') onProgress?.(Math.round(m.progress * 100)); } });
    await worker.setParameters({ tessedit_pageseg_mode: '4', preserve_interword_spaces: '1' });
    const { data } = await worker.recognize(canvas);
    return data.text || '';
  }
  async function pdfText(file, onProgress) {
    await loadScript(CDN.pdf);
    pdfjsLib.GlobalWorkerOptions.workerSrc = CDN.pdfWorker;
    const doc = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
    let text = '';
    for (let n = 1; n <= Math.min(doc.numPages, 3); n++) {
      const page = await doc.getPage(n), tc = await page.getTextContent(), rows = [];
      for (const it of tc.items) {
        if (!it.str?.trim()) continue;
        const y = Math.round(it.transform[5]), x = it.transform[4];
        let r = rows.find(r => Math.abs(r.y - y) <= 3); if (!r) rows.push(r = { y, parts: [] });
        r.parts.push({ x, s: it.str });
      }
      text += rows.sort((a, b) => b.y - a.y).map(r => r.parts.sort((a, b) => a.x - b.x).map(p => p.s).join(' ')).join('\n') + '\n';
    }
    if (text.replace(/\s/g, '').length > 60) return text;
    const page = await doc.getPage(1), vp = page.getViewport({ scale: 2.4 });
    const c = document.createElement('canvas'); c.width = vp.width; c.height = vp.height;
    await page.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
    return ocr(prepare(c, c.width, c.height), onProgress);
  }
  async function deviceText(file, onProgress) {
    if (isPdf(file)) return pdfText(file, onProgress);
    const b = await bitmap(file);
    return ocr(prepare(b, b.width || b.naturalWidth, b.height || b.naturalHeight), onProgress);
  }

  /* ---------- what kind of expense is it? ---------- */
  // [pattern, category, vendor name to show (optional)]
  const VENDORS = [
    [/home\s*rent|house\s*rent|flat\s*rent|apartment\s*rent|tenancy|ejari|rent\s*cheque/i, 'Home rent'],
    [/school|tuition|grade\s*\d|\bkg\s*[12]\b|nursery|term\s*fee/i, 'School fees'],
    [/shoot\s*days?|day\s*allowance|per\s*diem/i, 'Shoot day allowance'],
    [/black\s*magic\s*cloud/i, 'Software & subscriptions', 'Blackmagic Cloud'],
    [/etisalat|\be&\b|home\s*internet|elife/i, 'Phone & internet', 'Etisalat'], [/\bdu\b|mobile\s*bill/i, 'Phone & internet', 'du'],
    [/\bemirates\b(?!\s*(nbd|islamic|post|gold))|\bek\s?\d{3}\b/i, 'Air tickets', 'Emirates'], [/fly\s?dubai|\bfz\s?\d{3}\b/i, 'Air tickets', 'flydubai'],
    [/etihad/i, 'Air tickets', 'Etihad'], [/air\s?arabia|\bg9\s?\d{3}\b/i, 'Air tickets', 'Air Arabia'], [/qatar\s*airways/i, 'Air tickets', 'Qatar Airways'],
    [/air\s?india|indigo|vistara|akasa|spicejet/i, 'Air tickets'], [/saudia|gulf\s*air|oman\s*air|turkish\s*airlines|lufthansa|british\s*airways/i, 'Air tickets'],
    [/e-?ticket|boarding\s*pass|itinerary\s*receipt|pnr|booking\s*ref|flight\s*no/i, 'Air tickets'],
    [/careem/i, 'Taxi & transport', 'Careem'], [/\buber\b/i, 'Taxi & transport', 'Uber'], [/\brta\b|dubai\s*taxi|hala\s*taxi|nol\b|metro/i, 'Taxi & transport', 'RTA / Taxi'],
    [/salik/i, 'Parking & tolls', 'Salik'], [/parkin|mawaqif|parking/i, 'Parking & tolls'],
    [/\benoc\b|eppco|adnoc|emarat|petrol|diesel|fuel|super\s*98|special\s*95/i, 'Fuel'],
    [/artlist/i, 'Music & licensing', 'Artlist'], [/musicbed/i, 'Music & licensing', 'Musicbed'], [/epidemic\s*sound/i, 'Music & licensing', 'Epidemic Sound'],
    [/audio\s*jungle|envato/i, 'Music & licensing', 'Envato'], [/soundstripe|premiumbeat|audio\s*network|pond5|motion\s*array|music\s*licen|sync\s*licen/i, 'Music & licensing'],
    [/adobe/i, 'Software & subscriptions', 'Adobe'], [/black\s*magic|davinci/i, 'Software & subscriptions', 'Blackmagic Design'], [/frame\.?io/i, 'Software & subscriptions', 'Frame.io'],
    [/dropbox|google\s*(one|workspace)|microsoft\s*365|apple\.com\/bill|icloud|openai|chatgpt|anthropic|claude|wetransfer|vimeo|canva|figma|notion|zoom|slack|backblaze/i, 'Software & subscriptions'],
    [/shutterstock|getty|storyblocks|artgrid|envato\s*elements|stock\s*footage/i, 'Stock footage & assets'],
    [/marriott|hilton|hyatt|rotana|jumeirah|accor|novotel|ibis|rove|premier\s*inn|booking\.com|agoda|airbnb|hotel|check-?in|nights?\b/i, 'Hotel & stay'],
    [/talabat|deliveroo|noon\s*food|zomato|restaurant|cafe|café|coffee|starbucks|costa|tim\s*hortons|mcdonald|kfc|catering|lunch|dinner|breakfast/i, 'Food & meals'],
    [/rental|\brent(al)?\s*(of|for)?\s*(camera|lens|gear|light|drone|equipment)|hire\b|camera\s*house|gear\s*rental|lens\s*rental|dji\s*rental|lighting\s*hire/i, 'Equipment rental'],
    [/sharaf\s*dg|jumbo|virgin\s*megastore|amazon|noon\.com|b\s*&\s*h|bhphoto|adorama|dji\b|sony|canon|nikon|sandisk|battery|memory\s*card|ssd|hard\s*drive/i, 'Equipment & gear'],
    [/dcaa|gcaa|permit|dubai\s*film|dftc|location\s*fee|municipality|nolle|license\s*fee/i, 'Permits & fees'],
    [/aramex|dhl|fedex|ups\b|courier|printing|print\s*shop/i, 'Courier & printing'],
    [/virgin\s*mobile|recharge|data\s*plan/i, 'Phone & internet'],
    [/ace\s*hardware|dragon\s*mart|props?\b|wardrobe|costume|ikea|home\s*centre|daiso/i, 'Props & set'],
    [/freelance|crew|day\s*rate|assistant|talent|model\s*fee|voice\s*over|vo\s*artist/i, 'Crew & freelancers'],
    [/carrefour|lulu|spinneys|waitrose|union\s*coop|choithrams|supermarket|grocery/i, 'Food & meals'],
  ];
  const CURRENCY = [
    [/\bAED\b|\bDHS?\b|dirham|د\.إ/i, 'AED'], [/\bUSD\b|US\s?\$|\$\s?\d/i, 'USD'], [/\bEUR\b|€/i, 'EUR'], [/\bGBP\b|£/i, 'GBP'],
    [/\bINR\b|₹|\bRs\.?\s?\d/i, 'INR'], [/\bSAR\b|riyal/i, 'SAR'], [/\bOMR\b/i, 'OMR'], [/\bQAR\b/i, 'QAR'], [/\bKWD\b/i, 'KWD'], [/\bBHD\b/i, 'BHD'],
  ];
  const MONTHS = 'janfebmaraprmayjunjulaugsepoctnovdec';
  const pad = n => String(n).padStart(2, '0');
  function dates(t) {
    const out = [];
    const ok = (y, m, d) => { y = +y; m = +m; d = +d; if (y < 100) y += 2000; if (y >= 2015 && y <= 2035 && m >= 1 && m <= 12 && d >= 1 && d <= 31) out.push(`${y}-${pad(m)}-${pad(d)}`); };
    for (const m of t.matchAll(/\b(\d{1,2})[\/.\-](\d{1,2})[\/.\-](20\d{2}|\d{2})\b/g)) ok(m[3], m[2], m[1]);         // 23/11/2026 (UAE style, day first)
    for (const m of t.matchAll(/\b(20\d{2})[\/.\-](\d{1,2})[\/.\-](\d{1,2})\b/g)) ok(m[1], m[2], m[3]);
    for (const m of t.matchAll(/\b(\d{1,2})(?:st|nd|rd|th)?[\s\-]*(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[\s\-,]*(20\d{2}|\d{2})\b/gi)) ok(m[3], MONTHS.indexOf(m[2].toLowerCase()) / 3 + 1, m[1]);
    for (const m of t.matchAll(/\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s+(\d{1,2}),?\s+(20\d{2})\b/gi)) ok(m[3], MONTHS.indexOf(m[1].toLowerCase()) / 3 + 1, m[2]);
    return out;
  }
  function amounts(line) {
    const out = [];
    for (const m of line.matchAll(/(?:^|[^\d.,])(\d{1,3}(?:,\d{3})+(?:\.\d{1,2})?|\d+(?:\.\d{1,2})?)(?![\d,]*\.\d{3})/g)) {
      const v = +m[1].replace(/,/g, ''); if (isFinite(v) && v > 0 && v < 1e7) out.push(v);
    }
    return out;
  }
  const TOTAL_WORDS = /(grand\s*total|total\s*amount|amount\s*(due|paid|payable)|net\s*(amount|total|payable)|total\s*(incl|inc|including)|balance\s*due|total\s*paid|you\s*paid|amount\s*charged|total\s*fare|total\s*price|invoice\s*total|\btotal\b)/i;
  const NOT_TOTAL = /(sub\s*-?\s*total|vat\s*(amount|\d)|tax\s*amount|discount|change|cash\s*tendered|tip)/i;

  function parseText(text) {
    const lines = text.split(/\n/).map(s => s.replace(/\s+/g, ' ').trim()).filter(Boolean);
    let total = null, best = -1;
    lines.forEach((l, i) => {
      if (!TOTAL_WORDS.test(l) || NOT_TOTAL.test(l)) return;
      let nums = amounts(l); if (!nums.length && lines[i + 1]) nums = amounts(lines[i + 1]);
      if (!nums.length) return;
      const v = nums[nums.length - 1];
      const w = (/grand|amount\s*(due|paid|payable)|net|incl|balance|paid|charged|fare|invoice\s*total/i.test(l) ? 3 : 1) + i / lines.length;
      if (w > best || (Math.abs(w - best) < 0.6 && v > total)) { best = w; total = v; }
    });
    if (total == null) {   // no "total" word — take the largest money-looking number
      const all = lines.flatMap(l => /\d{6,}|phone|tel|trn|vat\s*no|invoice\s*no|po\s*box/i.test(l) ? [] : amounts(l).filter(v => String(v).includes('.')));
      if (all.length) total = Math.max(...all);
    }
    let currency = '';
    for (const [re, c] of CURRENCY) if (re.test(text)) { currency = c; break; }
    const ds = dates(text);
    let vendor = '', category = '';
    for (const [re, cat, name] of VENDORS) if (re.test(text)) { category ||= cat; if (name && !vendor) vendor = name; if (vendor && category) break; }
    if (!vendor) {
      const top = lines.slice(0, 6).find(l => /[A-Za-z]{3,}/.test(l) && !/(tax\s*invoice|invoice|receipt|tel|phone|trn|date|www\.|@|^vat)/i.test(l) && l.length < 48);
      vendor = (top || '').replace(/[^\w&.'\- ]+/g, '').trim();
      if (vendor === vendor.toUpperCase()) vendor = vendor.toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
    }
    const today = new Date().toISOString().slice(0, 10);
    const date = ds.filter(d => d <= today).sort().pop() || ds[0] || '';
    return { vendor, date, amount: total, currency, category, docType: /e-?ticket|boarding|itinerary/i.test(text) ? 'Air ticket' : /tax\s*invoice|invoice/i.test(text) ? 'Invoice' : 'Receipt' };
  }

  /* ---------- AI reading (optional key) ---------- */
  const b64 = buf => { let s = '', a = new Uint8Array(buf); for (let i = 0; i < a.length; i += 0x8000) s += String.fromCharCode.apply(null, a.subarray(i, i + 0x8000)); return btoa(s); };
  async function aiPart(file) {
    if (isPdf(file)) return { mime: 'application/pdf', data: b64(await file.arrayBuffer()) };
    const b = await bitmap(file), w = b.width || b.naturalWidth, h = b.height || b.naturalHeight, sc = Math.min(1, 1800 / Math.max(w, h));
    const c = document.createElement('canvas'); c.width = Math.round(w * sc); c.height = Math.round(h * sc);
    c.getContext('2d').drawImage(b, 0, 0, c.width, c.height);
    return { mime: 'image/jpeg', data: c.toDataURL('image/jpeg', 0.88).split(',')[1] };
  }
  function prompt(ctx) {
    return `You are the expense assistant of a video production company in Dubai (UAE). The attached file is a bill, receipt, invoice, e-ticket, booking or licence that one person paid for and will claim back from the office.
Work out what it is and what it was for. Answer with ONE JSON object only, no prose:
{"docType":string,"vendor":string,"date":"YYYY-MM-DD","amount":number,"currency":"AED"|"USD"|"EUR"|"GBP"|"INR"|"SAR"|string,"category":string,"project":string,"reason":string,"confidence":"high"|"medium"|"low"}
Rules:
- amount = the final total actually paid (including VAT), a plain number.
- date = the purchase / invoice date (for flights: the booking or issue date, not the flight date).
- category: choose exactly one of: ${ctx.categories.join(' | ')}.
- project: choose one of these projects if the bill or the person's habits point to it, otherwise "": ${ctx.projects.map(p => JSON.stringify(p)).join(', ') || '(none yet)'}.
- reason: short and specific, like the person writes them, e.g. "Music licence for the brand film", "Flight DXB-COK for the shoot", "Taxi to location". For flights include the route.
${ctx.habits ? 'How this person usually files expenses (vendor → category / project), use it to decide:\n' + ctx.habits : ''}
${ctx.examples ? 'Recent reasons they wrote:\n' + ctx.examples : ''}
Never guess numbers that are not on the document; use "" or 0 instead.`;
  }
  async function callGemini(file, key, text) {
    const p = await aiPart(file);
    let last = 'AI could not read the bill.';
    for (const model of ['gemini-flash-latest', 'gemini-2.5-flash', 'gemini-2.5-flash-lite']) {
      const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`, {
        method: 'POST', headers: { 'content-type': 'application/json', 'x-goog-api-key': key },
        body: JSON.stringify({ contents: [{ parts: [{ inline_data: { mime_type: p.mime, data: p.data } }, { text }] }], generationConfig: { responseMimeType: 'application/json', temperature: 0 } }),
      });
      const j = await res.json().catch(() => ({}));
      if (res.ok) { const t = (j.candidates?.[0]?.content?.parts || []).map(x => x.text || '').join(''); if (t) return t; last = 'AI gave no answer.'; continue; }
      if (res.status === 400 && /API key/i.test(j.error?.message || '')) throw new Error('The AI key was not accepted — check it in Settings.');
      if (res.status === 429) { last = 'Free AI limit reached for now.'; continue; }
      last = j.error?.message || last;
    }
    throw new Error(last);
  }
  async function callClaude(file, key, text) {
    const p = await aiPart(file);
    const block = p.mime === 'application/pdf' ? { type: 'document', source: { type: 'base64', media_type: p.mime, data: p.data } } : { type: 'image', source: { type: 'base64', media_type: p.mime, data: p.data } };
    const res = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'content-type': 'application/json', 'x-api-key': key, 'anthropic-version': '2023-06-01', 'anthropic-dangerous-direct-browser-access': 'true' },
      body: JSON.stringify({ model: 'claude-haiku-5-5', max_tokens: 600, messages: [{ role: 'user', content: [block, { type: 'text', text }] }] }),
    });
    const j = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(res.status === 401 ? 'The AI key was not accepted — check it in Settings.' : (j.error?.message || 'AI could not read the bill.'));
    return (j.content || []).map(c => c.text || '').join('');
  }

  /* ---------- memory: learn from past expenses ---------- */
  const norm = s => String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
  /** expenses: [{vendor, category, projectName, reason, date}] */
  function buildMemory(expenses) {
    const byVendor = new Map(), words = new Map();
    const bump = (m, k, v) => { if (!k || !v) return; const e = m.get(k) || {}; e[v] = (e[v] || 0) + 1; m.set(k, e); };
    for (const x of expenses) {
      const v = norm(x.vendor);
      if (v) { bump(byVendor, v + '|c', x.category); bump(byVendor, v + '|p', x.projectName); }
      for (const w of new Set(norm(x.reason).split(' ').filter(w => w.length > 3))) bump(words, w, x.category);
    }
    return { byVendor, words, recent: expenses.slice().sort((a, b) => (b.date || '').localeCompare(a.date || '')) };
  }
  const top = o => o ? Object.entries(o).sort((a, b) => b[1] - a[1])[0]?.[0] || '' : '';
  function habitsText(mem) {
    const rows = [];
    for (const [k, v] of mem.byVendor) if (k.endsWith('|c')) {
      const name = k.slice(0, -2), p = top(mem.byVendor.get(name + '|p'));
      rows.push(`${name} → ${top(v)}${p ? ' / ' + p : ''}`);
      if (rows.length >= 40) break;
    }
    return rows.join('\n');
  }

  /** Improves a guess with the person's own history and the project names they use. */
  function refine(r, ctx, rawText) {
    const mem = ctx.memory;
    const v = norm(r.vendor);
    if (v && mem) {
      let hitKey = [...mem.byVendor.keys()].find(k => k.endsWith('|c') && (k.slice(0, -2) === v || (v.length > 4 && k.includes(v)) || (k.length > 6 && v.includes(k.slice(0, -2)))));
      if (hitKey) {
        const name = hitKey.slice(0, -2);
        const cat = top(mem.byVendor.get(name + '|c')); if (cat && !r.categoryFromAI) r.category = cat;
        const proj = top(mem.byVendor.get(name + '|p')); if (proj && !r.project) r.project = proj;
        r.learned = true;
      }
    }
    if (!r.category && mem && r.reason) {
      const score = {};
      for (const w of norm(r.reason).split(' ')) { const e = mem.words.get(w); if (e) for (const [c, n] of Object.entries(e)) score[c] = (score[c] || 0) + n; }
      r.category = top(score);
    }
    // a project name printed on the bill wins
    const t = norm(rawText || '');
    if (t) for (const p of ctx.projects) { const pn = norm(p); if (pn.length > 3 && t.includes(pn)) { r.project = p; r.projectOnBill = true; break; } }
    if (r.category && !ctx.categories.includes(r.category)) r.category = ctx.categories.find(c => norm(c) === norm(r.category)) || 'Other';
    if (r.project && !ctx.projects.includes(r.project)) r.project = ctx.projects.find(p => norm(p) === norm(r.project)) || '';
    if (!r.reason) {
      const lbl = { 'Air tickets': 'Flight', 'Taxi & transport': 'Taxi', 'Hotel & stay': 'Hotel stay', 'Music & licensing': 'Music licence', 'Fuel': 'Fuel', 'Parking & tolls': 'Parking', 'Food & meals': 'Meal', 'Equipment rental': 'Equipment rental', 'Software & subscriptions': 'Subscription', 'Phone & internet': 'Bill', 'Home rent': 'Home rent', 'School fees': 'School fees' }[r.category];
      r.reason = [lbl || r.category || '', r.vendor ? (lbl ? '— ' : '') + r.vendor : ''].filter(Boolean).join(' ');
      if (r.project && r.reason) r.reason += ' for ' + r.project;
    }
    return r;
  }

  /**
   * Reads one file. ctx = { categories: [names], projects: [names], memory, aiKey, onProgress }
   * Returns { vendor, date, amount, currency, category, project, reason, docType, ai, learned }.
   */
  async function read(file, ctx) {
    const habits = ctx.memory ? habitsText(ctx.memory) : '';
    const examples = ctx.memory ? ctx.memory.recent.slice(0, 15).map(x => `- ${x.reason}${x.projectName ? ' [' + x.projectName + ']' : ''} (${x.category})`).join('\n') : '';
    if (ctx.aiKey) {
      try {
        const p = prompt({ ...ctx, habits, examples });
        const txt = /^sk-ant-/.test(ctx.aiKey) ? await callClaude(file, ctx.aiKey, p) : await callGemini(file, ctx.aiKey, p);
        const m = /\{[\s\S]*\}/.exec(txt); if (!m) throw new Error('AI gave no answer.');
        const a = JSON.parse(m[0]);
        const r = {
          vendor: String(a.vendor || '').trim(), date: /^\d{4}-\d{2}-\d{2}$/.test(a.date || '') ? a.date : '',
          amount: +String(a.amount ?? '').replace(/[^\d.]/g, '') || null, currency: String(a.currency || '').toUpperCase().slice(0, 3),
          category: String(a.category || ''), project: String(a.project || ''), reason: String(a.reason || '').trim(), docType: String(a.docType || ''),
          ai: true, categoryFromAI: !!a.category, confidence: a.confidence || '',
        };
        return refine(r, ctx, '');
      } catch (e) {
        if (/not accepted/.test(e.message)) throw e;
        console.warn('AI read failed, reading on the phone instead', e);
      }
    }
    const text = await deviceText(file, ctx.onProgress);
    const r = parseText(text);
    return refine({ ...r, project: '', reason: '' }, ctx, text);
  }

  return { read, parseText, buildMemory, refine, VENDORS };
})();
if (typeof module !== 'undefined') module.exports = SCAN;
