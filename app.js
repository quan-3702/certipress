const $ = (id) => document.getElementById(id);
const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const FREE_NAMES = 5;
const FREE_TPLS = 3;
const STORE_BILLING = 'https://store.microsoft.com/billing';
const PRO_SKU = 'certipress_pro';
const PRO_IDS = [PRO_SKU];
let STORE_URL = 'https://apps.microsoft.com/detail/9NSLV1NJP0J5';
let proAvailable = false; // Pro limits apply only when the Store actually sells the add-on
let proSku = PRO_SKU;
let isPro = false;
try { isPro = localStorage.getItem('certipress-pro') === '1'; } catch (_) {}
const unlocked = () => isPro || !proAvailable;

async function billing() {
  if (!('getDigitalGoodsService' in window)) return null;
  try { return await window.getDigitalGoodsService(STORE_BILLING); } catch (_) { return null; }
}
function setPro(v) {
  isPro = v;
  try { localStorage.setItem('certipress-pro', v ? '1' : '0'); } catch (_) {}
  const b = $('proBtn');
  b.textContent = v ? '★ Pro' : '★ Get Pro';
  b.classList.toggle('is-pro', v);
  b.hidden = !proAvailable && !v;
  document.body.classList.toggle('unlocked', unlocked());
  if (typeof renderAll === 'function') renderAll();
}
async function checkPro() {
  const svc = await billing();
  if (!svc) return false;
  try {
    const details = await svc.getDetails(PRO_IDS).catch(() => []);
    proAvailable = details.length > 0;
    if (details[0]) { proSku = details[0].itemId; if (!PRO_IDS.includes(proSku)) PRO_IDS.push(proSku); }
    const list = await svc.listPurchases();
    const owned = list.some((p) => PRO_IDS.includes(p.itemId));
    setPro(owned);
    return owned;
  } catch (_) { return isPro; }
}
function proMessage(text) { const m = $('proMsg'); m.hidden = !text; m.textContent = text || ''; }
async function openPro(reason) {
  proMessage(reason || '');
  const svc = await billing();
  if (!svc) {
    $('buyBtn').textContent = 'Get CertiPress on Microsoft Store';
    $('restoreBtn').hidden = true;
  } else {
    $('restoreBtn').hidden = false;
    try {
      const [d] = await svc.getDetails([proSku]);
      const price = d && d.price ? new Intl.NumberFormat(undefined, { style: 'currency', currency: d.price.currency }).format(Number(d.price.value)) : '';
      $('buyBtn').textContent = price ? `Unlock Pro – ${price}` : 'Unlock Pro';
    } catch (_) { $('buyBtn').textContent = 'Unlock Pro'; }
  }
  if (!$('proDialog').open) $('proDialog').showModal();
}
async function buyPro() {
  const svc = await billing();
  if (!svc) { window.open(STORE_URL, '_blank'); return; }
  const methods = [{ supportedMethods: STORE_BILLING, data: { sku: proSku } }];
  try {
    let req;
    try { req = new PaymentRequest(methods); }
    catch (_) { req = new PaymentRequest(methods, { total: { label: 'Total', amount: { currency: 'USD', value: '0' } } }); }
    const res = await req.show();
    await res.complete('success');
  } catch (e) {
    if (e && e.name === 'AbortError') return;
    proMessage('Purchase could not be completed: ' + (e.message || e));
    return;
  }
  if (await checkPro()) { proMessage('Thank you! Pro is unlocked. ★'); setTimeout(() => $('proDialog').close(), 1500); }
}
function needsPro(reason) {
  if (unlocked()) return false;
  openPro(reason);
  return true;
}


// ---------- options ----------
const PAPERS = { letter: { name: 'US Letter (8.5 × 11 in)', w: 279.4, h: 215.9 }, a4: { name: 'A4 (210 × 297 mm)', w: 297, h: 210 } };
const TPLS = [
  { id: 'classic', name: 'Classic' }, { id: 'modern', name: 'Modern' }, { id: 'ribbon', name: 'Ribbon' },
  { id: 'elegant', name: 'Elegant' }, { id: 'stars', name: 'Stars' }, { id: 'diploma', name: 'Diploma' },
  { id: 'coupon', name: 'Gift coupon' }, { id: 'simple', name: 'Simple' },
];
const COLORS = ['#1e3a8a', '#7f1d1d', '#14532d', '#6b21a8', '#9a6a00', '#0f766e', '#be185d', '#1f2937'];
const HEADS = ['Certificate of Achievement', 'Certificate of Appreciation', 'Certificate of Completion', 'Certificate of Participation', 'Certificate of Excellence', 'Employee of the Month', 'Student of the Month', 'Volunteer Award', 'Diploma', 'Gift Certificate'];
const FONTS = {
  classic: { head: 'Georgia, "Times New Roman", serif', name: '"Segoe Script", "Lucida Handwriting", cursive', body: 'Georgia, "Times New Roman", serif', hk: 0.6, nk: 0.64, bk: 0.5 },
  modern: { head: '"Segoe UI", Arial, sans-serif', name: '"Segoe UI", Arial, sans-serif', body: '"Segoe UI", Arial, sans-serif', hk: 0.58, nk: 0.56, bk: 0.5 },
  formal: { head: '"Palatino Linotype", "Book Antiqua", Georgia, serif', name: '"Palatino Linotype", "Book Antiqua", Georgia, serif', body: '"Palatino Linotype", "Book Antiqua", Georgia, serif', hk: 0.6, nk: 0.56, bk: 0.5 },
};
const tplIndex = (id) => Math.max(0, TPLS.findIndex((t) => t.id === id));
const tplLocked = (id) => !unlocked() && tplIndex(id) >= FREE_TPLS;

// ---------- state ----------
const today = () => { try { return new Date().toLocaleDateString(undefined, { year: 'numeric', month: 'long', day: 'numeric' }); } catch (_) { return ''; } };
let st = { paper: 'letter', land: true, tpl: 'classic', color: COLORS[0], font: 'classic', head: HEADS[0], pre: 'This certificate is proudly presented to', names: '', reason: 'in recognition of outstanding effort and dedication.', date: today(), showDate: true, sig1: '', sig1t: '', sig2: '', sig2t: '', seal: true, logo: '' };
try { Object.assign(st, JSON.parse(localStorage.getItem('certipress') || '{}')); } catch (_) {}
let saveTimer;
function save() { clearTimeout(saveTimer); saveTimer = setTimeout(() => { try { localStorage.setItem('certipress', JSON.stringify(st)); } catch (_) { toast('Could not save – the logo picture may be too large'); } }, 250); }
const nameList = () => String(st.names || '').split(/\r?\n/).map((s) => s.trim()).filter(Boolean);
function pageSize() { const p = PAPERS[st.paper] || PAPERS.letter; return st.land ? { w: p.w, h: p.h } : { w: p.h, h: p.w }; }

// ---------- text measuring ----------
let mctx = null;
function textW(s, fs, family, k, weight) { // width in the same unit as fs
  if (typeof document !== 'undefined') {
    try { if (!mctx) mctx = document.createElement('canvas').getContext('2d'); mctx.font = `${weight || 400} 100px ${family}`; return mctx.measureText(s).width * fs / 100; } catch (_) {}
  }
  return s.length * fs * k;
}
function fitSize(s, fs, maxW, family, k, weight) { const w = textW(s, fs, family, k, weight); return w > maxW ? fs * maxW / w : fs; }
function wrapText(s, fs, maxW, family, k) {
  const out = [];
  for (const para of String(s || '').split(/\r?\n/)) {
    let line = '';
    for (const word of para.split(/\s+/).filter(Boolean)) {
      const t = line ? line + ' ' + word : word;
      if (line && textW(t, fs, family, k) > maxW) { out.push(line); line = word; } else line = t;
    }
    if (line) out.push(line);
  }
  return out;
}

// ---------- borders ----------
function border(tpl, w, h, c) {
  const f = (n) => n.toFixed(2);
  const rect = (i, sw, col, extra) => `<rect x="${i}" y="${i}" width="${f(w - 2 * i)}" height="${f(h - 2 * i)}" fill="none" stroke="${col || c}" stroke-width="${sw}"${extra || ''}/>`;
  const gold = '#b8892b';
  if (tpl === 'modern') return { s: `<rect x="0" y="0" width="34" height="${h}" fill="${c}"/><rect x="34" y="0" width="5" height="${h}" fill="${gold}"/><rect x="42" y="0" width="1.2" height="${h}" fill="${c}" opacity=".35"/>`, left: 46, pad: 16 };
  if (tpl === 'ribbon') return { s: `<rect x="0" y="0" width="${w}" height="20" fill="${c}"/><rect x="0" y="20" width="${w}" height="2.5" fill="${gold}"/><rect x="0" y="${f(h - 20)}" width="${w}" height="20" fill="${c}"/><rect x="0" y="${f(h - 22.5)}" width="${w}" height="2.5" fill="${gold}"/>`, left: 0, pad: 26, top: 26 };
  if (tpl === 'elegant') {
    const i = 10, r = 12, x1 = w - i, y1 = h - i;
    const path = (d) => `M${i + r},${i} H${f(x1 - r)} A${r},${r} 0 0 0 ${f(x1)},${i + r} V${f(y1 - r)} A${r},${r} 0 0 0 ${f(x1 - r)},${f(y1)} H${i + r} A${r},${r} 0 0 0 ${i},${f(y1 - r)} V${i + r} A${r},${r} 0 0 0 ${i + r},${i} Z`;
    let s = `<path d="${path()}" fill="none" stroke="${c}" stroke-width="1.4"/>` + rect(14.5, 0.4, gold);
    [[i, i], [x1, i], [i, y1], [x1, y1]].forEach(([x, y]) => { s += `<circle cx="${f(x)}" cy="${f(y)}" r="3.2" fill="${gold}"/>`; });
    return { s, left: 0, pad: 26 };
  }
  if (tpl === 'stars') {
    const star = (cx, cy, r, col) => { let p = ''; for (let k = 0; k < 10; k++) { const a = Math.PI * k / 5 - Math.PI / 2, rr = k % 2 ? r * 0.42 : r; p += `${k ? 'L' : 'M'}${f(cx + rr * Math.cos(a))},${f(cy + rr * Math.sin(a))}`; } return `<path d="${p}Z" fill="${col}"/>`; };
    const cols = [c, '#f59e0b', '#ef4444', '#10b981', '#3b82f6'];
    let s = rect(7, 2.2, c, ' stroke-dasharray="7 4" rx="6"'), n = 0;
    for (let x = 22; x < w - 14; x += 26) { s += star(x, 15, 4.2, cols[n++ % 5]) + star(x + 9, h - 15, 4.2, cols[n++ % 5]); }
    for (let y = 40; y < h - 30; y += 28) { s += star(15, y, 4.2, cols[n++ % 5]) + star(w - 15, y + 10, 4.2, cols[n++ % 5]); }
    return { s, left: 0, pad: 30, top: 26 };
  }
  if (tpl === 'diploma') {
    let s = rect(8, 0.5, c) + rect(18, 0.5, c), d = 5;
    const nx = Math.round((w - 26) / d), ny = Math.round((h - 26) / d);
    for (let k = 0; k <= nx; k++) { const x = 13 + k * (w - 26) / nx; s += `<circle cx="${f(x)}" cy="13" r="3.4" fill="none" stroke="${c}" stroke-width=".3"/><circle cx="${f(x)}" cy="${f(h - 13)}" r="3.4" fill="none" stroke="${c}" stroke-width=".3"/>`; }
    for (let k = 1; k < ny; k++) { const y = 13 + k * (h - 26) / ny; s += `<circle cx="13" cy="${f(y)}" r="3.4" fill="none" stroke="${c}" stroke-width=".3"/><circle cx="${f(w - 13)}" cy="${f(y)}" r="3.4" fill="none" stroke="${c}" stroke-width=".3"/>`; }
    return { s, left: 0, pad: 30 };
  }
  if (tpl === 'coupon') return { s: `<rect x="9" y="9" width="${f(w - 18)}" height="${f(h - 18)}" fill="${c}" opacity=".07"/>` + rect(9, 0.9, c, ' stroke-dasharray="4 3"') + rect(13, 0.35, c), left: 0, pad: 24 };
  if (tpl === 'simple') return { s: `<rect x="20" y="16" width="${f(w - 40)}" height="1.6" fill="${c}"/><rect x="20" y="${f(h - 17.6)}" width="${f(w - 40)}" height="1.6" fill="${c}"/>`, left: 0, pad: 24 };
  let s = rect(8, 2.4, c) + rect(12.2, 0.5, gold) + rect(14.2, 0.5, c); // classic
  [[8, 8], [w - 8, 8], [8, h - 8], [w - 8, h - 8]].forEach(([x, y]) => { s += `<rect x="${f(x - 3.4)}" y="${f(y - 3.4)}" width="6.8" height="6.8" fill="${gold}" transform="rotate(45 ${f(x)} ${f(y)})"/>`; });
  return { s, left: 0, pad: 26 };
}
function sealSvg(cx, cy, r, c) {
  const gold = '#c9992e'; let p = '';
  for (let k = 0; k < 32; k++) { const a = Math.PI * k / 16, rr = k % 2 ? r : r * 0.88; p += `${k ? 'L' : 'M'}${(cx + rr * Math.cos(a)).toFixed(2)},${(cy + rr * Math.sin(a)).toFixed(2)}`; }
  return `<path d="M${cx - r * 0.55},${cy + r * 0.5} l${-r * 0.25},${r * 1.25} l${r * 0.5},${-r * 0.3} l${r * 0.3},${r * 0.45} Z" fill="${c}"/><path d="M${cx + r * 0.55},${cy + r * 0.5} l${r * 0.25},${r * 1.25} l${-r * 0.5},${-r * 0.3} l${-r * 0.3},${r * 0.45} Z" fill="${c}"/><path d="${p}Z" fill="${gold}"/><circle cx="${cx}" cy="${cy}" r="${(r * 0.7).toFixed(2)}" fill="none" stroke="#fff" stroke-width=".5"/><path d="${(() => { let q = ''; for (let k = 0; k < 10; k++) { const a = Math.PI * k / 5 - Math.PI / 2, rr = (k % 2 ? 0.2 : 0.46) * r; q += `${k ? 'L' : 'M'}${(cx + rr * Math.cos(a)).toFixed(2)},${(cy + rr * Math.sin(a)).toFixed(2)}`; } return q; })()}Z" fill="#fff"/>`;
}

// ---------- one certificate (SVG in millimetres) ----------
function certificate(name, screen) {
  const { w, h } = pageSize(), F = FONTS[st.font] || FONTS.classic, c = st.color || COLORS[0];
  const b = border(st.tpl, w, h, c), u = Math.min(w, h) / 210; // u scales letters with the paper
  const x0 = b.left + b.pad, x1 = w - b.pad, cw = x1 - x0, cx = (x0 + x1) / 2;
  const top = (b.top || 20) + 6, T = (t, x, y, fs, fam, extra) => `<text x="${x.toFixed(2)}" y="${y.toFixed(2)}" font-size="${fs.toFixed(2)}" font-family='${fam}' text-anchor="middle"${extra || ''}>${esc(t)}</text>`;
  let s = b.s, y = top;
  if (st.logo) { const lh = 26 * u; s += `<image href="${st.logo}" x="${(cx - 40 * u).toFixed(2)}" y="${y.toFixed(2)}" width="${(80 * u).toFixed(2)}" height="${lh.toFixed(2)}" preserveAspectRatio="xMidYMid meet"/>`; y += lh + 3 * u; } else y += (st.land ? 16 : 30) * u;
  const head = String(st.head || '').trim().toUpperCase();
  if (head) { const fs = fitSize(head, 17 * u, cw * 0.94, F.head, F.hk * 1.25, 700) ; y += fs; s += T(head, cx, y, fs, F.head, ` font-weight="700" fill="${c}"`); y += 5 * u; s += `<rect x="${(cx - 22 * u).toFixed(2)}" y="${y.toFixed(2)}" width="${(44 * u).toFixed(2)}" height="${(0.9 * u).toFixed(2)}" fill="#b8892b"/>`; y += 15 * u; }
  if (st.pre) { const fs = fitSize(st.pre, 6.2 * u, cw * 0.9, F.body, F.bk); s += T(st.pre, cx, y, fs, F.body, ' font-style="italic" fill="#333"'); }
  y += (st.land ? (st.logo ? 26 : 30) : 36) * u;
  const nfs = name ? fitSize(name, 21 * u, cw * 0.86, F.name, F.nk, st.font === 'classic' ? 400 : 700) : 0;
  if (name) s += T(name, cx, y - 2 * u, nfs, F.name, st.font === 'classic' ? '' : ' font-weight="700"');
  else if (screen) s += T('Recipient name', cx, y - 2 * u, 15 * u, F.name, ' fill="#b6bcc4"');
  const lw = Math.min(cw * 0.7, Math.max(90 * u, nfs ? textW(name, nfs, F.name, F.nk) + 20 * u : 0));
  s += `<rect x="${(cx - lw / 2).toFixed(2)}" y="${(y + 3 * u).toFixed(2)}" width="${lw.toFixed(2)}" height="${(0.5 * u).toFixed(2)}" fill="#333"/>`;
  y += (st.land ? (st.logo ? 17 : 20) : 24) * u;
  const rfs = 5.8 * u, lines = wrapText(st.reason, rfs, cw * 0.8, F.body, F.bk).slice(0, st.land ? (st.logo ? 3 : 4) : 7);
  lines.forEach((l, i) => { s += T(l, cx, y + i * rfs * 1.45, rfs, F.body, ' fill="#222"'); });
  // bottom row: date, seal, signatures
  const slots = [];
  if (st.showDate) slots.push({ top: st.date, label: 'Date' });
  if (st.seal) slots.push({ seal: true });
  slots.push({ label: st.sig1 || 'Signature', sub: st.sig1t });
  if (st.sig2 || st.sig2t) slots.push({ label: st.sig2 || 'Signature', sub: st.sig2t });
  const by = h - (b.top ? b.top + 12 : 30) - 6 * u, sw = cw / slots.length;
  slots.forEach((sl, i) => {
    const sx = x0 + sw * (i + 0.5), ll = Math.min(62 * u, sw - 8);
    if (sl.seal) { s += sealSvg(+sx.toFixed(2), +(by - 8 * u).toFixed(2), +(11 * u).toFixed(2), c); return; }
    if (sl.top) s += T(sl.top, sx, by - 2.5 * u, fitSize(sl.top, 5.6 * u, ll, F.body, F.bk), F.body);
    s += `<rect x="${(sx - ll / 2).toFixed(2)}" y="${by.toFixed(2)}" width="${ll.toFixed(2)}" height="${(0.4 * u).toFixed(2)}" fill="#333"/>`;
    s += T(sl.label, sx, by + 6.2 * u, fitSize(sl.label, 5 * u, ll + 6, F.body, F.bk), F.body, sl.top ? ' fill="#555"' : ' font-weight="700"');
    if (sl.sub) s += T(sl.sub, sx, by + 11.8 * u, fitSize(sl.sub, 4.4 * u, ll + 6, F.body, F.bk), F.body, ' fill="#555"');
  });
  return { w, h, svg: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${w} ${h}" width="${w}mm" height="${h}mm" fill="#111"><rect width="${w}" height="${h}" fill="#fff"/>${s}</svg>` };
}
function build(screen) { const n = nameList(); return n.length ? n.map((x) => certificate(x, screen)) : [certificate('', screen)]; }
function limitReason() {
  if (unlocked()) return '';
  if (tplLocked(st.tpl)) return `The "${TPLS[tplIndex(st.tpl)].name}" design is part of Pro. The first ${FREE_TPLS} designs are free.`;
  if (st.logo) return 'Adding your own logo is part of Pro. Remove the logo to print with the free version.';
  if (nameList().length > FREE_NAMES) return `The free version prints up to ${FREE_NAMES} names at a time. You can print as many times as you like.`;
  return '';
}
if (typeof module !== 'undefined') module.exports = { certificate, build, border, wrapText, fitSize, textW, nameList, pageSize, limitReason, TPLS, HEADS, COLORS, PAPERS, getSt: () => st, setSt: (o) => { Object.assign(st, o); }, setProState: (avail, pro) => { proAvailable = avail; isPro = pro; } };

if (typeof document !== 'undefined') (function () {
  let pages = [], pageNo = 0, toastTimer;
  window.toast = function (msg) { const t = $('toast'); t.textContent = msg; t.hidden = false; clearTimeout(toastTimer); toastTimer = setTimeout(() => { t.hidden = true; }, 5000); };
  function renderPage() {
    const p = pages[pageNo];
    $('page').innerHTML = p.svg.replace(/ width="[\d.]+mm" height="[\d.]+mm"/, ` style="aspect-ratio:${p.w}/${p.h};width:min(100%, calc((100vh - 170px) * ${p.w / p.h}))"`);
    $('pager').textContent = pages.length > 1 ? `Certificate ${pageNo + 1} of ${pages.length}` : 'Preview';
    $('prev').disabled = pageNo === 0; $('next').disabled = pageNo >= pages.length - 1;
  }
  function renderTpls() {
    $('tpls').innerHTML = TPLS.map((t) => { const b = border(t.id, 120, 90, st.color); return `<button data-t="${t.id}" class="${t.id === st.tpl ? 'on' : ''}" title="${t.name}" aria-label="${t.name} design"><svg viewBox="0 0 120 90"><rect width="120" height="90" fill="#fff"/>${b.s}</svg>${tplLocked(t.id) ? '<span class="badge">PRO</span>' : ''}</button>`; }).join('');
    $('colors').innerHTML = COLORS.map((c) => `<button data-c="${c}" class="${c === st.color ? 'on' : ''}" style="background:${c}" aria-label="Colour"></button>`).join('');
  }
  window.renderAll = function () {
    ['paper', 'font', 'pre', 'date', 'sig1', 'sig1t', 'sig2', 'sig2t', 'head', 'names', 'reason'].forEach((k) => { if (document.activeElement !== $(k)) $(k).value = st[k]; });
    $('land').value = st.land ? '1' : '0'; $('showDate').checked = !!st.showDate; $('seal').checked = !!st.seal;
    $('logoDel').hidden = !st.logo; $('logoBtn').textContent = st.logo ? 'Change logo…' : 'Add a logo picture…';
    renderTpls();
    pages = build(true); pageNo = Math.min(pageNo, pages.length - 1);
    const n = nameList().length;
    $('hint').textContent = n ? `${n} certificate${n === 1 ? '' : 's'} will be printed – one for each name.` : 'No names yet: one certificate with an empty line is printed, ready to fill in by hand.';
    renderPage();
  };
  const change = () => { save(); renderAll(); };
  function printNow() {
    const r = limitReason(); if (r) { openPro(r); return; }
    const list = build(false), p = list[0];
    $('pageRule').textContent = `@page { size: ${p.w}mm ${p.h}mm; margin: 0; } @media print { #printArea svg { width: ${p.w}mm !important; height: ${p.h}mm !important; } }`;
    $('printArea').innerHTML = list.map((x) => x.svg).join('');
    window.print();
  }
  function savePicture() {
    if (needsPro('Saving a certificate as a picture is part of Pro. Printing, and saving as PDF from the print window, are free.')) return;
    const r = limitReason(); if (r) { openPro(r); return; }
    const p = build(false)[pageNo] || build(false)[0], img = new Image(), k = 8; // 8 px per mm ≈ 200 dpi
    img.onload = () => { const cv = document.createElement('canvas'); cv.width = Math.round(p.w * k); cv.height = Math.round(p.h * k); cv.getContext('2d').drawImage(img, 0, 0, cv.width, cv.height); cv.toBlob((bl) => { const a = document.createElement('a'); a.href = URL.createObjectURL(bl); const nm = (nameList()[pageNo] || 'certificate').replace(/[^\w\- ]+/g, '').trim() || 'certificate'; a.download = nm + '.png'; a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 60000); toast(`Saved ${a.download} to your Downloads folder`); }, 'image/png'); };
    img.onerror = () => toast('The picture could not be created.');
    img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(p.svg);
  }
  function pickLogo() {
    if (needsPro('Adding your own logo is part of Pro.')) return;
    const i = $('fileIn'); i.value = '';
    i.onchange = () => { const f = i.files[0]; if (!f) return; const img = new Image(), url = URL.createObjectURL(f);
      img.onload = () => { const m = 500 / Math.max(img.width, img.height, 500), cv = document.createElement('canvas'); cv.width = Math.round(img.width * m); cv.height = Math.round(img.height * m); cv.getContext('2d').drawImage(img, 0, 0, cv.width, cv.height); st.logo = cv.toDataURL('image/png'); URL.revokeObjectURL(url); change(); };
      img.onerror = () => { URL.revokeObjectURL(url); toast('This file is not a picture that can be used.'); };
      img.src = url; };
    i.click();
  }

  Object.keys(PAPERS).forEach((k) => $('paper').add(new Option(PAPERS[k].name, k)));
  HEADS.forEach((t) => { const o = document.createElement('option'); o.value = t; $('headList').appendChild(o); });
  ['paper', 'font'].forEach((k) => { $(k).onchange = () => { st[k] = $(k).value; change(); }; });
  $('land').onchange = () => { st.land = $('land').value === '1'; change(); };
  ['pre', 'date', 'sig1', 'sig1t', 'sig2', 'sig2t', 'head', 'reason'].forEach((k) => { $(k).oninput = () => { st[k] = $(k).value; change(); }; });
  $('names').oninput = () => { st.names = $('names').value; pageNo = Math.min(pageNo, Math.max(0, nameList().length - 1)); change(); };
  $('showDate').onchange = () => { st.showDate = $('showDate').checked; change(); };
  $('seal').onchange = () => { st.seal = $('seal').checked; change(); };
  $('tpls').addEventListener('click', (e) => { const b = e.target.closest('button'); if (b) { st.tpl = b.dataset.t; change(); if (tplLocked(st.tpl)) toast('This design is part of Pro. You can preview it here.'); } });
  $('colors').addEventListener('click', (e) => { const b = e.target.closest('button'); if (b) { st.color = b.dataset.c; change(); } });
  $('logoBtn').onclick = pickLogo; $('logoDel').onclick = () => { st.logo = ''; change(); };
  $('prev').onclick = () => { if (pageNo > 0) { pageNo--; renderPage(); } };
  $('next').onclick = () => { if (pageNo < pages.length - 1) { pageNo++; renderPage(); } };
  $('printBtn').onclick = printNow; $('pngBtn').onclick = savePicture;
  $('moreBtn').onclick = (e) => { e.stopPropagation(); $('menu').hidden = !$('menu').hidden; };
  document.addEventListener('click', () => { $('menu').hidden = true; });
  $('sampleBtn').onclick = () => { st.names = 'Emily Johnson\nMichael Carter\nSophia Nguyen'; st.sig1 = 'Margaret Ellis'; st.sig1t = 'Principal'; st.sig2 = 'David Brooks'; st.sig2t = 'Class Teacher'; pageNo = 0; change(); };
  $('resetBtn').onclick = () => { if (confirm('Clear all names and texts and start again?')) { try { localStorage.removeItem('certipress'); } catch (_) {} location.reload(); } };
  $('proBtn').onclick = () => { if (!isPro) openPro(''); };
  $('buyBtn').onclick = buyPro;
  $('restoreBtn').onclick = async () => { proMessage((await checkPro()) ? 'Pro is unlocked. ★' : 'No purchase was found on this Microsoft account.'); };
  $('closePro').onclick = () => $('proDialog').close();
  window.addEventListener('resize', () => renderPage());
  setPro(isPro);
  checkPro();
  if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js').catch(() => {});
})();
