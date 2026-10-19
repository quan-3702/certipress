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
// ---------- Import a list from Excel, CSV or the clipboard (shared by the Press apps) ----------
// importList({ title, fields: [{ key, label, required, guess: /regex/ }], onDone(rows) })
// rows = [{ key: value, ... }] – only non-empty rows.
const DataIn = (() => {
  const XLSX_URL = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
  let xlsxReady = null;
  function loadXlsx() {
    if (window.XLSX) return Promise.resolve(window.XLSX);
    if (xlsxReady) return xlsxReady;
    xlsxReady = new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = XLSX_URL; s.onload = () => res(window.XLSX); s.onerror = () => { xlsxReady = null; rej(new Error('Excel reader could not be loaded – connect to the internet once, or save the sheet as CSV.')); };
      document.head.appendChild(s);
    });
    return xlsxReady;
  }
  function parseDelimited(text) {
    text = String(text || '').replace(/^﻿/, '').replace(/\r\n?/g, '\n');
    const first = text.split('\n').slice(0, 5).join('\n');
    const count = (ch) => (first.match(new RegExp(ch === '\t' ? '\t' : '\\' + ch, 'g')) || []).length;
    const delim = ['\t', ';', ','].map((d) => [d, count(d)]).sort((a, b) => b[1] - a[1])[0];
    const d = delim[1] ? delim[0] : '\t';
    const rows = []; let row = [], cell = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (q) {
        if (c === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; } else cell += c;
      } else if (c === '"' && cell === '') q = true;
      else if (c === d) { row.push(cell); cell = ''; }
      else if (c === '\n') { row.push(cell); rows.push(row); row = []; cell = ''; }
      else cell += c;
    }
    if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
    return rows.map((r) => r.map((x) => x.trim())).filter((r) => r.some((x) => x));
  }
  async function readFile(file) {
    if (/\.(xlsx|xlsm|xls|ods)$/i.test(file.name)) {
      const X = await loadXlsx();
      const wb = X.read(await file.arrayBuffer(), { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      return X.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' }).map((r) => r.map((x) => String(x == null ? '' : x).trim())).filter((r) => r.some((x) => x));
    }
    const buf = await file.arrayBuffer();
    let text = new TextDecoder('utf-8').decode(buf);
    if (text.includes('�')) text = new TextDecoder('windows-1252').decode(buf); // old Excel CSV
    return parseDelimited(text);
  }
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  function css() {
    if (document.getElementById('dataInCss')) return;
    const st = document.createElement('style'); st.id = 'dataInCss';
    st.textContent = `
.di-dlg{max-width:760px;width:calc(100vw - 32px)}
.di-dlg h2{margin:0 0 4px}
.di-src{display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin:10px 0}
.di-dlg textarea{width:100%;min-height:90px;font:13px/1.4 Consolas,monospace;padding:8px;border-radius:8px;border:1px solid #8886;background:transparent;color:inherit;box-sizing:border-box}
.di-map{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:8px;margin:10px 0}
.di-map label{display:flex;flex-direction:column;gap:3px;font-size:13px;font-weight:600}
.di-map select{padding:6px;border-radius:8px;border:1px solid #8886;background:transparent;color:inherit}
.di-map select option{color:#111}
.di-prev{max-height:220px;overflow:auto;border:1px solid #8885;border-radius:8px}
.di-prev table{border-collapse:collapse;width:100%;font-size:13px}
.di-prev th,.di-prev td{border-bottom:1px solid #8884;padding:4px 8px;text-align:left;white-space:nowrap;max-width:220px;overflow:hidden;text-overflow:ellipsis}
.di-prev th{position:sticky;top:0;background:Canvas}
.di-msg{font-size:13px;opacity:.8;margin:6px 0}
.di-err{color:#dc2626}
.di-btns{display:flex;gap:8px;justify-content:flex-end;margin-top:12px}`;
    document.head.appendChild(st);
  }
  function importList(opts) {
    css();
    let dlg = document.getElementById('dataInDlg');
    if (dlg) dlg.remove();
    dlg = document.createElement('dialog'); dlg.id = 'dataInDlg'; dlg.className = 'dlg di-dlg';
    dlg.innerHTML = `<h2>${esc(opts.title || 'Import a list')}</h2>
      <p class="di-msg">Open an Excel or CSV file, or copy the cells in Excel / Google Sheets and paste them below.</p>
      <div class="di-src"><button type="button" class="primary" data-a="file">📂 Open Excel or CSV file…</button><span class="di-msg" data-r="fname"></span></div>
      <textarea data-r="paste" placeholder="…or paste here (Ctrl+V)"></textarea>
      <div data-r="mapBox" hidden>
        <label class="check"><input type="checkbox" data-r="head" checked> The first row has the column names</label>
        <div class="di-map" data-r="map"></div>
        <div class="di-prev" data-r="prev"></div>
        <p class="di-msg" data-r="count"></p>
      </div>
      <p class="di-msg di-err" data-r="err"></p>
      <div class="di-btns"><button type="button" class="ghost" data-a="cancel">Cancel</button><button type="button" class="primary" data-a="ok" disabled>Import</button></div>
      <input type="file" accept=".xlsx,.xls,.xlsm,.ods,.csv,.txt,.tsv" hidden data-r="file">`;
    document.body.appendChild(dlg);
    const R = (k) => dlg.querySelector(`[data-r="${k}"]`), A = (k) => dlg.querySelector(`[data-a="${k}"]`);
    let table = [];
    const cols = () => { const w = Math.max(0, ...table.map((r) => r.length)); return Array.from({ length: w }, (_, i) => (R('head').checked && table[0] && table[0][i]) || `Column ${String.fromCharCode(65 + (i % 26))}`); };
    function guessMap() {
      const names = cols(), used = new Set();
      R('map').innerHTML = opts.fields.map((f) => {
        let pick = -1;
        if (R('head').checked && f.guess) pick = names.findIndex((n, i) => !used.has(i) && f.guess.test(n));
        if (pick < 0 && f.required) pick = names.findIndex((_, i) => !used.has(i));
        if (pick >= 0) used.add(pick);
        return `<label>${esc(f.label)}${f.required ? ' *' : ''}<select data-k="${f.key}"><option value="-1">${f.required ? '– choose –' : '(not used)'}</option>${names.map((n, i) => `<option value="${i}"${i === pick ? ' selected' : ''}>${esc(n)}</option>`).join('')}</select></label>`;
      }).join('');
      R('map').querySelectorAll('select').forEach((s) => (s.onchange = preview));
      preview();
    }
    function mapped() {
      const m = {}; R('map').querySelectorAll('select').forEach((s) => (m[s.dataset.k] = Number(s.value)));
      const body = R('head').checked ? table.slice(1) : table;
      return body.map((r) => { const o = {}; opts.fields.forEach((f) => { o[f.key] = m[f.key] >= 0 ? (r[m[f.key]] || '').trim() : ''; }); return o; })
        .filter((o) => opts.fields.some((f) => f.required && o[f.key]));
    }
    function preview() {
      const rows = mapped(), shown = opts.fields.filter((f) => Number(R('map').querySelector(`[data-k="${f.key}"]`).value) >= 0);
      R('prev').innerHTML = `<table><tr>${shown.map((f) => `<th>${esc(f.label)}</th>`).join('')}</tr>${rows.slice(0, 8).map((o) => `<tr>${shown.map((f) => `<td>${esc(o[f.key])}</td>`).join('')}</tr>`).join('')}</table>`;
      R('count').textContent = rows.length ? `${rows.length} row${rows.length > 1 ? 's' : ''} will be imported${rows.length > 8 ? ' (the first 8 are shown)' : ''}.` : 'No rows yet – choose the right columns.';
      A('ok').disabled = !rows.length;
    }
    function setTable(t, label) {
      table = t; R('err').textContent = '';
      if (!table.length) { R('err').textContent = 'No rows were found.'; R('mapBox').hidden = true; A('ok').disabled = true; return; }
      R('fname').textContent = label || '';
      // a first row that looks like headings
      const h = table[0].join(' ');
      R('head').checked = opts.fields.some((f) => f.guess && f.guess.test(h)) || /name|address|street|city|zip|email|title|date|company/i.test(h);
      R('mapBox').hidden = false; guessMap();
    }
    A('file').onclick = () => R('file').click();
    R('file').onchange = async () => {
      const f = R('file').files[0]; if (!f) return;
      R('fname').textContent = 'Reading…';
      try { setTable(await readFile(f), f.name); } catch (e) { R('fname').textContent = ''; R('err').textContent = e.message || 'That file could not be read.'; }
      R('file').value = '';
    };
    R('paste').oninput = () => { if (R('paste').value.trim()) setTable(parseDelimited(R('paste').value), ''); };
    R('head').onchange = guessMap;
    A('cancel').onclick = () => dlg.close();
    A('ok').onclick = () => { const rows = mapped(); dlg.close(); opts.onDone(rows); };
    dlg.addEventListener('close', () => setTimeout(() => dlg.remove(), 50));
    dlg.showModal();
  }
  return { importList, parseDelimited, readFile };
})();
/*! qrcode-generator 1.4.4 (c) Kazuhiko Arase, MIT license */
var qrcode=function(){var t=function(t,r){var e=t,n=g[r],o=null,i=0,a=null,u=[],f={},c=function(t,r){o=function(t){for(var r=new Array(t),e=0;e<t;e+=1){r[e]=new Array(t);for(var n=0;n<t;n+=1)r[e][n]=null}return r}(i=4*e+17),l(0,0),l(i-7,0),l(0,i-7),s(),h(),d(t,r),e>=7&&v(t),null==a&&(a=p(e,n,u)),w(a,r)},l=function(t,r){for(var e=-1;e<=7;e+=1)if(!(t+e<=-1||i<=t+e))for(var n=-1;n<=7;n+=1)r+n<=-1||i<=r+n||(o[t+e][r+n]=0<=e&&e<=6&&(0==n||6==n)||0<=n&&n<=6&&(0==e||6==e)||2<=e&&e<=4&&2<=n&&n<=4)},h=function(){for(var t=8;t<i-8;t+=1)null==o[t][6]&&(o[t][6]=t%2==0);for(var r=8;r<i-8;r+=1)null==o[6][r]&&(o[6][r]=r%2==0)},s=function(){for(var t=B.getPatternPosition(e),r=0;r<t.length;r+=1)for(var n=0;n<t.length;n+=1){var i=t[r],a=t[n];if(null==o[i][a])for(var u=-2;u<=2;u+=1)for(var f=-2;f<=2;f+=1)o[i+u][a+f]=-2==u||2==u||-2==f||2==f||0==u&&0==f}},v=function(t){for(var r=B.getBCHTypeNumber(e),n=0;n<18;n+=1){var a=!t&&1==(r>>n&1);o[Math.floor(n/3)][n%3+i-8-3]=a}for(n=0;n<18;n+=1){a=!t&&1==(r>>n&1);o[n%3+i-8-3][Math.floor(n/3)]=a}},d=function(t,r){for(var e=n<<3|r,a=B.getBCHTypeInfo(e),u=0;u<15;u+=1){var f=!t&&1==(a>>u&1);u<6?o[u][8]=f:u<8?o[u+1][8]=f:o[i-15+u][8]=f}for(u=0;u<15;u+=1){f=!t&&1==(a>>u&1);u<8?o[8][i-u-1]=f:u<9?o[8][15-u-1+1]=f:o[8][15-u-1]=f}o[i-8][8]=!t},w=function(t,r){for(var e=-1,n=i-1,a=7,u=0,f=B.getMaskFunction(r),c=i-1;c>0;c-=2)for(6==c&&(c-=1);;){for(var g=0;g<2;g+=1)if(null==o[n][c-g]){var l=!1;u<t.length&&(l=1==(t[u]>>>a&1)),f(n,c-g)&&(l=!l),o[n][c-g]=l,-1==(a-=1)&&(u+=1,a=7)}if((n+=e)<0||i<=n){n-=e,e=-e;break}}},p=function(t,r,e){for(var n=A.getRSBlocks(t,r),o=b(),i=0;i<e.length;i+=1){var a=e[i];o.put(a.getMode(),4),o.put(a.getLength(),B.getLengthInBits(a.getMode(),t)),a.write(o)}var u=0;for(i=0;i<n.length;i+=1)u+=n[i].dataCount;if(o.getLengthInBits()>8*u)throw"code length overflow. ("+o.getLengthInBits()+">"+8*u+")";for(o.getLengthInBits()+4<=8*u&&o.put(0,4);o.getLengthInBits()%8!=0;)o.putBit(!1);for(;!(o.getLengthInBits()>=8*u||(o.put(236,8),o.getLengthInBits()>=8*u));)o.put(17,8);return function(t,r){for(var e=0,n=0,o=0,i=new Array(r.length),a=new Array(r.length),u=0;u<r.length;u+=1){var f=r[u].dataCount,c=r[u].totalCount-f;n=Math.max(n,f),o=Math.max(o,c),i[u]=new Array(f);for(var g=0;g<i[u].length;g+=1)i[u][g]=255&t.getBuffer()[g+e];e+=f;var l=B.getErrorCorrectPolynomial(c),h=k(i[u],l.getLength()-1).mod(l);for(a[u]=new Array(l.getLength()-1),g=0;g<a[u].length;g+=1){var s=g+h.getLength()-a[u].length;a[u][g]=s>=0?h.getAt(s):0}}var v=0;for(g=0;g<r.length;g+=1)v+=r[g].totalCount;var d=new Array(v),w=0;for(g=0;g<n;g+=1)for(u=0;u<r.length;u+=1)g<i[u].length&&(d[w]=i[u][g],w+=1);for(g=0;g<o;g+=1)for(u=0;u<r.length;u+=1)g<a[u].length&&(d[w]=a[u][g],w+=1);return d}(o,n)};f.addData=function(t,r){var e=null;switch(r=r||"Byte"){case"Numeric":e=M(t);break;case"Alphanumeric":e=x(t);break;case"Byte":e=m(t);break;case"Kanji":e=L(t);break;default:throw"mode:"+r}u.push(e),a=null},f.isDark=function(t,r){if(t<0||i<=t||r<0||i<=r)throw t+","+r;return o[t][r]},f.getModuleCount=function(){return i},f.make=function(){if(e<1){for(var t=1;t<40;t++){for(var r=A.getRSBlocks(t,n),o=b(),i=0;i<u.length;i++){var a=u[i];o.put(a.getMode(),4),o.put(a.getLength(),B.getLengthInBits(a.getMode(),t)),a.write(o)}var g=0;for(i=0;i<r.length;i++)g+=r[i].dataCount;if(o.getLengthInBits()<=8*g)break}e=t}c(!1,function(){for(var t=0,r=0,e=0;e<8;e+=1){c(!0,e);var n=B.getLostPoint(f);(0==e||t>n)&&(t=n,r=e)}return r}())},f.createTableTag=function(t,r){t=t||2;var e="";e+='<table style="',e+=" border-width: 0px; border-style: none;",e+=" border-collapse: collapse;",e+=" padding: 0px; margin: "+(r=void 0===r?4*t:r)+"px;",e+='">',e+="<tbody>";for(var n=0;n<f.getModuleCount();n+=1){e+="<tr>";for(var o=0;o<f.getModuleCount();o+=1)e+='<td style="',e+=" border-width: 0px; border-style: none;",e+=" border-collapse: collapse;",e+=" padding: 0px; margin: 0px;",e+=" width: "+t+"px;",e+=" height: "+t+"px;",e+=" background-color: ",e+=f.isDark(n,o)?"#000000":"#ffffff",e+=";",e+='"/>';e+="</tr>"}return e+="</tbody>",e+="</table>"},f.createSvgTag=function(t,r,e,n){var o={};"object"==typeof arguments[0]&&(t=(o=arguments[0]).cellSize,r=o.margin,e=o.alt,n=o.title),t=t||2,r=void 0===r?4*t:r,(e="string"==typeof e?{text:e}:e||{}).text=e.text||null,e.id=e.text?e.id||"qrcode-description":null,(n="string"==typeof n?{text:n}:n||{}).text=n.text||null,n.id=n.text?n.id||"qrcode-title":null;var i,a,u,c,g=f.getModuleCount()*t+2*r,l="";for(c="l"+t+",0 0,"+t+" -"+t+",0 0,-"+t+"z ",l+='<svg version="1.1" xmlns="http://www.w3.org/2000/svg"',l+=o.scalable?"":' width="'+g+'px" height="'+g+'px"',l+=' viewBox="0 0 '+g+" "+g+'" ',l+=' preserveAspectRatio="xMinYMin meet"',l+=n.text||e.text?' role="img" aria-labelledby="'+y([n.id,e.id].join(" ").trim())+'"':"",l+=">",l+=n.text?'<title id="'+y(n.id)+'">'+y(n.text)+"</title>":"",l+=e.text?'<description id="'+y(e.id)+'">'+y(e.text)+"</description>":"",l+='<rect width="100%" height="100%" fill="white" cx="0" cy="0"/>',l+='<path d="',a=0;a<f.getModuleCount();a+=1)for(u=a*t+r,i=0;i<f.getModuleCount();i+=1)f.isDark(a,i)&&(l+="M"+(i*t+r)+","+u+c);return l+='" stroke="transparent" fill="black"/>',l+="</svg>"},f.createDataURL=function(t,r){t=t||2,r=void 0===r?4*t:r;var e=f.getModuleCount()*t+2*r,n=r,o=e-r;return I(e,e,(function(r,e){if(n<=r&&r<o&&n<=e&&e<o){var i=Math.floor((r-n)/t),a=Math.floor((e-n)/t);return f.isDark(a,i)?0:1}return 1}))},f.createImgTag=function(t,r,e){t=t||2,r=void 0===r?4*t:r;var n=f.getModuleCount()*t+2*r,o="";return o+="<img",o+=' src="',o+=f.createDataURL(t,r),o+='"',o+=' width="',o+=n,o+='"',o+=' height="',o+=n,o+='"',e&&(o+=' alt="',o+=y(e),o+='"'),o+="/>"};var y=function(t){for(var r="",e=0;e<t.length;e+=1){var n=t.charAt(e);switch(n){case"<":r+="&lt;";break;case">":r+="&gt;";break;case"&":r+="&amp;";break;case'"':r+="&quot;";break;default:r+=n}}return r};return f.createASCII=function(t,r){if((t=t||1)<2)return function(t){t=void 0===t?2:t;var r,e,n,o,i,a=1*f.getModuleCount()+2*t,u=t,c=a-t,g={"██":"█","█ ":"▀"," █":"▄","  ":" "},l={"██":"▀","█ ":"▀"," █":" ","  ":" "},h="";for(r=0;r<a;r+=2){for(n=Math.floor((r-u)/1),o=Math.floor((r+1-u)/1),e=0;e<a;e+=1)i="█",u<=e&&e<c&&u<=r&&r<c&&f.isDark(n,Math.floor((e-u)/1))&&(i=" "),u<=e&&e<c&&u<=r+1&&r+1<c&&f.isDark(o,Math.floor((e-u)/1))?i+=" ":i+="█",h+=t<1&&r+1>=c?l[i]:g[i];h+="\n"}return a%2&&t>0?h.substring(0,h.length-a-1)+Array(a+1).join("▀"):h.substring(0,h.length-1)}(r);t-=1,r=void 0===r?2*t:r;var e,n,o,i,a=f.getModuleCount()*t+2*r,u=r,c=a-r,g=Array(t+1).join("██"),l=Array(t+1).join("  "),h="",s="";for(e=0;e<a;e+=1){for(o=Math.floor((e-u)/t),s="",n=0;n<a;n+=1)i=1,u<=n&&n<c&&u<=e&&e<c&&f.isDark(o,Math.floor((n-u)/t))&&(i=0),s+=i?g:l;for(o=0;o<t;o+=1)h+=s+"\n"}return h.substring(0,h.length-1)},f.renderTo2dContext=function(t,r){r=r||2;for(var e=f.getModuleCount(),n=0;n<e;n++)for(var o=0;o<e;o++)t.fillStyle=f.isDark(n,o)?"black":"white",t.fillRect(n*r,o*r,r,r)},f};t.stringToBytes=(t.stringToBytesFuncs={default:function(t){for(var r=[],e=0;e<t.length;e+=1){var n=t.charCodeAt(e);r.push(255&n)}return r}}).default,t.createStringToBytes=function(t,r){var e=function(){for(var e=S(t),n=function(){var t=e.read();if(-1==t)throw"eof";return t},o=0,i={};;){var a=e.read();if(-1==a)break;var u=n(),f=n()<<8|n();i[String.fromCharCode(a<<8|u)]=f,o+=1}if(o!=r)throw o+" != "+r;return i}(),n="?".charCodeAt(0);return function(t){for(var r=[],o=0;o<t.length;o+=1){var i=t.charCodeAt(o);if(i<128)r.push(i);else{var a=e[t.charAt(o)];"number"==typeof a?(255&a)==a?r.push(a):(r.push(a>>>8),r.push(255&a)):r.push(n)}}return r}};var r,e,n,o,i,a=1,u=2,f=4,c=8,g={L:1,M:0,Q:3,H:2},l=0,h=1,s=2,v=3,d=4,w=5,p=6,y=7,B=(r=[[],[6,18],[6,22],[6,26],[6,30],[6,34],[6,22,38],[6,24,42],[6,26,46],[6,28,50],[6,30,54],[6,32,58],[6,34,62],[6,26,46,66],[6,26,48,70],[6,26,50,74],[6,30,54,78],[6,30,56,82],[6,30,58,86],[6,34,62,90],[6,28,50,72,94],[6,26,50,74,98],[6,30,54,78,102],[6,28,54,80,106],[6,32,58,84,110],[6,30,58,86,114],[6,34,62,90,118],[6,26,50,74,98,122],[6,30,54,78,102,126],[6,26,52,78,104,130],[6,30,56,82,108,134],[6,34,60,86,112,138],[6,30,58,86,114,142],[6,34,62,90,118,146],[6,30,54,78,102,126,150],[6,24,50,76,102,128,154],[6,28,54,80,106,132,158],[6,32,58,84,110,136,162],[6,26,54,82,110,138,166],[6,30,58,86,114,142,170]],e=1335,n=7973,i=function(t){for(var r=0;0!=t;)r+=1,t>>>=1;return r},(o={}).getBCHTypeInfo=function(t){for(var r=t<<10;i(r)-i(e)>=0;)r^=e<<i(r)-i(e);return 21522^(t<<10|r)},o.getBCHTypeNumber=function(t){for(var r=t<<12;i(r)-i(n)>=0;)r^=n<<i(r)-i(n);return t<<12|r},o.getPatternPosition=function(t){return r[t-1]},o.getMaskFunction=function(t){switch(t){case l:return function(t,r){return(t+r)%2==0};case h:return function(t,r){return t%2==0};case s:return function(t,r){return r%3==0};case v:return function(t,r){return(t+r)%3==0};case d:return function(t,r){return(Math.floor(t/2)+Math.floor(r/3))%2==0};case w:return function(t,r){return t*r%2+t*r%3==0};case p:return function(t,r){return(t*r%2+t*r%3)%2==0};case y:return function(t,r){return(t*r%3+(t+r)%2)%2==0};default:throw"bad maskPattern:"+t}},o.getErrorCorrectPolynomial=function(t){for(var r=k([1],0),e=0;e<t;e+=1)r=r.multiply(k([1,C.gexp(e)],0));return r},o.getLengthInBits=function(t,r){if(1<=r&&r<10)switch(t){case a:return 10;case u:return 9;case f:case c:return 8;default:throw"mode:"+t}else if(r<27)switch(t){case a:return 12;case u:return 11;case f:return 16;case c:return 10;default:throw"mode:"+t}else{if(!(r<41))throw"type:"+r;switch(t){case a:return 14;case u:return 13;case f:return 16;case c:return 12;default:throw"mode:"+t}}},o.getLostPoint=function(t){for(var r=t.getModuleCount(),e=0,n=0;n<r;n+=1)for(var o=0;o<r;o+=1){for(var i=0,a=t.isDark(n,o),u=-1;u<=1;u+=1)if(!(n+u<0||r<=n+u))for(var f=-1;f<=1;f+=1)o+f<0||r<=o+f||0==u&&0==f||a==t.isDark(n+u,o+f)&&(i+=1);i>5&&(e+=3+i-5)}for(n=0;n<r-1;n+=1)for(o=0;o<r-1;o+=1){var c=0;t.isDark(n,o)&&(c+=1),t.isDark(n+1,o)&&(c+=1),t.isDark(n,o+1)&&(c+=1),t.isDark(n+1,o+1)&&(c+=1),0!=c&&4!=c||(e+=3)}for(n=0;n<r;n+=1)for(o=0;o<r-6;o+=1)t.isDark(n,o)&&!t.isDark(n,o+1)&&t.isDark(n,o+2)&&t.isDark(n,o+3)&&t.isDark(n,o+4)&&!t.isDark(n,o+5)&&t.isDark(n,o+6)&&(e+=40);for(o=0;o<r;o+=1)for(n=0;n<r-6;n+=1)t.isDark(n,o)&&!t.isDark(n+1,o)&&t.isDark(n+2,o)&&t.isDark(n+3,o)&&t.isDark(n+4,o)&&!t.isDark(n+5,o)&&t.isDark(n+6,o)&&(e+=40);var g=0;for(o=0;o<r;o+=1)for(n=0;n<r;n+=1)t.isDark(n,o)&&(g+=1);return e+=Math.abs(100*g/r/r-50)/5*10},o),C=function(){for(var t=new Array(256),r=new Array(256),e=0;e<8;e+=1)t[e]=1<<e;for(e=8;e<256;e+=1)t[e]=t[e-4]^t[e-5]^t[e-6]^t[e-8];for(e=0;e<255;e+=1)r[t[e]]=e;var n={glog:function(t){if(t<1)throw"glog("+t+")";return r[t]},gexp:function(r){for(;r<0;)r+=255;for(;r>=256;)r-=255;return t[r]}};return n}();function k(t,r){if(void 0===t.length)throw t.length+"/"+r;var e=function(){for(var e=0;e<t.length&&0==t[e];)e+=1;for(var n=new Array(t.length-e+r),o=0;o<t.length-e;o+=1)n[o]=t[o+e];return n}(),n={getAt:function(t){return e[t]},getLength:function(){return e.length},multiply:function(t){for(var r=new Array(n.getLength()+t.getLength()-1),e=0;e<n.getLength();e+=1)for(var o=0;o<t.getLength();o+=1)r[e+o]^=C.gexp(C.glog(n.getAt(e))+C.glog(t.getAt(o)));return k(r,0)},mod:function(t){if(n.getLength()-t.getLength()<0)return n;for(var r=C.glog(n.getAt(0))-C.glog(t.getAt(0)),e=new Array(n.getLength()),o=0;o<n.getLength();o+=1)e[o]=n.getAt(o);for(o=0;o<t.getLength();o+=1)e[o]^=C.gexp(C.glog(t.getAt(o))+r);return k(e,0).mod(t)}};return n}var A=function(){var t=[[1,26,19],[1,26,16],[1,26,13],[1,26,9],[1,44,34],[1,44,28],[1,44,22],[1,44,16],[1,70,55],[1,70,44],[2,35,17],[2,35,13],[1,100,80],[2,50,32],[2,50,24],[4,25,9],[1,134,108],[2,67,43],[2,33,15,2,34,16],[2,33,11,2,34,12],[2,86,68],[4,43,27],[4,43,19],[4,43,15],[2,98,78],[4,49,31],[2,32,14,4,33,15],[4,39,13,1,40,14],[2,121,97],[2,60,38,2,61,39],[4,40,18,2,41,19],[4,40,14,2,41,15],[2,146,116],[3,58,36,2,59,37],[4,36,16,4,37,17],[4,36,12,4,37,13],[2,86,68,2,87,69],[4,69,43,1,70,44],[6,43,19,2,44,20],[6,43,15,2,44,16],[4,101,81],[1,80,50,4,81,51],[4,50,22,4,51,23],[3,36,12,8,37,13],[2,116,92,2,117,93],[6,58,36,2,59,37],[4,46,20,6,47,21],[7,42,14,4,43,15],[4,133,107],[8,59,37,1,60,38],[8,44,20,4,45,21],[12,33,11,4,34,12],[3,145,115,1,146,116],[4,64,40,5,65,41],[11,36,16,5,37,17],[11,36,12,5,37,13],[5,109,87,1,110,88],[5,65,41,5,66,42],[5,54,24,7,55,25],[11,36,12,7,37,13],[5,122,98,1,123,99],[7,73,45,3,74,46],[15,43,19,2,44,20],[3,45,15,13,46,16],[1,135,107,5,136,108],[10,74,46,1,75,47],[1,50,22,15,51,23],[2,42,14,17,43,15],[5,150,120,1,151,121],[9,69,43,4,70,44],[17,50,22,1,51,23],[2,42,14,19,43,15],[3,141,113,4,142,114],[3,70,44,11,71,45],[17,47,21,4,48,22],[9,39,13,16,40,14],[3,135,107,5,136,108],[3,67,41,13,68,42],[15,54,24,5,55,25],[15,43,15,10,44,16],[4,144,116,4,145,117],[17,68,42],[17,50,22,6,51,23],[19,46,16,6,47,17],[2,139,111,7,140,112],[17,74,46],[7,54,24,16,55,25],[34,37,13],[4,151,121,5,152,122],[4,75,47,14,76,48],[11,54,24,14,55,25],[16,45,15,14,46,16],[6,147,117,4,148,118],[6,73,45,14,74,46],[11,54,24,16,55,25],[30,46,16,2,47,17],[8,132,106,4,133,107],[8,75,47,13,76,48],[7,54,24,22,55,25],[22,45,15,13,46,16],[10,142,114,2,143,115],[19,74,46,4,75,47],[28,50,22,6,51,23],[33,46,16,4,47,17],[8,152,122,4,153,123],[22,73,45,3,74,46],[8,53,23,26,54,24],[12,45,15,28,46,16],[3,147,117,10,148,118],[3,73,45,23,74,46],[4,54,24,31,55,25],[11,45,15,31,46,16],[7,146,116,7,147,117],[21,73,45,7,74,46],[1,53,23,37,54,24],[19,45,15,26,46,16],[5,145,115,10,146,116],[19,75,47,10,76,48],[15,54,24,25,55,25],[23,45,15,25,46,16],[13,145,115,3,146,116],[2,74,46,29,75,47],[42,54,24,1,55,25],[23,45,15,28,46,16],[17,145,115],[10,74,46,23,75,47],[10,54,24,35,55,25],[19,45,15,35,46,16],[17,145,115,1,146,116],[14,74,46,21,75,47],[29,54,24,19,55,25],[11,45,15,46,46,16],[13,145,115,6,146,116],[14,74,46,23,75,47],[44,54,24,7,55,25],[59,46,16,1,47,17],[12,151,121,7,152,122],[12,75,47,26,76,48],[39,54,24,14,55,25],[22,45,15,41,46,16],[6,151,121,14,152,122],[6,75,47,34,76,48],[46,54,24,10,55,25],[2,45,15,64,46,16],[17,152,122,4,153,123],[29,74,46,14,75,47],[49,54,24,10,55,25],[24,45,15,46,46,16],[4,152,122,18,153,123],[13,74,46,32,75,47],[48,54,24,14,55,25],[42,45,15,32,46,16],[20,147,117,4,148,118],[40,75,47,7,76,48],[43,54,24,22,55,25],[10,45,15,67,46,16],[19,148,118,6,149,119],[18,75,47,31,76,48],[34,54,24,34,55,25],[20,45,15,61,46,16]],r=function(t,r){var e={};return e.totalCount=t,e.dataCount=r,e},e={};return e.getRSBlocks=function(e,n){var o=function(r,e){switch(e){case g.L:return t[4*(r-1)+0];case g.M:return t[4*(r-1)+1];case g.Q:return t[4*(r-1)+2];case g.H:return t[4*(r-1)+3];default:return}}(e,n);if(void 0===o)throw"bad rs block @ typeNumber:"+e+"/errorCorrectionLevel:"+n;for(var i=o.length/3,a=[],u=0;u<i;u+=1)for(var f=o[3*u+0],c=o[3*u+1],l=o[3*u+2],h=0;h<f;h+=1)a.push(r(c,l));return a},e}(),b=function(){var t=[],r=0,e={getBuffer:function(){return t},getAt:function(r){var e=Math.floor(r/8);return 1==(t[e]>>>7-r%8&1)},put:function(t,r){for(var n=0;n<r;n+=1)e.putBit(1==(t>>>r-n-1&1))},getLengthInBits:function(){return r},putBit:function(e){var n=Math.floor(r/8);t.length<=n&&t.push(0),e&&(t[n]|=128>>>r%8),r+=1}};return e},M=function(t){var r=a,e=t,n={getMode:function(){return r},getLength:function(t){return e.length},write:function(t){for(var r=e,n=0;n+2<r.length;)t.put(o(r.substring(n,n+3)),10),n+=3;n<r.length&&(r.length-n==1?t.put(o(r.substring(n,n+1)),4):r.length-n==2&&t.put(o(r.substring(n,n+2)),7))}},o=function(t){for(var r=0,e=0;e<t.length;e+=1)r=10*r+i(t.charAt(e));return r},i=function(t){if("0"<=t&&t<="9")return t.charCodeAt(0)-"0".charCodeAt(0);throw"illegal char :"+t};return n},x=function(t){var r=u,e=t,n={getMode:function(){return r},getLength:function(t){return e.length},write:function(t){for(var r=e,n=0;n+1<r.length;)t.put(45*o(r.charAt(n))+o(r.charAt(n+1)),11),n+=2;n<r.length&&t.put(o(r.charAt(n)),6)}},o=function(t){if("0"<=t&&t<="9")return t.charCodeAt(0)-"0".charCodeAt(0);if("A"<=t&&t<="Z")return t.charCodeAt(0)-"A".charCodeAt(0)+10;switch(t){case" ":return 36;case"$":return 37;case"%":return 38;case"*":return 39;case"+":return 40;case"-":return 41;case".":return 42;case"/":return 43;case":":return 44;default:throw"illegal char :"+t}};return n},m=function(r){var e=f,n=t.stringToBytes(r),o={getMode:function(){return e},getLength:function(t){return n.length},write:function(t){for(var r=0;r<n.length;r+=1)t.put(n[r],8)}};return o},L=function(r){var e=c,n=t.stringToBytesFuncs.SJIS;if(!n)throw"sjis not supported.";!function(t,r){var e=n("友");if(2!=e.length||38726!=(e[0]<<8|e[1]))throw"sjis not supported."}();var o=n(r),i={getMode:function(){return e},getLength:function(t){return~~(o.length/2)},write:function(t){for(var r=o,e=0;e+1<r.length;){var n=(255&r[e])<<8|255&r[e+1];if(33088<=n&&n<=40956)n-=33088;else{if(!(57408<=n&&n<=60351))throw"illegal char at "+(e+1)+"/"+n;n-=49472}n=192*(n>>>8&255)+(255&n),t.put(n,13),e+=2}if(e<r.length)throw"illegal char at "+(e+1)}};return i},D=function(){var t=[],r={writeByte:function(r){t.push(255&r)},writeShort:function(t){r.writeByte(t),r.writeByte(t>>>8)},writeBytes:function(t,e,n){e=e||0,n=n||t.length;for(var o=0;o<n;o+=1)r.writeByte(t[o+e])},writeString:function(t){for(var e=0;e<t.length;e+=1)r.writeByte(t.charCodeAt(e))},toByteArray:function(){return t},toString:function(){var r="";r+="[";for(var e=0;e<t.length;e+=1)e>0&&(r+=","),r+=t[e];return r+="]"}};return r},S=function(t){var r=t,e=0,n=0,o=0,i={read:function(){for(;o<8;){if(e>=r.length){if(0==o)return-1;throw"unexpected end of file./"+o}var t=r.charAt(e);if(e+=1,"="==t)return o=0,-1;t.match(/^\s$/)||(n=n<<6|a(t.charCodeAt(0)),o+=6)}var i=n>>>o-8&255;return o-=8,i}},a=function(t){if(65<=t&&t<=90)return t-65;if(97<=t&&t<=122)return t-97+26;if(48<=t&&t<=57)return t-48+52;if(43==t)return 62;if(47==t)return 63;throw"c:"+t};return i},I=function(t,r,e){for(var n=function(t,r){var e=t,n=r,o=new Array(t*r),i={setPixel:function(t,r,n){o[r*e+t]=n},write:function(t){t.writeString("GIF87a"),t.writeShort(e),t.writeShort(n),t.writeByte(128),t.writeByte(0),t.writeByte(0),t.writeByte(0),t.writeByte(0),t.writeByte(0),t.writeByte(255),t.writeByte(255),t.writeByte(255),t.writeString(","),t.writeShort(0),t.writeShort(0),t.writeShort(e),t.writeShort(n),t.writeByte(0);var r=a(2);t.writeByte(2);for(var o=0;r.length-o>255;)t.writeByte(255),t.writeBytes(r,o,255),o+=255;t.writeByte(r.length-o),t.writeBytes(r,o,r.length-o),t.writeByte(0),t.writeString(";")}},a=function(t){for(var r=1<<t,e=1+(1<<t),n=t+1,i=u(),a=0;a<r;a+=1)i.add(String.fromCharCode(a));i.add(String.fromCharCode(r)),i.add(String.fromCharCode(e));var f,c,g,l=D(),h=(f=l,c=0,g=0,{write:function(t,r){if(t>>>r!=0)throw"length over";for(;c+r>=8;)f.writeByte(255&(t<<c|g)),r-=8-c,t>>>=8-c,g=0,c=0;g|=t<<c,c+=r},flush:function(){c>0&&f.writeByte(g)}});h.write(r,n);var s=0,v=String.fromCharCode(o[s]);for(s+=1;s<o.length;){var d=String.fromCharCode(o[s]);s+=1,i.contains(v+d)?v+=d:(h.write(i.indexOf(v),n),i.size()<4095&&(i.size()==1<<n&&(n+=1),i.add(v+d)),v=d)}return h.write(i.indexOf(v),n),h.write(e,n),h.flush(),l.toByteArray()},u=function(){var t={},r=0,e={add:function(n){if(e.contains(n))throw"dup key:"+n;t[n]=r,r+=1},size:function(){return r},indexOf:function(r){return t[r]},contains:function(r){return void 0!==t[r]}};return e};return i}(t,r),o=0;o<r;o+=1)for(var i=0;i<t;i+=1)n.setPixel(i,o,e(i,o));var a=D();n.write(a);for(var u=function(){var t=0,r=0,e=0,n="",o={},i=function(t){n+=String.fromCharCode(a(63&t))},a=function(t){if(t<0);else{if(t<26)return 65+t;if(t<52)return t-26+97;if(t<62)return t-52+48;if(62==t)return 43;if(63==t)return 47}throw"n:"+t};return o.writeByte=function(n){for(t=t<<8|255&n,r+=8,e+=1;r>=6;)i(t>>>r-6),r-=6},o.flush=function(){if(r>0&&(i(t<<6-r),t=0,r=0),e%3!=0)for(var o=3-e%3,a=0;a<o;a+=1)n+="="},o.toString=function(){return n},o}(),f=a.toByteArray(),c=0;c<f.length;c+=1)u.writeByte(f[c]);return u.flush(),"data:image/gif;base64,"+u};return t}();qrcode.stringToBytesFuncs["UTF-8"]=function(t){return function(t){for(var r=[],e=0;e<t.length;e++){var n=t.charCodeAt(e);n<128?r.push(n):n<2048?r.push(192|n>>6,128|63&n):n<55296||n>=57344?r.push(224|n>>12,128|n>>6&63,128|63&n):(e++,n=65536+((1023&n)<<10|1023&t.charCodeAt(e)),r.push(240|n>>18,128|n>>12&63,128|n>>6&63,128|63&n))}return r}(t)},function(t){"function"==typeof define&&define.amd?define([],t):"object"==typeof exports&&(module.exports=t())}((function(){return qrcode}));
// QR code as an SVG path in a square box (x, y, size in the drawing's units). Dark modules only.
function qrSvg(text, x, y, size, color) {
  if (typeof qrcode === 'undefined' || !text) return '';
  qrcode.stringToBytes = qrcode.stringToBytesFuncs['UTF-8'];
  const q = qrcode(0, 'M'); q.addData(String(text)); q.make();
  const n = q.getModuleCount(), m = 2, k = size / (n + m * 2);
  let d = '';
  for (let r = 0; r < n; r++) for (let c = 0; c < n; c++) if (q.isDark(r, c)) d += `M${(x + (c + m) * k).toFixed(3)} ${(y + (r + m) * k).toFixed(3)}h${k.toFixed(3)}v${k.toFixed(3)}h-${k.toFixed(3)}z`;
  return `<rect x="${x}" y="${y}" width="${size}" height="${size}" fill="#fff"/><path d="${d}" fill="${color || '#000'}"/>`;
}
// ---------- tiny ZIP (stored) and PDF writers, no libraries ----------
const CRC_T = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
function crc32(b) { let c = 0xffffffff; for (let i = 0; i < b.length; i++) c = CRC_T[(c ^ b[i]) & 255] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; }
function utf8(s) { return new TextEncoder().encode(s); }
// Minimal ZIP writer (stored, no compression) – enough for .docx files.
function makeZip(files) {
  const parts = [], central = [];
  let off = 0;
  const u16 = (v) => [v & 255, (v >>> 8) & 255], u32 = (v) => [v & 255, (v >>> 8) & 255, (v >>> 16) & 255, (v >>> 24) & 255];
  for (const [name, content] of files) {
    const nb = utf8(name), data = typeof content === 'string' ? utf8(content) : content, crc = crc32(data);
    const head = [...u32(0x04034b50), ...u16(20), ...u16(0x0800), ...u16(0), ...u16(0), ...u16(0x21), ...u32(crc), ...u32(data.length), ...u32(data.length), ...u16(nb.length), ...u16(0)];
    parts.push(new Uint8Array(head), nb, data);
    central.push(new Uint8Array([...u32(0x02014b50), ...u16(20), ...u16(20), ...u16(0x0800), ...u16(0), ...u16(0), ...u16(0x21), ...u32(crc), ...u32(data.length), ...u32(data.length), ...u16(nb.length), ...u16(0), ...u16(0), ...u16(0), ...u16(0), ...u32(0), ...u32(off)]), nb);
    off += head.length + nb.length + data.length;
  }
  const cdSize = central.reduce((n, p) => n + p.length, 0);
  const end = new Uint8Array([...u32(0x06054b50), ...u16(0), ...u16(0), ...u16(files.length), ...u16(files.length), ...u32(cdSize), ...u32(off), ...u16(0)]);
  const all = parts.concat(central, [end]);
  const out = new Uint8Array(all.reduce((n, p) => n + p.length, 0));
  let p = 0; for (const a of all) { out.set(a, p); p += a.length; }
  return out;
}

// Minimal PDF writer: one JPEG picture per page, page size in millimetres.
function jpegPdf(pages) { // pages: [{ jpeg: Uint8Array, px: [w, h], mm: [w, h] }]
  const enc = (s) => new TextEncoder().encode(s), parts = [], offs = [];
  let len = 0;
  const add = (b) => { const u = typeof b === 'string' ? enc(b) : b; parts.push(u); len += u.length; };
  const obj = (n, body) => { offs[n] = len; add(`${n} 0 obj\n`); body(); add('\nendobj\n'); };
  add('%PDF-1.4\n%\xE2\xE3\xCF\xD3\n');
  const n = pages.length, kids = pages.map((_, i) => `${3 + i * 3} 0 R`).join(' ');
  obj(1, () => add('<< /Type /Catalog /Pages 2 0 R >>'));
  obj(2, () => add(`<< /Type /Pages /Count ${n} /Kids [${kids}] >>`));
  pages.forEach((p, i) => {
    const pg = 3 + i * 3, im = pg + 1, ct = pg + 2, W = p.mm[0] * 72 / 25.4, H = p.mm[1] * 72 / 25.4;
    obj(pg, () => add(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${W.toFixed(2)} ${H.toFixed(2)}] /Resources << /XObject << /Im0 ${im} 0 R >> >> /Contents ${ct} 0 R >>`));
    obj(im, () => { add(`<< /Type /XObject /Subtype /Image /Width ${p.px[0]} /Height ${p.px[1]} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${p.jpeg.length} >>\nstream\n`); add(p.jpeg); add('\nendstream'); });
    const c = `q ${W.toFixed(2)} 0 0 ${H.toFixed(2)} 0 0 cm /Im0 Do Q`;
    obj(ct, () => add(`<< /Length ${c.length} >>\nstream\n${c}\nendstream`));
  });
  const xref = len, total = 3 + n * 3;
  let x = `xref\n0 ${total}\n0000000000 65535 f \n`;
  for (let k = 1; k < total; k++) x += String(offs[k]).padStart(10, '0') + ' 00000 n \n';
  add(x + `trailer\n<< /Size ${total} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF`);
  const out = new Uint8Array(len); let p2 = 0; for (const u of parts) { out.set(u, p2); p2 += u.length; } return out;
}
// An SVG string drawn to a JPEG at `pxPerMm`.
function svgToJpeg(svg, wMm, hMm, pxPerMm) {
  return new Promise((res, rej) => {
    const img = new Image();
    img.onload = () => {
      const c = document.createElement('canvas'); c.width = Math.round(wMm * pxPerMm); c.height = Math.round(hMm * pxPerMm);
      const x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); x.drawImage(img, 0, 0, c.width, c.height);
      c.toBlob(async (b) => b ? res({ jpeg: new Uint8Array(await b.arrayBuffer()), px: [c.width, c.height], mm: [wMm, hMm] }) : rej(new Error('jpeg')), 'image/jpeg', 0.92);
    };
    img.onerror = () => rej(new Error('The page could not be drawn.'));
    img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
  });
}
const safeName = (s, d) => (String(s || '').replace(/[\\/:*?"<>|]+/g, '').replace(/\s+/g, ' ').trim().slice(0, 80) || d);
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
let st = { paper: 'letter', land: true, tpl: 'classic', color: COLORS[0], font: 'classic', head: HEADS[0], pre: 'This certificate is proudly presented to', names: '', reason: 'in recognition of outstanding effort and dedication.', date: today(), showDate: true, sig1: '', sig1t: '', sig2: '', sig2t: '', seal: true, logo: '', people: {}, numOn: false, numPrefix: 'No. ', numStart: 1, qrOn: false };
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
const certNo = (i) => `${st.numPrefix || ''}${String((Number(st.numStart) || 1) + i).padStart(3, '0')}`;
const personOf = (name) => (st.people && st.people[name]) || {};
function certificate(name, screen, idx) {
  idx = idx || 0;
  const who = personOf(name), reasonTxt = who.r || st.reason, dateTxt = who.d || st.date;
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
  const rfs = 5.8 * u, lines = wrapText(reasonTxt, rfs, cw * 0.8, F.body, F.bk).slice(0, st.land ? (st.logo ? 3 : 4) : 7);
  lines.forEach((l, i) => { s += T(l, cx, y + i * rfs * 1.45, rfs, F.body, ' fill="#222"'); });
  // bottom row: date, seal, signatures
  const slots = [];
  if (st.showDate) slots.push({ top: dateTxt, label: 'Date' });
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
  // certificate number and QR code, in the bottom corners inside the border
  const no = st.numOn ? certNo(idx) : '', inner = b.top ? b.top + 3 : 12;
  if (no) s += `<text x="${(w - inner - 2).toFixed(2)}" y="${(h - inner).toFixed(2)}" font-size="${(3.6 * u).toFixed(2)}" font-family='${F.body}' text-anchor="end" fill="#555">${esc(no)}</text>`;
  if (st.qrOn && (name || !screen)) {
    const q = 17 * u, info = [String(st.head || '').trim(), name, no, st.showDate ? dateTxt : '', st.sig1 ? `${st.sig1}${st.sig1t ? ', ' + st.sig1t : ''}` : ''].filter(Boolean).join('\n');
    s += qrSvg(info, +(inner + 1).toFixed(2), +(h - inner - q).toFixed(2), +q.toFixed(2), '#111');
  } else if (st.qrOn && screen) s += `<rect x="${(inner + 1).toFixed(2)}" y="${(h - inner - 17 * u).toFixed(2)}" width="${(17 * u).toFixed(2)}" height="${(17 * u).toFixed(2)}" fill="none" stroke="#b6bcc4" stroke-dasharray="1.5 1.5"/>`;
  return { w, h, svg: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${w} ${h}" width="${w}mm" height="${h}mm" fill="#111"><rect width="${w}" height="${h}" fill="#fff"/>${s}</svg>` };
}
function build(screen) { const n = nameList(); return n.length ? n.map((x, i) => certificate(x, screen, i)) : [certificate('', screen, 0)]; }
function limitReason() {
  if (unlocked()) return '';
  if (tplLocked(st.tpl)) return `The "${TPLS[tplIndex(st.tpl)].name}" design is part of Pro. The first ${FREE_TPLS} designs are free.`;
  if (st.logo) return 'Adding your own logo is part of Pro. Remove the logo to print with the free version.';
  if (st.qrOn) return 'QR codes on certificates are part of Pro. Untick “Add a QR code” to print with the free version.';
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
    $('numOn').checked = !!st.numOn; $('qrOn').checked = !!st.qrOn; $('numRow').hidden = !st.numOn;
    if (document.activeElement !== $('numPrefix')) $('numPrefix').value = st.numPrefix || ''; if (document.activeElement !== $('numStart')) $('numStart').value = st.numStart || 1;
    $('logoDel').hidden = !st.logo; $('logoBtn').textContent = st.logo ? 'Change logo…' : 'Add a logo picture…';
    renderTpls();
    pages = build(true); pageNo = Math.min(pageNo, pages.length - 1);
    const n = nameList().length;
    const own = Object.keys(st.people || {}).filter((k) => nameList().includes(k)).length;
    $('hint').textContent = (n ? `${n} certificate${n === 1 ? '' : 's'} will be printed – one for each name.` : 'No names yet: one certificate with an empty line is printed, ready to fill in by hand.') + (own ? ` ${own} have their own text or date from the imported list.` : '');
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
  // import from Excel / CSV
  $('importBtn').onclick = () => DataIn.importList({
    title: 'Import names from Excel or CSV',
    fields: [
      { key: 'n', label: 'Name', required: true, guess: /name|tên|nombre|nom|student|recipient|participant/i },
      { key: 'r', label: 'What it is for (optional)', guess: /reason|award|for|course|lý do|khóa|achievement|description/i },
      { key: 'd', label: 'Date (optional)', guess: /date|ngày|fecha/i },
    ],
    onDone: (rows) => {
      const people = {}; rows.forEach((o) => { if (o.r || o.d) people[o.n] = { r: o.r, d: o.d }; });
      st.names = rows.map((o) => o.n).join('\n'); st.people = people; pageNo = 0; change();
      toast(`${rows.length} name${rows.length === 1 ? '' : 's'} imported.`);
      if (rows.length > FREE_NAMES && !unlocked()) setTimeout(() => toast(`The free version prints ${FREE_NAMES} at a time – Pro prints the whole list.`), 2500);
    },
  });
  $('numOn').onchange = () => { st.numOn = $('numOn').checked; change(); };
  $('qrOn').onchange = () => { st.qrOn = $('qrOn').checked; change(); if (st.qrOn && !unlocked()) toast('QR codes are part of Pro. You can preview them here.'); };
  $('numPrefix').oninput = () => { st.numPrefix = $('numPrefix').value; change(); };
  $('numStart').oninput = () => { const v = parseInt($('numStart').value, 10); st.numStart = v >= 0 ? v : 1; change(); };
  // one PDF per person, in a ZIP
  $('zipBtn').onclick = async () => {
    if (needsPro('Saving each certificate as its own PDF is part of Pro. Printing, and saving as PDF from the print window, are free.')) return;
    const r = limitReason(); if (r) { openPro(r); return; }
    const list = build(false), names = nameList();
    if (!names.length) { toast('Type or import the names first.'); return; }
    toast(`Making ${list.length} PDF files…`);
    const files = [], used = {};
    for (let i = 0; i < list.length; i++) {
      const p = list[i], pg = await svgToJpeg(p.svg, p.w, p.h, 8);
      let nm = safeName(names[i], 'certificate'); used[nm] = (used[nm] || 0) + 1; if (used[nm] > 1) nm += ' ' + used[nm];
      files.push([nm + '.pdf', jpegPdf([pg])]);
    }
    const zip = makeZip(files), a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([zip], { type: 'application/zip' })); a.download = 'certificates.zip'; a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 60000);
    toast(`Saved certificates.zip (${list.length} PDF files) to your Downloads folder`);
  };
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
