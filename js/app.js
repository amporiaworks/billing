const $ = s => document.querySelector(s);
const store = {
  get(k, d){ try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch(e){ return d; } },
  set(k, v){ try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch(e){ return false; } }
};
let products = store.get('products', []);
const DEFAULTS = {shop:'Amporia Works',
  addr:'Division of Student Welfare, Lovely Professional University, Jalandhar-Delhi G.T. Road, Phagwara, Punjab 144411',
  phone:'+91 76529 49538  |  amporiaworks@gmail.com',
  cur:'₹', note:'Thank you for shopping with Amporia Works!', upi:'var2005@ptyes', payee:'Amporia Works', brand:'#5F2517', accent:'#CAB08D', logo:DEFAULT_LOGO, v:2};
let settings = Object.assign({}, DEFAULTS, store.get('settings', {}));
if (settings.cur === '$') settings.cur = '₹';
if (settings.v !== 2){ settings.brand = DEFAULTS.brand; settings.accent = DEFAULTS.accent; if (!settings.logo) settings.logo = DEFAULTS.logo; settings.v = 2; }
if (settings.shop === 'My Electronics Store'){ settings.shop = DEFAULTS.shop; settings.addr ||= DEFAULTS.addr; settings.phone ||= DEFAULTS.phone; }
let cart = [];
let invoiced = false;

const esc = s => String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const fmt = n => settings.cur + (Number(n)||0).toLocaleString('en-IN',{minimumFractionDigits:2,maximumFractionDigits:2});
function needs(names){ // the file, PDF and QR tools come from the internet
  const missing = names.filter(n => typeof window[n] === 'undefined');
  if (missing.length){ toast('Some tools did not load. Check the internet connection and reload the page.'); return false; }
  return true;
}
function toast(msg){ const s=$('#status'); s.textContent=msg; s.classList.add('show'); clearTimeout(toast.t); toast.t=setTimeout(()=>s.classList.remove('show'),2600); }

/* ---------- price list import ---------- */
function cleanNum(v){
  if (typeof v === 'number') return v;
  const s = String(v ?? '').replace(/rs\.?|inr|usd|[₹$€£,\s]|\/-/gi,'');
  return /^-?\d+(\.\d+)?$/.test(s) ? parseFloat(s) : NaN;
}
const NAME_KEYS = ['description','component','product','item','particular','name'];
const PRICE_KEYS = ['unit price','price','rate','mrp','cost','selling','amount'];
const CODE_KEYS = ['part no','part number','code','sku','model','p/n'];
function findCol(h, keys, skip=[]){
  for (const k of keys){ const i = h.findIndex((x,j) => !skip.includes(j) && x.includes(k)); if (i>=0) return i; }
  return -1;
}
function rowsToProducts(rows){
  let hdr=-1, ni=-1, pi=-1, ci=-1;
  for (let r=0; r<Math.min(rows.length,15); r++){
    const h = rows[r].map(x => String(x).toLowerCase().trim());
    const c = findCol(h, CODE_KEYS);
    const p = findCol(h, PRICE_KEYS, [c]);
    const n = findCol(h, NAME_KEYS, [c,p]);
    if (n>=0 && p>=0){ hdr=r; ni=n; pi=p; ci=c; break; }
  }
  if (hdr < 0){ // no header row: guess the most-text column and the right-most number column
    const cols = Math.max(0, ...rows.map(r => r.length));
    const txt = Array(cols).fill(0), num = Array(cols).fill(0);
    rows.forEach(r => r.forEach((v,i) => { if (v===''||v==null) return; isNaN(cleanNum(v)) ? txt[i]++ : num[i]++; }));
    ni = txt.indexOf(Math.max(...txt));
    num.forEach((c,i) => { if (i!==ni && c>0 && (pi<0 || c>=num[pi])) pi=i; });
  }
  if (ni<0 || pi<0) return [];
  const out = [];
  for (let r=hdr+1; r<rows.length; r++){
    const row = rows[r], name = String(row[ni] ?? '').trim(), price = cleanNum(row[pi]);
    if (!name || isNaN(price) || !/[a-z]/i.test(name)) continue;
    out.push({name, code: ci>=0 ? String(row[ci] ?? '').trim() : '', price});
  }
  return out;
}
async function loadExcel(file){
  const wb = XLSX.read(await file.arrayBuffer(), {type:'array'});
  return wb.SheetNames.flatMap(n => rowsToProducts(XLSX.utils.sheet_to_json(wb.Sheets[n], {header:1, defval:'', raw:true})));
}
async function loadPdf(file){
  pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  const pdf = await pdfjsLib.getDocument({data: await file.arrayBuffer()}).promise;
  const out = [];
  const skip = /^(sub\s*)?total|grand total|page \d/i;
  for (let p=1; p<=pdf.numPages; p++){
    const tc = await (await pdf.getPage(p)).getTextContent();
    const lines = [];
    tc.items.forEach(it => {
      const s = it.str.trim(); if (!s) return;
      const y = it.transform[5], x = it.transform[4];
      let ln = lines.find(l => Math.abs(l.y - y) < 3);
      if (!ln){ ln = {y, cells:[]}; lines.push(ln); }
      ln.cells.push({x, s});
    });
    lines.sort((a,b) => b.y - a.y).forEach(l => {
      const cells = l.cells.sort((a,b) => a.x - b.x).map(c => c.s);
      const nums = cells.map(cleanNum);
      let pi = -1; nums.forEach((n,i) => { if (!isNaN(n)) pi = i; });
      const name = cells.filter((c,i) => i!==pi && /[a-z]/i.test(c)).join(' ').trim();
      if (pi < 0 || !name || skip.test(name)) return;
      out.push({name, code:'', price: nums[pi]});
    });
  }
  return out;
}
$('#uploadBtn').onclick = () => $('#file').click();
$('#file').onchange = async e => {
  const f = e.target.files[0]; if (!f) return;
  toast('Reading ' + f.name + '…');
  try {
    const isPdf = /\.pdf$/i.test(f.name);
    if (!needs([isPdf ? 'pdfjsLib' : 'XLSX'])) { e.target.value=''; return; }
    const list = isPdf ? await loadPdf(f) : await loadExcel(f);
    if (!list.length){ toast('No items with prices found. Check that the file has a name column and a price column.'); return; }
    products = list;
    toast(store.set('products', products) ? 'Loaded ' + list.length + ' items from ' + f.name
      : 'Loaded ' + list.length + ' items, but the list is too large to remember. Load it again next time.');
    renderCount(); renderResults();
  } catch(err){
    console.error(err);
    toast('Could not read this file. Try saving it as .xlsx.');
  }
  e.target.value = '';
};

/* ---------- search ---------- */
let lastList = [];
function renderCount(){ $('#count').textContent = products.length ? products.length + ' items in price list' : 'No price list loaded'; }
function renderResults(){
  const words = $('#q').value.toLowerCase().trim().split(/\s+/).filter(Boolean);
  const box = $('#results');
  if (!products.length){ box.innerHTML = '<div class="empty">Load your price list (Excel, CSV or PDF) to start searching.</div>'; lastList=[]; return; }
  lastList = (words.length ? products.filter(p => { const t=(p.name+' '+p.code).toLowerCase(); return words.every(w => t.includes(w)); }) : products).slice(0,80);
  if (!lastList.length){ box.innerHTML = '<div class="empty">No match. Try fewer words, or add it below as a custom item.</div>'; return; }
  box.innerHTML = lastList.map((p,i) =>
    `<button class="res${i===0&&words.length?' first':''}" data-i="${i}"><span class="nm">${esc(p.name)}${p.code?`<span class="cd">${esc(p.code)}</span>`:''}</span><span class="pr">${fmt(p.price)}</span><span class="add">+</span></button>`
  ).join('');
}
$('#results').onclick = e => { const b = e.target.closest('.res'); if (b) addToCart(lastList[+b.dataset.i]); };
$('#q').oninput = renderResults;
$('#q').onkeydown = e => { if (e.key==='Enter' && lastList[0] && $('#q').value.trim()){ addToCart(lastList[0]); $('#q').select(); } };

/* ---------- cart ---------- */
let feedbackAudio;
function playAddBeep(){
  const AudioContext = window.AudioContext || window.webkitAudioContext;
  if (!AudioContext) return;
  try {
    feedbackAudio ||= new AudioContext();
    const play = () => {
      const now = feedbackAudio.currentTime;
      const oscillator = feedbackAudio.createOscillator();
      const gain = feedbackAudio.createGain();
      oscillator.frequency.value = 880;
      gain.gain.setValueAtTime(0.04, now);
      gain.gain.exponentialRampToValueAtTime(0.001, now + 0.07);
      oscillator.connect(gain); gain.connect(feedbackAudio.destination);
      oscillator.start(now); oscillator.stop(now + 0.07);
    };
    if (feedbackAudio.state === 'running') play();
    else feedbackAudio.resume().then(play).catch(err => console.warn('Could not play item-added feedback.', err));
  } catch(err){
    console.warn('Could not play item-added feedback.', err);
  }
}
function isMobileDevice(){
  const ua = navigator.userAgent || '';
  return /Android|iPhone|iPad|iPod|IEMobile|Windows Phone|Mobile/i.test(ua) ||
    (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
}
function itemAddedFeedback(){
  if (isMobileDevice() && typeof navigator.vibrate === 'function'){
    try { if (navigator.vibrate(40) !== false) return; }
    catch(err){ console.warn('Could not vibrate for item-added feedback.', err); }
  }
  playAddBeep();
}
function addToCart(p){
  const ex = cart.find(c => c.name===p.name && c.price===p.price);
  ex ? ex.qty++ : cart.push({name:p.name, code:p.code||'', price:p.price, qty:1});
  itemAddedFeedback();
  invoiced = false; renderCart(); toast('Added ' + p.name);
}
$('#cAdd').onclick = () => {
  const n = $('#cName').value.trim(), pr = parseFloat($('#cPrice').value);
  if (!n || isNaN(pr)){ toast('Enter an item name and a price.'); return; }
  addToCart({name:n, price:pr}); $('#cName').value=''; $('#cPrice').value='';
};
function totals(){
  const sub = cart.reduce((s,c) => s + c.price*c.qty, 0);
  const disc = Math.min(Math.max(parseFloat($('#discount').value)||0, 0), sub);
  return {sub, disc, total: sub - disc, count: cart.reduce((s,c)=>s+c.qty,0)};
}
function renderCart(){
  $('#cartBody').innerHTML = cart.map((c,i) => `<tr>
    <td>${esc(c.name)}${c.code?`<div class="cd" style="font-size:12px;color:var(--muted)">${esc(c.code)}</div>`:''}</td>
    <td class="num"><input class="qty" type="number" min="1" step="1" value="${c.qty}" data-i="${i}" data-f="qty"></td>
    <td class="num"><input type="number" min="0" step="0.01" value="${c.price}" data-i="${i}" data-f="price"></td>
    <td class="num">${fmt(c.price*c.qty)}</td>
    <td><button class="x" data-del="${i}" aria-label="Remove">×</button></td></tr>`).join('');
  $('#cartEmpty').style.display = cart.length ? 'none' : 'block';
  updateTotals();
  $('#savePdf').disabled = $('#saveImg').disabled = !cart.length;
  $('#payBtn').disabled = !cart.length || !settings.upi;
}
function updateTotals(){
  const t = totals();
  $('#itemsLbl').textContent = t.count + (t.count===1?' item':' items');
  $('#subLbl').textContent = 'Subtotal ' + fmt(t.sub) + (t.disc ? '  −' + fmt(t.disc) : '');
  $('#total').textContent = fmt(t.total);
}
$('#cartBody').onchange = e => {
  const i = e.target.dataset.i; if (i===undefined) return;
  const v = parseFloat(e.target.value);
  if (e.target.dataset.f==='qty') cart[i].qty = Math.max(1, Math.round(v)||1);
  else cart[i].price = Math.max(0, v||0);
  invoiced = false; renderCart();
};
$('#cartBody').onclick = e => { const d = e.target.dataset.del; if (d!==undefined){ cart.splice(+d,1); renderCart(); } };
$('#discount').oninput = () => { invoiced = false; updateTotals(); };
['#custName','#custPhone'].forEach(id => $(id).addEventListener('input', () => { invoiced = false; }));
$('#newBill').onclick = () => {
  if (cart.length && !invoiced && !confirm('This bill has not been saved. Start a new bill anyway?')) return;
  cart=[]; invoiced=false; $('#custName').value=''; $('#custPhone').value=''; $('#discount').value=0; $('#q').value='';
  renderCart(); renderResults(); $('#q').focus();
};

/* ---------- UPI payment QR ---------- */
function upiLink(amount){
  const enc = encodeURIComponent;
  return 'upi://pay?pa=' + settings.upi.trim() + '&pn=' + enc(settings.payee || settings.shop) +
    '&am=' + amount.toFixed(2) + '&cu=INR&tn=' + enc((settings.payee || settings.shop) + ' bill');
}
function qrImage(amount, size){
  return new QRious({value: upiLink(amount), size, level:'M', background:'#ffffff', foreground:'#000000'}).toDataURL('image/png');
}
$('#payBtn').onclick = () => {
  const t = totals();
  if (!settings.upi){ toast('Add your UPI ID in Shop settings first.'); return; }
  if (t.total <= 0){ toast('The total must be more than zero.'); return; }
  if (!needs(['QRious'])) return;
  $('#payAmt').textContent = fmt(t.total);
  $('#payQr').src = qrImage(t.total, 520);
  $('#payTo').textContent = (settings.payee || settings.shop) + '\n' + settings.upi;
  $('#payTo').style.whiteSpace = 'pre-line';
  $('#payDlg').showModal();
};
$('#payClose').onclick = () => $('#payDlg').close();
$('#discount').addEventListener('input', () => { $('#payBtn').disabled = !cart.length || !settings.upi; });

/* ---------- invoice export ---------- */
async function buildInvoice(){
  const t = totals(), d = new Date();
  const name = $('#custName').value.trim(), phone = $('#custPhone').value.trim();
  $('#invoice').innerHTML = `
    <div class="top">
      <div class="brandrow">${settings.logo ? `<img src="${settings.logo}" alt="">` : ''}<div>${settings.logo ? '' : `<div class="shop">${esc(settings.shop)}</div>`}<div class="sub">${esc([settings.addr, settings.phone].filter(Boolean).join('\n'))}</div></div></div>
      <div class="meta"><div class="t">INVOICE</div><div>${d.toLocaleDateString('en-IN')} ${d.toLocaleTimeString('en-IN', {hour:'2-digit',minute:'2-digit'})}</div></div>
    </div>
    <div class="bar"></div>
    ${name||phone ? `<div class="bill"><b>Bill to:</b> ${esc(name)}${name&&phone?' — ':''}${esc(phone)}</div>` : '<div class="bill"></div>'}
    <table><thead><tr><th style="width:40px">#</th><th>Item</th><th class="r">Qty</th><th class="r">Rate</th><th class="r">Amount</th></tr></thead>
    <tbody>${cart.map((c,i) => `<tr><td>${i+1}</td><td>${esc(c.name)}${c.code?` <span style="color:#5f6b66">(${esc(c.code)})</span>`:''}</td><td class="r">${c.qty}</td><td class="r">${fmt(c.price)}</td><td class="r">${fmt(c.price*c.qty)}</td></tr>`).join('')}</tbody></table>
    <div class="payrow">
      ${settings.upi && t.total > 0 && typeof QRious !== 'undefined' ? `<div class="qr"><img src="${qrImage(t.total, 300)}" alt=""><div class="cap"><b>Scan to pay ${fmt(t.total)}</b>Any UPI app<br>${esc(settings.upi)}</div></div>` : '<div></div>'}
      <div class="tot">
        <div><span>Subtotal</span><span>${fmt(t.sub)}</span></div>
        ${t.disc ? `<div><span>Discount</span><span>−${fmt(t.disc)}</span></div>` : ''}
        <div class="g"><span>Total</span><span>${fmt(t.total)}</span></div>
      </div>
    </div>
    ${settings.note ? `<div class="note">${esc(settings.note)}</div>` : ''}`;
  await document.fonts.ready;
  return html2canvas($('#invoice'), {scale:2, backgroundColor:'#ffffff', scrollX:0, scrollY:0,
    onclone: doc => { const el = doc.getElementById('invoice'); el.style.left = '0'; el.style.position = 'absolute'; }});
}
async function saveFile(filename, blob){
  const url = URL.createObjectURL(blob), a = document.createElement('a');
  a.href = url; a.download = filename; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
  toast('Saved ' + filename); return true;
}
function fileBase(){ const n = $('#custName').value.trim().replace(/[^\w\- ]/g,'').replace(/\s+/g,'-'); const d = new Date(), z = x => String(x).padStart(2,'0'); return 'Invoice-' + d.getFullYear() + z(d.getMonth()+1) + z(d.getDate()) + '-' + z(d.getHours()) + z(d.getMinutes()) + (n?'-'+n:''); }
let saving = false;
async function runSave(kind){
  if (!cart.length || saving) return;
  if (!needs(kind === 'pdf' ? ['html2canvas','jspdf'] : ['html2canvas'])) return;
  saving = true; $('#savePdf').disabled = $('#saveImg').disabled = true;
  toast('Preparing invoice…');
  try {
    const canvas = await buildInvoice();
    let blob, name = fileBase();
    if (kind === 'pdf'){
      const pdf = new jspdf.jsPDF({unit:'mm', format:'a4'});
      const w = 190, h = canvas.height * w / canvas.width, pageH = 277, img = canvas.toDataURL('image/png');
      let y = 0;
      pdf.addImage(img, 'PNG', 10, 10, w, h);
      while (h - (y += pageH) > 0){ pdf.addPage(); pdf.addImage(img, 'PNG', 10, 10 - y, w, h); }
      blob = pdf.output('blob'); name += '.pdf';
    } else {
      blob = await new Promise(r => canvas.toBlob(r, 'image/png')); name += '.png';
    }
    if (await saveFile(name, blob)) invoiced = true;
  } catch(err){
    console.error(err); toast('Could not create the invoice. Please try again.');
  } finally {
    saving = false; $('#savePdf').disabled = $('#saveImg').disabled = !cart.length;
  }
}
$('#saveImg').onclick = () => runSave('img');
$('#savePdf').onclick = () => runSave('pdf');

/* ---------- settings ---------- */
function textOn(hex){ const n=parseInt(hex.slice(1),16), r=n>>16, g=n>>8&255, b=n&255; return (r*299+g*587+b*114)/1000 > 150 ? '#1c211f' : '#ffffff'; }
function applySettings(){
  $('#shopTitle').textContent = settings.shop || 'Counter Billing';
  document.title = settings.shop + ' billing';
  const root = document.documentElement.style, inv = $('#invoice').style;
  root.setProperty('--pcb', settings.brand); root.setProperty('--pcb-text', textOn(settings.brand));
  root.setProperty('--copper', settings.accent);
  inv.setProperty('--brand', settings.brand); inv.setProperty('--brand-text', textOn(settings.brand));
  inv.setProperty('--accent', settings.accent);
  updateTotals(); renderResults();
}
let pendingLogo = null;
function showLogoPrev(src){ const p=$('#sLogoPrev'); if (src){ p.src=src; p.style.display='block'; } else p.style.display='none'; }
$('#sLogo').onchange = e => {
  const f = e.target.files[0]; if (!f) return;
  const img = new Image();
  img.onload = () => { // shrink so it stores well and prints sharp
    const s = Math.min(1, 480/img.width, 240/img.height), c = document.createElement('canvas');
    c.width = Math.round(img.width*s); c.height = Math.round(img.height*s);
    c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
    pendingLogo = c.toDataURL('image/png'); showLogoPrev(pendingLogo); URL.revokeObjectURL(img.src);
  };
  img.onerror = () => toast('That file is not an image Chrome can open. Try PNG or JPG.');
  img.src = URL.createObjectURL(f);
};
$('#sLogoReset').onclick = () => { pendingLogo = DEFAULTS.logo; $('#sLogo').value=''; showLogoPrev(pendingLogo); };
$('#sLogoClear').onclick = () => { pendingLogo = ''; $('#sLogo').value=''; showLogoPrev(''); };
$('#settingsBtn').onclick = () => {
  $('#sShop').value=settings.shop; $('#sAddr').value=settings.addr; $('#sPhone').value=settings.phone;
  $('#sCur').value=settings.cur; $('#sNote').value=settings.note;
  $('#sUpi').value=settings.upi||''; $('#sPayee').value=settings.payee||'';
  $('#sBrand').value=settings.brand; $('#sAccent').value=settings.accent;
  pendingLogo = null; $('#sLogo').value=''; showLogoPrev(settings.logo);
  $('#settings').showModal();
};
$('#sCancel').onclick = () => $('#settings').close();
$('#sSave').onclick = () => {
  const upi = $('#sUpi').value.trim();
  if (upi && !/^[a-zA-Z0-9._-]{2,}@[a-zA-Z]{2,}$/.test(upi)){ toast('That UPI ID does not look right. It should look like name@bank.'); return; }
  settings = {shop:$('#sShop').value.trim()||DEFAULTS.shop, addr:$('#sAddr').value.trim(), phone:$('#sPhone').value.trim(),
    cur:$('#sCur').value, note:$('#sNote').value.trim(),
    upi:$('#sUpi').value.trim(), payee:$('#sPayee').value.trim()||DEFAULTS.payee,
    brand:$('#sBrand').value, accent:$('#sAccent').value, logo: pendingLogo===null ? settings.logo : pendingLogo, v:2};
  store.set('settings', settings); $('#settings').close(); applySettings(); renderCart(); toast('Settings saved');
};

renderCount(); applySettings(); renderCart();
if (window.matchMedia('(pointer:fine)').matches) $('#q').focus();
