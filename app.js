/* Parcel Ledger
 * Reads dispatch scans + the Ledger tab from the Dispatch Google Sheet (same Apps Script, same logins),
 * and, optionally, godown scan-outs and stock from the Godown Stock Firebase project (read-only).
 */
(function () {
'use strict';
const $ = id => document.getElementById(id);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const key = s => String(s ?? '').replace(/[\s'"]/g, '').toUpperCase();
const today = () => new Date().toLocaleDateString('en-CA');
const addDays = (d, n) => { const x = new Date(d + 'T00:00:00'); x.setDate(x.getDate() + n); return x.toLocaleDateString('en-CA'); };
const daysSince = s => s ? Math.floor((Date.now() - new Date(s + 'T00:00:00').getTime()) / 864e5) : null;
const num = v => { if (v == null) return null; const s = String(v).replace(/[₹,\s]/g, ''); if (s === '') return null; const n = Number(s); return isNaN(n) ? null : n; };
const inr = n => n == null || n === '' || isNaN(n) ? '—' : '₹' + Number(n).toLocaleString('en-IN', { maximumFractionDigits: 2 });
const inr0 = n => '₹' + Math.round(Number(n || 0)).toLocaleString('en-IN');
const fmtD = s => { if (!s) return '—'; const d = new Date(String(s).slice(0, 10) + 'T00:00:00'); return isNaN(d) ? s : d.toLocaleDateString('en-IN', { day: '2-digit', month: 'short' }); };
let toastT; function toast(m) { const t = $('toast'); t.textContent = m; t.hidden = false; clearTimeout(toastT); toastT = setTimeout(() => t.hidden = true, 3600); }
const store = {
  get(k, d) { try { const v = localStorage.getItem('pl_' + k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } },
  set(k, v) { try { localStorage.setItem('pl_' + k, JSON.stringify(v)); } catch (e) {} }
};

/* ---------- constants ---------- */
const PLATFORMS = ['Flipkart', 'Meesho', 'Amazon', 'Website / Direct', 'Other'];
const STATUSES = [['dispatched', 'Dispatched', 'p-neutral'], ['transit', 'In transit', 'p-info'], ['ofd', 'Out for delivery', 'p-info'], ['delivered', 'Delivered', 'p-ok'], ['rto', 'RTO (courier return)', 'p-warn'], ['return', 'Customer return', 'p-warn'], ['cancelled', 'Cancelled', 'p-neutral'], ['lost', 'Lost / damaged in transit', 'p-bad']];
const SMAP = Object.fromEntries(STATUSES.map(s => [s[0], { l: s[1], c: s[2] }]));
const RET = { good: ['Back in stock', 'p-ok'], damaged: ['Returned damaged', 'p-bad'], wrong: ['Wrong item returned', 'p-bad'], missing: ['Return not received', 'p-bad'] };
const CLAIM_T = { short: 'Short / wrong settlement', lost: 'Lost in transit', damaged: 'Damaged in transit', 'return-damaged': 'Return received damaged / used', 'return-wrong': 'Wrong / empty return', 'return-missing': 'Return never received', unpaid: 'Delivered but not paid', weight: 'Weight / shipping fee dispute', other: 'Other' };
const CLAIM_S = { open: ['Open', 'p-warn'], approved: ['Approved', 'p-info'], paid: ['Recovered', 'p-ok'], rejected: ['Rejected', 'p-bad'] };
const UNPAID_DAYS = 15, CLAIM_UNPAID_DAYS = 45, STUCK_DAYS = 7;
const FIREBASE_CONFIG = {
  apiKey: 'AIzaSyCj7b7i1hvmtOUUO3_nf7rypGWhfsKrFmA', authDomain: 'godown-stock-9fcaa.firebaseapp.com', projectId: 'godown-stock-9fcaa',
  storageBucket: 'godown-stock-9fcaa.firebasestorage.app', messagingSenderId: '456756323654', appId: '1:456756323654:web:44b2d79601217723eab147'
};
const AUTH_EMAIL_SUFFIX = '@godown-stock.local';

/* ---------- state ---------- */
let url = store.get('url', ''), token = store.get('token', ''), user = store.get('user', null);
let ledger = {}, scans = {}, cancelled = [], packets = {}, rates = {}, sync = null;
let gd = { on: false, products: {}, moves: [], at: null };
let plist = [];

/* ---------- API ---------- */
async function api(method, params) {
  if (!url) throw new Error('Not connected to the Google Sheet.');
  const body = Object.assign({ token }, params);
  let res;
  if (method === 'GET') { const u = new URL(url); Object.entries(body).forEach(([k, v]) => u.searchParams.set(k, v)); res = await fetch(u.toString()); }
  else res = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(body) });
  if (!res.ok) throw new Error('Sheet replied ' + res.status);
  const j = await res.json();
  if (j.auth) { logout(j.error); throw new Error(j.error); }
  if (!j.ok) throw new Error(j.error || 'Sheet error');
  return j;
}

/* ---------- login / connect ---------- */
function showLogin(msg) {
  $('appView').hidden = true; $('loginView').hidden = false;
  $('lMsg').textContent = msg || ''; $('dUrl').value = url;
  $('devDot').className = 'dot ' + (url ? 'on' : 'off'); $('devState').textContent = url ? 'Connected to the Google Sheet' : 'Connect to the Google Sheet';
  $('devBox').open = !url;
}
function logout(msg) { token = ''; user = null; store.set('token', ''); store.set('user', null); showLogin(msg); }
$('devForm').addEventListener('submit', async e => {
  e.preventDefault(); const u = $('dUrl').value.trim(); $('dMsg').className = 'msg'; $('dMsg').textContent = 'Testing…';
  try {
    const r = await fetch(u + (u.includes('?') ? '&' : '?') + 'action=ping'); const j = await r.json();
    if (!j.ok || j.app !== 'dispatch-scanner') throw new Error('That link is not the Dispatch Sheet script.');
    if (!j.ledger) throw new Error('The Sheet still has the old script. Paste the new Code.gs, run setup, and deploy a new version.');
    url = u; store.set('url', u); $('dMsg').className = 'msg ok'; $('dMsg').textContent = 'Connected.'; showLogin('');
  } catch (err) { $('dMsg').textContent = err.message || 'Could not reach that link.'; }
});
$('loginForm').addEventListener('submit', async e => {
  e.preventDefault(); if (!url) { $('lMsg').textContent = 'Connect to the Google Sheet first (below).'; $('devBox').open = true; return; }
  $('lBtn').disabled = true; $('lMsg').textContent = '';
  try {
    const r = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ action: 'login', name: $('lName').value, pin: $('lPin').value }) });
    const j = await r.json(); if (!j.ok) throw new Error(j.error || 'Login failed.');
    token = j.token; user = j.user; store.set('token', token); store.set('user', user); $('lPin').value = ''; start();
  } catch (err) { $('lMsg').textContent = err.message; } finally { $('lBtn').disabled = false; }
});
$('logout').onclick = () => logout('Logged out.');
const isAdmin = () => user && user.role === 'admin';

/* ---------- load ---------- */
$('gFrom').value = store.get('from', addDays(today(), -45)); $('gTo').value = today();
$('reload').onclick = () => { store.set('from', $('gFrom').value); load(); };
async function load() {
  const from = $('gFrom').value || addDays(today(), -45), to = $('gTo').value || today();
  $('loadInfo').textContent = 'Loading…'; $('reload').disabled = true;
  try {
    const j = await api('GET', { action: 'ledger', from, to });
    ledger = {}; (j.ledger || []).forEach(l => { ledger[key(l.awb)] = l; });
    scans = {}; (j.scans || []).forEach(s => { if (s.status === 'Cancelled') return; const k = key(s.awb); if (!scans[k] || (s.ts || 0) > (scans[k].ts || 0)) scans[k] = s; });
    cancelled = j.cancelled || [];
    packets = {}; (j.products || []).forEach(p => { packets[String(p.code).trim().toUpperCase()] = p; });
    rates = {}; (j.rates || []).forEach(r => { rates[key(r.sku)] = r; });
    sync = j.sync || null;
    $('loadInfo').textContent = `${Object.keys(scans).length} scans, ${Object.keys(ledger).length} ledger rows · ${new Date().toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' })}`;
    $('banner').hidden = true;
  } catch (err) { $('loadInfo').textContent = ''; $('banner').hidden = false; $('banner').textContent = 'Could not load from the Sheet: ' + err.message; }
  finally { $('reload').disabled = false; }
  if (gd.on) await loadGodown();
  derive(); render();
}

/* ---------- Godown Stock (Firebase, read-only) ---------- */
let fb = null;
function fbInit() {
  if (fb) return fb;
  if (typeof firebase === 'undefined') return null;
  try { const app = firebase.apps.find(a => a.name === 'godown') || firebase.initializeApp(FIREBASE_CONFIG, 'godown'); fb = { app, auth: app.auth(), db: app.firestore() }; } catch (e) { fb = null; }
  return fb;
}
async function loadGodown() {
  const f = fbInit(); if (!f || !f.auth.currentUser) { gd.on = false; renderGodownBox(); return; }
  try {
    const [ps, ms] = await Promise.all([f.db.collection('products').get(), f.db.collection('movements').get()]);
    gd.products = {}; ps.forEach(d => gd.products[d.id] = d.data());
    gd.moves = []; ms.forEach(d => (d.data().entries || []).forEach(e => gd.moves.push(Object.assign({ day: d.id }, e))));
    gd.at = new Date(); gd.on = true;
  } catch (err) { gd.on = false; $('gMsg').textContent = 'Godown Stock: ' + (err.code === 'permission-denied' ? 'this login cannot read stock.' : err.message); }
  renderGodownBox();
}
function renderGodownBox() {
  const f = fbInit(), u = f && f.auth.currentUser;
  $('gdOut').hidden = !!u; $('gdIn').hidden = !u;
  $('gInfo').textContent = u ? `Connected as ${u.email.replace(AUTH_EMAIL_SUFFIX, '')}. ${Object.keys(gd.products).length} products, ${gd.moves.length} stock movements read${gd.at ? ' at ' + gd.at.toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' }) : ''}.` : '';
}
$('gLogin').onclick = async () => {
  const f = fbInit(); if (!f) { $('gMsg').textContent = 'Firebase did not load. Check the internet and reload.'; return; }
  const id = $('gUser').value.trim().toLowerCase(), pw = $('gPass').value;
  if (!id || !pw) { $('gMsg').textContent = 'Enter the Godown Stock user ID and password.'; return; }
  $('gMsg').textContent = 'Connecting…';
  try { await f.auth.signInWithEmailAndPassword(id.includes('@') ? id : id + AUTH_EMAIL_SUFFIX, pw); $('gPass').value = ''; $('gMsg').textContent = ''; gd.on = true; await loadGodown(); derive(); render(); }
  catch (err) { $('gMsg').textContent = /wrong-password|invalid-credential|user-not-found|invalid-login/.test(err.code || '') ? 'User ID or password is wrong.' : (err.message || 'Could not log in.'); }
};
$('gLogout').onclick = async () => { const f = fbInit(); if (f) await f.auth.signOut(); gd = { on: false, products: {}, moves: [], at: null }; renderGodownBox(); derive(); render(); };
$('gLink').value = store.get('gLink', '');
function setGLink() { const v = $('gLink').value.trim(); $('lnkGodown').hidden = !v; if (v) $('lnkGodown').href = v; }
$('gLink').addEventListener('change', () => { store.set('gLink', $('gLink').value.trim()); setGLink(); });
setGLink();

/* ---------- merge into parcels ---------- */
function rateFor(sku, plat) {
  if (!sku) return null;
  const parts = String(sku).split(',').map(s => key(s)).filter(Boolean);
  for (const p of parts) { const r = rates[p]; if (r) { const v = plat === 'Flipkart' ? r.fk : plat === 'Meesho' ? r.ms : null; if (v != null) return v; } }
  return null;
}
function godownOutIndex() {
  const m = {};
  gd.moves.forEach(e => { if (e.type === 'out' && e.awb && !e.localSale) { const k = key(e.awb); (m[k] = m[k] || []).push(e); } });
  return m;
}
function derive() {
  const gOut = gd.on ? godownOutIndex() : null;
  const keys = new Set([...Object.keys(ledger), ...Object.keys(scans)]);
  plist = [];
  keys.forEach(k => {
    const l = ledger[k] || {}, s = scans[k] || null;
    const pk = s ? packets[String(s.packet || '').trim().toUpperCase()] : null;
    const p = {
      _id: k, awb: (s && s.awb) || l.awb, l, s,
      platform: (s && s.platform) || l.platform || '', scanned: !!s, scanBy: s ? s.by : '', ts: s ? s.ts : 0,
      packet: s ? s.packet : '', packetSku: pk ? pk.sku : '', product: (s && s.product) || l.product || (pk ? pk.name : ''),
      sku: l.sku || (s && s.orderSku) || '', qty: num(l.qty) || 1, dispatchDate: (s && s.date) || l.orderDate || '',
      orderId: l.orderId || (s && s.orderId) || '', orderItemId: l.orderItemId || '', courier: l.courier || (s && s.courier) || '',
      status: l.status || 'dispatched', statusDate: l.statusDate || '', returnReceived: l.returnReceived || '',
      salePrice: num(l.salePrice), expected: num(l.expected), fees: num(l.fees), received: num(l.received),
      settlementDate: l.settlementDate || '', utr: l.utr || '', notes: l.notes || '', cancelReason: l.cancelReason || '',
      claim: l.claimType ? { type: l.claimType, amount: num(l.claimAmount), id: l.claimId || '', raised: l.claimRaised || '', status: l.claimStatus || 'open', recovered: num(l.recovered) } : null
    };
    p._auto = rateFor(p.sku || p.packetSku, p.platform); if (p._auto != null) p._auto = Math.round(p._auto * p.qty * 100) / 100;
    p._exp = p.expected != null ? p.expected : p._auto;
    p._settle = p.received != null ? (p._exp != null && p.received < p._exp - 0.5 ? 'short' : 'settled') : (p.status === 'delivered' ? 'pending' : '');
    p._short = p._settle === 'short' ? Math.round((p._exp - p.received) * 100) / 100 : 0;
    p._wrong = !!(p.sku && p.packetSku && !String(p.sku).split(',').some(x => key(x) === key(p.packetSku)));
    p._stage = (p._settle === 'settled' || p._settle === 'short') ? 'settled' : p.status === 'delivered' ? 'delivered' : (p.status === 'transit' || p.status === 'ofd') ? 'transit' : p.status === 'dispatched' ? 'dispatched' : '';
    p._value = p._exp != null ? p._exp : p.salePrice;
    p._billed = !!(l.orderId || l.orderItemId || l.salePrice != null || (l.source || '').length);
    p._gOut = gOut ? (gOut[k] || null) : undefined;
    p._suggest = claimReason(p);
    plist.push(p);
  });
}
function claimReason(p) {
  if (p.claim) return null;
  if (p._settle === 'short') return ['short', 'Settled ' + inr(p._short) + ' less than expected', p._short];
  if (p.status === 'lost') return ['lost', 'Marked lost / damaged in transit', p._value];
  if (p.returnReceived === 'damaged') return ['return-damaged', 'Return came back damaged / used', p._value];
  if (p.returnReceived === 'wrong') return ['return-wrong', 'Wrong item or empty box returned', p._value];
  if (p.returnReceived === 'missing') return ['return-missing', 'Return never reached the godown', p._value];
  if (p.status === 'delivered' && p._settle === 'pending' && (daysSince(p.statusDate || p.dispatchDate) ?? 0) > CLAIM_UNPAID_DAYS) return ['unpaid', 'Delivered ' + daysSince(p.statusDate || p.dispatchDate) + ' days ago, still unpaid', p._exp];
  return null;
}

/* ---------- tabs ---------- */
const VIEWS = ['ledger', 'claims', 'recon', 'import', 'rates', 'settings'];
let tab = 'ledger';
function show(t) {
  if (!VIEWS.includes(t)) t = 'ledger'; tab = t;
  document.querySelectorAll('.tabs button').forEach(b => b.setAttribute('aria-selected', b.dataset.tab === t));
  VIEWS.forEach(v => $('v-' + v).hidden = v !== t);
  try { history.replaceState(null, '', '#' + t); } catch (e) {}
  render();
}
document.querySelectorAll('.tabs button').forEach(b => b.onclick = () => show(b.dataset.tab));
function render() {
  setCnt('cLedger', plist.filter(p => p._settle === 'short' || (p._settle === 'pending' && (daysSince(p.statusDate || p.dispatchDate) ?? 0) > UNPAID_DAYS)).length);
  setCnt('cClaims', plist.filter(p => p._suggest || (p.claim && p.claim.status === 'open')).length);
  setCnt('cRates', missingRates().length);
  ({ ledger: renderLedger, claims: renderClaims, recon: renderRecon, import: () => {}, rates: renderRates, settings: renderSettings })[tab]();
}
function setCnt(id, n) { const el = $(id); el.hidden = !n; el.textContent = n; el.className = 'cnt' + (n ? ' hot' : ''); }

/* ---------- LEDGER ---------- */
PLATFORMS.forEach(p => { $('fPlat').add(new Option(p, p)); $('f_platform').add(new Option(p, p)); });
STATUSES.forEach(s => { $('fStatus').add(new Option(s[1], s[0])); $('f_status').add(new Option(s[1], s[0])); });
Object.entries(CLAIM_T).forEach(([k, v]) => $('c_type').add(new Option(v, k)));
const filt = { q: '', plat: '', status: '', settle: '', stage: '', alert: '', ids: null, idsLabel: '' };
let sets = {};
function renderLedger() {
  const st = { dispatched: [], transit: [], delivered: [], settled: [] }; plist.forEach(p => { if (st[p._stage]) st[p._stage].push(p); });
  const sum = (a, f) => a.reduce((s, p) => s + (num(f(p)) || 0), 0);
  const stages = [['dispatched', 'Dispatched', sum(st.dispatched, p => p._value), 'value'], ['transit', 'In transit', sum(st.transit, p => p._value), 'value'], ['delivered', 'Delivered, unpaid', sum(st.delivered, p => p._exp), 'expected'], ['settled', 'Settled', sum(st.settled, p => p.received), 'received']];
  $('pipe').innerHTML = stages.map(([k, l, v, w], i) => `<button type="button" class="stage" data-stage="${k}" aria-pressed="${filt.stage === k}"><span class="no">0${i + 1}</span><span class="lbl">${l}</span><span class="bigN">${st[k].length}</span><span class="val">${inr0(v)} ${w}</span></button>`).join('');
  const late = plist.filter(p => p._settle === 'pending' && (daysSince(p.statusDate || p.dispatchDate) ?? 0) > UNPAID_DAYS);
  const short = plist.filter(p => p._settle === 'short');
  const stuck = plist.filter(p => ['dispatched', 'transit', 'ofd'].includes(p.status) && p.scanned && (daysSince(p.dispatchDate) ?? 0) > STUCK_DAYS);
  const retWait = plist.filter(p => ['rto', 'return'].includes(p.status) && !p.returnReceived);
  const ns = plist.filter(p => !p.scanned && !['cancelled'].includes(p.status) && p.status !== 'dispatched');
  const wrong = plist.filter(p => p._wrong);
  sets = { late, short, stuck, retWait, ns, wrong };
  const al = [['late', 'warn', late.length, `Delivered, unpaid > ${UNPAID_DAYS} days`, inr0(sum(late, p => p._exp)) + ' expected'], ['short', 'bad', short.length, 'Short-paid', inr0(sum(short, p => p._short)) + ' less than expected'], ['retWait', 'warn', retWait.length, 'Returns on the way', 'Mark when they reach the godown'], ['stuck', 'info', stuck.length, `Not delivered after ${STUCK_DAYS} days`, 'Check with courier'], ['ns', 'bad', ns.length, 'Shipped per portal, never scanned', 'Packed without a scan, or wrong AWB'], ['wrong', 'bad', wrong.length, 'Portal SKU ≠ packet scanned', 'Wrong product may have gone']];
  $('alerts').innerHTML = al.map(([k, c, n, t, s]) => `<button type="button" class="alert ${c} ${n ? '' : 'zero'}" data-alert="${k}" aria-pressed="${filt.alert === k}"><span class="n">${n}</span><span class="t"><b>${t}</b>${s}</span></button>`).join('');
  const q = filt.q.trim().toLowerCase();
  const list = plist.filter(p => {
    if (filt.ids && !filt.ids.has(p._id)) return false;
    if (filt.plat && p.platform !== filt.plat) return false;
    if (filt.status && p.status !== filt.status) return false;
    if (filt.settle && p._settle !== filt.settle) return false;
    if (filt.stage && p._stage !== filt.stage) return false;
    if (filt.alert && !sets[filt.alert].includes(p)) return false;
    if (q && ![p.cancelReason, p.awb, p.orderId, p.orderItemId, p.product, p.sku, p.packet, p.courier, p.utr, p.notes, p.claim && p.claim.id].some(x => String(x || '').toLowerCase().includes(q))) return false;
    return true;
  }).sort((a, b) => String(b.dispatchDate || '').localeCompare(String(a.dispatchDate || '')) || (b.ts || 0) - (a.ts || 0));
  const shown = list.slice(0, 500);
  $('lrows').innerHTML = shown.map(p => {
    const s = SMAP[p.status] || SMAP.dispatched, diff = (p._exp != null && p.received != null) ? p.received - p._exp : null;
    const sp = p._settle === 'settled' ? '<span class="pill p-ok">Settled</span>' : p._settle === 'short' ? `<span class="pill p-bad">Short ${inr(p._short)}</span>` : p._settle === 'pending' ? `<span class="pill p-warn">Pending${p.statusDate ? ' · ' + daysSince(p.statusDate) + 'd' : ''}</span>` : '<span class="small">—</span>';
    const rc = [p.returnReceived && RET[p.returnReceived] ? `<span class="pill ${RET[p.returnReceived][1]}">${RET[p.returnReceived][0]}</span>` : '', p.claim ? `<span class="pill ${(CLAIM_S[p.claim.status] || CLAIM_S.open)[1]}">Claim ${(CLAIM_S[p.claim.status] || CLAIM_S.open)[0].toLowerCase()}</span>` : p._suggest ? '<span class="pill p-warn">Claim due</span>' : ''].filter(Boolean).join(' ') || '<span class="small">—</span>';
    return `<tr class="click" tabindex="0" data-id="${esc(p._id)}"><td><span class="awb">${esc(p.awb)}</span>${p.scanned ? '' : '<span class="tag p-bad">Not scanned</span>'}${p._gOut === null && p.scanned ? '<span class="tag p-warn">No godown scan-out</span>' : ''}<div class="small">${esc(p.orderId || p.orderItemId || '')}</div></td>
      <td><span class="pill ${esc(p.platform)}">${esc(p.platform || '—')}</span><div class="small">${esc(p.courier)}</div></td>
      <td>${esc(p.product) || '—'}<div class="small mono">${esc(p.sku || p.packet || '')}</div>${p._wrong ? '<div class="small diff-neg">Packet scanned is ' + esc(p.packetSku) + '</div>' : ''}</td>
      <td class="mono">${fmtD(p.dispatchDate)}</td><td><span class="pill ${s.c}">${esc(s.l.replace(/ \(.*\)/, ''))}</span>${p.statusDate ? `<div class="small">${fmtD(p.statusDate)}</div>` : ''}${p.cancelReason ? `<div class="small">${esc(p.cancelReason.replace(/_/g, ' '))}</div>` : ''}</td>
      <td class="n">${inr(p.salePrice)}</td><td class="n">${inr(p._exp)}${p.expected == null && p._auto != null ? '<div class="small">from rate</div>' : ''}</td>
      <td class="n">${inr(p.received)}${diff != null && Math.abs(diff) >= 0.5 ? `<div class="small ${diff < 0 ? 'diff-neg' : 'diff-pos'}">${diff < 0 ? '−' : '+'}${inr(Math.abs(diff)).slice(1)}</div>` : ''}</td>
      <td>${sp}${p.settlementDate ? `<div class="small">${fmtD(p.settlementDate)}</div>` : ''}</td><td>${rc}</td></tr>`;
  }).join('');
  const em = $('lempty');
  if (!plist.length) { em.hidden = false; em.textContent = 'No parcels in these dates. Scan dispatches in the Dispatch Scan Register, or import a portal report.'; }
  else if (!list.length) { em.hidden = false; em.textContent = 'No parcels match these filters.'; } else em.hidden = true;
  $('countLine').textContent = `Showing ${shown.length} of ${plist.length} parcels` + (list.length > shown.length ? ` (first 500 of ${list.length} matches; search to narrow)` : '');
  const act = [filt.idsLabel, filt.q && `“${filt.q}”`, filt.plat, filt.status && SMAP[filt.status].l, filt.settle && $('fSettle').selectedOptions[0].text, filt.stage && 'Stage: ' + filt.stage, filt.alert && (al.find(a => a[0] === filt.alert) || [])[3]].filter(Boolean);
  $('chipline').hidden = !act.length; $('chipText').textContent = act.join(' · ');
}
$('q').addEventListener('input', e => { filt.q = e.target.value; renderLedger(); });
$('fPlat').onchange = e => { filt.plat = e.target.value; renderLedger(); };
$('fStatus').onchange = e => { filt.status = e.target.value; renderLedger(); };
$('fSettle').onchange = e => { filt.settle = e.target.value; renderLedger(); };
function clearFilt() { Object.assign(filt, { q: '', plat: '', status: '', settle: '', stage: '', alert: '', ids: null, idsLabel: '' }); ['q', 'fPlat', 'fStatus', 'fSettle'].forEach(i => $(i).value = ''); }
$('clearF').onclick = () => { clearFilt(); renderLedger(); };
$('pipe').onclick = e => { const b = e.target.closest('[data-stage]'); if (!b) return; filt.stage = filt.stage === b.dataset.stage ? '' : b.dataset.stage; filt.alert = ''; renderLedger(); };
$('alerts').onclick = e => { const b = e.target.closest('[data-alert]'); if (!b) return; filt.alert = filt.alert === b.dataset.alert ? '' : b.dataset.alert; filt.stage = ''; renderLedger(); };
$('lrows').onclick = e => { const r = e.target.closest('[data-id]'); if (r) openParcel(r.dataset.id); };
$('lrows').onkeydown = e => { if (e.key === 'Enter') { const r = e.target.closest('[data-id]'); if (r) openParcel(r.dataset.id); } };
$('btnAdd').onclick = () => openParcel(null);
function showIds(ids, label) { clearFilt(); filt.ids = new Set(ids); filt.idsLabel = label; show('ledger'); }

/* ---------- parcel dialog ---------- */
const PF = ['awb', 'platform', 'orderId', 'orderItemId', 'sku', 'qty', 'courier', 'salePrice', 'status', 'statusDate', 'returnReceived', 'expected', 'fees', 'received', 'settlementDate', 'utr', 'notes'];
const NUMF = new Set(['qty', 'salePrice', 'expected', 'fees', 'received', 'claimAmount', 'recovered']);
let editId = null, claimOn = false;
document.querySelectorAll('[data-close]').forEach(b => b.onclick = () => b.closest('dialog').close());
function openParcel(id, claimType) {
  editId = id; const p = id ? plist.find(x => x._id === id) : { platform: 'Flipkart', status: 'dispatched', l: {} };
  PF.forEach(f => {
    let v = f === 'expected' ? p.expected : f === 'qty' ? (p.l && p.l.qty) : p[f];
    $('f_' + f).value = v != null ? v : '';
  });
  $('f_awb').readOnly = !!id;
  $('f_expected').placeholder = p._auto != null ? 'from rate: ' + p._auto.toFixed(2) : 'add a rate in Rates';
  $('dlgTitle').textContent = id ? 'AWB ' + p.awb : 'Add parcel';
  const info = [];
  if (id) {
    info.push(p.scanned ? `<div>Scanned ${fmtD(p.dispatchDate)}${p.s && p.s.time ? ' ' + esc(p.s.time.slice(0, 5)) : ''} by ${esc(p.scanBy)} · packet <span class="mono">${esc(p.packet)}</span>${p.packetSku ? ' (' + esc(p.packetSku) + ')' : ''}</div>` : '<div class="diff-neg">Not scanned at dispatch.</div>');
    if (p._gOut !== undefined) info.push(p._gOut ? `<div>Godown scan-out: ${p._gOut.reduce((a, e) => a + (Number(e.qty) || 0), 0)} packet(s) by ${esc(p._gOut[0].staff || '')}</div>` : (p.scanned ? '<div class="diff-neg">No godown scan-out with this AWB.</div>' : ''));
    if (p.cancelReason) info.push(`<div>Flipkart cancel reason: <b>${esc(p.cancelReason.replace(/_/g, ' '))}</b></div>`);
    if (p._wrong) info.push(`<div class="diff-neg">Portal SKU ${esc(p.sku)} does not match the packet scanned (${esc(p.packetSku)}).</div>`);
    if (p.l && p.l.source) info.push(`<div class="small">Data from: ${esc(p.l.source)}${p.l.updated ? ' · updated ' + esc(p.l.updated.replace('T', ' ').slice(0, 16)) + (p.l.updatedBy ? ' by ' + esc(p.l.updatedBy) : '') : ''}</div>`);
  } else info.push('<div class="small">Manual entry. Parcels scanned in the Dispatch Scan Register appear here by themselves.</div>');
  $('dlgInfo').innerHTML = info.join('');
  const c = p.claim; claimOn = !!c || !!claimType;
  const sug = claimType ? claimReason(p) : null;
  $('c_type').value = c ? c.type : (claimType || 'short');
  $('c_amount').value = c ? (c.amount ?? '') : (sug && sug[2] != null ? Math.round(sug[2] * 100) / 100 : '');
  $('c_id').value = c ? c.id : ''; $('c_raised').value = c ? c.raised : today(); $('c_status').value = c ? c.status : 'open'; $('c_recovered').value = c ? (c.recovered ?? '') : '';
  syncClaim();
  const dz = $('delZone'); dz.innerHTML = id && isAdmin() && ledger[id] && !p.scanned ? '<button type="button" class="danger" id="delP">Delete ledger row</button>' : '';
  if ($('delP')) $('delP').onclick = () => { dz.innerHTML = '<span class="confirm">Delete this ledger row? <button type="button" class="danger" id="delY">Yes</button></span>'; $('delY').onclick = async () => { try { await api('POST', { action: 'ledgerDelete', awb: p.awb }); delete ledger[id]; derive(); render(); $('dlg').close(); toast('Deleted'); } catch (e) { toast(e.message); } }; };
  $('dlg').showModal();
  setTimeout(() => (claimType ? $('c_id') : id ? $('f_status') : $('f_awb')).focus(), 30);
}
function syncClaim() { $('claimFields').hidden = !claimOn; $('claimToggle').textContent = claimOn ? 'Remove claim' : 'Raise a claim'; }
$('claimToggle').onclick = () => { claimOn = !claimOn; if (claimOn && !$('c_raised').value) $('c_raised').value = today(); syncClaim(); };
$('f_status').onchange = e => { if (['delivered', 'rto', 'return', 'lost'].includes(e.target.value) && !$('f_statusDate').value) $('f_statusDate').value = today(); };
$('f_received').oninput = () => { if ($('f_received').value && !$('f_settlementDate').value) $('f_settlementDate').value = today(); };
$('f_returnReceived').onchange = e => { const v = e.target.value; if (v && !['rto', 'return'].includes($('f_status').value)) $('f_status').value = 'return'; const m = { damaged: 'return-damaged', wrong: 'return-wrong', missing: 'return-missing' }; if (m[v] && !claimOn) { claimOn = true; $('c_type').value = m[v]; syncClaim(); } };
$('pform').addEventListener('submit', async e => {
  e.preventDefault();
  const f = {}; PF.forEach(k => { const v = $('f_' + k).value.trim(); f[k] = NUMF.has(k) ? num(v) : v; });
  if (!f.awb) { $('f_awb').focus(); return; }
  const id = editId || key(f.awb);
  if (!editId && (ledger[id] || scans[id])) { toast('AWB ' + f.awb + ' is already here. Opening it.'); $('dlg').close(); openParcel(id); return; }
  if (claimOn) Object.assign(f, { claimType: $('c_type').value, claimAmount: num($('c_amount').value), claimId: $('c_id').value.trim(), claimRaised: $('c_raised').value, claimStatus: $('c_status').value, recovered: num($('c_recovered').value) });
  else Object.assign(f, { claimType: '', claimAmount: null, claimId: '', claimRaised: '', claimStatus: '', recovered: null });
  $('saveP').disabled = true;
  try {
    await api('POST', { action: 'ledgerSave', item: { awb: f.awb, f } });
    ledger[id] = Object.assign({}, ledger[id] || {}, f, { updated: new Date().toISOString().slice(0, 19), updatedBy: user ? user.name : '' });
    derive(); render(); $('dlg').close(); toast('Saved to the Sheet');
  } catch (err) { toast('Not saved: ' + err.message); }
  finally { $('saveP').disabled = false; }
});

/* ---------- CLAIMS ---------- */
$('claimFilter').onchange = renderClaims;
function renderClaims() {
  const cl = plist.filter(p => p.claim), sug = plist.filter(p => p._suggest);
  const sum = (a, f) => a.reduce((s, p) => s + (num(f(p)) || 0), 0);
  const by = s => cl.filter(p => (p.claim.status || 'open') === s);
  const open = by('open'), appr = by('approved'), paid = by('paid'), rej = by('rejected');
  $('claimTiles').innerHTML = [['Should be claimed', sug.length, sug.length ? 'warn' : '', inr0(sum(sug, p => p._suggest[2])) + ' at stake'], ['Open claims', open.length, '', inr0(sum(open, p => p.claim.amount)) + ' claimed'], ['Approved, money awaited', appr.length, '', inr0(sum(appr, p => p.claim.amount))], ['Recovered', paid.length, 'ok', inr0(sum(paid, p => p.claim.recovered ?? p.claim.amount))], ['Rejected', rej.length, rej.length ? 'bad' : '', inr0(sum(rej, p => p.claim.amount))]].map(([l, n, c, s]) => `<div class="tile ${c}"><span class="lbl">${l}</span><b>${n}</b><div class="small">${s}</div></div>`).join('');
  sug.sort((a, b) => (b._suggest[2] || 0) - (a._suggest[2] || 0));
  $('suggestRows').innerHTML = sug.length ? sug.map(p => `<tr><td class="awb">${esc(p.awb)}<div class="small">${esc(p.orderId)}</div></td><td><span class="pill ${esc(p.platform)}">${esc(p.platform || '—')}</span></td><td>${esc(p.product) || '—'}</td><td>${esc(p._suggest[1])}</td><td class="n">${inr(p._suggest[2])}</td><td class="n"><button type="button" class="sm primary" data-raise="${esc(p._id)}" data-t="${p._suggest[0]}">Raise claim</button></td></tr>`).join('') : '<tr><td colspan="6" class="empty">Nothing to claim right now.</td></tr>';
  const f = $('claimFilter').value; const list = cl.filter(p => !f || (p.claim.status || 'open') === f).sort((a, b) => String(b.claim.raised).localeCompare(String(a.claim.raised)));
  $('claimRows').innerHTML = list.length ? list.map(p => { const s = CLAIM_S[p.claim.status] || CLAIM_S.open; const age = p.claim.raised && p.claim.status === 'open' ? daysSince(p.claim.raised) : null; return `<tr class="click" tabindex="0" data-id="${esc(p._id)}"><td class="awb">${esc(p.awb)}</td><td><span class="pill ${esc(p.platform)}">${esc(p.platform || '—')}</span></td><td>${esc(CLAIM_T[p.claim.type] || p.claim.type)}</td><td class="mono">${esc(p.claim.id) || '<span class="small">add ID</span>'}</td><td>${fmtD(p.claim.raised)}${age != null ? `<div class="small">${age} days open</div>` : ''}</td><td><span class="pill ${s[1]}">${s[0]}</span></td><td class="n">${inr(p.claim.amount)}</td><td class="n">${inr(p.claim.recovered)}</td></tr>`; }).join('') : `<tr><td colspan="8" class="empty">${cl.length ? 'No claims with this status.' : 'No claims raised yet.'}</td></tr>`;
}
$('suggestRows').onclick = e => { const b = e.target.closest('[data-raise]'); if (b) openParcel(b.dataset.raise, b.dataset.t); };
$('claimRows').onclick = e => { const r = e.target.closest('[data-id]'); if (r) openParcel(r.dataset.id); };
$('claimRows').onkeydown = e => { if (e.key === 'Enter') { const r = e.target.closest('[data-id]'); if (r) openParcel(r.dataset.id); } };

/* ---------- RECONCILE ---------- */
function renderRecon() {
  const from = $('gFrom').value || '0000', to = $('gTo').value || '9999';
  const inR = d => d && d >= from && d <= to;
  const P = plist.filter(p => inR(p.dispatchDate)), S = P.filter(p => p.scanned);
  const sum = (a, f) => a.reduce((s, p) => s + (num(f(p)) || 0), 0);
  const btn = (arr, label) => arr.length ? ` <button type="button" class="link" data-ids="${arr.map(p => esc(p._id)).join('|')}" data-lbl="${esc(label)}">view</button>` : '';
  const badge = ok => `<span class="pill ${ok ? 'p-ok' : 'p-bad'}">${ok ? 'Matches' : 'Check'}</span>`;
  const cards = [];
  // 1 godown out = dispatched
  if (gd.on) {
    const gm = gd.moves.filter(e => inR(String(e.timestamp || e.day || '').slice(0, 10)));
    const gOutAwb = {}; gm.forEach(e => { if (e.type === 'out' && e.awb && !e.localSale) gOutAwb[key(e.awb)] = (gOutAwb[key(e.awb)] || 0) + (Number(e.qty) || 0); });
    const dispNoG = S.filter(p => !p._gOut);
    const gNoDisp = Object.keys(gOutAwb).filter(k => !scans[k]);
    const local = gm.filter(e => e.type === 'out' && e.localSale).reduce((a, e) => a + (Number(e.qty) || 0), 0);
    const outNoAwb = gm.filter(e => e.type === 'out' && !e.awb && !e.localSale).reduce((a, e) => a + (Number(e.qty) || 0), 0);
    const negAdj = -gm.filter(e => e.type === 'adjust' && e.qty < 0).reduce((a, e) => a + e.qty, 0);
    cards.push(['1', 'Godown out = dispatched', 'Every AWB scanned out of the godown should be scanned at dispatch, and the other way round.', !dispNoG.length && !gNoDisp.length,
      `<div class="kv"><span>AWBs scanned at dispatch</span><span class="v">${S.length}</span>
      <span>AWBs scanned out of godown</span><span class="v">${Object.keys(gOutAwb).length}</span>
      <span>Dispatched, no godown scan-out${btn(dispNoG, 'Dispatched, no godown scan-out')}</span><span class="v ${dispNoG.length ? 'bad' : 'ok'}">${dispNoG.length}</span>
      <span>Left godown, never dispatched</span><span class="v ${gNoDisp.length ? 'bad' : 'ok'}">${gNoDisp.length}</span>
      ${gNoDisp.length ? `<span class="small mono" style="grid-column:1/-1">${gNoDisp.slice(0, 12).map(esc).join(', ')}${gNoDisp.length > 12 ? ' …' : ''}</span>` : ''}
      <span>Packets out without an AWB</span><span class="v ${outNoAwb ? 'warn' : ''}">${outNoAwb}</span>
      <span>Local sale packets</span><span class="v">${local}</span>
      <span>Packets removed by adjustment</span><span class="v ${negAdj ? 'warn' : ''}">${negAdj}</span></div>`]);
  } else cards.push(['1', 'Godown out = dispatched', 'Compares Godown Stock scan-outs with dispatch scans, AWB by AWB.', null, '<p class="hint">Connect Godown Stock in the <b>Connections</b> tab to run this check.</p>']);
  // 2 dispatched = billed
  const unbilled = S.filter(p => !p._billed && p.status !== 'cancelled');
  const notScanned = P.filter(p => !p.scanned && !['cancelled', 'dispatched'].includes(p.status));
  const wrong = P.filter(p => p._wrong);
  const cxl = cancelled.filter(c => inR(String(c.cancelledAt || '').slice(0, 10)));
  cards.push(['2', 'Dispatched = billed', 'Every scanned parcel should have a portal order, and every shipped order should have been scanned.', !unbilled.length && !notScanned.length && !wrong.length,
    `<div class="kv"><span>Parcels scanned</span><span class="v">${S.length}</span>
    <span>Matched to a portal order</span><span class="v ok">${S.length - unbilled.length}</span>
    <span>Scanned, no portal order yet${btn(unbilled, 'Scanned, no portal order')}</span><span class="v ${unbilled.length ? 'warn' : ''}">${unbilled.length}</span>
    <span>Shipped per portal, never scanned${btn(notScanned, 'Shipped, never scanned')}</span><span class="v ${notScanned.length ? 'bad' : ''}">${notScanned.length}</span>
    <span>Portal SKU ≠ packet scanned${btn(wrong, 'Wrong product')}</span><span class="v ${wrong.length ? 'bad' : ''}">${wrong.length}</span>
    <span>Cancelled at pickup (dispatch app)</span><span class="v">${cxl.length}</span></div>`]);
  // 3 billed = received
  const del = P.filter(p => p.status === 'delivered');
  const pend = del.filter(p => p._settle === 'pending'), sh = P.filter(p => p._settle === 'short'), noExp = del.filter(p => p._exp == null);
  cards.push(['3', 'Billed = money received', 'For delivered parcels: what was sold, what should settle, and what reached the bank.', !sh.length && !pend.filter(p => (daysSince(p.statusDate || p.dispatchDate) ?? 0) > UNPAID_DAYS).length,
    `<div class="kv"><span>Delivered parcels</span><span class="v">${del.length}</span>
    <span>Sale value</span><span class="v">${inr0(sum(del, p => p.salePrice))}</span>
    <span>Expected settlement</span><span class="v">${inr0(sum(del, p => p._exp))}</span>
    <span>Fees / deductions</span><span class="v">${inr0(sum(P, p => p.fees))}</span>
    <span>Received</span><span class="v ok">${inr0(sum(P, p => p.received))}</span>
    <span>Still pending${btn(pend, 'Settlement pending')}</span><span class="v ${pend.length ? 'warn' : ''}">${inr0(sum(pend, p => p._exp))} · ${pend.length}</span>
    <span>Short-paid${btn(sh, 'Short-paid')}</span><span class="v ${sh.length ? 'bad' : ''}">${inr0(sum(sh, p => p._short))} · ${sh.length}</span>
    ${noExp.length ? `<span class="small">Delivered, no expected amount (add rates)${btn(noExp, 'No expected amount')}</span><span class="v small warn">${noExp.length}</span>` : ''}</div>`]);
  // 4 returns back in godown
  const rets = P.filter(p => ['rto', 'return'].includes(p.status));
  const rec = r => rets.filter(p => p.returnReceived === r);
  const retOpen = rets.filter(p => !p.returnReceived);
  const gRet = gd.on ? gd.moves.filter(e => e.type === 'return' && inR(String(e.timestamp || e.day || '').slice(0, 10))) : [];
  const gGood = gRet.filter(e => e.condition === 'good').reduce((a, e) => a + (Number(e.qty) || 0), 0), gBad = gRet.filter(e => e.condition !== 'good').reduce((a, e) => a + (Number(e.qty) || 0), 0);
  cards.push(['4', 'Returns = back in godown', 'Every RTO and customer return should reach the godown and be logged in Godown Stock → Returns.', !retOpen.filter(p => (daysSince(p.statusDate || p.dispatchDate) ?? 0) > 10).length && !rec('missing').length,
    `<div class="kv"><span>RTO + customer returns${btn(rets, 'Returns')}</span><span class="v">${rets.length}</span>
    <span>Received good</span><span class="v ok">${rec('good').length}</span>
    <span>Received damaged / wrong${btn(rec('damaged').concat(rec('wrong')), 'Bad returns')}</span><span class="v ${rec('damaged').length + rec('wrong').length ? 'bad' : ''}">${rec('damaged').length + rec('wrong').length}</span>
    <span>Never came back${btn(rec('missing'), 'Return missing')}</span><span class="v ${rec('missing').length ? 'bad' : ''}">${rec('missing').length}</span>
    <span>Not marked yet${btn(retOpen, 'Returns not marked')}</span><span class="v ${retOpen.length ? 'warn' : ''}">${retOpen.length}</span>
    ${gd.on ? `<span>Godown Stock returns logged (packets)</span><span class="v">${gGood} good · ${gBad} damaged</span>` : '<span class="small" style="grid-column:1/-1">Connect Godown Stock to compare with returns logged in the godown.</span>'}</div>`]);
  $('checks').innerHTML = cards.map(([n, t, s, ok, body]) => `<div class="panel check"><div class="head"><div><span class="no">Check ${n}</span><h2>${t}</h2></div>${ok == null ? '<span class="pill p-neutral">Not connected</span>' : badge(ok)}</div><p class="hint">${s}</p>${body}</div>`).join('');
  // stock table
  $('stockPanel').hidden = !gd.on;
  if (gd.on) {
    const per = {}; gd.moves.filter(e => inR(String(e.timestamp || e.day || '').slice(0, 10))).forEach(e => { const o = per[e.sku] = per[e.sku] || { out: 0, local: 0, ret: 0 }; const q = Number(e.qty) || 0; if (e.type === 'out') { if (e.localSale) o.local += q; else o.out += q; } else if (e.type === 'return') o.ret += q; });
    const label = p => [p.brand, p.name, p.variant, p.size && 'Size ' + p.size].filter(Boolean).join(' · ');
    const ks = Object.keys(gd.products).sort((a, b) => label(gd.products[a]).localeCompare(label(gd.products[b])));
    $('stockAsOf').textContent = gd.at ? 'Read ' + gd.at.toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' }) : '';
    $('stockRows').innerHTML = ks.map(k => { const p = gd.products[k], o = per[k] || { out: 0, local: 0, ret: 0 }; return `<tr><td>${esc(label(p))}</td><td class="mono small">${esc(k)}</td><td class="n ${(p.qty || 0) <= 0 ? 'diff-neg' : ''}"><b>${p.qty || 0}</b></td><td class="n">${o.out}</td><td class="n">${o.local}</td><td class="n">${o.ret}</td></tr>`; }).join('') || '<tr><td colspan="6" class="empty">No products in Godown Stock.</td></tr>';
  }
}
$('checks').onclick = e => { const b = e.target.closest('[data-ids]'); if (b) showIds(b.dataset.ids.split('|'), b.dataset.lbl); };

/* ---------- RATES ---------- */
function missingRates() {
  const need = {};
  plist.forEach(p => { if (p._exp == null && (p.sku || p.packetSku) && ['Flipkart', 'Meesho'].includes(p.platform)) { const s = String(p.sku || p.packetSku).split(',')[0].trim(); if (s) need[key(s)] = need[key(s)] || { sku: s, product: p.product, n: 0 }; if (s) need[key(s)].n++; } });
  return Object.values(need);
}
function renderRates() {
  const cnt = {}; plist.forEach(p => String(p.sku || p.packetSku || '').split(',').forEach(s => { if (s.trim()) cnt[key(s)] = (cnt[key(s)] || 0) + 1; }));
  const rows = Object.values(rates).sort((a, b) => a.sku.localeCompare(b.sku)).map(r => `<tr class="click" tabindex="0" data-r="${esc(r.sku)}"><td class="mono">${esc(r.sku)}</td><td>${esc(r.product)}</td><td class="n">${inr(r.fk)}</td><td class="n">${inr(r.ms)}</td><td class="n">${inr(r.cost)}</td><td class="n">${cnt[key(r.sku)] || 0}</td></tr>`);
  const miss = missingRates().sort((a, b) => b.n - a.n).map(m => `<tr class="click" tabindex="0" data-new="${esc(m.sku)}" data-prod="${esc(m.product)}"><td class="mono">${esc(m.sku)} <span class="tag p-warn">No rate</span></td><td>${esc(m.product)}</td><td class="n">—</td><td class="n">—</td><td class="n">—</td><td class="n">${m.n}</td></tr>`);
  $('rateRows').innerHTML = miss.concat(rows).join('') || '<tr><td colspan="6" class="empty">No rates yet. SKUs from your parcels show up here so you can add their rate.</td></tr>';
  $('addRate').hidden = !isAdmin();
}
function openRate(sku, product) {
  if (!isAdmin()) { toast('Only an admin can change rates.'); return; }
  const r = rates[key(sku)] || { sku: sku || '', product: product || '' };
  $('r_sku').value = r.sku; $('r_sku').readOnly = !!rates[key(sku)]; $('r_product').value = r.product || ''; $('r_fk').value = r.fk ?? ''; $('r_ms').value = r.ms ?? ''; $('r_cost').value = r.cost ?? '';
  $('rdlg').showModal();
}
$('addRate').onclick = () => openRate('', '');
$('rateRows').onclick = e => { const r = e.target.closest('[data-r],[data-new]'); if (r) openRate(r.dataset.r || r.dataset.new, r.dataset.prod); };
$('rform').addEventListener('submit', async e => {
  e.preventDefault(); const rate = { sku: $('r_sku').value.trim(), product: $('r_product').value.trim(), fk: num($('r_fk').value), ms: num($('r_ms').value), cost: num($('r_cost').value) };
  if (!rate.sku) return;
  try { await api('POST', { action: 'rate', rate }); rates[key(rate.sku)] = rate; derive(); render(); $('rdlg').close(); toast('Rate saved'); } catch (err) { toast(err.message); }
});

/* ---------- SETTINGS / sync ---------- */
function renderSettings() {
  $('sheetInfo').textContent = `Connected as ${user ? user.name : ''} (${user ? user.role : ''}). Ledger rows are saved in the Sheet's Ledger tab; rates in the Rates tab.`;
  const s = sync || {};
  $('fkInfo').innerHTML = !s.flipkart ? 'Flipkart keys are not saved yet. In Apps Script: Project Settings → Script Properties → add <code>FK_APP_ID</code> and <code>FK_APP_SECRET</code>, then run <b>checkFlipkart</b> and <b>setup</b>.'
    : s.last ? (s.last.ok ? `Last sync ${esc(s.last.at)}: ${s.last.shipments} Flipkart shipments read, ${s.last.updated} parcels updated, ${s.last.added} added${s.last.noAwb ? ', ' + s.last.noAwb + ' without AWB yet' : ''}. Runs every hour by itself.${s.last.warn ? '<br><span class="diff-neg">' + esc(s.last.warn) + '</span>' : ''}` : `<span class="diff-neg">Last sync ${esc(s.last.at)} failed: ${esc(s.last.error)}</span>`)
    : 'Keys saved. The first sync runs within the hour, or press Sync now.';
  $('syncNow').hidden = !(isAdmin() && s.flipkart);
  renderGodownBox();
}
$('syncNow').onclick = async () => {
  $('syncNow').disabled = true; $('syncMsg').textContent = 'Syncing with Flipkart… (up to 5 minutes)';
  try { const r = await api('POST', { action: 'syncLedger' }); $('syncMsg').textContent = `Done: ${r.updated} updated, ${r.added} added.`; await load(); }
  catch (err) { $('syncMsg').textContent = err.message; } finally { $('syncNow').disabled = false; }
};

/* ---------- IMPORT ---------- */
const drop = $('drop');
['dragover', 'dragenter'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach(t => drop.addEventListener(t, () => drop.classList.remove('over')));
drop.addEventListener('drop', e => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (f) readFile(f); });
$('file').onchange = e => { const f = e.target.files[0]; if (f) readFile(f); };
const HEAD = [
  ['awb', /^(awb|awb\s*(no|number|#|code)?|tracking\s*(id|no|number)|waybill|shipment\s*tracking\s*id)\b/],
  ['orderItemId', /order\s*item\s*id|sub\s*order\s*(no|id|number)?\b/],
  ['orderId', /^order\s*(id|no|number)\b/],
  ['sku', /^(seller\s*)?sku(\s*(id|code))?$/],
  ['utr', /utr|neft\s*(id|ref|no)|bank\s*ref|payment\s*(id|ref)|transaction\s*id|settlement\s*ref/],
  ['settlementDate', /settle(ment)?\s*date|payment\s*date|payout\s*date|paid\s*on|neft\s*date/],
  ['statusDate', /deliver(ed|y)\s*(date|on|at)|return(ed)?\s*(date|on)|return\s*created/],
  ['received', /bank\s*settlement|final\s*settlement|settlement\s*(amount|value)|settled\s*amount|net\s*(amount|settlement|payout)|payout\s*amount|amount\s*(paid|settled)|^received/],
  ['expected', /expected|projected/],
  ['fees', /total\s*(fees|deductions?|charges)|marketplace\s*fee|^commission/],
  ['salePrice', /order\s*value|selling\s*price|sale\s*(amount|price)|invoice\s*(amount|value)|total\s*sale|customer\s*price|final\s*price|^price$/],
  ['qty', /^(qty|quantity|units)\b/],
  ['status', /^(status|shipment\s*status|delivery\s*status|order\s*(item\s*)?status|return\s*(type|status)|sub\s*order\s*status|live\s*order\s*status|reason\s*for\s*credit)/],
  ['courier', /courier|carrier|logistics\s*partner|shipping\s*partner|^3pl/],
  ['platform', /^(platform|channel|marketplace)$/],
  ['product', /^(product(\s*(name|title))?|title|item\s*name|product\s*description)$/],
  ['orderDate', /^order\s*(date|placed)/],
  ['notes', /remark|note/]
];
function mapHeaders(h) { const map = {}, used = new Set(); HEAD.forEach(([f, re]) => { for (let i = 0; i < h.length; i++) { if (used.has(i)) continue; const s = String(h[i] || '').trim().toLowerCase().replace(/\s+/g, ' '); if (s && re.test(s)) { map[f] = i; used.add(i); break; } } }); return map; }
function normDate(v) {
  if (v == null || v === '') return ''; if (typeof v === 'number' && v > 20000 && v < 80000) return new Date(Math.round((v - 25569) * 864e5)).toISOString().slice(0, 10);
  v = String(v).trim(); let m = v.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/); if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  m = v.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/); if (m) { const y = m[3].length === 2 ? '20' + m[3] : m[3]; return `${y}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`; }
  m = v.match(/^(\d{1,2})[\s-]([A-Za-z]{3})[a-z]*[\s-,]+(\d{2,4})/); if (m) { const d = new Date(`${m[2]} ${m[1]} ${m[3].length === 2 ? '20' + m[3] : m[3]}`); if (!isNaN(d)) return d.toLocaleDateString('en-CA'); }
  const d = new Date(v); return isNaN(d) ? '' : d.toLocaleDateString('en-CA');
}
function normStatus(v) {
  v = String(v || '').toLowerCase(); if (!v) return '';
  if (/rto|origin|courier\s*return/.test(v)) return 'rto'; if (/return/.test(v)) return 'return'; if (/cancel/.test(v)) return 'cancelled'; if (/lost|damag/.test(v)) return 'lost';
  if (/out\s*for/.test(v)) return 'ofd'; if (/deliver/.test(v)) return 'delivered'; if (/transit|shipped|picked/.test(v)) return 'transit'; if (/dispatch|ready|packed|manifest|approved/.test(v)) return 'dispatched'; return '';
}
function normPlat(v) { v = String(v || '').trim(); if (!v) return ''; return PLATFORMS.find(p => p.toLowerCase().startsWith(v.toLowerCase().slice(0, 4))) || v; }
function pickTable(wb) {
  let best = null;
  wb.SheetNames.forEach(n => {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: '' });
    for (let r = 0; r < Math.min(15, rows.length); r++) { const m = mapHeaders(rows[r]); if (m.awb == null && m.orderId == null && m.orderItemId == null) continue; const sc = Object.keys(m).length; if (!best || sc > best.sc) best = { sc, rows: rows.slice(r + 1), map: m, sheet: n }; }
  });
  return best;
}
function buildImport(t, fixedPlat) {
  const acc = {}; let skipped = 0;
  for (const row of t.rows) {
    const f = {};
    for (const [k, i] of Object.entries(t.map)) {
      let v = row[i]; if (v == null || String(v).trim() === '') continue;
      if (NUMF.has(k)) { v = num(v); if (v == null) continue; }
      else if (/Date$/.test(k)) { v = normDate(v); if (!v) continue; }
      else if (k === 'status') { v = normStatus(v); if (!v) continue; }
      else if (k === 'platform') v = normPlat(v);
      else v = typeof v === 'number' && Number.isInteger(v) ? v.toLocaleString('en-US', { useGrouping: false }) : String(v).trim();
      f[k] = v;
    }
    if (fixedPlat && !f.platform) f.platform = fixedPlat;
    const id = f.awb ? 'a:' + key(f.awb) : f.orderItemId ? 'i:' + f.orderItemId : f.orderId ? 'o:' + f.orderId : null;
    if (!id) { skipped++; continue; }
    if (!acc[id]) acc[id] = f;
    else { const a = acc[id]; for (const k in f) { if ((k === 'received' || k === 'fees') && a[k] != null) a[k] = Math.round((a[k] + f[k]) * 100) / 100; else if (k === 'salePrice' && a[k] != null) continue; else a[k] = f[k]; } }
  }
  const items = Object.values(acc).map(f => {
    if (f.received != null && !f.status) f.status = 'delivered';
    f.source = (fixedPlat || f.platform || 'Portal') + ' report';
    return { awb: f.awb || '', orderId: f.orderId || '', orderItemId: f.orderItemId || '', f };
  });
  return { items, skipped };
}
async function readFile(file) {
  const res = $('impResult'); res.textContent = 'Reading ' + file.name + '…';
  if (!window.XLSX) { res.textContent = 'The file reader did not load. Reload the page.'; return; }
  let wb; try { wb = XLSX.read(await file.arrayBuffer(), { type: 'array' }); } catch (e) { res.textContent = 'Could not read that file. Save it as .xlsx or .csv and try again.'; return; }
  const t = pickTable(wb);
  if (!t) { res.textContent = 'Could not find an AWB / Tracking ID or Order ID column in this file.'; return; }
  const fixedPlat = $('impPlat').value || (/meesho/i.test(file.name) ? 'Meesho' : /flipkart|fk_|settlement/i.test(file.name) ? 'Flipkart' : '');
  const { items, skipped } = buildImport(t, fixedPlat);
  let matched = 0, added = 0, notFound = 0;
  for (let i = 0; i < items.length; i += 200) {
    res.textContent = `Saving to the Sheet… ${Math.min(i + 200, items.length)} of ${items.length}`;
    try { const r = await api('POST', { action: 'ledgerImport', items: items.slice(i, i + 200) }); matched += r.matched; added += r.added; notFound += r.skipped; }
    catch (err) { res.textContent = 'Stopped: ' + err.message + `. ${matched + added} rows were saved before that.`; return; }
  }
  res.textContent = `Done, from sheet “${t.sheet}”. ${matched} matched to your parcels, ${added} new AWBs added${notFound ? `, ${notFound} rows had only an order ID that is not linked to any AWB yet (import the orders report first)` : ''}${skipped ? `, ${skipped} rows had no AWB or order ID` : ''}. Columns read: ${Object.keys(t.map).join(', ')}.`;
  toast('Imported ' + (matched + added) + ' rows');
  load();
}

/* ---------- export ---------- */
$('btnExport').onclick = () => {
  if (!window.XLSX) { toast('Excel export did not load.'); return; }
  const rows = plist.map(p => ({ 'AWB': p.awb, 'Platform': p.platform, 'Order ID': p.orderId, 'Order Item ID': p.orderItemId, 'SKU': p.sku, 'Product': p.product, 'Packet': p.packet, 'Scanned': p.scanned ? 'Yes' : 'No', 'Godown scan-out': p._gOut === undefined ? '' : p._gOut ? 'Yes' : 'No', 'Dispatch date': p.dispatchDate, 'Status': (SMAP[p.status] || SMAP.dispatched).l, 'Status date': p.statusDate, 'Cancel reason': p.cancelReason, 'Return received': p.returnReceived ? RET[p.returnReceived][0] : '', 'Sale ₹': p.salePrice, 'Expected ₹': p._exp, 'Fees ₹': p.fees, 'Received ₹': p.received, 'Short ₹': p._short || null, 'Settlement date': p.settlementDate, 'UTR': p.utr, 'Claim': p.claim ? CLAIM_T[p.claim.type] : '', 'Claim status': p.claim ? (CLAIM_S[p.claim.status] || CLAIM_S.open)[0] : '', 'Claim ₹': p.claim ? p.claim.amount : null, 'Recovered ₹': p.claim ? p.claim.recovered : null, 'Notes': p.notes }));
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.length ? rows : [{ '': 'No rows' }]), 'Parcels');
  XLSX.writeFile(wb, `Parcel_ledger_${$('gFrom').value}_to_${$('gTo').value}.xlsx`);
};

/* ---------- boot ---------- */
function start() {
  $('loginView').hidden = true; $('appView').hidden = false;
  $('hdrSub').textContent = `Every AWB, from the packing table to the bank · ${user.name}`;
  const h = (location.hash || '').slice(1); show(VIEWS.includes(h) ? h : 'ledger');
  const f = fbInit();
  if (f) f.auth.onAuthStateChanged(u => { if (u && !gd.on) { gd.on = true; loadGodown().then(() => { derive(); render(); }); } renderGodownBox(); });
  load();
}
if (url && token && user) start(); else showLogin('');
if (typeof module !== 'undefined') module.exports = { buildImport, mapHeaders, normDate, normStatus };
})();
