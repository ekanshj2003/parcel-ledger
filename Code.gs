/**
 * Dispatch Scanner – Google Sheet backend (login + Flipkart orders)
 * -----------------------------------------------------------------
 * 1. Paste this whole file into Extensions > Apps Script of your Google Sheet.
 * 2. Choose "setup" in the function list and click Run once. Allow the permissions.
 *    This creates the tabs Dispatch, Products, Orders, Summary and Users.
 * 3. Open the Users tab. Change the Admin PIN (it starts as 1234) and add your packers.
 *    Role is "admin" (can delete scans and link packet codes) or "packer".
 *    Set Active to NO to block someone.
 * 4. Deploy > New deployment > Web app.  Execute as: Me.  Who has access: Anyone.
 *    Copy the Web app URL (ends with /exec) into the app's login page, under "Connect this device".
 *
 * Flipkart orders (optional):
 *    Project Settings (gear icon on the left) > Script Properties > Add script property:
 *      FK_APP_ID      = your Flipkart Self Access API key / Application ID
 *      FK_APP_SECRET  = your Flipkart Self Access secret
 *    Then run "setup" again. It checks the keys, pulls orders now, and refreshes them every 15 minutes.
 *    Never paste these keys into the code itself or share them with anyone.
 *
 * Settings tab: Meesho / Flipkart courier lists, whether a packet barcode may repeat, and the Drive folder name.
 * Cancelled parcels: moved out of Dispatch into the Cancelled tab (Undo puts them back). Every delete / cancel is listed in Changes Log.
 * Signed manifests are saved in Google Drive > Dispatch Manifests > <date>.
 *
 * Parcel Ledger app: setup also creates the Ledger and Rates tabs and, when the Flipkart keys are
 *    saved, syncs shipped / delivered / returned / cancelled Flipkart orders into Ledger every hour.
 *
 * After editing this code later: Deploy > Manage deployments > Edit > Version: New version > Deploy.
 */

const TZ = 'Asia/Kolkata';
const DISPATCH = 'Dispatch', PRODUCTS = 'Products', SUMMARY = 'Summary', USERS = 'Users', ORDERS = 'Orders', SETTINGS = 'Settings', MANIFESTS = 'Manifests', CHANGES = 'Changes Log', CANCELLED = 'Cancelled';
const XHEAD = ['Cancelled On', 'Cancelled Time', 'Reason', 'Cancelled By', 'AWB Number', 'Packet Barcode', 'Platform', 'Courier', 'Scanned On', 'Scanned Time', 'Scanned By', 'Scan Timestamp', 'Original Row'];
const HEAD = ['Date', 'Time', 'Platform', 'AWB Number', 'Packet Barcode', 'Product', 'Scanned By', 'Timestamp', 'Order ID', 'Order SKU', 'Check', 'Courier', 'Status', 'Cancelled At', 'Cancelled By', 'Cancel Reason'];
const CHEAD = ['Date', 'Time', 'Change', 'AWB Number', 'Packet Barcode', 'Platform', 'Courier', 'Done By', 'Reason', 'Timestamp', 'AWB After'];
const NOT_PICKED = /pick/i;      // cancel reasons like "Courier did not pick up": the same AWB is dispatched again later
const OHEAD = ['Platform', 'AWB', 'AWB 2', 'Order ID', 'Order Item ID', 'Shipment ID', 'SKU', 'Qty', 'Product', 'State', 'Dispatch By', 'Courier', 'Updated'];
const LOGIN_DAYS = 30;          // stay signed in for 30 days on a phone
const MAX_WRONG_PIN = 5;        // lock a name for 10 minutes after 5 wrong PINs
const KEEP_ORDER_DAYS = 15;     // keep synced orders this long so late scans still match
const MHEAD = ['Date', 'Time', 'Platform', 'Courier', 'Page', 'File', 'Link', 'Uploaded By', 'Parcels Scanned'];
const SETTING_DEFAULTS = [
  ['Meesho couriers', 'Valmo, Delhivery, Shadowfax, Xpressbees, Ecom Express, DTDC', 'Names shown in the app when Meesho is picked. Separate with commas.'],
  ['Flipkart couriers', 'Ekart, Delhivery, Shadowfax, Xpressbees', 'Used for manifests and when an order has no courier.'],
  ['Packet barcode used only once', 'YES', 'YES = a packet barcode can never be scanned twice. NO = same barcode allowed on many parcels (e.g. product EAN).'],
  ['Manifest folder name', 'Dispatch Manifests', 'Folder in Google Drive. One sub-folder per day is made inside it.']
];
const FK_API = 'https://api.flipkart.net/sellers';
const FK_TOKEN_URL = 'https://api.flipkart.net/oauth-service/oauth/token?grant_type=client_credentials&scope=Seller_Api';

/* ---------- one-time setup ---------- */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const header = (sh, cols) => { sh.setFrozenRows(1); sh.getRange(1, 1, 1, cols).setFontWeight('bold').setBackground('#e1f1ee'); };

  const d = ss.getSheetByName(DISPATCH) || ss.insertSheet(DISPATCH);
  if (d.getLastRow() === 0) d.appendRow(HEAD); else d.getRange(1, 1, 1, HEAD.length).setValues([HEAD]);
  d.getRange('A:E').setNumberFormat('@'); d.getRange('I:J').setNumberFormat('@'); d.getRange('L:P').setNumberFormat('@');
  header(d, HEAD.length); d.setColumnWidths(1, HEAD.length, 130);

  const p = ss.getSheetByName(PRODUCTS) || ss.insertSheet(PRODUCTS);
  if (p.getLastRow() === 0) p.appendRow(['Packet Barcode', 'Product Name', 'SKU']); else p.getRange(1, 1, 1, 3).setValues([['Packet Barcode', 'Product Name', 'SKU']]);
  p.getRange('A:A').setNumberFormat('@'); p.getRange('C:C').setNumberFormat('@');
  header(p, 3); p.setColumnWidths(1, 3, 200);

  const props0 = PropertiesService.getScriptProperties();
  if (props0.getProperty('FK_APP_ID') || ss.getSheetByName(ORDERS)) {      // Orders tab only when Flipkart is connected
    const o = ss.getSheetByName(ORDERS) || ss.insertSheet(ORDERS);
    if (o.getLastRow() === 0) o.appendRow(OHEAD);
    o.getRange('A:G').setNumberFormat('@');
    header(o, OHEAD.length); o.setColumnWidths(1, OHEAD.length, 130);
  }

  const u = ss.getSheetByName(USERS) || ss.insertSheet(USERS);
  if (u.getLastRow() === 0) { u.appendRow(['Name', 'PIN', 'Role', 'Active']); u.appendRow(['Admin', '1234', 'admin', 'YES']); }
  u.getRange('A:B').setNumberFormat('@');
  header(u, 4); u.setColumnWidths(1, 4, 140);
  u.getRange('C2:C500').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['admin', 'packer'], true).build());
  u.getRange('D2:D500').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['YES', 'NO'], true).build());

  const st = ss.getSheetByName(SETTINGS) || ss.insertSheet(SETTINGS);
  if (st.getLastRow() === 0) st.appendRow(['Setting', 'Value', 'What it does']);
  const have = st.getLastRow() > 1 ? st.getRange(2, 1, st.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
  SETTING_DEFAULTS.forEach(r => { if (!have.includes(r[0])) st.appendRow(r); });
  header(st, 3); st.setColumnWidths(1, 1, 220); st.setColumnWidths(2, 1, 380); st.setColumnWidths(3, 1, 420);

  const cl = ss.getSheetByName(CHANGES) || ss.insertSheet(CHANGES);
  if (cl.getLastRow() === 0) cl.appendRow(CHEAD); else cl.getRange(1, 1, 1, CHEAD.length).setValues([CHEAD]);
  cl.getRange('A:B').setNumberFormat('@'); cl.getRange('D:E').setNumberFormat('@');
  header(cl, CHEAD.length); cl.setColumnWidths(1, CHEAD.length, 130);

  const mf = ss.getSheetByName(MANIFESTS) || ss.insertSheet(MANIFESTS);
  if (mf.getLastRow() === 0) mf.appendRow(MHEAD);
  mf.getRange('A:B').setNumberFormat('@');
  header(mf, MHEAD.length); mf.setColumnWidths(1, MHEAD.length, 140);

  const cx = ss.getSheetByName(CANCELLED) || ss.insertSheet(CANCELLED);
  if (cx.getLastRow() === 0) cx.appendRow(XHEAD);
  cx.getRange('A:B').setNumberFormat('@'); cx.getRange('E:F').setNumberFormat('@'); cx.getRange('I:J').setNumberFormat('@');
  header(cx, XHEAD.length); cx.setColumnWidths(1, XHEAD.length, 130); cx.hideColumns(XHEAD.length);
  const moved = moveOldCancelled();
  if (moved) Logger.log('Moved ' + moved + ' cancelled parcels out of Dispatch into the Cancelled tab.');

  const fixed = fixLongNumbers();
  if (fixed) Logger.log('Repaired ' + fixed + ' long AWB / barcode numbers that showed like 1.49084E+15.');

  try { manifestRoot(); } catch (x) { Logger.log('Drive folder not created yet: ' + x.message); }

  const s = ss.getSheetByName(SUMMARY) || ss.insertSheet(SUMMARY);
  s.clear();
  s.getRange('A1').setFormula(
    "=QUERY(Dispatch!A:M,\"select A, C, L, count(D) where D is not null and M != 'Cancelled' group by A, C, L order by A desc, C label L 'Delivery Partner', count(D) 'Parcels'\",1)");
  s.getRange(1, 1, 1, 4).setFontWeight('bold'); s.setColumnWidths(1, 4, 160);

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SECRET')) props.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());

  // Flipkart: if keys are saved, test them, pull orders now, and refresh every 15 minutes.
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'syncFlipkart').forEach(t => ScriptApp.deleteTrigger(t));
  if (props.getProperty('FK_APP_ID') && props.getProperty('FK_APP_SECRET')) {
    ScriptApp.newTrigger('syncFlipkart').timeBased().everyMinutes(15).create();
    const r = syncFlipkart();
    Logger.log(r.ok ? 'Flipkart connected. Orders with AWB: ' + r.withAwb + ' (of ' + r.items + ' order items).' : 'Flipkart problem: ' + r.error);
  } else {
    Logger.log('Setup finished. (Flipkart connection is off.)');
  }
  setupLedger();
}

/* ---------- web entry points ---------- */
function doGet(e) {
  const q = e.parameter || {};
  try {
    if (q.action === 'ping') return out({ ok: true, app: 'dispatch-scanner', v: 7, ledger: 1 });
    const me = auth(q.token);
    if (q.action === 'list') return out({ ok: true, rows: rowsBetween(q.date, q.date) });
    if (q.action === 'range') return out({ ok: true, rows: rowsBetween(q.from, q.to) });
    if (q.action === 'products') return out({ ok: true, products: products() });
    if (q.action === 'orders') return out({ ok: true, orders: ordersList(), sync: syncStatus() });
    if (q.action === 'me') return out({ ok: true, user: me });
    if (q.action === 'ledger') return out(ledgerData(q.from, q.to));
    if (q.action === 'settings') return out({ ok: true, settings: settings() });
    if (q.action === 'used') return out(Object.assign({ ok: true }, usedSince(Number(q.since) || 0, Number(q.esince) || 0)));
    if (q.action === 'cancelled') return out({ ok: true, rows: cancelledBetween(q.date, q.date) });
    if (q.action === 'cancelledRange') return out({ ok: true, rows: cancelledBetween(q.from, q.to) });
    if (q.action === 'search') return out({ ok: true, rows: searchScans(q.q) });
    if (q.action === 'manifests') return out({ ok: true, manifests: manifestsOn(q.date) });
    return out({ ok: false, error: 'Unknown action' });
  } catch (err) { return fail(err); }
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse((e.postData && e.postData.contents) || '{}'); } catch (x) { return out({ ok: false, error: 'Bad request' }); }
  if (body.action === 'login') { try { return out(login(body.name, body.pin)); } catch (err) { return fail(err); } }
  if (body.action === 'manifest') {                     // file upload: outside the lock so scans never wait
    try { return out(saveManifest(body, auth(body.token))); } catch (err) { return fail(err); }
  }
  if (body.action === 'syncNow') {                      // runs outside the lock: it can take a few seconds
    try { admin(auth(body.token)); const r = syncFlipkart(); return out(Object.assign({ orders: ordersList(), sync: syncStatus() }, r)); }
    catch (err) { return fail(err); }
  }
  if (body.action === 'syncLedger') {                   // Flipkart → Ledger, outside the lock
    try { admin(auth(body.token)); return out(syncFlipkartLedger()); } catch (err) { return fail(err); }
  }
  const lock = LockService.getScriptLock();
  try {
    const me = auth(body.token);
    lock.waitLock(20000);
    if (body.action === 'add') return out(addScan(body.rec || {}, me));
    if (body.action === 'delete') { admin(me); return out(deleteScan(body.awb, me, body.ts)); }
    if (body.action === 'deleteMany') { admin(me); return out(deleteMany(body.items, me)); }
    if (body.action === 'cancel') return out(cancelScan(body.awb, body.reason, me));
    if (body.action === 'uncancel') { admin(me); return out(uncancelScan(body.awb, me, body.ts)); }
    if (body.action === 'product') { admin(me); return out(saveProduct(body.code, body.name, body.sku)); }
    if (body.action === 'deleteProduct') { admin(me); return out(deleteProduct(body.code)); }
    if (body.action === 'ledgerSave') return out(upsertLedger([body.item || {}], 'manual', me.name));
    if (body.action === 'ledgerImport') return out(upsertLedger(body.items || [], 'import', me.name));
    if (body.action === 'ledgerDelete') { admin(me); return out(deleteLedger(body.awb)); }
    if (body.action === 'rate') { admin(me); return out(saveRate(body.rate || {})); }
    return out({ ok: false, error: 'Unknown action' });
  } catch (err) { return fail(err); }
  finally { try { lock.releaseLock(); } catch (x) {} }
}

/* ---------- login ---------- */
function login(name, pin) {
  name = String(name || '').trim(); pin = String(pin || '').trim();
  if (!name || !pin) return { ok: false, error: 'Enter your name and PIN.' };
  const cache = CacheService.getScriptCache(), ck = 'fail_' + name.toLowerCase();
  const fails = Number(cache.get(ck) || 0);
  if (fails >= MAX_WRONG_PIN) return { ok: false, error: 'Too many wrong PINs. Try again in 10 minutes.' };
  const u = findUser(name);
  if (!u || u.pin !== pin) {
    cache.put(ck, String(fails + 1), 600);
    return { ok: false, error: 'Name or PIN is wrong.' };
  }
  if (!u.active) return { ok: false, error: 'This login is switched off. Ask the admin.' };
  cache.remove(ck);
  const exp = Date.now() + LOGIN_DAYS * 864e5;
  const payload = Utilities.base64EncodeWebSafe(JSON.stringify({ n: u.name, exp }));
  return { ok: true, token: payload + '.' + sign(payload), user: { name: u.name, role: u.role } };
}

function auth(token) {
  const parts = String(token || '').split('.');
  if (parts.length !== 2 || sign(parts[0]) !== parts[1]) throw authErr('Please log in again.');
  let p; try { p = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString()); } catch (x) { throw authErr('Please log in again.'); }
  if (!p.exp || p.exp < Date.now()) throw authErr('Your login expired. Please log in again.');
  const u = findUser(p.n);
  if (!u || !u.active) throw authErr('This login is switched off. Ask the admin.');
  return { name: u.name, role: u.role };
}
function admin(me) { if (me.role !== 'admin') throw new Error('Only an admin can do this.'); }
function authErr(msg) { const e = new Error(msg); e.auth = true; return e; }

function sign(s) {
  const secret = PropertiesService.getScriptProperties().getProperty('SECRET');
  if (!secret) throw new Error('Run setup once in Apps Script.');
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(s, secret)).replace(/=+$/, '');
}

function findUser(name) {
  const sh = sheet(USERS), n = sh.getLastRow();
  if (n < 2) return null;
  const want = String(name || '').trim().toLowerCase();
  const row = sh.getRange(2, 1, n - 1, 4).getValues().find(v => String(v[0]).trim().toLowerCase() === want);
  if (!row) return null;
  return { name: String(row[0]).trim(), pin: String(row[1]).trim(), role: String(row[2]).trim().toLowerCase() === 'admin' ? 'admin' : 'packer', active: String(row[3]).trim().toUpperCase() !== 'NO' };
}

/* ---------- scans ---------- */
function addScan(rec, me) {
  const awb = String(rec.awb || '').trim(), packet = String(rec.packet || '').trim();
  if (!awb || !packet) return { ok: false, error: 'AWB and packet are both required.' };
  const sh = sheet(DISPATCH);
  const dupOf = (col, value, kind) => { const r = findRow(sh, col, value); return r ? { ok: false, dup: true, kind, code: value, existing: toObj(sh.getRange(r, 1, 1, HEAD.length).getValues()[0]) } : null; };
  // AWB used before: blocked, unless every earlier scan of it was cancelled (parcel not picked up → goes out again)
  const prev = awbRows(sh, awb).map(r => toObj(sh.getRange(r, 1, 1, HEAD.length).getValues()[0]));
  let again = false;
  if (prev.length) {
    const live = prev.find(o => o.status !== 'Cancelled');
    if (live) return { ok: false, dup: true, kind: 'awb', code: awb, existing: live };
    const last = prev[prev.length - 1];
    if (!rec.redispatch && !NOT_PICKED.test(last.cancelReason)) return { ok: false, dup: true, kind: 'awb', code: awb, existing: last };
    again = true;
  }
  const d1 = dupOf(5, awb, 'awb-as-packet');                                   // an AWB can never be a used packet code
  if (d1) return d1;
  const d2 = dupOf(4, packet, 'packet-as-awb');                                // a packet code can never be an AWB
  if (d2) return d2;
  if (settings().uniquePacket) {
    const r3 = findRows(sh, 5, packet).find(r => String(sh.getRange(r, 13).getValue()) !== 'Cancelled');
    if (r3) return { ok: false, dup: true, kind: 'packet', code: packet, existing: toObj(sh.getRange(r3, 1, 1, HEAD.length).getValues()[0]) };
  }
  const ts = Number(rec.ts) || Date.now(), when = new Date(ts);
  const prod = products().find(p => p.code === packet);
  const clip = (v, n) => String(v || '').slice(0, n);
  const row = [Utilities.formatDate(when, TZ, 'yyyy-MM-dd'), Utilities.formatDate(when, TZ, 'HH:mm:ss'),
    rec.platform === 'Meesho' ? 'Meesho' : 'Flipkart', awb, packet, prod ? prod.name : clip(rec.product, 120), me.name, ts,
    clip(rec.orderId, 60), clip(rec.orderSku, 200), again ? 'DISPATCHED AGAIN' : clip(rec.check, 20), clip(rec.courier, 40)];
  sh.appendRow(asText(row, [3, 4, 8]));          // AWB, packet and order ID always stored as text
  return { ok: true, row: toObj(row) };
}

function deleteScan(awb, me, ts) {
  awb = String(awb || '').trim();
  const sh = sheet(DISPATCH), r = pickRow(sh, awb, ts);
  if (!r) return { ok: false, error: 'AWB not found.' };
  const row = toObj(sh.getRange(r, 1, 1, HEAD.length).getValues()[0]);
  sh.deleteRow(r);
  const ev = logChange('deleted', row, me, 'Wrong entry deleted', awbState(sh, awb));   // phones read this and allow the codes again
  return { ok: true, event: ev };
}

// Delete several wrong entries in one go (admin). Each one is logged like a single delete.
function deleteMany(items, me) {
  items = (Array.isArray(items) ? items : []).slice(0, 300);
  const done = [], failed = [], events = [];
  items.forEach(it => {
    const awb = String((it && it.awb) || '').trim();
    if (!awb) return;
    const r = deleteScan(awb, me, it.ts);
    if (r.ok) { done.push({ awb, ts: Number(it.ts) || 0 }); events.push(r.event); } else failed.push({ awb, error: r.error });
  });
  return { ok: true, done, failed, events };
}

// All rows of one AWB, oldest first.
function awbRows(sh, awb) { return findRows(sh, 4, awb).sort((x, y) => x - y); }
// The exact row (matched by its timestamp) or else the newest one.
function pickRow(sh, awb, ts) {
  const rows = awbRows(sh, awb);
  if (!rows.length) return 0;
  if (ts) { const hit = rows.find(r => Number(sh.getRange(r, 8).getValue()) === Number(ts)); if (hit) return hit; }
  return rows[rows.length - 1];
}
// What an AWB looks like now, for phones to update their list: none / active / cancelled.
function awbState(sh, awb) {
  const objs = awbRows(sh, awb).map(r => toObj(sh.getRange(r, 1, 1, HEAD.length).getValues()[0]));
  if (!objs.length) return { s: 'none' };
  const live = objs.find(o => o.status !== 'Cancelled');
  if (live) return { s: 'active', p: live.packet, d: live.date };
  const last = objs[objs.length - 1];
  return { s: 'cancelled', d: last.date, at: last.cancelledAt.slice(0, 10), r: last.cancelReason };
}

// Parcel cancelled at pickup: mark it, keep the AWB blocked, give the packet barcode back for reuse.
function cancelScan(awb, reason, me) {
  awb = String(awb || '').trim();
  const sh = sheet(DISPATCH), rows = awbRows(sh, awb);
  const r = rows.slice().reverse().find(x => String(sh.getRange(x, 13).getValue()) !== 'Cancelled');
  if (!r) {
    const old = cancelledRows(awb);
    if (old.length) { const o = old[old.length - 1]; return { ok: false, already: true, row: o, error: 'Already cancelled ' + o.cancelledAt + ' by ' + o.cancelledBy + '.' }; }
    return { ok: false, notFound: true, error: 'AWB ' + awb + ' was never scanned for dispatch.' };
  }
  reason = String(reason || 'Cancelled at pickup').slice(0, 80);
  const vals = sh.getRange(r, 1, 1, HEAD.length).getValues()[0];
  const done = moveToCancelled(vals, reason, me.name, new Date());
  sh.deleteRow(r);                                            // gone from Dispatch: the AWB and packet can be scanned again
  const ev = logChange('cancelled', done, me, reason, awbState(sh, awb));
  return { ok: true, row: done, event: ev };
}

// Put one Dispatch row into the Cancelled tab (keeps the original row so Undo can put it back exactly).
function moveToCancelled(vals, reason, byName, when) {
  const row = toObj(vals);
  const at = Utilities.formatDate(when, TZ, 'yyyy-MM-dd'), tm = Utilities.formatDate(when, TZ, 'HH:mm:ss');
  const clean = vals.slice(0, 12).map(v => v instanceof Date ? '' : v);
  clean[0] = row.date; clean[1] = row.time;
  sheet(CANCELLED).appendRow(asText([at, tm, reason, byName, row.awb, row.packet, row.platform, row.courier, row.date, row.time, row.by, row.ts, JSON.stringify(clean)], [4, 5]));
  return Object.assign(row, { status: 'Cancelled', cancelledAt: at + ' ' + tm.slice(0, 5), cancelledBy: byName, cancelReason: reason });
}

// Earlier versions kept cancelled parcels in Dispatch with Status = Cancelled. Move those out once.
function moveOldCancelled() {
  const sh = sheet(DISPATCH), n = sh.getLastRow();
  if (n < 2) return 0;
  const vals = sh.getRange(2, 1, n - 1, HEAD.length).getValues();
  const hit = [];
  vals.forEach((v, i) => { if (String(v[12]) === 'Cancelled') hit.push(i); });
  hit.forEach(i => {
    const v = vals[i], at = String(v[13] || '');
    const when = /^\d{4}-\d{2}-\d{2}/.test(at) ? new Date(at.slice(0, 10) + 'T' + (at.slice(11, 16) || '00:00') + ':00+05:30') : new Date();
    moveToCancelled(v, String(v[15] || 'Cancelled'), String(v[14] || ''), when);
  });
  hit.slice().reverse().forEach(i => sh.deleteRow(i + 2));   // bottom up, so row numbers stay right
  return hit.length;
}

function cancelledRows(awb) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CANCELLED);
  if (!sh || sh.getLastRow() < 2) return [];
  return findRows(sh, 5, awb).sort((x, y) => x - y).map(r => cxObj(sh.getRange(r, 1, 1, XHEAD.length).getValues()[0], r));
}
function cxObj(v, rowNo) {
  const d = x => x instanceof Date ? Utilities.formatDate(x, TZ, 'yyyy-MM-dd') : String(x);
  const t = x => x instanceof Date ? Utilities.formatDate(x, TZ, 'HH:mm:ss') : String(x);
  return { date: d(v[8]), time: t(v[9]), platform: String(v[6]), awb: String(v[4]), packet: String(v[5]), courier: String(v[7]), by: String(v[10]), ts: Number(v[11]) || 0,
    status: 'Cancelled', cancelledAt: d(v[0]) + ' ' + t(v[1]).slice(0, 5), cancelledBy: String(v[3]), cancelReason: String(v[2]), _row: rowNo };
}

function uncancelScan(awb, me, ts) {
  awb = String(awb || '').trim();
  const cx = sheet(CANCELLED), sh = sheet(DISPATCH);
  const list = cancelledRows(awb);
  if (!list.length) return { ok: false, error: 'This parcel is not in the Cancelled list.' };
  const c = (ts && list.find(o => o.ts === Number(ts))) || list[list.length - 1];
  const live = awbRows(sh, awb).find(x => String(sh.getRange(x, 13).getValue()) !== 'Cancelled');
  if (live) return { ok: false, error: 'This AWB was already scanned again on ' + sh.getRange(live, 1).getDisplayValue() + '. Nothing to undo.' };
  if (settings().uniquePacket) {
    const other = findRows(sh, 5, c.packet).find(x => String(sh.getRange(x, 13).getValue()) !== 'Cancelled');
    if (other) return { ok: false, error: 'Packet ' + c.packet + ' was already used again for AWB ' + sh.getRange(other, 4).getValue() + '. Cannot undo.' };
  }
  let orig = [];
  try { orig = JSON.parse(String(cx.getRange(c._row, XHEAD.length).getValue() || '[]')); } catch (x) {}
  if (!orig.length) orig = [c.date, c.time, c.platform, c.awb, c.packet, '', c.by, c.ts, '', '', '', c.courier];
  while (orig.length < 12) orig.push('');
  const row = orig.slice(0, 12).concat(['', '', '', '']);
  sh.appendRow(asText(row, [3, 4, 8]));
  cx.deleteRow(c._row);
  const done = toObj(row);
  const ev = logChange('uncancelled', done, me, 'Cancel undone', awbState(sh, awb));
  return { ok: true, row: done, event: ev };
}

// Find parcels by AWB or packet barcode (full or part of it), any date. Newest first, up to 50.
function searchScans(q) {
  q = String(q || '').trim();
  if (q.length < 3) return [];
  const out = [];
  const sh = sheet(DISPATCH), n = sh.getLastRow();
  if (n > 1) {
    const rows = {};
    sh.getRange(2, 4, n - 1, 2).createTextFinder(q).matchCase(false).findAll().forEach(c => { rows[c.getRow()] = true; });
    Object.keys(rows).forEach(r => out.push(toObj(sh.getRange(Number(r), 1, 1, HEAD.length).getValues()[0])));
  }
  const cx = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CANCELLED);
  if (cx && cx.getLastRow() > 1) {
    const rows = {};
    cx.getRange(2, 5, cx.getLastRow() - 1, 2).createTextFinder(q).matchCase(false).findAll().forEach(c => { rows[c.getRow()] = true; });
    Object.keys(rows).forEach(r => { const o = cxObj(cx.getRange(Number(r), 1, 1, XHEAD.length).getValues()[0], Number(r)); delete o._row; out.push(o); });
  }
  return out.sort((x, y) => (y.ts || 0) - (x.ts || 0)).slice(0, 50);
}

function cancelledBetween(from, to) {
  if (!from || !to) throw new Error('Dates missing.');
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CANCELLED);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, XHEAD.length).getValues().map((v, i) => { const o = cxObj(v, i + 2); delete o._row; return o; })
    .filter(o => o.cancelledAt.slice(0, 10) >= from && o.cancelledAt.slice(0, 10) <= to);
}

// Every delete / cancel / undo is written here. Phones read it to keep their list of used codes right.
function logChange(type, row, me, reason, after) {
  const props = PropertiesService.getScriptProperties();
  const ts = Math.max(Date.now(), Number(props.getProperty('LAST_CHANGE_TS') || 0) + 1);
  props.setProperty('LAST_CHANGE_TS', String(ts));
  const when = new Date(ts);
  let sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CHANGES);
  if (!sh) { sh = SpreadsheetApp.getActiveSpreadsheet().insertSheet(CHANGES); sh.appendRow(CHEAD); }
  sh.appendRow(asText([Utilities.formatDate(when, TZ, 'yyyy-MM-dd'), Utilities.formatDate(when, TZ, 'HH:mm:ss'), type, row.awb, row.packet, row.platform, row.courier, me.name, reason || '', ts, JSON.stringify(after || {})], [3, 4]));
  return { t: ts, type, awb: row.awb, packet: row.packet, date: row.date, at: Utilities.formatDate(when, TZ, 'yyyy-MM-dd'), reason: reason || '', after: after || {} };
}

/* ---------- packet codes ---------- */
function saveProduct(code, name, sku) {
  code = String(code || '').trim(); name = String(name || '').trim(); sku = String(sku || '').trim();
  if (!code || (!name && !sku)) return { ok: false, error: 'Enter the barcode and a product name or SKU.' };
  const sh = sheet(PRODUCTS), r = findRow(sh, 1, code);
  if (r) {
    const old = sh.getRange(r, 1, 1, 3).getValues()[0];
    sh.getRange(r, 1, 1, 3).setValues([asText([code, name || old[1], sku || old[2]], [0, 2])]);
  } else sh.appendRow(asText([code, name, sku], [0, 2]));
  if (name) {
    const d = sheet(DISPATCH), n = d.getLastRow();
    if (n > 1) {
      const rng = d.getRange(2, 5, n - 1, 2), vals = rng.getValues(); let changed = false;
      vals.forEach(v => { if (String(v[0]) === code && !v[1]) { v[1] = name; changed = true; } });
      if (changed) rng.setValues(vals);
    }
  }
  return { ok: true };
}

function deleteProduct(code) {
  const sh = sheet(PRODUCTS), r = findRow(sh, 1, String(code || '').trim());
  if (r) sh.deleteRow(r); return { ok: true };
}

/* ---------- Flipkart ---------- */
// Remove anything that sneaks in when copying: spaces, line breaks, quotes, invisible characters.
function cleanKey(v) { return String(v || '').replace(/[\s"'`​-‍﻿ ]/g, ''); }
function fkKeys() {
  const props = PropertiesService.getScriptProperties();
  return { id: cleanKey(props.getProperty('FK_APP_ID')), secret: cleanKey(props.getProperty('FK_APP_SECRET')) };
}

// Ask Flipkart for a token. Tries GET first (as in Flipkart's docs), then POST.
function fkTokenRequest(id, secret) {
  const tries = [];
  for (const method of ['get', 'post']) {
    const res = UrlFetchApp.fetch(FK_TOKEN_URL, {
      method, muteHttpExceptions: true,
      headers: { Authorization: 'Basic ' + Utilities.base64Encode(id + ':' + secret), Accept: 'application/json' }
    });
    const code = res.getResponseCode(), text = res.getContentText();
    let body = text;
    if (/<html/i.test(text)) {                            // an HTML error page: keep only its title / heading
      const m = text.match(/<title[^>]*>([^<]*)<\/title>/i) || text.match(/<h1[^>]*>([^<]*)<\/h1>/i);
      body = 'web page: ' + (m ? m[1].trim() : 'no title');
    }
    tries.push({ method: method.toUpperCase(), code, body: body.replace(/\s+/g, ' ').slice(0, 160) });
    if (code === 200) { try { const j = JSON.parse(text); if (j.access_token) return { ok: true, json: j, tries }; } catch (x) {} }
  }
  return { ok: false, tries };
}

function fkToken(force) {
  const cache = CacheService.getScriptCache();
  if (!force) { const t = cache.get('fk_token'); if (t) return t; }
  const k = fkKeys();
  if (!k.id || !k.secret) throw new Error('Flipkart keys are missing. Add FK_APP_ID and FK_APP_SECRET in Script Properties.');
  const r = fkTokenRequest(k.id, k.secret);
  if (!r.ok) {
    throw new Error('Flipkart did not accept the keys. ' + r.tries.map(t => t.method + ' → HTTP ' + t.code + ' (' + (t.body || 'no message') + ')').join('; ') + '. Run checkFlipkart for details.');
  }
  const ttl = Math.max(60, Math.min(21600, Number(r.json.expires_in || 3600) - 300));
  cache.put('fk_token', r.json.access_token, ttl);
  return r.json.access_token;
}

// Run this from the Apps Script editor if Flipkart rejects the keys.
// It never prints your secret – only its length and whether it looks right.
function checkFlipkart() {
  const props = PropertiesService.getScriptProperties();
  const rawId = String(props.getProperty('FK_APP_ID') || ''), rawSecret = String(props.getProperty('FK_APP_SECRET') || '');
  const k = fkKeys();
  const mask = s => s.length <= 6 ? '(too short)' : s.slice(0, 3) + '…' + s.slice(-2);
  Logger.log('FK_APP_ID: length ' + k.id.length + ', looks like ' + mask(k.id) + (rawId.length !== k.id.length ? '  (removed ' + (rawId.length - k.id.length) + ' extra spaces/quotes)' : ''));
  Logger.log('FK_APP_SECRET: length ' + k.secret.length + (rawSecret.length !== k.secret.length ? '  (removed ' + (rawSecret.length - k.secret.length) + ' extra spaces/quotes)' : ''));
  if (!k.id || !k.secret) { Logger.log('One of the keys is empty. Add both in Project Settings > Script Properties.'); return; }
  if (k.id === k.secret) Logger.log('WARNING: ID and secret are the same value. The secret is a different value from the API key.');
  let r = fkTokenRequest(k.id, k.secret);
  r.tries.forEach(t => Logger.log('Token ' + t.method + ': HTTP ' + t.code + ' – ' + t.body));
  if (r.ok) { Logger.log('SUCCESS: Flipkart accepted the keys. Now run setup.'); return; }
  const sw = fkTokenRequest(k.secret, k.id);
  if (sw.ok) { Logger.log('The keys are SWAPPED. Move the value in FK_APP_ID into FK_APP_SECRET and the value in FK_APP_SECRET into FK_APP_ID, then run setup.'); return; }
  Logger.log('Still rejected. Check in Seller Hub > Manage Profile > Developer Access > Self Access that you copied the Application ID and the Application Secret of that access (not the access name), and that it has not expired.');
}

function fkCall(method, path, body) {
  const url = /^https?:/.test(path) ? path : FK_API + path;
  for (let attempt = 0; attempt < 2; attempt++) {
    const opt = { method, muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + fkToken(attempt > 0), Accept: 'application/json' } };
    if (body) { opt.contentType = 'application/json'; opt.payload = JSON.stringify(body); }
    const res = UrlFetchApp.fetch(url, opt), code = res.getResponseCode();
    if (code === 401 && attempt === 0) continue;                 // token expired: get a new one and retry once
    if (code < 200 || code >= 300) throw new Error('Flipkart ' + method.toUpperCase() + ' ' + path.split('?')[0] + ' failed (HTTP ' + code + '): ' + res.getContentText().slice(0, 200));
    return JSON.parse(res.getContentText() || '{}');
  }
}

// Pull every pre-dispatch shipment, read its tracking ID (AWB), and update the Orders tab.
function syncFlipkart() {
  const props = PropertiesService.getScriptProperties();
  const cache = CacheService.getScriptCache();          // own flag, so scans never wait for a sync
  if (cache.get('fk_sync_running')) return { ok: false, error: 'A sync is already running. Try again in a minute.' };
  cache.put('fk_sync_running', '1', 300);
  try {
    // 1. shipments waiting for dispatch
    const shipments = [];
    let page = fkCall('post', '/v3/shipments/filter/', {
      filter: { type: 'preDispatch', states: ['APPROVED', 'PACKING_IN_PROGRESS', 'PACKED', 'FORM_FAILED', 'READY_TO_DISPATCH'] },
      pagination: { pageSize: 20 }
    });
    for (let guard = 0; guard < 200; guard++) {
      (page.shipments || []).forEach(s => { if (s && s.shipmentId) shipments.push(s); });
      if (!page.hasMore || !page.nextPageUrl) break;
      page = fkCall('get', page.nextPageUrl);
    }

    // 2. tracking IDs come from the shipment details (25 at a time, as Flipkart recommends)
    const detail = {};
    const ids = shipments.map(s => s.shipmentId).filter(Boolean);
    for (let i = 0; i < ids.length; i += 25) {
      const j = fkCall('get', '/v3/shipments/' + ids.slice(i, i + 25).map(encodeURIComponent).join(','));
      (j.shipments || []).forEach(s => { if (s && s.shipmentId) detail[s.shipmentId] = s; });
    }

    // 3. turn into rows, one per order item
    const now = new Date(), stamp = Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm');
    const fresh = {};
    shipments.forEach(s => {
      const d = detail[s.shipmentId] || s;
      const subs = (d.subShipments || s.subShipments || []);
      let awb = '', awb2 = '', courier = '', title = '';
      subs.forEach(sub => {
        const cd = sub.courierDetails || {};
        const del = cd.deliveryDetails || {}, pick = cd.pickupDetails || {};
        if (!awb && del.trackingId) { awb = String(del.trackingId); courier = del.vendorName || ''; }
        if (!awb2 && pick.trackingId && String(pick.trackingId) !== awb) awb2 = String(pick.trackingId);
        if (!title && sub.packageTitle) title = String(sub.packageTitle);
      });
      if (!awb && awb2) { awb = awb2; awb2 = ''; }
      (s.orderItems || []).forEach(it => {
        if (!it || !it.orderItemId) return;
        const key = 'Flipkart|' + it.orderItemId;
        fresh[key] = ['Flipkart', awb, awb2, String(it.orderId || ''), String(it.orderItemId || ''), String(s.shipmentId || ''),
          String(it.sku || ''), Number(it.quantity || 1), title, String(it.status || ''), String(s.dispatchByDate || '').slice(0, 16).replace('T', ' '), courier, stamp];
      });
    });

    // 4. merge with what is already in the sheet (keep recent orders even after they ship)
    const sh = sheet(ORDERS), n = sh.getLastRow();
    const old = n > 1 ? sh.getRange(2, 1, n - 1, OHEAD.length).getValues() : [];
    const cutoff = Utilities.formatDate(new Date(now.getTime() - KEEP_ORDER_DAYS * 864e5), TZ, 'yyyy-MM-dd');
    const merged = {};
    old.forEach(r => {
      const key = r[0] + '|' + r[4];
      if (String(r[12]).slice(0, 10) >= cutoff) merged[key] = r;
    });
    Object.keys(fresh).forEach(k => {
      const prev = merged[k], row = fresh[k];
      if (prev && !row[1] && prev[1]) { row[1] = prev[1]; row[2] = prev[2]; row[11] = prev[11]; }   // never lose a known AWB
      merged[k] = row;
    });
    const rows = Object.keys(merged).map(k => merged[k]);
    if (n > 1) sh.getRange(2, 1, n - 1, OHEAD.length).clearContent();
    if (rows.length) sh.getRange(2, 1, rows.length, OHEAD.length).setValues(rows);

    const fkRows = Object.keys(fresh).map(k => fresh[k]);
    const result = { ok: true, shipments: shipments.length, items: fkRows.length, withAwb: fkRows.filter(r => r[1]).length };
    props.setProperty('FK_LAST', JSON.stringify({ at: stamp, ok: true, items: result.items, withAwb: result.withAwb }));
    return result;
  } catch (err) {
    const msg = String(err.message || err);
    props.setProperty('FK_LAST', JSON.stringify({ at: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), ok: false, error: msg.slice(0, 300) }));
    return { ok: false, error: msg };
  } finally { cache.remove('fk_sync_running'); }
}

function syncStatus() {
  const props = PropertiesService.getScriptProperties();
  const has = !!(props.getProperty('FK_APP_ID') && props.getProperty('FK_APP_SECRET'));
  let last = null; try { last = JSON.parse(props.getProperty('FK_LAST') || 'null'); } catch (x) {}
  return { flipkart: has, last };
}

function ordersList() {
  const sh = sheet(ORDERS), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, OHEAD.length).getValues()
    .filter(r => r[1] || r[2])
    .map(r => ({ p: String(r[0]), awb: String(r[1]), awb2: String(r[2]), oid: String(r[3]), sku: String(r[6]), qty: Number(r[7]) || 1, title: String(r[8]), st: String(r[9]), c: String(r[11] || '') }));
}

/* ---------- settings ---------- */
function settings() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS);
  const map = {};
  SETTING_DEFAULTS.forEach(r => { map[r[0]] = r[1]; });
  if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(r => { if (r[0]) map[String(r[0]).trim()] = String(r[1]); });
  const list = s => String(s || '').split(',').map(x => x.trim()).filter(Boolean);
  return {
    meeshoCouriers: list(map['Meesho couriers']),
    flipkartCouriers: list(map['Flipkart couriers']),
    uniquePacket: String(map['Packet barcode used only once']).trim().toUpperCase() !== 'NO',
    folderName: String(map['Manifest folder name'] || 'Dispatch Manifests').trim() || 'Dispatch Manifests'
  };
}

/* ---------- every code ever used (the app keeps a copy to block repeats instantly) ---------- */
function usedSince(since, esince) {
  const sh = sheet(DISPATCH), n = sh.getLastRow();
  const awbs = {}, packets = {}, cancelled = {}, reasons = {}, live = {};
  let latest = since;
  if (n > 1) {
    sh.getRange(2, 1, n - 1, 16).getValues().forEach(v => {
      const ts = Number(v[7]) || 0;
      if (ts <= since) return;
      const date = v[0] instanceof Date ? Utilities.formatDate(v[0], TZ, 'yyyy-MM-dd') : String(v[0]);
      const awb = String(v[3]), isCancelled = String(v[12]) === 'Cancelled';
      if (awb !== '') awbs[awb] = date;
      if (v[4] !== '' && !isCancelled) packets[String(v[4])] = date;     // a cancelled parcel's packet can be used again
      if (isCancelled) { const at = v[13] instanceof Date ? Utilities.formatDate(v[13], TZ, 'yyyy-MM-dd') : String(v[13]).slice(0, 10); cancelled[awb] = at; reasons[awb] = String(v[15] || ''); }
      else live[awb] = true;
      if (ts > latest) latest = ts;
    });
  }
  Object.keys(live).forEach(k => { delete cancelled[k]; delete reasons[k]; });   // dispatched again after a cancel
  // deletes / cancels / undos since the phone last asked
  const events = [];
  let elatest = Math.max(esince, Number(PropertiesService.getScriptProperties().getProperty('LAST_CHANGE_TS') || 0));
  const cl = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CHANGES);
  if (since > 0 && cl && cl.getLastRow() > 1) {
    cl.getRange(2, 1, cl.getLastRow() - 1, CHEAD.length).getValues().forEach(v => {
      const t = Number(v[9]) || 0;
      let after = {}; try { after = JSON.parse(String(v[10] || '{}')); } catch (x) {}
      if (t > esince) events.push({ t, type: String(v[2]), awb: String(v[3]), packet: String(v[4]), reason: String(v[8] || ''), after, at: v[0] instanceof Date ? Utilities.formatDate(v[0], TZ, 'yyyy-MM-dd') : String(v[0]) });
    });
  }
  return { v: 7, awbs, packets, cancelled, reasons, live: Object.keys(live), latest, events, elatest, full: since === 0 };
}

/* ---------- signed manifests to Google Drive ---------- */
function manifestRoot() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('MANIFEST_ROOT');
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (x) {} }
  const name = settings().folderName;
  const it = DriveApp.getFoldersByName(name);
  const f = it.hasNext() ? it.next() : DriveApp.createFolder(name);
  props.setProperty('MANIFEST_ROOT', f.getId());
  return f;
}
function dayFolder(date) {
  const root = manifestRoot(), it = root.getFoldersByName(date);
  return it.hasNext() ? it.next() : root.createFolder(date);
}

function saveManifest(body, me) {
  const now = new Date(), date = Utilities.formatDate(now, TZ, 'yyyy-MM-dd'), time = Utilities.formatDate(now, TZ, 'HH:mm:ss');
  const platform = body.platform === 'Meesho' ? 'Meesho' : 'Flipkart';
  const courier = String(body.courier || '').replace(/[^\w &.-]/g, '').slice(0, 40) || 'Courier';
  const type = String(body.type || '');
  if (!/^(image\/(jpeg|png|webp|heic|heif)|application\/pdf)$/.test(type)) return { ok: false, error: 'Only photos or PDF files can be uploaded.' };
  if (!body.data) return { ok: false, error: 'The file is empty.' };
  const bytes = Utilities.base64Decode(String(body.data));
  if (bytes.length > 20 * 1024 * 1024) return { ok: false, error: 'The file is bigger than 20 MB.' };
  const page = Math.max(1, Math.min(99, Number(body.page) || 1));
  const ext = type === 'application/pdf' ? 'pdf' : type.split('/')[1].replace('jpeg', 'jpg');
  const name = [date, platform, courier.replace(/\s+/g, '-'), time.replace(/:/g, '').slice(0, 4), 'p' + page].join('_') + '.' + ext;
  const file = dayFolder(date).createFile(Utilities.newBlob(bytes, type, name));
  file.setDescription('Signed manifest · ' + platform + ' · ' + courier + ' · uploaded by ' + me.name);

  // how many parcels were scanned today for this platform + courier, to compare with the manifest
  const parcels = rowsBetween(date, date).filter(r => r.status !== 'Cancelled' && r.platform === platform && (!r.courier || courier === 'Courier' || r.courier.toLowerCase().indexOf(courier.toLowerCase()) === 0 || courier.toLowerCase().indexOf(r.courier.toLowerCase()) === 0)).length;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { sheet(MANIFESTS).appendRow([date, time, platform, courier, page, name, file.getUrl(), me.name, parcels]); }
  finally { lock.releaseLock(); }
  return { ok: true, manifest: { date, time, platform, courier, page, name, url: file.getUrl(), by: me.name, parcels } };
}

function manifestsOn(date) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MANIFESTS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, MHEAD.length).getValues()
    .map(v => ({ date: v[0] instanceof Date ? Utilities.formatDate(v[0], TZ, 'yyyy-MM-dd') : String(v[0]), time: v[1] instanceof Date ? Utilities.formatDate(v[1], TZ, 'HH:mm:ss') : String(v[1]),
      platform: String(v[2]), courier: String(v[3]), page: Number(v[4]) || 1, name: String(v[5]), url: String(v[6]), by: String(v[7]), parcels: Number(v[8]) || 0 }))
    .filter(m => m.date === date);
}

/* ---------- helpers ---------- */
function rowsBetween(from, to) {
  if (!from || !to) throw new Error('Dates missing.');
  const sh = sheet(DISPATCH), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, HEAD.length).getValues().map(toObj).filter(r => r.date >= from && r.date <= to);
}
function products() {
  const sh = sheet(PRODUCTS), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 3).getValues().filter(v => v[0] !== '').map(v => ({ code: String(v[0]), name: String(v[1]), sku: String(v[2] || '') }));
}
function toObj(v) {
  const date = v[0] instanceof Date ? Utilities.formatDate(v[0], TZ, 'yyyy-MM-dd') : String(v[0]);
  const time = v[1] instanceof Date ? Utilities.formatDate(v[1], TZ, 'HH:mm:ss') : String(v[1]);
  return { date, time, platform: String(v[2]), awb: String(v[3]), packet: String(v[4]), product: String(v[5]), by: String(v[6]), ts: Number(v[7]) || 0,
    orderId: String(v[8] || ''), orderSku: String(v[9] || ''), check: String(v[10] || ''), courier: String(v[11] || ''),
    status: String(v[12] || ''), cancelledAt: v[13] instanceof Date ? Utilities.formatDate(v[13], TZ, 'yyyy-MM-dd HH:mm') : String(v[13] || ''), cancelledBy: String(v[14] || ''), cancelReason: String(v[15] || '') };
}
function findRow(sh, col, value) {
  if (!value || sh.getLastRow() < 2) return 0;
  const cell = sh.getRange(2, col, sh.getLastRow() - 1, 1).createTextFinder(value).matchEntireCell(true).matchCase(true).findNext();
  return cell ? cell.getRow() : 0;
}
// Long codes such as 1490840123456789 must stay text. Otherwise Sheets turns them into numbers,
// shows 1.49084E+15, and searching / duplicate checks stop matching. An apostrophe in front keeps them as text.
function asText(values, cols) {
  return values.map((v, i) => {
    if (cols.indexOf(i) < 0 || v === '' || v == null) return v;
    const s = typeof v === 'number' ? numToText(v) : String(v);
    return /^[+\-]?[\d.,eE]+$/.test(s) ? "'" + s : s;
  });
}
function numToText(n) { return Number.isInteger(n) ? n.toLocaleString('en-US', { useGrouping: false, maximumFractionDigits: 0 }) : String(n); }

// One-time repair of codes that were already turned into numbers. Runs inside setup.
function fixLongNumbers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let count = 0;
  [[DISPATCH, [4, 5, 9]], [CHANGES, [4, 5]], [PRODUCTS, [1]], [ORDERS, [2, 3, 4, 5, 6]]].forEach(([name, cols]) => {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return;
    cols.forEach(c => {
      const rng = sh.getRange(2, c, sh.getLastRow() - 1, 1);
      rng.setNumberFormat('@');
      const vals = rng.getValues();
      let changed = false;
      const out = vals.map(r => {
        if (typeof r[0] === 'number') { changed = true; count++; return ["'" + numToText(r[0])]; }
        const s = String(r[0]);
        return [/^[+\-]?[\d.,eE]+$/.test(s) ? "'" + s : r[0]];    // keep digit-only text as text when written back
      });
      if (changed) rng.setValues(out);
    });
  });
  return count;
}

function findRows(sh, col, value) {
  if (!value || sh.getLastRow() < 2) return [];
  return sh.getRange(2, col, sh.getLastRow() - 1, 1).createTextFinder(value).matchEntireCell(true).matchCase(true).findAll().map(c => c.getRow());
}
function sheet(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" missing. Run setup once in Apps Script.');
  return sh;
}
function out(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function fail(err) { return out({ ok: false, error: String(err.message || err), auth: !!err.auth }); }

/* =====================================================================
 * PARCEL LEDGER (added for the Parcel Ledger app)
 * Ledger tab: one row per AWB – order, status, sale price, settlement,
 * return and claim. Rates tab: settlement per packet for each SKU.
 * Flipkart: every hour the script reads shipped / delivered / returned /
 * cancelled orders and fills status, sale price, SKU and AWB.
 * Flipkart has no settlement API: import the Settlement report in the app.
 * ===================================================================== */
const LEDGER = 'Ledger', RATES = 'Rates';
const LHEAD = ['AWB', 'Platform', 'Order ID', 'Order Item ID', 'SKU', 'Product', 'Qty', 'Order Date', 'Courier', 'Status', 'Status Date',
  'Sale Price', 'Expected', 'Fees', 'Received', 'Settlement Date', 'UTR', 'Return Received', 'Claim Type', 'Claim Amount', 'Claim ID',
  'Claim Raised', 'Claim Status', 'Recovered', 'Notes', 'Source', 'Updated', 'Updated By', 'Cancel Reason'];
const LKEYS = ['awb', 'platform', 'orderId', 'orderItemId', 'sku', 'product', 'qty', 'orderDate', 'courier', 'status', 'statusDate',
  'salePrice', 'expected', 'fees', 'received', 'settlementDate', 'utr', 'returnReceived', 'claimType', 'claimAmount', 'claimId',
  'claimRaised', 'claimStatus', 'recovered', 'notes', 'source', 'updated', 'updatedBy', 'cancelReason'];
const LTEXT = [0, 2, 3, 4, 16, 20];                           // AWB, order IDs, SKU, UTR, claim ID: always text
const LNUM = ['qty', 'salePrice', 'expected', 'fees', 'received', 'claimAmount', 'recovered'];
const LDATE = ['orderDate', 'statusDate', 'settlementDate', 'claimRaised'];
const RHEAD = ['SKU', 'Product', 'Flipkart settlement per packet', 'Meesho settlement per packet', 'Cost per packet'];
const LEDGER_SYNC_DAYS = 30;                                  // Flipkart orders placed in the last 30 days
const STATUS_RANK = { dispatched: 0, transit: 1, ofd: 2, delivered: 3, return: 4, rto: 4, cancelled: 4, lost: 5 };

function setupLedger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const header = (sh, cols) => { sh.setFrozenRows(1); sh.getRange(1, 1, 1, cols).setFontWeight('bold').setBackground('#e1f1ee'); };
  const l = ss.getSheetByName(LEDGER) || ss.insertSheet(LEDGER);
  if (l.getLastRow() === 0) l.appendRow(LHEAD); else l.getRange(1, 1, 1, LHEAD.length).setValues([LHEAD]);
  ['A:A', 'C:E', 'Q:Q', 'U:U'].forEach(a => l.getRange(a).setNumberFormat('@'));
  header(l, LHEAD.length); l.setColumnWidths(1, LHEAD.length, 120);
  const r = ss.getSheetByName(RATES) || ss.insertSheet(RATES);
  if (r.getLastRow() === 0) r.appendRow(RHEAD); else r.getRange(1, 1, 1, RHEAD.length).setValues([RHEAD]);
  r.getRange('A:A').setNumberFormat('@');
  header(r, RHEAD.length); r.setColumnWidths(1, RHEAD.length, 190);

  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'syncFlipkartLedger').forEach(t => ScriptApp.deleteTrigger(t));
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('FK_APP_ID') && props.getProperty('FK_APP_SECRET')) {
    ScriptApp.newTrigger('syncFlipkartLedger').timeBased().everyHours(1).create();
    const s = syncFlipkartLedger();
    Logger.log(s.ok ? 'Ledger: Flipkart sync done. ' + s.shipments + ' shipments read, ' + s.updated + ' parcels updated, ' + s.added + ' added' + (s.noAwb ? ', ' + s.noAwb + ' without AWB' : '') + ' (' + s.seconds + ' s).' : 'Ledger: Flipkart problem: ' + s.error);
  } else Logger.log('Ledger tabs ready. (Flipkart keys not saved, so no automatic sync.)');
}

const awbKey = v => String(v == null ? '' : v).replace(/[\s'"]/g, '').toUpperCase();
const oidKey = v => String(v == null ? '' : v).replace(/[\s'"]/g, '');
function lFmtDate(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v == null ? '' : v); }
function lToObj(v) {
  const o = {};
  LKEYS.forEach((k, i) => {
    let x = v[i];
    if (LNUM.indexOf(k) >= 0) o[k] = (x === '' || x == null || isNaN(Number(x))) ? null : Number(x);
    else if (LDATE.indexOf(k) >= 0) o[k] = lFmtDate(x);
    else if (k === 'updated') o[k] = x instanceof Date ? Utilities.formatDate(x, TZ, "yyyy-MM-dd'T'HH:mm:ss") : String(x || '');
    else o[k] = typeof x === 'number' ? numToText(x) : String(x == null ? '' : x);
  });
  return o;
}
function lToRow(o) { return asText(LKEYS.map(k => (o[k] == null ? '' : o[k])), LTEXT); }

function ledgerRows(from) {
  const sh = sheet(LEDGER), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, LHEAD.length).getValues().map(lToObj)
    .filter(o => o.awb && (!from || !(o.orderDate || o.updated) || (o.orderDate || o.updated.slice(0, 10)) >= from || ['', 'dispatched', 'transit', 'ofd'].indexOf(o.status) >= 0 || (o.claimStatus && o.claimStatus !== 'paid' && o.claimStatus !== 'rejected')));
}
function rates() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RATES);
  if (!sh || sh.getLastRow() < 2) return [];
  const n = v => (v === '' || v == null || isNaN(Number(v))) ? null : Number(v);
  return sh.getRange(2, 1, sh.getLastRow() - 1, RHEAD.length).getValues().filter(v => v[0] !== '')
    .map(v => ({ sku: String(v[0]).trim(), product: String(v[1] || ''), fk: n(v[2]), ms: n(v[3]), cost: n(v[4]) }));
}
function saveRate(r) {
  const sku = String(r.sku || '').trim(); if (!sku) return { ok: false, error: 'SKU is required.' };
  const sh = sheet(RATES), row = findRow(sh, 1, sku);
  const vals = asText([sku, r.product || '', r.fk == null ? '' : r.fk, r.ms == null ? '' : r.ms, r.cost == null ? '' : r.cost], [0]);
  if (row) sh.getRange(row, 1, 1, RHEAD.length).setValues([vals]); else sh.appendRow(vals);
  return { ok: true };
}

/**
 * Upsert many ledger changes in one read + one write.
 * items: [{awb, orderId, orderItemId, f:{field:value}}]
 * mode 'manual': empty values clear the field. 'import': empty values are ignored.
 * mode 'api': like import, and status only moves forward.
 */
function upsertLedger(items, mode, byName) {
  const sh = sheet(LEDGER), n = sh.getLastRow();
  const rows = n > 1 ? sh.getRange(2, 1, n - 1, LHEAD.length).getValues().map(lToObj) : [];
  const byAwb = {}, byOid = {};
  rows.forEach((o, i) => { if (o.awb) byAwb[awbKey(o.awb)] = i; if (o.orderItemId) byOid[oidKey(o.orderItemId)] = i; if (o.orderId && byOid[oidKey(o.orderId)] == null) byOid[oidKey(o.orderId)] = i; });
  const changed = {}, stamp = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss");
  let matched = 0, added = 0, skipped = 0;
  (items || []).forEach(it => {
    const f = it.f || {};
    const awb = String(it.awb || f.awb || '').trim();
    let i = awb ? byAwb[awbKey(awb)] : undefined;
    if (i == null && it.orderItemId) i = byOid[oidKey(it.orderItemId)];
    if (i == null && it.orderId) i = byOid[oidKey(it.orderId)];
    if (i == null) {
      if (!awb) { skipped++; return; }
      rows.push({ awb }); i = rows.length - 1; byAwb[awbKey(awb)] = i; added++;
    } else matched++;
    const o = rows[i], before = o.status || '';
    if (it.orderItemId && !f.orderItemId) f.orderItemId = it.orderItemId;
    if (it.orderId && !f.orderId) f.orderId = it.orderId;
    Object.keys(f).forEach(k => {
      if (LKEYS.indexOf(k) < 0 || k === 'updated' || k === 'updatedBy' || (k === 'awb' && o.awb)) return;
      let v = f[k];
      if (v === '' || v == null) { if (mode === 'manual') o[k] = (LNUM.indexOf(k) >= 0 ? null : ''); return; }
      if (LNUM.indexOf(k) >= 0) { v = Number(v); if (isNaN(v)) return; }
      if (k === 'status' && mode !== 'manual' && before && (STATUS_RANK[v] == null ? 0 : STATUS_RANK[v]) < (STATUS_RANK[before] || 0)) return;
      if (k === 'source' && o.source && String(o.source).indexOf(v) >= 0) return;
      if (k === 'source' && o.source) v = o.source + ', ' + v;
      o[k] = v;
    });
    if (o.status !== before && !f.statusDate && mode !== 'manual') o.statusDate = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    if (o.orderItemId) byOid[oidKey(o.orderItemId)] = i;
    if (o.orderId && byOid[oidKey(o.orderId)] == null) byOid[oidKey(o.orderId)] = i;
    o.updated = stamp; o.updatedBy = byName || '';
    changed[i] = true;
  });
  const idx = Object.keys(changed).map(Number);
  if (idx.length) {
    const startNew = n - 1;            // rows at index >= this are new
    const oldIdx = idx.filter(i => i < startNew), newIdx = idx.filter(i => i >= startNew).sort((a, b) => a - b);
    if (oldIdx.length > 40) {           // many edits: rewrite the existing block once
      sh.getRange(2, 1, startNew, LHEAD.length).setValues(rows.slice(0, startNew).map(lToRow));
    } else oldIdx.forEach(i => sh.getRange(i + 2, 1, 1, LHEAD.length).setValues([lToRow(rows[i])]));
    if (newIdx.length) sh.getRange(n + 1, 1, newIdx.length, LHEAD.length).setValues(newIdx.map(i => lToRow(rows[i])));
  }
  return { ok: true, matched, added, skipped };
}

function ledgerData(from, to) {
  return { ok: true, ledger: ledgerRows(from), scans: rowsBetween(from, to), cancelled: cancelledBetween(from, to),
    products: products(), rates: rates(), sync: ledgerSyncStatus() };
}
function ledgerSyncStatus() {
  const props = PropertiesService.getScriptProperties();
  let last = null; try { last = JSON.parse(props.getProperty('LEDGER_LAST') || 'null'); } catch (x) {}
  return { flipkart: !!(props.getProperty('FK_APP_ID') && props.getProperty('FK_APP_SECRET')), last };
}

/* ---------- Flipkart → Ledger ---------- */
function fkLedgerStatus(items, courierReturn) {
  let best = 'dispatched';
  const rank = s => STATUS_RANK[s] == null ? 0 : STATUS_RANK[s];
  (items || []).forEach(it => {
    const st = String(it.status || '').toUpperCase();
    let s = 'dispatched';
    if (st === 'PICKUP_COMPLETE' || st === 'SHIPPED') s = 'transit';
    else if (st === 'DELIVERED') s = 'delivered';
    else if (st === 'RETURN_REQUESTED' || st === 'RETURNED') s = (it.courierReturn || courierReturn) ? 'rto' : 'return';
    else if (st === 'CANCELLED') s = (it.courierReturn || courierReturn) ? 'rto' : 'cancelled';
    if (rank(s) >= rank(best)) best = s;
  });
  return best;
}
// Read every page of one Flipkart search. Flipkart's "next page" links can be longer than
// Apps Script allows; when that happens the date window is split in half and each half is read again.
function fkIso(d) { return Utilities.formatDate(d, TZ, "yyyy-MM-dd'T'HH:mm:ss"); }
function fkFilterRange(base, field, from, to, deadline, depth) {
  const filter = Object.assign({}, base);
  if (field) filter[field] = { from: fkIso(from), to: fkIso(to) };
  let page = fkCall('post', '/v3/shipments/filter/', { filter, pagination: { pageSize: 20 } });
  const out = [];
  let truncated = false;
  for (let guard = 0; guard < 260; guard++) {
    (page.shipments || []).forEach(s => { if (s && s.shipmentId) out.push(s); });
    if (!page.hasMore || !page.nextPageUrl || Date.now() > deadline) break;
    const next = /^https?:/.test(page.nextPageUrl) ? page.nextPageUrl : FK_API + page.nextPageUrl;
    if (next.length > 1900) { truncated = true; break; }
    try { page = fkCall('get', page.nextPageUrl); }
    catch (e) { if (/URL length|Limit Exceeded/i.test(String(e.message || e))) { truncated = true; break; } throw e; }
  }
  if (truncated && field && (to - from) > 2 * 3600e3 && (depth || 0) < 10 && Date.now() < deadline) {
    const mid = new Date(Math.floor((from.getTime() + to.getTime()) / 2));
    return fkFilterRange(base, field, from, mid, deadline, (depth || 0) + 1).concat(fkFilterRange(base, field, mid, to, deadline, (depth || 0) + 1));
  }
  return out;
}
// Try the date filters Flipkart may need for this search type; keep the first that returns orders.
function fkSearch(label, base, fields, from, to, deadline, log, errors) {
  let lastErr = '';
  for (const field of fields) {
    try {
      const got = fkFilterRange(base, field, from, to, deadline, 0);
      log.push(label + ' [' + (field || 'no date filter') + ']: ' + got.length);
      if (got.length) return got;
    } catch (e) { lastErr = String(e.message || e).slice(0, 160); log.push(label + ' [' + (field || 'no date filter') + ']: error ' + lastErr); }
    if (Date.now() > deadline) break;
  }
  if (lastErr) errors.push(label + ': ' + lastErr);
  return [];
}
function syncFlipkartLedger() {
  const props = PropertiesService.getScriptProperties(), cache = CacheService.getScriptCache();
  if (!(props.getProperty('FK_APP_ID') && props.getProperty('FK_APP_SECRET'))) return { ok: false, error: 'Flipkart keys are not saved in Script Properties.' };
  if (cache.get('fk_ledger_running')) return { ok: false, error: 'A ledger sync is already running. Try again in a few minutes.' };
  cache.put('fk_ledger_running', '1', 360);
  const t0 = Date.now(), deadline = t0 + 4.5 * 60 * 1000;
  try {
    const now = new Date(), fromD = new Date(now.getTime() - LEDGER_SYNC_DAYS * 864e5);
    const shipments = [], errors = [], log = [];
    sheet(LEDGER).getRange(1, 1, 1, LHEAD.length).setValues([LHEAD]);
    fkSearch('shipped/delivered', { type: 'postDispatch', states: ['PICKUP_COMPLETE', 'SHIPPED', 'DELIVERED'] },
      ['orderDate', null], fromD, now, deadline, log, errors).forEach(s => shipments.push(s));
    ['marketplaceCancellation', 'sellerCancellation', 'buyerCancellation'].forEach(ct =>
      fkSearch('cancelled ' + ct, { type: 'cancelled', states: ['CANCELLED'], cancellationType: ct },
        ['cancellationDate', 'orderDate'], fromD, now, deadline, log, errors).forEach(s => { s._ct = ct; shipments.push(s); }));
    const seeded = seedLedgerFromOrders();
    if (seeded) log.push('orders copied from Orders tab: ' + seeded);
    const ids = ledgerItemIdsToTrack();
    let byId = 0;
    for (let i = 0; i < ids.length && Date.now() < deadline; i += 40) {
      try {
        const j = fkCall('get', '/v3/shipments?orderItemIds=' + ids.slice(i, i + 40).map(encodeURIComponent).join(','));
        (j.shipments || []).forEach(sh => { if (sh && sh.shipmentId) { shipments.push(sh); byId++; } });
      } catch (e) { errors.push('status by order ID: ' + String(e.message || e).slice(0, 160)); break; }
    }
    log.push('status checked by order item ID: ' + ids.length + ' orders, ' + byId + ' shipments returned');
    if (!shipments.length) {
      props.setProperty('LEDGER_LAST', JSON.stringify({ at: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), ok: !errors.length, shipments: 0, updated: 0, added: 0, noAwb: 0, warn: log.join(' | ').slice(0, 300), error: errors.join(' | ').slice(0, 300) }));
      return { ok: !errors.length, shipments: 0, updated: 0, added: 0, noAwb: 0, seconds: Math.round((Date.now() - t0) / 1000), log, error: errors.join(' | ') || undefined };
    }

    // AWBs already known (Ledger or Orders tab), so details are fetched only for new shipments
    const known = {};
    const lsh = sheet(LEDGER), ln = lsh.getLastRow();
    if (ln > 1) lsh.getRange(2, 1, ln - 1, 4).getValues().forEach(v => { if (v[0] && v[3]) known[oidKey(v[3])] = String(v[0]); });
    const osh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS);
    if (osh && osh.getLastRow() > 1) osh.getRange(2, 1, osh.getLastRow() - 1, 5).getValues().forEach(v => { if (v[1] && v[4] && !known[oidKey(v[4])]) known[oidKey(v[4])] = String(v[1]); });

    const seen = {}, list = [];
    shipments.forEach(s => { if (!seen[s.shipmentId]) { seen[s.shipmentId] = 1; list.push(s); } });
    const need = list.filter(s => !(s.orderItems || []).some(it => known[oidKey(it.orderItemId)])).map(s => s.shipmentId);
    const detail = {};
    for (let i = 0; i < need.length && Date.now() < deadline; i += 25) {
      try {
        const j = fkCall('get', '/v3/shipments/' + need.slice(i, i + 25).map(encodeURIComponent).join(','));
        (j.shipments || []).forEach(s => { if (s && s.shipmentId) detail[s.shipmentId] = s; });
      } catch (e) { errors.push('details: ' + String(e.message || e).slice(0, 120)); }
    }

    const items = [], reasons = {};
    list.forEach(s => {
      const its = s.orderItems || [];
      if (!its.length) return;
      let awb = '', courier = '', title = '';
      its.forEach(it => { if (!awb && known[oidKey(it.orderItemId)]) awb = known[oidKey(it.orderItemId)]; });
      const d = detail[s.shipmentId];
      (d && d.subShipments || s.subShipments || []).forEach(sub => {
        const cd = sub.courierDetails || {}, del = cd.deliveryDetails || {}, pick = cd.pickupDetails || {};
        if (!awb && del.trackingId) awb = String(del.trackingId);
        if (!awb && pick.trackingId) awb = String(pick.trackingId);
        if (!courier) courier = del.vendorName || pick.vendorName || '';
        if (!title && sub.packageTitle) title = String(sub.packageTitle);
      });
      const price = its.reduce((a, it) => { const pc = it.priceComponents || {}; return a + (Number(pc.customerPrice != null ? pc.customerPrice : pc.totalPrice) || 0); }, 0);
      const f = {
        platform: 'Flipkart', orderId: String(its[0].orderId || ''), orderItemId: String(its[0].orderItemId || ''),
        sku: its.map(it => String(it.sku || '')).filter(Boolean).join(', '), qty: its.reduce((a, it) => a + (Number(it.quantity) || 1), 0),
        orderDate: String(its[0].orderDate || '').slice(0, 10), status: fkLedgerStatus(its, false), source: 'Flipkart API'
      };
      if (price) f.salePrice = Math.round(price * 100) / 100;
      const cx = its.find(it => it.cancellationReason || it.cancellationSubReason);
      if (cx) {
        const why = [cx.cancellationReason, cx.cancellationSubReason].filter(Boolean).join(' – ');
        f.cancelReason = (cx.courierReturn ? 'Courier return: ' : '') + why;
        const who = { marketplaceCancellation: 'Flipkart', sellerCancellation: 'Seller', buyerCancellation: 'Buyer' }[s._ct] || 'Cancelled';
        const k = who + ' | ' + f.cancelReason + (awb ? '' : ' (before AWB)');
        reasons[k] = (reasons[k] || 0) + 1;
      }
      if (courier) f.courier = courier;
      if (title) f.product = title;
      items.push({ awb, orderItemId: f.orderItemId, orderId: f.orderId, f });
    });
    const r = upsertLedger(items, 'api', 'Flipkart sync');
    Object.keys(reasons).sort((a, b) => reasons[b] - reasons[a]).forEach(k => log.push('cancel reason ' + k.replace(' | ', ': ') + ' → ' + reasons[k]));
    const res = { ok: true, shipments: list.length, updated: r.matched, added: r.added, noAwb: r.skipped, log, errors: errors.length ? errors : undefined, seconds: Math.round((Date.now() - t0) / 1000) };
    props.setProperty('LEDGER_LAST', JSON.stringify({ at: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), ok: true, shipments: list.length, updated: r.matched, added: r.added, noAwb: r.skipped, warn: (errors.length ? errors.join(' | ') + ' | ' : '') + log.join(' | ').slice(0, 300) }));
    return res;
  } catch (err) {
    const msg = String(err.message || err);
    props.setProperty('LEDGER_LAST', JSON.stringify({ at: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), ok: false, error: msg.slice(0, 300) }));
    return { ok: false, error: msg };
  } finally { cache.remove('fk_ledger_running'); }
}

function deleteLedger(awb) {
  const sh = sheet(LEDGER), n = sh.getLastRow(); if (n < 2) return { ok: true };
  const vals = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) if (awbKey(vals[i][0]) === awbKey(awb)) sh.deleteRow(i + 2);
  return { ok: true };
}

// Run this from the Apps Script editor to see exactly what the Flipkart ledger sync does.
function runLedgerSync() {
  const r = syncFlipkartLedger();
  (r.log || []).forEach(l => Logger.log('Search ' + l));
  Logger.log(r.ok ? 'Done: ' + r.shipments + ' shipments read, ' + r.updated + ' parcels updated, ' + r.added + ' added' + (r.noAwb ? ', ' + r.noAwb + ' without AWB' : '') + '.' : 'Problem: ' + r.error);
}

// Copy Flipkart orders the script already knows (Orders tab: AWB + order item ID, before dispatch) into the Ledger.
function seedLedgerFromOrders() {
  const osh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS);
  if (!osh || osh.getLastRow() < 2) return 0;
  const lsh = sheet(LEDGER), have = {};
  if (lsh.getLastRow() > 1) lsh.getRange(2, 1, lsh.getLastRow() - 1, 4).getValues().forEach(v => { if (v[3]) have[oidKey(v[3])] = 1; });
  const items = [];
  osh.getRange(2, 1, osh.getLastRow() - 1, OHEAD.length).getValues().forEach(r => {
    const awb = String(r[1] || '').trim(), oi = String(r[4] || '').trim();
    if (!awb || !oi || have[oidKey(oi)] || String(r[0]) !== 'Flipkart') return;
    items.push({ awb, orderItemId: oi, orderId: String(r[3] || ''), f: { platform: 'Flipkart', orderId: String(r[3] || ''), orderItemId: oi,
      sku: String(r[6] || ''), qty: Number(r[7]) || 1, product: String(r[8] || ''), courier: String(r[11] || ''), status: 'dispatched', source: 'Flipkart API' } });
  });
  if (!items.length) return 0;
  return upsertLedger(items, 'api', 'Flipkart sync').added;
}
// Flipkart order item IDs in the Ledger whose story is not finished yet (not cancelled, not delivered-and-paid long ago).
function ledgerItemIdsToTrack() {
  const sh = sheet(LEDGER), n = sh.getLastRow();
  if (n < 2) return [];
  const cutoff = Utilities.formatDate(new Date(Date.now() - 60 * 864e5), TZ, 'yyyy-MM-dd');
  const ids = {};
  sh.getRange(2, 1, n - 1, LHEAD.length).getValues().map(lToObj).forEach(o => {
    if (o.platform !== 'Flipkart' || !o.orderItemId || o.status === 'cancelled') return;
    if (o.orderDate && o.orderDate < cutoff) return;
    if (o.status === 'delivered' && o.received != null && o.statusDate && o.statusDate < Utilities.formatDate(new Date(Date.now() - 15 * 864e5), TZ, 'yyyy-MM-dd')) return;
    ids[oidKey(o.orderItemId)] = 1;
  });
  return Object.keys(ids);
}
