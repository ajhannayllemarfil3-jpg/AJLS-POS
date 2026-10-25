/**
 * AJL's Lumpia and Bread Roll POS - Google Apps Script backend.
 * Bind this script to the Google Sheet (Extensions > Apps Script), then deploy as a Web App.
 * @OnlyCurrentDoc
 */
const T = {
  'Products': ['Product ID','Product Name','Description','Selling Price','Unit Cost','Current Stock','Low Stock Threshold','Status','Date Created'],
  'Sales': ['Transaction ID','Actual Sales Date','Recorded At','Total Amount','Payment Method','Customer Name','Notes','Record Type','Cash Received','Change'],
  'Sale Items': ['Item ID','Transaction ID','Product ID','Product Name','Quantity','Unit Price','Unit Cost','Line Total'],
  'Expenses': ['Expense ID','Actual Expense Date','Recorded At','Category','Description','Amount','Receipt Reference','Notes'],
  'Inventory': ['Inventory Record ID','Product ID','Movement Date','Movement Type','Quantity','Reason','Related Transaction ID','Recorded At'],
  'Settings': ['Key','Value'],
  'Audit Logs': ['Log ID','Action','Record Type','Record ID','Previous Values','Updated Values','Timestamp']
};
const PAY = ['Cash','GCash','Maya','Bank Transfer','Other'];
const SETTING_KEYS = ['Business Name','Tagline','Phone','Email','Address','Facebook','Low Stock Default','Receipt Footer'];
const tz_ = () => Session.getScriptTimeZone();
const now_ = () => Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm:ss');
const day_ = () => Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
const r2_ = n => Math.round(n * 100) / 100;
const sh_ = n => SpreadsheetApp.getActive().getSheetByName(n);

function setup() { ensure_(); }

function ensure_() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(T).forEach(n => {
    if (ss.getSheetByName(n)) return;
    const sh = ss.insertSheet(n), h = T[n];
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#1d5a3a').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    h.forEach((k, i) => { if (/Date|Recorded At/.test(k)) sh.getRange(1, i + 1, sh.getMaxRows(), 1).setNumberFormat('@'); });
    if (n === 'Products') {
      add_(n, {'Product ID':'P001','Product Name':'Lumpia','Description':'Crispy fried lumpia','Selling Price':5,'Unit Cost':0,'Current Stock':0,'Low Stock Threshold':20,'Status':'Active','Date Created':day_()});
      add_(n, {'Product ID':'P002','Product Name':'Bread Roll','Description':'Soft bread roll','Selling Price':10,'Unit Cost':0,'Current Stock':0,'Low Stock Threshold':20,'Status':'Active','Date Created':day_()});
    }
    if (n === 'Settings') {
      [["Business Name","AJL's Lumpia and Bread Roll"],['Tagline','Fresh lumpia and bread rolls'],['Phone',''],['Email',''],['Address',''],['Facebook',''],['Low Stock Default','20'],['Receipt Footer','Salamat po! Thank you!']]
        .forEach(r => add_(n, {Key: r[0], Value: r[1]}));
    }
  });
  const d = ss.getSheetByName('Sheet1');
  if (d && ss.getSheets().length > 1 && d.getLastRow() === 0) ss.deleteSheet(d);
}

/* ---------- sheet helpers ---------- */
function rows_(n) {
  const v = sh_(n).getDataRange().getValues(), h = v.shift();
  return v.filter(r => r[0] !== '').map(r => {
    const o = {};
    h.forEach((k, i) => { const x = r[i]; o[k] = x instanceof Date ? Utilities.formatDate(x, tz_(), 'yyyy-MM-dd') : x; });
    return o;
  });
}
function add_(n, o) { sh_(n).appendRow(T[n].map(k => o[k] === undefined ? '' : o[k])); }
function find_(n, id) {
  const s = sh_(n), last = s.getLastRow(); if (last < 2) return 0;
  const v = s.getRange(1, 1, last, 1).getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === String(id)) return i + 1;
  return 0;
}
function set_(n, id, o) {
  const r = find_(n, id); if (!r) throw new Error('Record not found: ' + id);
  sh_(n).getRange(r, 1, 1, T[n].length).setValues([T[n].map(k => o[k] === undefined ? '' : o[k])]);
}
function del_(n, id) { const r = find_(n, id); if (r) sh_(n).deleteRow(r); }
function audit_(action, type, id, prev, next) {
  add_('Audit Logs', {'Log ID':'LOG-' + Utilities.getUuid().slice(0, 8).toUpperCase(), 'Action':action, 'Record Type':type, 'Record ID':id,
    'Previous Values': prev ? JSON.stringify(prev) : '', 'Updated Values': next ? JSON.stringify(next) : '', 'Timestamp': now_()});
}

/* ---------- validators ---------- */
function num_(v, name, min) { const n = Number(v); if (v === '' || v === null || !isFinite(n) || n < min) throw new Error(name + ' is invalid.'); return r2_(n); }
function int_(v, name) { const n = Number(v); if (!Number.isInteger(n) || n < 1) throw new Error(name + ' must be a whole number above 0.'); return n; }
function date_(v, name) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(v || '') || isNaN(new Date(v))) throw new Error(name + ' is invalid.');
  if (v > day_()) throw new Error(name + ' cannot be in the future.');
  return v;
}
function id_(v) { if (!/^[A-Za-z0-9_-]{4,60}$/.test(v || '')) throw new Error('Missing or invalid record ID.'); return v; }
function pay_(v) { if (PAY.indexOf(v) < 0) throw new Error('Invalid payment method.'); return v; }

/* ---------- inventory ---------- */
function move_(pid, type, qty, reason, date, tx, invId) {
  const p = rows_('Products').find(x => x['Product ID'] === pid);
  if (!p) throw new Error('Unknown product ' + pid);
  p['Current Stock'] = (Number(p['Current Stock']) || 0) + qty;
  set_('Products', pid, p);
  add_('Inventory', {'Inventory Record ID': invId || 'INV-' + Utilities.getUuid().slice(0, 8).toUpperCase(), 'Product ID': pid, 'Movement Date': date,
    'Movement Type': type, 'Quantity': qty, 'Reason': reason, 'Related Transaction ID': tx || '', 'Recorded At': now_()});
}
function netStock_(tx) {
  const n = {};
  rows_('Inventory').filter(m => m['Related Transaction ID'] === tx && (m['Movement Type'] === 'Sale' || m['Movement Type'] === 'Correction'))
    .forEach(m => { n[m['Product ID']] = (n[m['Product ID']] || 0) + Number(m.Quantity); });
  return n;
}
function reverse_(tx, why) {
  const n = netStock_(tx);
  Object.keys(n).forEach(pid => { if (n[pid] !== 0) move_(pid, 'Correction', -n[pid], why, day_(), tx); });
  return Object.keys(n).some(k => n[k] !== 0);
}

/* ---------- sales ---------- */
function buildItems_(id, items, old) {
  if (!Array.isArray(items) || !items.length) throw new Error('Add at least one product.');
  const prods = rows_('Products');
  return items.map((it, i) => {
    const p = prods.find(x => x['Product ID'] === it.productId); if (!p) throw new Error('Unknown product.');
    const q = int_(it.quantity, 'Quantity'), pr = num_(it.unitPrice, 'Price', 0);
    const o = (old || []).find(x => x['Product ID'] === p['Product ID']);
    return {'Item ID': id + '-' + (i + 1), 'Transaction ID': id, 'Product ID': p['Product ID'], 'Product Name': p['Product Name'], 'Quantity': q,
      'Unit Price': pr, 'Unit Cost': o ? Number(o['Unit Cost']) || 0 : Number(p['Unit Cost']) || 0, 'Line Total': r2_(q * pr)};
  });
}
const H = {};
H.saveSale = d => {
  const id = id_(d.id);
  if (find_('Sales', id)) return {duplicate: true, id: id};
  const type = d.recordType === 'Historical Sale' ? 'Historical Sale' : 'New Sale';
  const date = type === 'New Sale' ? day_() : date_(d.date, 'Sales date');
  const pm = pay_(d.paymentMethod), its = buildItems_(id, d.items), total = r2_(its.reduce((t, i) => t + i['Line Total'], 0));
  let cash = '', chg = '';
  if (type === 'New Sale' && pm === 'Cash') { cash = num_(d.cashReceived, 'Cash received', 0); if (cash < total) throw new Error('Cash received is less than the total.'); chg = r2_(cash - total); }
  const rec = {'Transaction ID': id, 'Actual Sales Date': date, 'Recorded At': now_(), 'Total Amount': total, 'Payment Method': pm, 'Customer Name': String(d.customer || '').slice(0, 100),
    'Notes': String(d.notes || '').slice(0, 300), 'Record Type': type, 'Cash Received': cash, 'Change': chg};
  add_('Sales', rec);
  its.forEach(i => add_('Sale Items', i));
  if (type === 'New Sale' || d.deductStock) its.forEach(i => move_(i['Product ID'], 'Sale', -i.Quantity, type === 'New Sale' ? 'POS sale' : 'Historical sale (stock deducted by choice)', date, id));
  audit_('Create', 'Sale', id, null, {sale: rec, items: its});
  return {id: id, total: total, date: date};
};
H.updateSale = d => {
  const id = id_(d.id), old = rows_('Sales').find(s => s['Transaction ID'] === id);
  if (!old) throw new Error('Sale not found.');
  const oldItems = rows_('Sale Items').filter(i => i['Transaction ID'] === id);
  const date = date_(d.date, 'Sales date'), pm = pay_(d.paymentMethod), its = buildItems_(id, d.items, oldItems);
  const total = r2_(its.reduce((t, i) => t + i['Line Total'], 0));
  const had = reverse_(id, 'Reversal for correction of ' + id);
  const deduct = d.deductStock === undefined ? had : !!d.deductStock;
  oldItems.forEach(i => del_('Sale Items', i['Item ID']));
  its.forEach(i => add_('Sale Items', i));
  if (deduct) its.forEach(i => move_(i['Product ID'], 'Sale', -i.Quantity, 'Re-applied after correction', date, id));
  const rec = Object.assign({}, old, {'Actual Sales Date': date, 'Total Amount': total, 'Payment Method': pm, 'Customer Name': String(d.customer || '').slice(0, 100), 'Notes': String(d.notes || '').slice(0, 300)});
  if (old['Cash Received'] !== '') rec['Change'] = r2_(Math.max(0, Number(old['Cash Received']) - total));
  set_('Sales', id, rec);
  audit_('Update', 'Sale', id, {sale: old, items: oldItems}, {sale: rec, items: its});
  return {id: id, total: total};
};
H.deleteSale = d => {
  const id = id_(d.id), old = rows_('Sales').find(s => s['Transaction ID'] === id);
  if (!old) throw new Error('Sale not found.');
  const oldItems = rows_('Sale Items').filter(i => i['Transaction ID'] === id);
  reverse_(id, 'Reversal for deletion of ' + id);
  oldItems.forEach(i => del_('Sale Items', i['Item ID']));
  del_('Sales', id);
  audit_('Delete', 'Sale', id, {sale: old, items: oldItems}, null);
  return {id: id};
};

/* ---------- products ---------- */
H.saveProduct = d => {
  const name = String(d.name || '').trim(); if (!name) throw new Error('Product name is required.');
  const price = num_(d.price, 'Selling price', 0), cost = num_(d.cost || 0, 'Unit cost', 0), low = num_(d.lowStock || 0, 'Low-stock threshold', 0);
  const status = d.status === 'Inactive' ? 'Inactive' : 'Active', ps = rows_('Products');
  if (d.id) {
    const p = ps.find(x => x['Product ID'] === d.id); if (!p) throw new Error('Product not found.');
    const prev = Object.assign({}, p);
    Object.assign(p, {'Product Name': name, 'Description': String(d.description || ''), 'Selling Price': price, 'Unit Cost': cost, 'Low Stock Threshold': low, 'Status': status});
    set_('Products', d.id, p); audit_('Update', 'Product', d.id, prev, p); return {id: d.id};
  }
  if (ps.some(x => String(x['Product Name']).toLowerCase() === name.toLowerCase())) throw new Error('A product with this name already exists.');
  const id = 'P' + String(ps.length + 1).padStart(3, '0');
  const p = {'Product ID': id, 'Product Name': name, 'Description': String(d.description || ''), 'Selling Price': price, 'Unit Cost': cost, 'Current Stock': 0, 'Low Stock Threshold': low, 'Status': status, 'Date Created': day_()};
  add_('Products', p);
  const st = Number(d.stock) || 0; if (st > 0) move_(id, 'Beginning', st, 'Opening stock', day_(), '');
  audit_('Create', 'Product', id, null, p); return {id: id};
};

/* ---------- expenses ---------- */
function exp_(d, id) {
  const cat = String(d.category || '').trim(), desc = String(d.description || '').trim();
  if (!cat) throw new Error('Choose an expense category.'); if (!desc) throw new Error('Description is required.');
  return {'Expense ID': id, 'Actual Expense Date': date_(d.date, 'Expense date'), 'Category': cat, 'Description': desc.slice(0, 200), 'Amount': num_(d.amount, 'Amount', 0.01),
    'Receipt Reference': String(d.receipt || '').slice(0, 60), 'Notes': String(d.notes || '').slice(0, 300)};
}
H.saveExpense = d => {
  const id = id_(d.id); if (find_('Expenses', id)) return {duplicate: true, id: id};
  const rec = Object.assign(exp_(d, id), {'Recorded At': now_()});
  add_('Expenses', rec); audit_('Create', 'Expense', id, null, rec); return {id: id};
};
H.updateExpense = d => {
  const id = id_(d.id), old = rows_('Expenses').find(e => e['Expense ID'] === id); if (!old) throw new Error('Expense not found.');
  const rec = Object.assign(exp_(d, id), {'Recorded At': old['Recorded At']});
  set_('Expenses', id, rec); audit_('Update', 'Expense', id, old, rec); return {id: id};
};
H.deleteExpense = d => {
  const id = id_(d.id), old = rows_('Expenses').find(e => e['Expense ID'] === id); if (!old) throw new Error('Expense not found.');
  del_('Expenses', id); audit_('Delete', 'Expense', id, old, null); return {id: id};
};

/* ---------- inventory + settings ---------- */
H.adjustStock = d => {
  const id = id_(d.id); if (find_('Inventory', id)) return {duplicate: true, id: id};
  const type = d.type; if (['Beginning', 'Addition', 'Adjustment'].indexOf(type) < 0) throw new Error('Invalid movement type.');
  const date = date_(d.date, 'Movement date'), qty = Number(d.quantity), reason = String(d.reason || '').trim();
  if (!isFinite(qty) || qty === 0 || !Number.isInteger(qty)) throw new Error('Quantity must be a whole number (not 0).');
  if (type !== 'Adjustment' && qty < 0) throw new Error('Quantity must be positive for this movement type.');
  if (type === 'Adjustment' && !reason) throw new Error('A reason is required for stock adjustments.');
  const p = rows_('Products').find(x => x['Product ID'] === d.productId); if (!p) throw new Error('Choose a product.');
  if ((Number(p['Current Stock']) || 0) + qty < 0) throw new Error('This would make stock negative.');
  move_(d.productId, type, qty, reason || (type === 'Beginning' ? 'Beginning inventory' : 'Stock added'), date, '', id);
  audit_('Create', 'Inventory', id, null, d); return {id: id};
};
H.saveSettings = d => {
  const prev = {};
  rows_('Settings').forEach(r => prev[r.Key] = r.Value);
  SETTING_KEYS.forEach(k => {
    if (d[k] === undefined) return;
    const row = {Key: k, Value: String(d[k]).slice(0, 300)};
    if (find_('Settings', k)) set_('Settings', k, row); else add_('Settings', row);
  });
  audit_('Update', 'Settings', 'settings', prev, d); return {};
};

/* ---------- HTTP entry points ---------- */
const out_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
function doGet(e) {
  try {
    ensure_();
    if (((e && e.parameter) || {}).action === 'ping') return out_({ok: true});
    const st = {}; rows_('Settings').forEach(r => st[r.Key] = r.Value);
    return out_({ok: true, data: {Products: rows_('Products'), Sales: rows_('Sales'), SaleItems: rows_('Sale Items'), Expenses: rows_('Expenses'),
      Inventory: rows_('Inventory'), Settings: st, AuditLogs: rows_('Audit Logs').slice(-200)}});
  } catch (err) { return out_({ok: false, error: String(err.message || err)}); }
}
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000); ensure_();
    const b = JSON.parse(e.postData.contents), h = H[b.action];
    if (!h) throw new Error('Unknown action.');
    return out_(Object.assign({ok: true}, h(b.data || {})));
  } catch (err) { return out_({ok: false, error: String(err.message || err)}); }
  finally { try { lock.releaseLock(); } catch (x) {} }
}
