/**
 * Kinhfolk OWNER backend (Google Apps Script). Separate from the staff stock script.
 * Reads Square (read-only), stores misc charges + fixed costs in this Sheet, and only
 * answers requests that carry a Google sign-in token from an owner email.
 *
 * Script Properties (Project Settings > Script properties):
 *   SQUARE_TOKEN        Square access token (read-only permissions)
 *   SQUARE_LOCATION_ID  built in: LNFXFREH9063S (main) and L2WJ62AQ2QPYE (catering); a property overrides it
 *   OWNER_EMAILS        comma-separated, e.g. you@gmail.com,partner@gmail.com
 *   GOOGLE_CLIENT_ID    the same OAuth client ID you put in owner/index.html
 *   LABOR_BURDEN        built in as 1.10 (adds ~10% for employer payroll taxes); a property overrides it
 *   TZ                  optional, default America/New_York
 */
var SQ = 'https://connect.squareup.com/v2';
var SQ_VERSION = '2025-10-16';

function P_(k, d) { return PropertiesService.getScriptProperties().getProperty(k) || d; }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
/* SQUARE_LOCATION_ID can list several locations separated by commas, e.g. CAFEID,CATERINGID */
function locs_() { return P_('SQUARE_LOCATION_ID', 'LNFXFREH9063S,L2WJ62AQ2QPYE').split(',').map(function (x) { return x.trim(); }).filter(String); }
function tz_() { return P_('TZ', 'America/New_York'); }

/* ---------- sign-in check ---------- */
function checkOwner_(idToken) {
  if (!idToken) throw new Error('Sign in required.');
  var r = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken), { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('Sign-in expired. Reload and sign in again.');
  var t = JSON.parse(r.getContentText());
  if (t.aud !== P_('GOOGLE_CLIENT_ID')) throw new Error('Wrong app.');
  if (String(t.email_verified) !== 'true') throw new Error('Email not verified.');
  var ok = P_('OWNER_EMAILS', '').toLowerCase().split(',').map(function (s) { return s.trim(); });
  if (ok.indexOf(String(t.email).toLowerCase()) < 0) throw new Error('This account is not an owner.');
  return t.email;
}

/* ---------- web app entry points ---------- */
function doGet(e) {
  try {
    var p = e.parameter || {};
    checkOwner_(p.idToken);
    if (p.action === 'summary') return json_(summary_(p.month));
    return json_({ ok: false, error: 'Unknown action.' });
  } catch (err) { return json_({ ok: false, error: String(err.message || err) }); }
}
function doPost(e) {
  try {
    var b = JSON.parse(e.postData.contents);
    checkOwner_(b.idToken);
    if (b.action === 'addMisc') { addMisc_(b.month, b.item); return json_({ ok: true, misc: misc_(b.month) }); }
    if (b.action === 'deleteMisc') { deleteMisc_(b.month, b.id); return json_({ ok: true, misc: misc_(b.month) }); }
    return json_({ ok: false, error: 'Unknown action.' });
  } catch (err) { return json_({ ok: false, error: String(err.message || err) }); }
}

/* ---------- Square helpers ---------- */
function sq_(method, path, body) {
  var r = UrlFetchApp.fetch(SQ + path, {
    method: method, muteHttpExceptions: true, contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + P_('SQUARE_TOKEN'), 'Square-Version': SQ_VERSION },
    payload: body ? JSON.stringify(body) : undefined
  });
  var j = JSON.parse(r.getContentText() || '{}');
  if (r.getResponseCode() >= 300) throw new Error('Square: ' + ((j.errors && j.errors[0] && j.errors[0].detail) || r.getResponseCode()));
  return j;
}
function cents_(m) { return m && typeof m.amount === 'number' ? m.amount : 0; }

function monthRange_(ym) {                 // [start, end) as ISO strings in the shop time zone
  var y = +ym.slice(0, 4), m = +ym.slice(5, 7);
  var s = new Date(y, m - 1, 1), e = new Date(y, m, 1);
  return [Utilities.formatDate(s, tz_(), "yyyy-MM-dd'T'00:00:00XXX"), Utilities.formatDate(e, tz_(), "yyyy-MM-dd'T'00:00:00XXX")];
}
/* Counts every order that is paid in full, even when Square leaves it OPEN (unfinished fulfillment),
   by the day it was created. Unpaid open tabs and canceled orders are skipped. */
function searchOrders_(startIso, endIso) {
  var out = [], cursor;
  do {
    var j = sq_('POST', '/orders/search', {
      location_ids: locs_(), limit: 500, cursor: cursor,
      query: { filter: { state_filter: { states: ['OPEN', 'COMPLETED'] }, date_time_filter: { created_at: { start_at: startIso, end_at: endIso } } },
               sort: { sort_field: 'CREATED_AT', sort_order: 'ASC' } }
    });
    out = out.concat(j.orders || []); cursor = j.cursor;
  } while (cursor);
  return out.filter(isPaid_);
}
function isPaid_(o) {
  if (o.state === 'COMPLETED') return true;
  var paid = (o.tenders || []).length > 0;
  var due = o.net_amount_due_money ? o.net_amount_due_money.amount : 0;
  return paid && due === 0;
}

/* Net sales = order total minus tax, tips and service charges (check it against Square's Sales Summary). */
function orderNet_(o) { return (cents_(o.total_money) - cents_(o.total_tax_money) - cents_(o.total_tip_money) - cents_(o.total_service_charge_money)) / 100; }

function weekIndex_(iso, monthStart) {      // Week 1 = days 1-7, Week 2 = 8-14 ...
  var d = new Date(iso); var day = +Utilities.formatDate(d, tz_(), 'd'); return Math.floor((day - 1) / 7);
}

function summary_(ym) {
  if (!/^\d{4}-\d{2}$/.test(ym || '')) throw new Error('Bad month.');
  var r = monthRange_(ym), orders = searchOrders_(r[0], r[1]);
  var weeks = [0, 0, 0, 0, 0], custIds = {};
  orders.forEach(function (o) {
    weeks[Math.min(4, weekIndex_(o.created_at))] += orderNet_(o);
    if (o.customer_id) custIds[o.customer_id] = 1;
  });
  var nWeeks = Math.max(1, Math.ceil(daysSoFar_(ym) / 7));
  var sales = weeks.slice(0, nWeeks).map(function (v) { return Math.round(v * 100) / 100; });

  /* New vs returning: a customer created inside this month is "new", otherwise "returning".
     Only orders that have a customer attached can be counted (loyalty, card on file, receipt email). */
  var ids = Object.keys(custIds), nw = 0, rt = 0;
  for (var i = 0; i < ids.length; i += 100) {
    var j = sq_('POST', '/customers/bulk-retrieve', { customer_ids: ids.slice(i, i + 100) });
    Object.keys(j.responses || {}).forEach(function (id) {
      var c = j.responses[id].customer; if (!c) return;
      if (c.created_at >= r[0] && c.created_at < r[1]) nw++; else rt++;
    });
  }

  /* Items ordered this week (Monday to now) */
  var ws = weekStart_(), items = {};
  searchOrders_(ws, Utilities.formatDate(new Date(), tz_(), "yyyy-MM-dd'T'HH:mm:ssXXX")).forEach(function (o) {
    (o.line_items || []).forEach(function (li) { items[li.name] = (items[li.name] || 0) + (+li.quantity || 0); });
  });
  var top = Object.keys(items).map(function (k) { return [k, items[k]]; }).sort(function (a, b) { return b[1] - a[1]; }).slice(0, 12);

  return { ok: true, sales: sales, labor: labor_(r[0], r[1]), nw: nw, rt: rt, items: top,
           misc: misc_(ym), fixed: fixed_(), f: fraction_(ym) };
}
function daysSoFar_(ym) {
  var now = new Date(), cur = Utilities.formatDate(now, tz_(), 'yyyy-MM');
  if (ym < cur) return 31; if (ym > cur) return 1;
  return +Utilities.formatDate(now, tz_(), 'd');
}
function fraction_(ym) {                    // share of the month passed, to prorate fixed costs
  var y = +ym.slice(0, 4), m = +ym.slice(5, 7), dim = new Date(y, m, 0).getDate();
  var cur = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM');
  return ym < cur ? 1 : ym > cur ? 0 : Math.min(1, daysSoFar_(ym) / dim);
}
function weekStart_() {
  var now = new Date(), dow = +Utilities.formatDate(now, tz_(), 'u');           // Monday = 1
  var s = new Date(now.getTime() - (dow - 1) * 86400000);
  return Utilities.formatDate(s, tz_(), "yyyy-MM-dd'T'00:00:00XXX");
}

/* ---------- labor estimate from Square timecards x the hourly rate set on each timecard ---------- */
function labor_(startIso, endIso) {
  var total = 0, cursor;
  do {
    var j = sq_('POST', '/labor/timecards/search', {
      limit: 200, cursor: cursor,
      query: { filter: { location_ids: locs_(), start: { start_at: startIso, end_at: endIso } } }
    });
    (j.timecards || []).forEach(function (t) {
      if (!t.start_at || !t.end_at || !t.wage || !t.wage.hourly_rate) return;   // still clocked in, or no rate set
      var hrs = (new Date(t.end_at) - new Date(t.start_at)) / 3600000;
      (t.breaks || []).forEach(function (b) { if (!b.is_paid && b.start_at && b.end_at) hrs -= (new Date(b.end_at) - new Date(b.start_at)) / 3600000; });
      total += Math.max(0, hrs) * cents_(t.wage.hourly_rate) / 100;
    });
    cursor = j.cursor;
  } while (cursor);
  return Math.round(total * parseFloat(P_('LABOR_BURDEN', '1.10')));
}

/* ---------- Sheet tabs: Misc (month, id, date, label, amount) and Fixed (name, monthly amount) ---------- */
function tab_(name, header) {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(header); sh.setFrozenRows(1); }
  return sh;
}
function misc_(ym) {
  var v = tab_('Misc', ['month', 'id', 'date', 'label', 'amount']).getDataRange().getValues().slice(1);
  return v.filter(function (r) { return r[0] === ym; }).map(function (r) {
    return { id: String(r[1]), date: r[2] instanceof Date ? Utilities.formatDate(r[2], tz_(), 'yyyy-MM-dd') : String(r[2]), label: String(r[3]), amount: +r[4] };
  });
}
function addMisc_(ym, it) {
  if (!it || !it.label || !(+it.amount > 0)) throw new Error('Enter a name and an amount.');
  tab_('Misc', ['month', 'id', 'date', 'label', 'amount']).appendRow([ym, String(it.id || Date.now()), String(it.date || ''), String(it.label).slice(0, 60), +it.amount]);
}
function deleteMisc_(ym, id) {
  var sh = tab_('Misc', ['month', 'id', 'date', 'label', 'amount']), v = sh.getDataRange().getValues();
  for (var i = v.length - 1; i >= 1; i--) if (v[i][0] === ym && String(v[i][1]) === String(id)) { sh.deleteRow(i + 1); return; }
}
function fixed_() {
  var sh = tab_('Fixed', ['name', 'monthly amount']);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().filter(function (r) { return r[0] && +r[1]; }).map(function (r) { return [String(r[0]), +r[1]]; });
}

/* Run once from the editor to approve permissions and test the Square connection. */
function testSquare() {
  var ym = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM');
  var r = monthRange_(ym), o = searchOrders_(r[0], r[1]);
  Logger.log('Orders this month: ' + o.length + ', labor est: ' + labor_(r[0], r[1]));
}
