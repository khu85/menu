/**
 * Stock Count + weekly COGS backend (Google Apps Script, bound to a Google Sheet).
 *
 * Tabs (all created for you by setup()):
 *   Items      Section | Item | Unit | Unit Cost | Vendor | Note     <- your inventory (fill in Vendor)
 *   Weeks      one row per week: open or closed, with its total COGS
 *   Entries    per week and item: Beginning, Ending, and (once closed) Unit Cost and Used
 *   Additions  every delivery you confirm: time, item, quantity
 *
 * GET            -> { ok, items, weeks, current }          (current = the open week, or null)
 * GET ?week=ID   -> same, plus { week } for that week (used to look at past weeks)
 * POST (JSON text body), one of:
 *   { action: "setEnding",   week, key, count }            count null clears it
 *   { action: "setBeginning", week, key, count }
 *   { action: "addBatch",    week, batch, entries: [{ key, qty, note }] }
 *   { action: "removeAddition", week, id }
 *   { action: "closeWeek",   week, basis }                 locks the week, opens the next one, sends the order list to Discord
 *   { action: "startWeek" }                                only when no week is open
 *   { action: "sendOrder",   week, basis }                 sends that week's order list to Discord now
 *   { action: "addAlert",    week, id, key, by, count, note }  staff low stock alert: saved, and sent to Discord right away
 *   { action: "resolveAlert", week, id }                   marks an alert as dealt with
 * Adding that item to the week (Added tab) clears its low stock alert on its own, even when it was
 * reported in an earlier week. Alerts stay active from week to week until the item is added or marked OK.
 *
 * Discord: the webhook address is set in DISCORD_WEBHOOK_URL below.
 * Months: a week belongs to the month its Thursday falls in (so Oct 29 - Nov 1 is an October week).
 */

/** Order list basis when the page does not say: 'week' (last week's usage) or 'month' (average of the last 4 weeks). */
var ORDER_BASIS = 'week';
var AVG_WEEKS = 4;

/** Discord webhook for the order list. (A DISCORD_WEBHOOK_URL Script property, if set, is used instead.) */
var DISCORD_WEBHOOK_URL = 'https://discord.com/api/webhooks/1567366015380095116/GRXrQ2VZ-eEUtuq3qaVWXSABaxvpue2VyB7i15Uj3QcNeX3X4EX0c50Q5l6BuUSDdFId';

var NAMES = { items: 'Items', weeks: 'Weeks', entries: 'Entries', adds: 'Additions', alerts: 'Alerts' };
var HEADERS = {
  items: ['Section', 'Item', 'Unit', 'Unit Cost', 'Vendor', 'Note'],
  weeks: ['Week', 'Label', 'Status', 'Started', 'Closed', 'Total COGS', 'Month'],
  entries: ['Week', 'Key', 'Beginning', 'Ending', 'Unit Cost', 'Used'],
  adds: ['Id', 'Week', 'Time', 'Key', 'Qty', 'Note'],
  alerts: ['Id', 'Week', 'Time', 'Key', 'By', 'Note', 'Status', 'Cleared', 'Reason', 'Count']
};
var TIME_FORMAT = 'yyyy-mm-dd hh:mm:ss';
var MAX_TEXT = 120;
var MAX_QTY = 99999;

/** Run once from the editor: creates the tabs and grants permission. */
function setup() {
  Object.keys(NAMES).forEach(function (k) {
    getSheet_(k);
  });
}

function doGet(e) {
  try {
    var weeks = readWeeks_();
    var open = null;
    weeks.forEach(function (w) {
      if (w.status === 'open') {
        open = w;
      }
    });
    var out = {
      ok: true,
      items: readItems_(),
      weeks: weeks,
      current: open ? readWeek_(open.id) : null,
      usage: readUsage_(weeks)
    };
    var want = e && e.parameter && e.parameter.week;
    if (want) {
      out.week = readWeek_(String(want));
    }
    return json_(out);
  } catch (err) {
    return json_({ ok: false, error: message_(err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var body = parseBody_(e);
    lock.waitLock(20000);
    switch (body.action) {
      case 'setEnding':
        return json_(setEntry_(body, 'ending'));
      case 'setBeginning':
        return json_(setEntry_(body, 'beginning'));
      case 'addBatch':
        return json_(addBatch_(body));
      case 'removeAddition':
        return json_(removeAddition_(body));
      case 'closeWeek':
        return json_(closeWeek_(body));
      case 'startWeek':
        return json_(startWeek_());
      case 'addAlert':
        return json_(addAlert_(body));
      case 'resolveAlert':
        return json_(resolveAlert_(body));
      case 'editWeek':
        return json_(editWeek_(body));
      case 'sendOrder':
        return json_(sendOrder_(body));
      default:
        return json_({ ok: false, error: 'Unknown action.' });
    }
  } catch (err) {
    return json_({ ok: false, error: message_(err) });
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

/* ---------- actions ---------- */

function setEntry_(body, field) {
  var week = requireOpenWeek_(body.week);
  var key = knownKey_(body.key);
  var col = field === 'ending' ? 4 : 3;
  var sheet = getSheet_('entries');
  var row = findEntryRow_(sheet, week.id, key);
  var blank = body.count === null || body.count === undefined || String(body.count).trim() === '';
  var value = '';
  if (!blank) {
    value = Number(body.count);
    if (!isFinite(value) || value < 0 || value > MAX_QTY) {
      throw new Error('count must be a number from 0 to ' + MAX_QTY + '.');
    }
    value = Math.round(value * 1000) / 1000;
  }
  if (!row) {
    row = sheet.getLastRow() + 1;
    sheet.getRange(row, 1, 1, 2).setNumberFormat('@');
    sheet.getRange(row, 1, 1, 2).setValues([[week.id, key]]);
  }
  sheet.getRange(row, col).setValue(value);
  return { ok: true, key: key, count: blank ? null : value };
}

function addBatch_(body) {
  var week = requireOpenWeek_(body.week);
  var batch = String(body.batch || '').replace(/[^A-Za-z0-9_-]/g, '').slice(0, 40);
  if (!batch) {
    throw new Error('batch id is required.');
  }
  var entries = Array.isArray(body.entries) ? body.entries.slice(0, 200) : [];
  if (!entries.length) {
    throw new Error('There is nothing to add.');
  }
  var sheet = getSheet_('adds');
  var seen = {};
  var last = sheet.getLastRow();
  if (last > 1) {
    sheet.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) {
      seen[String(r[0])] = true;
    });
  }
  var now = new Date();
  var rows = [];
  entries.forEach(function (en, i) {
    var id = batch + '-' + i;
    if (seen[id]) {
      return; // already saved (a retry), skip
    }
    var key = knownKey_(en.key);
    var qty = Number(en.qty);
    if (!isFinite(qty) || qty <= 0 || qty > MAX_QTY) {
      throw new Error('Each quantity must be more than 0.');
    }
    rows.push([id, week.id, now, key, Math.round(qty * 1000) / 1000, text_(en.note)]);
  });
  if (rows.length) {
    var start = sheet.getLastRow() + 1;
    sheet.getRange(start, 1, rows.length, 2).setNumberFormat('@');
    sheet.getRange(start, 3, rows.length, 1).setNumberFormat(TIME_FORMAT);
    sheet.getRange(start, 1, rows.length, 6).setValues(rows);
    clearAlerts_(week.id, rows.map(function (r) {
      return r[3];
    }), 'restocked');
  }
  return { ok: true, saved: rows.length };
}

function removeAddition_(body) {
  var week = requireOpenWeek_(body.week);
  var id = String(body.id || '');
  var sheet = getSheet_('adds');
  var last = sheet.getLastRow();
  if (last > 1) {
    var values = sheet.getRange(2, 1, last - 1, 2).getValues();
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][0]) === id && weekIdOf_(values[i][1]) === week.id) {
        sheet.getRange(i + 2, 1, 1, HEADERS.adds.length).clearContent();
        return { ok: true, removed: id };
      }
    }
  }
  return { ok: true, removed: null };
}

function closeWeek_(body) {
  var week = requireOpenWeek_(body.week);
  var items = readItems_();
  var detail = readWeek_(week.id);
  var missing = items.filter(function (it) {
    var v = detail.end[keyOf_(it)];
    return v === undefined || v === null;
  }).length;
  if (missing) {
    throw new Error(missing + (missing === 1 ? ' item has' : ' items have') + ' no ending count yet.');
  }

  var added = {};
  detail.additions.forEach(function (a) {
    added[a.key] = (added[a.key] || 0) + a.qty;
  });

  // Write the cost and used snapshot into Entries in one pass.
  var sheet = getSheet_('entries');
  var last = sheet.getLastRow();
  var all = last > 1 ? sheet.getRange(2, 1, last - 1, HEADERS.entries.length).getValues() : [];
  var index = {};
  all.forEach(function (r, i) {
    if (weekIdOf_(r[0]) === week.id) {
      index[String(r[1])] = i;
    }
  });
  var total = 0;
  var fresh = [];
  items.forEach(function (it) {
    var key = keyOf_(it);
    var begin = Number(detail.begin[key]) || 0;
    var end = Number(detail.end[key]);
    var used = Math.round((begin + (added[key] || 0) - end) * 1000) / 1000;
    var cost = it.cost === null ? 0 : it.cost;
    total += used * cost;
    if (index[key] !== undefined) {
      all[index[key]][4] = cost;
      all[index[key]][5] = used;
    } else {
      fresh.push([week.id, key, begin, end, cost, used]);
    }
  });
  if (all.length) {
    sheet.getRange(2, 1, all.length, HEADERS.entries.length).setValues(all);
  }
  if (fresh.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, fresh.length, 6).setValues(fresh);
  }
  total = Math.round(total * 100) / 100;

  var now = new Date();
  var wsheet = getSheet_('weeks');
  var wrow = findWeekRow_(wsheet, week.id);
  wsheet.getRange(wrow, 3).setValue('closed');
  wsheet.getRange(wrow, 5).setValue(now);
  wsheet.getRange(wrow, 6).setValue(total);
  var thu = thursdayOnOrBefore_(now);
  wsheet.getRange(wrow, 2).setValue(Utilities.formatDate(thu, 'UTC', 'MMM d') + ' - ' + Utilities.formatDate(utcDay_(ymd_(now)), 'UTC', 'MMM d'));
  wsheet.getRange(wrow, 7).setNumberFormat('@').setValue(iso10_(thu).slice(0, 7));

  var next = newWeek_(detail.end);
  var discord = sendOrderSafe_(week.id, body.basis);
  return { ok: true, total: total, next: next, discord: discord };
}

function startWeek_() {
  var open = readWeeks_().filter(function (w) {
    return w.status === 'open';
  });
  if (open.length) {
    throw new Error('A week is already open.');
  }
  // Begin with the ending counts of the most recent closed week, if there is one.
  var closed = readWeeks_().filter(function (w) {
    return w.status === 'closed';
  });
  var carry = {};
  if (closed.length) {
    carry = readWeek_(closed[closed.length - 1].id).end;
  }
  return { ok: true, next: newWeek_(carry) };
}

/** Creates the next open week; beginning counts come from `carry` (key -> number). */
function newWeek_(carry) {
  var now = new Date();
  var tz = Session.getScriptTimeZone();
  var base = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  var taken = {};
  readWeeks_().forEach(function (w) {
    taken[w.id] = true;
  });
  var id = base;
  var n = 2;
  while (taken[id]) {
    id = base + '-' + n++;
  }
  var label = 'Week of ' + Utilities.formatDate(now, tz, 'MMM d');
  var wsheet = getSheet_('weeks');
  var wrow = wsheet.getLastRow() + 1;
  wsheet.getRange(wrow, 1, 1, 2).setNumberFormat('@');
  wsheet.getRange(wrow, 4, 1, 2).setNumberFormat(TIME_FORMAT);
  wsheet.getRange(wrow, 7).setNumberFormat('@');
  wsheet.getRange(wrow, 1, 1, 7).setValues([[id, label, 'open', now, '', '', '']]);

  var rows = readItems_().map(function (it) {
    var key = keyOf_(it);
    var v = carry && typeof carry[key] === 'number' ? carry[key] : 0;
    return [id, key, v, '', '', ''];
  });
  if (rows.length) {
    var sheet = getSheet_('entries');
    var start = sheet.getLastRow() + 1;
    sheet.getRange(start, 1, rows.length, 2).setNumberFormat('@');
    sheet.getRange(start, 1, rows.length, 6).setValues(rows);
  }
  return id;
}

/**
 * Fix a mistake in any week, open or closed.
 *   { action: 'editWeek', week, key, field: 'begin' | 'end' | 'added', value }
 * 'added' sets the week's total received for that item (the difference is saved as an "Edited" line).
 * Used, cost and COGS are recalculated, and an ending count carries into the next week's beginning.
 */
function editWeek_(body) {
  var weeks = readWeeks_();
  var id = String(body.week || '');
  var idx = -1;
  weeks.forEach(function (w, i) {
    if (w.id === id) {
      idx = i;
    }
  });
  if (idx < 0) {
    throw new Error('Week not found.');
  }
  var week = weeks[idx];
  var key = knownKey_(body.key);
  var field = String(body.field || '');
  var blank = body.value === null || body.value === undefined || String(body.value).trim() === '';
  var value = blank ? null : Number(body.value);
  if (!blank && (!isFinite(value) || value < 0 || value > MAX_QTY)) {
    throw new Error('The number must be from 0 to ' + MAX_QTY + '.');
  }
  if (value !== null) {
    value = Math.round(value * 1000) / 1000;
  }
  if (field === 'begin' || field === 'end') {
    if (value === null && (field === 'begin' || week.status === 'closed')) {
      throw new Error('Enter a number.');
    }
    var sheet = getSheet_('entries');
    var row = findEntryRow_(sheet, id, key);
    if (!row) {
      row = sheet.getLastRow() + 1;
      sheet.getRange(row, 1, 1, 2).setNumberFormat('@');
      sheet.getRange(row, 1, 1, 2).setValues([[id, key]]);
    }
    sheet.getRange(row, field === 'end' ? 4 : 3).setValue(value === null ? '' : value);
    if (field === 'end' && value !== null && idx + 1 < weeks.length) {
      var nrow = findEntryRow_(sheet, weeks[idx + 1].id, key);
      if (nrow) {
        sheet.getRange(nrow, 3).setValue(value);
      }
      if (weeks[idx + 1].status === 'closed') {
        recalcWeek_(weeks[idx + 1].id);
      }
    }
  } else if (field === 'added') {
    if (value === null) {
      throw new Error('Enter a number.');
    }
    var detail = readWeek_(id);
    var have = 0;
    detail.additions.forEach(function (a) {
      if (a.key === key) {
        have += a.qty;
      }
    });
    var delta = Math.round((value - have) * 1000) / 1000;
    if (delta !== 0) {
      var adds = getSheet_('adds');
      var start = adds.getLastRow() + 1;
      adds.getRange(start, 1, 1, 2).setNumberFormat('@');
      adds.getRange(start, 3, 1, 1).setNumberFormat(TIME_FORMAT);
      adds.getRange(start, 1, 1, 6).setValues([['edit-' + new Date().getTime().toString(36) + Math.floor(Math.random() * 1000), id, new Date(), key, delta, 'Edited']]);
      if (delta > 0) {
        clearAlerts_(id, [key], 'restocked');
      }
    }
  } else {
    throw new Error('Unknown field.');
  }
  if (week.status === 'closed') {
    recalcWeek_(id);
  }
  return { ok: true };
}

/** Recomputes used and the COGS total of a closed week (unit costs stay as they were when it closed). */
function recalcWeek_(id) {
  var items = readItems_();
  var detail = readWeek_(id);
  var added = {};
  detail.additions.forEach(function (a) {
    added[a.key] = (added[a.key] || 0) + a.qty;
  });
  var sheet = getSheet_('entries');
  var last = sheet.getLastRow();
  var all = last > 1 ? sheet.getRange(2, 1, last - 1, HEADERS.entries.length).getValues() : [];
  var total = 0;
  var costOf = {};
  items.forEach(function (it) {
    costOf[keyOf_(it)] = it.cost === null ? 0 : it.cost;
  });
  all.forEach(function (r) {
    if (weekIdOf_(r[0]) !== id) {
      return;
    }
    var key = String(r[1]);
    if (r[3] === '' || r[3] === null) {
      return;
    }
    var cost = r[4] === '' || r[4] === null ? (costOf[key] || 0) : Number(r[4]);
    var used = Math.round(((Number(r[2]) || 0) + (added[key] || 0) - Number(r[3])) * 1000) / 1000;
    r[4] = cost;
    r[5] = used;
    total += used * cost;
  });
  if (all.length) {
    sheet.getRange(2, 1, all.length, HEADERS.entries.length).setValues(all);
  }
  var wsheet = getSheet_('weeks');
  wsheet.getRange(findWeekRow_(wsheet, id), 6).setValue(Math.round(total * 100) / 100);
}

/* ---------- reading ---------- */

function readItems_() {
  var sheet = getSheet_('items');
  var last = sheet.getLastRow();
  if (last < 2) {
    return [];
  }
  var values = sheet.getRange(2, 1, last - 1, HEADERS.items.length).getValues();
  var items = [];
  var seen = {};
  values.forEach(function (r) {
    var item = text_(r[1]);
    if (!item) {
      return;
    }
    var cost = r[3] === '' || r[3] === null ? null : Number(r[3]);
    var it = {
      section: text_(r[0]) || 'Other',
      item: item,
      unit: text_(r[2]),
      cost: cost !== null && isFinite(cost) && cost >= 0 ? cost : null,
      vendor: text_(r[4]),
      note: text_(r[5])
    };
    var k = keyOf_(it);
    if (!seen[k]) {
      seen[k] = true;
      items.push(it);
    }
  });
  return items;
}

function readWeeks_() {
  var sheet = getSheet_('weeks');
  var last = sheet.getLastRow();
  if (last < 2) {
    return [];
  }
  var out = [];
  sheet.getRange(2, 1, last - 1, HEADERS.weeks.length).getValues().forEach(function (r) {
    if (String(r[0]) === '') {
      return;
    }
    var month = r[6] instanceof Date ? Utilities.formatDate(r[6], Session.getScriptTimeZone(), 'yyyy-MM') : String(r[6] || '').trim();
    if (!/^\d{4}-\d{2}$/.test(month)) {
      month = monthOfWeek_(String(r[2]), r[3], r[4]);
    }
    out.push({
      id: weekIdOf_(r[0]),
      label: String(r[1]),
      status: String(r[2]),
      started: iso_(r[3]),
      closed: iso_(r[4]),
      total: r[5] === '' || r[5] === null ? null : Number(r[5]),
      month: month
    });
  });
  return out;
}

function readWeek_(id) {
  var meta = null;
  readWeeks_().forEach(function (w) {
    if (w.id === id) {
      meta = w;
    }
  });
  if (!meta) {
    return null;
  }
  var out = {
    id: meta.id, label: meta.label, status: meta.status, started: meta.started, closed: meta.closed,
    total: meta.total, begin: {}, end: {}, cost: {}, used: {}, additions: [], alerts: readAlerts_(meta.id, meta.status === 'open')
  };
  var es = getSheet_('entries');
  var el = es.getLastRow();
  if (el > 1) {
    es.getRange(2, 1, el - 1, HEADERS.entries.length).getValues().forEach(function (r) {
      if (weekIdOf_(r[0]) !== id) {
        return;
      }
      var key = String(r[1]);
      if (r[2] !== '') {
        out.begin[key] = Number(r[2]);
      }
      if (r[3] !== '') {
        out.end[key] = Number(r[3]);
      }
      if (r[4] !== '') {
        out.cost[key] = Number(r[4]);
      }
      if (r[5] !== '') {
        out.used[key] = Number(r[5]);
      }
    });
  }
  var as = getSheet_('adds');
  var al = as.getLastRow();
  if (al > 1) {
    as.getRange(2, 1, al - 1, HEADERS.adds.length).getValues().forEach(function (r) {
      if (String(r[0]) === '' || weekIdOf_(r[1]) !== id) {
        return;
      }
      out.additions.push({ id: String(r[0]), time: iso_(r[2]), key: String(r[3]), qty: Number(r[4]), note: String(r[5] || '') });
    });
  }
  return out;
}


/* ---------- months: a week belongs to the month its Thursday falls in ---------- */

function ymd_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
function utcDay_(s) {
  var p = s.split('-');
  return new Date(Date.UTC(+p[0], +p[1] - 1, +p[2]));
}
function iso10_(d) {
  return d.toISOString().slice(0, 10);
}
function thursdayOnOrBefore_(date) {
  var d = utcDay_(ymd_(date));
  while (d.getUTCDay() !== 4) {
    d = new Date(d.getTime() - 86400000);
  }
  return d;
}
function thursdayOnOrAfter_(date) {
  var d = utcDay_(ymd_(date));
  while (d.getUTCDay() !== 4) {
    d = new Date(d.getTime() + 86400000);
  }
  return d;
}
function asDate_(v) {
  return v instanceof Date ? v : new Date(v);
}
/** Closed weeks use the Thursday on or before the close date; the open week uses the next Thursday. */
function monthOfWeek_(status, started, closed) {
  var d;
  if (status === 'closed' && closed) {
    d = thursdayOnOrBefore_(asDate_(closed));
  } else {
    d = thursdayOnOrAfter_(started ? asDate_(started) : new Date());
  }
  return iso10_(d).slice(0, 7);
}

/* ---------- usage history and order suggestions ---------- */

/** { weekId: { key: [used, unitCost] } } for the most recent closed weeks. */
function readUsage_(weeks) {
  var ids = {};
  weeks.filter(function (w) {
    return w.status === 'closed';
  }).slice(-60).forEach(function (w) {
    ids[w.id] = true;
  });
  var out = {};
  var sheet = getSheet_('entries');
  var last = sheet.getLastRow();
  if (last < 2) {
    return out;
  }
  sheet.getRange(2, 1, last - 1, HEADERS.entries.length).getValues().forEach(function (r) {
    var id = weekIdOf_(r[0]);
    if (!ids[id] || r[5] === '') {
      return;
    }
    if (!out[id]) {
      out[id] = {};
    }
    out[id][String(r[1])] = [Number(r[5]), r[4] === '' ? 0 : Number(r[4])];
  });
  return out;
}

/**
 * What to order for a week. For each counted item: forecast = last week's usage ('week')
 * or the average of the last AVG_WEEKS weeks ('month'); suggested = forecast - ending, rounded up.
 */
function orderPlan_(weekId, basis) {
  basis = basis === 'month' ? 'month' : (basis === 'week' ? 'week' : ORDER_BASIS);
  var items = readItems_();
  var weeks = readWeeks_();
  var target = readWeek_(weekId);
  if (!target) {
    throw new Error('Week not found.');
  }
  var usage = readUsage_(weeks);
  var hist = [];
  for (var i = 0; i < weeks.length; i++) {
    if (weeks[i].status === 'closed' && weeks[i].id !== weekId) {
      hist.push(weeks[i].id);
    } else if (weeks[i].id === weekId) {
      break;
    }
  }
  var vendors = {};
  var order = [];
  var total = 0;
  items.forEach(function (it) {
    var key = keyOf_(it);
    var end = target.end[key];
    if (typeof end !== 'number') {
      return;
    }
    var used;
    if (target.status === 'closed' && typeof target.used[key] === 'number') {
      used = target.used[key];
    } else {
      var added = 0;
      target.additions.forEach(function (a) {
        if (a.key === key) {
          added += a.qty;
        }
      });
      used = (target.begin[key] || 0) + added - end;
    }
    var samples = [Math.max(0, used)];
    if (basis === 'month') {
      for (var j = hist.length - 1; j >= 0 && samples.length < AVG_WEEKS; j--) {
        var u = usage[hist[j]] && usage[hist[j]][key];
        if (u) {
          samples.push(Math.max(0, u[0]));
        }
      }
    }
    var forecast = samples.reduce(function (a, b) {
      return a + b;
    }, 0) / samples.length;
    var need = Math.ceil(forecast - end - 1e-9);
    if (need > 0) {
      order.push({ key: key, item: it, used: used, end: end, forecast: forecast, need: need });
    }
  });
  return { week: target, basis: basis, rows: order };
}

/** "3 x 500ml", "2 x lb": the amount, an x, then the unit (a size such as "500 ml" is closed up). */
function qtyText_(n, unit) {
  var u = String(unit || '').replace(/^(\d+(?:\.\d+)?)\s+([A-Za-z])/, '$1$2');
  return u ? n + ' x ' + u : String(n);
}

function basisText_(basis) {
  return basis === 'month' ? 'average weekly usage over the last ' + AVG_WEEKS + ' weeks' : "last week's usage";
}

function orderMessages_(plan) {
  var head = '**ORDER LIST** \u00B7 ' + plan.week.label + '\n_Based on ' + basisText_(plan.basis) + ' minus what is left_';
  if (!plan.rows.length) {
    return [head + '\n\u2713 Nothing to order. Every item has enough left.'];
  }
  var by = {};
  var names = [];
  plan.rows.forEach(function (r) {
    var v = r.item.vendor || 'No vendor set';
    if (!by[v]) {
      by[v] = [];
      names.push(v);
    }
    by[v].push(r);
  });
  names.sort(function (a, b) {
    if (a === 'No vendor set') {
      return 1;
    }
    if (b === 'No vendor set') {
      return -1;
    }
    return a < b ? -1 : (a > b ? 1 : 0);
  });
  var blocks = [head];
  names.forEach(function (v) {
    var cost = 0;
    var lines = [];
    var section = '';
    by[v].forEach(function (r) {
      cost += r.need * (r.item.cost || 0);
      if (r.item.section !== section) {
        section = r.item.section;
        lines.push('__' + section + '__');
      }
      lines.push('\u2022 ' + r.item.item + ': **' + qtyText_(r.need, r.item.unit) + '**');
    });
    blocks.push('**' + v + '**' + (cost > 0 ? ' (about $' + cost.toFixed(2) + ')' : '') + '\n' + lines.join('\n'));
  });
  var msgs = [];
  var cur = '';
  blocks.forEach(function (b) {
    if (cur && (cur.length + b.length + 2) > 1900) {
      msgs.push(cur);
      cur = '';
    }
    cur += (cur ? '\n\n' : '') + b.slice(0, 1900);
  });
  if (cur) {
    msgs.push(cur);
  }
  return msgs;
}


/* ---------- low stock alerts ---------- */

function alertsSheet_() {
  var sheet = getSheet_('alerts');
  if (sheet.getMaxColumns() < HEADERS.alerts.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), HEADERS.alerts.length - sheet.getMaxColumns());
  }
  if (String(sheet.getRange(1, HEADERS.alerts.length).getValue()) === '') {
    sheet.getRange(1, 1, 1, HEADERS.alerts.length).setValues([HEADERS.alerts]).setFontWeight('bold');
  }
  return sheet;
}

/** Alerts for a week. An open week also lists every alert from earlier weeks that is still active. */
function readAlerts_(weekId, carry) {
  var sheet = alertsSheet_();
  var last = sheet.getLastRow();
  var out = [];
  if (last < 2) {
    return out;
  }
  sheet.getRange(2, 1, last - 1, HEADERS.alerts.length).getValues().forEach(function (r) {
    if (String(r[0]) === '') {
      return;
    }
    var own = weekIdOf_(r[1]) === weekId;
    if (!own && !(carry && String(r[6] || 'active') === 'active')) {
      return;
    }
    out.push({
      id: String(r[0]),
      week: weekIdOf_(r[1]),
      count: r[9] === '' || r[9] === undefined ? null : Number(r[9]),
      time: iso_(r[2]),
      key: String(r[3]),
      by: String(r[4] || ''),
      note: String(r[5] || ''),
      status: String(r[6] || 'active'),
      cleared: iso_(r[7]),
      reason: String(r[8] || '')
    });
  });
  return out;
}

/** Marks active alerts for the given item keys (or all of them when keys is null) as cleared. */
function clearAlerts_(weekId, keys, reason) {
  var sheet = alertsSheet_();
  var last = sheet.getLastRow();
  if (last < 2) {
    return 0;
  }
  var range = sheet.getRange(2, 1, last - 1, HEADERS.alerts.length);
  var vals = range.getValues();
  var want = null;
  if (keys) {
    want = {};
    keys.forEach(function (k) {
      want[k] = true;
    });
  }
  var now = new Date();
  var n = 0;
  vals.forEach(function (r) {
    if (String(r[0]) !== '' && String(r[6]) === 'active' && (!want || want[String(r[3])])) {
      r[6] = 'cleared';
      r[7] = now;
      r[8] = reason;
      n++;
    }
  });
  if (n) {
    range.setValues(vals);
  }
  return n;
}

function addAlert_(body) {
  var week = requireOpenWeek_(body.week);
  var key = knownKey_(body.key);
  var id = String(body.id || '').replace(/[^A-Za-z0-9_-]/g, '').slice(0, 40) || ('a' + new Date().getTime().toString(36));
  var existing = readAlerts_(week.id, true);
  var dup = null;
  existing.forEach(function (a) {
    if (a.id === id) {
      dup = a;
    } else if (!dup && a.key === key && a.status === 'active') {
      dup = a;
    }
  });
  if (dup) {
    return { ok: true, duplicate: true };
  }
  var count = null;
  if (body.count !== undefined && body.count !== null && String(body.count).trim() !== '') {
    count = Number(body.count);
    if (isNaN(count) || count < 0) {
      throw new Error('The current count must be a number, 0 or more.');
    }
    count = Math.round(count * 1000) / 1000;
  }
  var sheet = alertsSheet_();
  var row = sheet.getLastRow() + 1;
  var now = new Date();
  var note = String(body.note || '').trim().slice(0, 200);
  sheet.getRange(row, 1, 1, 2).setNumberFormat('@');
  sheet.getRange(row, 3, 1, 1).setNumberFormat(TIME_FORMAT);
  sheet.getRange(row, 1, 1, HEADERS.alerts.length).setValues([[id, week.id, now, key, text_(body.by), note, 'active', '', '', count === null ? '' : count]]);
  var discord = sendLowSafe_(key, text_(body.by), note, now, count);
  return { ok: true, duplicate: false, discord: discord };
}

function resolveAlert_(body) {
  requireOpenWeek_(body.week);
  var id = String(body.id || '');
  var sheet = alertsSheet_();
  var last = sheet.getLastRow();
  if (last > 1) {
    var vals = sheet.getRange(2, 1, last - 1, HEADERS.alerts.length).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (String(vals[i][0]) === id && String(vals[i][6]) === 'active') {
        sheet.getRange(i + 2, 7, 1, 3).setValues([['cleared', new Date(), 'resolved by hand']]);
        return { ok: true };
      }
    }
  }
  return { ok: true };
}

/** Expected use for the rest of the week: last week's use (or the recent average), shared out over the open days left (Thu-Sun). */
function projectedNeed_(key) {
  var weeks = readWeeks_();
  var closed = weeks.filter(function (w) {
    return w.status === 'closed';
  });
  if (!closed.length) {
    return null;
  }
  var usage = readUsage_(weeks);
  var take = ORDER_BASIS === 'month' ? AVG_WEEKS : 1;
  var smp = [];
  for (var i = closed.length - 1; i >= 0 && smp.length < take; i--) {
    var u = usage[closed[i].id] && usage[closed[i].id][key];
    if (u) {
      smp.push(Math.max(0, u[0]));
    }
  }
  if (!smp.length) {
    return null;
  }
  var f = smp.reduce(function (a, b) {
    return a + b;
  }, 0) / smp.length;
  var dow = Number(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'u')) % 7; // 0 = Sunday
  var left = { 4: 4, 5: 3, 6: 2, 0: 1 }[dow];
  return Math.ceil(f * (left || 4) / 4 - 1e-9);
}

function sendLowSafe_(key, by, note, when, count) {
  try {
    var url = webhookUrl_();
    if (!url) {
      return 'not_set';
    }
    var it = null;
    readItems_().forEach(function (x) {
      if (keyOf_(x) === key) {
        it = x;
      }
    });
    var time = Utilities.formatDate(when, Session.getScriptTimeZone(), 'EEE h:mm a');
    var need = null;
    try {
      need = projectedNeed_(key);
    } catch (ignore) {
      need = null;
    }
    var unit = it && it.unit ? ' ' + String(it.unit).replace(/^(\d+(?:\.\d+)?)\s+([A-Za-z])/, '$1$2') : '';
    var line = [];
    if (count !== null && count !== undefined) {
      line.push('Count: **' + count + unit + '**');
    }
    if (need !== null) {
      line.push('Projected need to finish week: **' + need + unit + '**');
    }
    var msg = '\u26A0 **LOW STOCK** \u00B7 **' + (it ? it.item : key) + '**' + (it ? ' _(' + it.section + ')_' : '') +
      (line.length ? '\n' + line.join(' \u00B7 ') : '') +
      '\nReported' + (by ? ' by ' + by : '') + ' \u00B7 ' + time +
      (note ? '\n> ' + note : '');
    postDiscord_(url, msg);
    return 'sent';
  } catch (err) {
    return 'error: ' + message_(err);
  }
}

/* ---------- Discord ---------- */

function webhookUrl_() {
  return String(PropertiesService.getScriptProperties().getProperty('DISCORD_WEBHOOK_URL') || DISCORD_WEBHOOK_URL || '').trim();
}

function postDiscord_(url, content) {
  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ content: content, allowed_mentions: { parse: [] } }),
    muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Discord answered ' + code + (code === 429 ? ' (too many messages, wait a minute)' : ''));
  }
}

/** Returns 'sent', 'not_set', or 'error: ...'. Never throws, so closing a week always works. */
function sendOrderSafe_(weekId, basis) {
  try {
    var url = webhookUrl_();
    if (!url) {
      return 'not_set';
    }
    var msgs = orderMessages_(orderPlan_(weekId, basis));
    msgs.forEach(function (m, i) {
      if (i) {
        Utilities.sleep(1200);
      }
      postDiscord_(url, m);
    });
    return 'sent';
  } catch (err) {
    return 'error: ' + message_(err);
  }
}

function sendOrder_(body) {
  var weeks = readWeeks_();
  var id = body.week ? String(body.week) : '';
  if (!id) {
    weeks.forEach(function (w) {
      if (w.status === 'open' || (!id && w.status === 'closed')) {
        id = w.id;
      }
    });
  }
  if (!id) {
    throw new Error('There is no week to send.');
  }
  var result = sendOrderSafe_(id, body.basis);
  if (result === 'not_set') {
    throw new Error('Discord is not set up yet. Add DISCORD_WEBHOOK_URL in Script properties.');
  }
  if (result !== 'sent') {
    throw new Error(result.replace(/^error: /, ''));
  }
  return { ok: true, discord: 'sent' };
}

/** Run once from the editor to allow Discord and check the webhook works. */
function testDiscord() {
  var url = webhookUrl_();
  if (!url) {
    throw new Error('Add DISCORD_WEBHOOK_URL under Project Settings > Script properties first.');
  }
  postDiscord_(url, '\u2713 Kinhfolk Stock Count is connected to this channel.');
}

/* ---------- one-time upgrade for a sheet you already have ---------- */

/**
 * Run once from the editor after pasting this version. Safe to run again.
 *  - renames "Cape Roaster" to "Ca phe Roaster" on Items, Entries and Additions
 *  - adds the Month column to Weeks and fills it in (Oct 29 - Nov 1 becomes 2026-10)
 */
function upgrade() {
  getSheet_('alerts');
  var rename = function (v) {
    return typeof v === 'string' ? v.replace(/Cape Roaster/g, 'Ca phe Roaster') : v;
  };
  [['items', 2], ['entries', 2], ['adds', 4]].forEach(function (p) {
    var sheet = getSheet_(p[0]);
    var last = sheet.getLastRow();
    if (last < 2) {
      return;
    }
    var rng = sheet.getRange(2, p[1], last - 1, 1);
    rng.setValues(rng.getValues().map(function (r) {
      return [rename(r[0])];
    }));
  });
  [['adds', 2], ['entries', 1], ['weeks', 1]].forEach(function (p) {
    var sh = getSheet_(p[0]);
    var n = sh.getLastRow();
    if (n < 2) {
      return;
    }
    var col = sh.getRange(2, p[1], n - 1, 1);
    var fixed = col.getValues().map(function (r) {
      return [weekIdOf_(r[0])];
    });
    col.setNumberFormat('@');
    col.setValues(fixed);
  });
  var ws = getSheet_('weeks');
  ws.getRange(1, 7).setValue('Month').setFontWeight('bold');
  var last = ws.getLastRow();
  if (last > 1) {
    var vals = ws.getRange(2, 1, last - 1, 7).getValues();
    var out = vals.map(function (r) {
      var cur = r[6] instanceof Date ? Utilities.formatDate(r[6], Session.getScriptTimeZone(), 'yyyy-MM') : String(r[6] || '').trim();
      return [/^\d{4}-\d{2}$/.test(cur) ? cur : (String(r[0]) === '' ? '' : monthOfWeek_(String(r[2]), r[3], r[4]))];
    });
    var col = ws.getRange(2, 7, last - 1, 1);
    col.setNumberFormat('@');
    col.setValues(out);
  }
}

/* ---------- helpers ---------- */

function requireOpenWeek_(weekId) {
  var found = null;
  readWeeks_().forEach(function (w) {
    if (w.status === 'open') {
      found = w;
    }
  });
  if (!found) {
    throw new Error('There is no open week.');
  }
  if (weekId && String(weekId) !== found.id) {
    throw new Error('That week is already closed.');
  }
  return found;
}

function knownKey_(raw) {
  var key = String(raw === undefined || raw === null ? '' : raw).trim();
  var ok = false;
  readItems_().forEach(function (it) {
    if (keyOf_(it) === key) {
      ok = true;
    }
  });
  if (!ok) {
    throw new Error('That item is not on the Items tab.');
  }
  return key;
}

function findEntryRow_(sheet, weekId, key) {
  var last = sheet.getLastRow();
  if (last < 2) {
    return 0;
  }
  var v = sheet.getRange(2, 1, last - 1, 2).getValues();
  for (var i = 0; i < v.length; i++) {
    if (weekIdOf_(v[i][0]) === weekId && String(v[i][1]) === key) {
      return i + 2;
    }
  }
  return 0;
}

function findWeekRow_(sheet, id) {
  var last = sheet.getLastRow();
  var v = last > 1 ? sheet.getRange(2, 1, last - 1, 1).getValues() : [];
  for (var i = 0; i < v.length; i++) {
    if (weekIdOf_(v[i][0]) === id) {
      return i + 2;
    }
  }
  throw new Error('Week not found.');
}

/** Sheets sometimes turns a week id like 2026-11-03 into a date; turn it back into text. */
function weekIdOf_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v === undefined || v === null ? '' : v);
}

function keyOf_(it) {
  return it.section + ' / ' + it.item;
}

function text_(v) {
  return String(v === undefined || v === null ? '' : v).trim().slice(0, MAX_TEXT);
}

function iso_(v) {
  return v instanceof Date ? v.toISOString() : (v === '' || v === null ? null : String(v));
}

function getSheet_(which) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = NAMES[which];
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, HEADERS[which].length).setValues([HEADERS[which]]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** Accepts JSON text in the POST body, or ordinary form fields as a fallback. */
function parseBody_(e) {
  if (e && e.postData && e.postData.contents) {
    try {
      var parsed = JSON.parse(e.postData.contents);
      if (parsed && typeof parsed === 'object') {
        return parsed;
      }
    } catch (ignore) {
      // not JSON, fall through to form fields
    }
  }
  if (e && e.parameter && e.parameter.action) {
    return e.parameter;
  }
  throw new Error('Send a JSON body such as {"action":"closeWeek"}.');
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function message_(err) {
  return String((err && err.message) || err);
}

/* ---------- one-time starter data, from your COGS sheet (Oct 29 - Nov 1) ---------- */

/**
 * Run once from the editor. It REPLACES everything on all four tabs with:
 *  - your 27 items, units and unit costs (Vendor is blank: fill it in),
 *  - last week from your COGS as a closed week (so History has something in it),
 *  - a new open week whose beginning counts are last week's ending counts.
 */
function loadInventory() {
  Object.keys(NAMES).forEach(function (k) {
    var sheet = getSheet_(k);
    var last = sheet.getLastRow();
    if (last > 1) {
      sheet.getRange(2, 1, last - 1, HEADERS[k].length).clearContent();
    }
  });
  getSheet_('items').getRange(2, 1, INVENTORY.length, HEADERS.items.length).setValues(INVENTORY);

  var tz = Session.getScriptTimeZone();
  var started = new Date(2026, 9, 29, 8, 0, 0);
  var closed = new Date(2026, 10, 1, 20, 0, 0);
  var costs = {};
  INVENTORY.forEach(function (r) {
    costs[r[0] + ' / ' + r[1]] = r[3];
  });
  var total = 0;
  var entryRows = [];
  var addRows = [];
  var ends = {};
  LAST_WEEK.forEach(function (r, i) {
    var used = Math.round((r[1] + r[2] - r[3]) * 1000) / 1000;
    total += used * costs[r[0]];
    entryRows.push(['2026-10-29', r[0], r[1], r[3], costs[r[0]], used]);
    ends[r[0]] = r[3];
    if (r[2] > 0) {
      addRows.push(['seed-' + i, '2026-10-29', new Date(2026, 9, 29, 12, 0, 0), r[0], r[2], 'From COGS sheet']);
    }
  });
  total = Math.round(total * 100) / 100;
  getSheet_('weeks').getRange(2, 7).setNumberFormat('@');
  getSheet_('weeks').getRange(2, 1, 1, 7).setValues([['2026-10-29', 'Oct 29 - Nov 1', 'closed', started, closed, total, '2026-10']]);
  getSheet_('entries').getRange(2, 1, entryRows.length, 6).setValues(entryRows);
  getSheet_('adds').getRange(2, 1, addRows.length, 6).setValues(addRows);
  newWeek_(ends);
}

var INVENTORY = [
 [
  "Beans + Matcha",
  "Riverphin Countryside - Medium",
  "lb",
  12.7,
  "",
  ""
 ],
 [
  "Beans + Matcha",
  "Ca phe Roaster - True Robusta",
  "lb",
  10.0,
  "",
  ""
 ],
 [
  "Beans + Matcha",
  "Ca phe Roaster - Single Origin",
  "lb",
  12.0,
  "",
  ""
 ],
 [
  "Beans + Matcha",
  "Matcha 250g",
  "each",
  70.0,
  "",
  ""
 ],
 [
  "Milk + Juices",
  "Milk (Half Gallon)",
  "half gallon",
  3.0,
  "",
  ""
 ],
 [
  "Milk + Juices",
  "Heavy Cream",
  "half gallon",
  11.0,
  "",
  ""
 ],
 [
  "Milk + Juices",
  "Oat Milk",
  "carton",
  3.0,
  "",
  "Count front fridge"
 ],
 [
  "Milk + Juices",
  "Almond Milk",
  "carton",
  3.33,
  "",
  ""
 ],
 [
  "Milk + Juices",
  "Coconut Milk (Chaokoh)",
  "can",
  1.38,
  "",
  ""
 ],
 [
  "Milk + Juices",
  "Coconut Milk (Arroy D)",
  "can",
  2.54,
  "",
  ""
 ],
 [
  "Syrups + Garnish",
  "Pandan Syrup",
  "500 ml",
  1.62,
  "",
  ""
 ],
 [
  "Syrups + Garnish",
  "Earl Grey Syrup",
  "500 ml",
  1.69,
  "",
  ""
 ],
 [
  "Syrups + Garnish",
  "Pho Syrup",
  "500 ml",
  1.79,
  "",
  ""
 ],
 [
  "Syrups + Garnish",
  "Condensed Milk",
  "can",
  1.8,
  "",
  ""
 ],
 [
  "Syrups + Garnish",
  "Coconut Concentrate",
  "bottle",
  3.19,
  "",
  ""
 ],
 [
  "Cups",
  "Hot Cups (10oz)",
  "sleeve (25)",
  6.75,
  "",
  ""
 ],
 [
  "Cups",
  "Hot Cups (8oz)",
  "sleeve (25)",
  2.5,
  "",
  ""
 ],
 [
  "Cups",
  "Hot Lids",
  "sleeve (50)",
  1.75,
  "",
  ""
 ],
 [
  "Cups",
  "Cold Cups",
  "sleeve (100)",
  4.5,
  "",
  ""
 ],
 [
  "Cups",
  "Cold Lids",
  "sleeve (100)",
  1.75,
  "",
  ""
 ],
 [
  "Retail",
  "Marou Kumquat",
  "each",
  2.95,
  "",
  ""
 ],
 [
  "Retail",
  "Marou Lam Dong",
  "each",
  2.65,
  "",
  ""
 ],
 [
  "Retail",
  "Marou Tien Giang",
  "each",
  2.65,
  "",
  ""
 ],
 [
  "Retail",
  "Ca phe Roaster - Single Origin",
  "each",
  10.0,
  "",
  ""
 ],
 [
  "Retail",
  "Ca phe Roaster - 6th Street",
  "each",
  10.0,
  "",
  ""
 ],
 [
  "Retail",
  "Ca phe Roaster - Kensington Kindeng",
  "each",
  12.0,
  "",
  ""
 ],
 [
  "Retail",
  "Workers Coffee - Anaerobic Natural",
  "each",
  11.65,
  "",
  ""
 ]
];

// [key, beginning, added, ending]
var LAST_WEEK = [["Beans + Matcha / Riverphin Countryside - Medium", 5, 20, 12.22], ["Beans + Matcha / Ca phe Roaster - True Robusta", 12.572, 25, 22], ["Beans + Matcha / Ca phe Roaster - Single Origin", 19.866, 0, 10.548], ["Beans + Matcha / Matcha 250g", 4.25, 8, 9.2], ["Milk + Juices / Milk (Half Gallon)", 11.75, 19, 9.25], ["Milk + Juices / Heavy Cream", 2, 5, 4], ["Milk + Juices / Oat Milk", 17, 72, 31.5], ["Milk + Juices / Almond Milk", 14, 0, 9.5], ["Milk + Juices / Coconut Milk (Chaokoh)", 16, 0, 0], ["Milk + Juices / Coconut Milk (Arroy D)", 37, 0, 33], ["Syrups + Garnish / Pandan Syrup", 6, 6, 2], ["Syrups + Garnish / Earl Grey Syrup", 1.5, 8, 3.5], ["Syrups + Garnish / Pho Syrup", 3.25, 5, 3.5], ["Syrups + Garnish / Condensed Milk", 94, 0, 77], ["Syrups + Garnish / Coconut Concentrate", 7, 0, 5.25], ["Cups / Hot Cups (10oz)", 13, 0, 8.5], ["Cups / Hot Cups (8oz)", 3, 0, 0.25], ["Cups / Hot Lids", 24, 0, 21], ["Cups / Cold Cups", 12, 20, 18.5], ["Cups / Cold Lids", 7.5, 20, 17], ["Retail / Marou Kumquat", 37, 0, 32], ["Retail / Marou Lam Dong", 20, 0, 20], ["Retail / Marou Tien Giang", 16, 0, 14], ["Retail / Ca phe Roaster - Single Origin", 0, 0, 0], ["Retail / Ca phe Roaster - 6th Street", 13, 0, 10], ["Retail / Ca phe Roaster - Kensington Kindeng", 13, 0, 11], ["Retail / Workers Coffee - Anaerobic Natural", 5, 0, 4]];
