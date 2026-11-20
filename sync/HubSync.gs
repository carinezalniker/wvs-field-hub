/**
 * WVS Field Hub — Repsly sync
 * ---------------------------------------------------------------------------
 * Reads the Stores, Orders and Reps tabs of "WVS Store Master (Repsly sync)"
 * and writes them into the Field Hub's Firestore database.
 *
 * It signs in as whoever owns the script (you), so there is no key to store
 * and nothing to leak. Run syncToHub() once by hand, then set a daily trigger.
 *
 * Install: Extensions → Apps Script, then File → + → Script and name it
 * HubSync. Paste this in there. Do NOT touch the existing Code.gs — that is
 * the Repsly→Sheet sync and this one runs after it, off the same tabs.
 * Then add the datastore scope to appsscript.json (see the notes in chat).
 */

var PROJECT_ID = 'wvs-field-hub';
var SHEET_ID   = '1oOAmIoJFTH3FGIOeVmKh-Vs5zSDcNDj3BrVTMYfxHdU';

/* Tags Repsly carries that are brands, not banners — ignored when working out
   which retailer a store belongs to. */
var BRAND_TAGS = ['exso','fix','lita','south side','southside','pretzels','takis','crackfers','ev3'];

/* Banner is decided by the first rule that matches the remaining tags. */
var BANNERS = [
  [/spar\s*tops|tops\s*at\s*spar/i, 'TOPS at SPAR'],
  [/\bspar\b/i,                     'Spar'],
  [/shell/i,                        'Shell'],
  [/makro/i,                        'Makro'],
  [/pnp|pick\s*n\s*pay/i,           'Pick n Pay'],
];

function syncToHub() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var token = ScriptApp.getOAuthToken();
  var now = new Date().toISOString();

  var reps = readReps_(ss);
  var stores = buildStores_(ss, reps);
  var orders = buildOrders_(ss, reps);

  var written = 0, log = [];
  for (var id in stores) {
    writeDoc_(token, 'storelists/' + id,
      { stores: stores[id], updatedAt: now, source: 'Repsly sync' });
    log.push('storelists/' + id + ' (' + stores[id].length + ')');
    written++;
  }
  for (var rid in orders) {
    writeDoc_(token, 'orders/' + rid,
      { orders: orders[rid], updatedAt: now, source: 'Repsly sync' });
    log.push('orders/' + rid + ' (' + orders[rid].length + ')');
    written++;
  }
  noteSync_(ss, written, log.join(', '));
  return log.join('\n');
}

/* ---------- reading the sheet ---------- */

function readReps_(ss) {
  var rows = ss.getSheetByName('Reps').getDataRange().getValues();
  var head = rows.shift().map(String);
  var iCode = head.indexOf('Rep Code'), iName = head.indexOf('Rep Name'),
      iHub = head.indexOf('Hub Rep ID');
  var byCode = {}, byName = {};
  rows.forEach(function (r) {
    var hub = String(r[iHub] || '').trim();
    if (!hub) return;
    byCode[String(r[iCode]).trim()] = hub;
    byName[String(r[iName]).trim().toLowerCase()] = hub;
  });
  return { byCode: byCode, byName: byName };
}

function buildStores_(ss, reps) {
  var rows = ss.getSheetByName('Stores').getDataRange().getValues();
  var h = rows.shift().map(String);
  var i = {
    code: h.indexOf('Client Code'), name: h.indexOf('Store Name'),
    tags: h.indexOf('Retailer'), active: h.indexOf('Active'),
    terr: h.indexOf('Territory'), rep: h.indexOf('Rep Name'),
    town: h.indexOf('Town'), brands: h.indexOf('Brands Stocked (90 days)'),
    lo: h.indexOf('Last Order Date'), lv: h.indexOf('Last Visit Date')
  };
  var out = {}, seen = {};
  rows.forEach(function (r) {
    var name = String(r[i.name] || '').trim();
    if (!name) return;

    var store = {
      c: String(r[i.code] || '').trim(),
      n: name,
      r: bannerOf_(r[i.tags]),
      t: String(r[i.town] || '').trim()
    };
    var brands = splitTags_(r[i.brands]).map(function (b) { return b.toLowerCase(); });
    if (brands.length) store.b = brands;
    var lv = toISO_(r[i.lv]), lo = toISO_(r[i.lo]);
    if (lv) store.lv = lv;
    if (lo) store.lo = lo;
    if (!isTrue_(r[i.active])) store.x = 1;
    if (r[i.terr]) store.p = String(r[i.terr]).trim();

    // A row can name several reps; the store belongs on each of their lists.
    var ids = splitTags_(r[i.rep])
      .map(function (n) { return reps.byName[n.toLowerCase()]; })
      .filter(function (x) { return !!x; });
    if (!ids.length) ids = ['unassigned'];

    ids.forEach(function (id) {
      var k = id + '|' + store.c;
      if (seen[k]) return;          // same Repsly record twice
      seen[k] = 1;
      (out[id] = out[id] || []).push(store);
    });
  });
  return out;
}

function buildOrders_(ss, reps) {
  var rows = ss.getSheetByName('Orders').getDataRange().getValues();
  var h = rows.shift().map(String);
  var i = {
    id: h.indexOf('Order ID'), date: h.indexOf('Order Date'),
    code: h.indexOf('Client Code'), store: h.indexOf('Store Name'),
    rep: h.indexOf('Rep Code'), product: h.indexOf('Product Name'),
    brand: h.indexOf('Brand'), qty: h.indexOf('Quantity'), val: h.indexOf('Value')
  };
  var byRep = {}, index = {};
  rows.forEach(function (r) {
    var oid = String(r[i.id] || '').trim();
    if (!oid) return;
    var rep = reps.byCode[String(r[i.rep]).trim()];
    if (!rep) return;
    var key = rep + '|' + oid;
    var o = index[key];
    if (!o) {
      o = index[key] = {
        id: oid, d: toISO_(r[i.date]),
        c: String(r[i.code] || '').trim(),
        s: String(r[i.store] || '').trim(),
        l: []
      };
      (byRep[rep] = byRep[rep] || []).push(o);
    }
    o.l.push([
      String(r[i.product] || '').trim(),
      String(r[i.brand] || '').trim().toLowerCase(),
      Number(r[i.qty]) || 0,
      Math.round((Number(r[i.val]) || 0) * 100) / 100
    ]);
  });
  // newest first, same as the hub expects
  for (var k in byRep) byRep[k].sort(function (a, b) { return (b.d || '').localeCompare(a.d || ''); });
  return byRep;
}

/* ---------- small helpers ---------- */

function splitTags_(v) {
  return String(v == null ? '' : v).split(',')
    .map(function (s) { return s.trim(); })
    .filter(function (s) { return !!s; });
}

function bannerOf_(v) {
  var tags = splitTags_(v).filter(function (t) {
    return BRAND_TAGS.indexOf(t.toLowerCase()) === -1;
  });
  var joined = tags.join(' ');
  for (var i = 0; i < BANNERS.length; i++) {
    if (BANNERS[i][0].test(joined)) return BANNERS[i][1];
  }
  return 'Independent';
}

function isTrue_(v) {
  return v === true || String(v).trim().toUpperCase() === 'TRUE';
}

/** Dates arrive as Date objects, as 2026-07-23, or as 23/07/2026. */
function toISO_(v) {
  if (!v && v !== 0) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, 'Africa/Johannesburg', 'yyyy-MM-dd');
  }
  var s = String(v).trim();
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  var m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) return m[3] + '-' + pad_(m[2]) + '-' + pad_(m[1]);
  return '';
}

function pad_(n) { return ('0' + n).slice(-2); }

/* ---------- writing to Firestore ---------- */

/** Firestore has no arrays inside arrays; the hub reads them back as {__a:[…]}. */
function toValue_(v) {
  if (v === null || v === undefined) return { nullValue: null };
  if (typeof v === 'boolean') return { booleanValue: v };
  if (typeof v === 'number') {
    return Number.isInteger(v) ? { integerValue: String(v) } : { doubleValue: v };
  }
  if (Array.isArray(v)) {
    return { arrayValue: { values: v.map(function (x) {
      return Array.isArray(x) ? toValue_({ __a: x }) : toValue_(x);
    }) } };
  }
  if (typeof v === 'object') {
    var f = {};
    for (var k in v) f[k] = toValue_(v[k]);
    return { mapValue: { fields: f } };
  }
  return { stringValue: String(v) };
}

function writeDoc_(token, path, data) {
  var fields = {};
  for (var k in data) fields[k] = toValue_(data[k]);
  var url = 'https://firestore.googleapis.com/v1/projects/' + PROJECT_ID +
            '/databases/(default)/documents/' + path;
  var res = UrlFetchApp.fetch(url, {
    method: 'patch',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ fields: fields }),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) {
    throw new Error('Writing ' + path + ' failed (' + res.getResponseCode() + '): ' +
                    res.getContentText().slice(0, 400));
  }
}

function noteSync_(ss, written, detail) {
  var sh = ss.getSheetByName('Sync Log');
  if (!sh) return;
  // same six columns the Repsly sync logs
  sh.appendRow([new Date(), '', '', '', 'HUB OK', written + ' docs — ' + detail.slice(0, 350)]);
}

/**
 * Optional: run once to sync to the hub every morning at 6am.
 * Not needed if you add syncToHub() to the end of syncAll(), which keeps the
 * hub in step with every hourly Repsly pull instead.
 */
function createDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncToHub') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncToHub').timeBased().atHour(6).everyDays(1).create();
}
