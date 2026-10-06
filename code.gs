/**
 * Mileage Logbook — Google Sheets backend (Google Apps Script)
 * ------------------------------------------------------------
 * Turns this Google Sheet into the database for the Mileage Logbook app.
 *
 * SETUP (once):
 *   1. In the menu above choose  Run ▶  with the function "setup" selected.
 *      Approve the permissions. The execution log shows your SECRET KEY.
 *   2. Deploy → New deployment → type "Web app"
 *        Execute as:      Me
 *        Who has access:  Anyone
 *      Copy the Web app URL (ends with /exec).
 *   3. In the app: Settings → Google Sheets → paste the URL and the secret key.
 *
 * Only requests that carry the secret key are accepted.
 * Need a new key (e.g. a phone was lost)? Run "newSecretKey" and enter the new key on your devices.
 */

const TRIPS_SHEET = 'Trips';
const DELETED_SHEET = 'Deleted trips';
const APP_SHEET = 'App';
const HEADERS = ['ID', 'Date', 'Trip type', 'Description', 'Start time', 'End time', 'Start KM', 'End KM', 'KM driven', 'Created', 'Updated'];
const COL = { id: 1, date: 2, type: 3, description: 4, startTime: 5, endTime: 6, startKm: 7, endKm: 8, km: 9, created: 10, updated: 11 };
const NCOLS = HEADERS.length;

/* ---------- one-time setup ---------- */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trips = getOrCreateSheet_(TRIPS_SHEET, HEADERS);
  const deleted = getOrCreateSheet_(DELETED_SHEET, HEADERS.concat(['Deleted at']));
  const app = getOrCreateSheet_(APP_SHEET, ['Key', 'Value']);
  [trips, deleted].forEach(function (s) {
    s.getRange('B:B').setNumberFormat('yyyy-mm-dd');
    s.getRange('E:F').setNumberFormat('@');
    s.getRange('G:I').setNumberFormat('0.0');
    s.getRange('J:L').setNumberFormat('@');
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, s.getLastColumn()).setFontWeight('bold');
    s.hideColumns(1); // the ID column is for the app only
  });
  app.getRange('B:B').setNumberFormat('@');

  // Remove the empty default sheet, if any
  const first = ss.getSheetByName('Sheet1') || ss.getSheetByName('Blad1');
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);

  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('SECRET');
  if (!key) { key = Utilities.getUuid().replace(/-/g, ''); props.setProperty('SECRET', key); }
  Logger.log('Setup complete. Your secret key is:  ' + key);
}

function newSecretKey() {
  const key = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('SECRET', key);
  Logger.log('New secret key:  ' + key + '   (enter it again on every device)');
}

/* ---------- web app entry points ---------- */

function doGet() {
  return ContentService.createTextOutput('Mileage Logbook backend is running. Use the app to connect.');
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }

  const secret = PropertiesService.getScriptProperties().getProperty('SECRET');
  if (!secret || !req || req.key !== secret) return json_({ ok: false, error: 'unauthorized', message: 'Secret key not accepted.' });

  const action = ACTIONS[req.action];
  if (!action) return json_({ ok: false, error: 'unknown_action' });

  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (err) { return json_({ ok: false, error: 'busy', message: 'The Sheet is busy, try again.' }); }
  try {
    const result = action(req.data || {});
    return json_(Object.assign({ ok: true }, result));
  } catch (err) {
    return json_({ ok: false, error: err.code || 'server', message: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

const ACTIONS = {
  ping: function () {
    return { name: SpreadsheetApp.getActiveSpreadsheet().getName() };
  },

  list: function () {
    const r = readTrips_();
    return { trips: r.trips, problems: r.problems, active: getApp_('active'), settings: getApp_('settings') };
  },

  upsert: function (d) {
    upsertTrip_(d.trip);
    sortTrips_();
    return {};
  },

  upsertMany: function (d) {
    const added = upsertMany_(d.trips || []);
    sortTrips_();
    return { added: added };
  },

  delete: function (d) {
    deleteTrip_(String(d.id || ''));
    return {};
  },

  /* Start (active = {...}) or discard (active = null) the active trip. Only one active trip at a time. */
  setActive: function (d) {
    const current = getApp_('active');
    if (d.active) {
      if (current && current.startedAt !== d.active.startedAt) return { conflict: true, active: current };
      setApp_('active', d.active);
    } else {
      if (current && d.expect && current.startedAt !== d.expect) return { conflict: true, active: current };
      setApp_('active', null);
    }
    return {};
  },

  /* Save the finished trip and clear the active trip in one step. Safe to retry. */
  endTrip: function (d) {
    upsertTrip_(d.trip);
    sortTrips_();
    const current = getApp_('active');
    if (current && current.startedAt === d.startedAt) setApp_('active', null);
    return {};
  },

  setSettings: function (d) {
    const s = d.settings || {};
    setApp_('settings', { vehicle: str_(s.vehicle, 100), plate: str_(s.plate, 30), driver: str_(s.driver, 100), company: str_(s.company, 100) });
    return {};
  },
};

/* ---------- trips ---------- */

function readTrips_() {
  const sheet = sheetOrFail_(TRIPS_SHEET);
  const n = sheet.getLastRow() - 1;
  const trips = [], problems = [];
  if (n < 1) return { trips: trips, problems: problems };
  const range = sheet.getRange(2, 1, n, NCOLS);
  const values = range.getValues(), shown = range.getDisplayValues();
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();

  for (let i = 0; i < n; i++) {
    const v = values[i], s = shown[i], row = i + 2;
    if (v.slice(1, 9).every(function (x) { return x === '' || x === null; })) continue; // empty row
    let id = String(v[COL.id - 1] || '');
    if (!id) { id = Utilities.getUuid(); sheet.getRange(row, COL.id).setValue(id); } // row typed in by hand

    const date = v[COL.date - 1] instanceof Date ? Utilities.formatDate(v[COL.date - 1], tz, 'yyyy-MM-dd') : normDate_(s[COL.date - 1]);
    const type = normType_(s[COL.type - 1]);
    const startKm = num_(v[COL.startKm - 1]), endKm = num_(v[COL.endKm - 1]);
    const startTime = normTime_(v[COL.startTime - 1], s[COL.startTime - 1], tz);
    const endTime = normTime_(v[COL.endTime - 1], s[COL.endTime - 1], tz);

    const errs = [];
    if (!date) errs.push('date not recognised');
    if (!type) errs.push('trip type must be Business or Private');
    if (isNaN(startKm)) errs.push('start km missing');
    if (isNaN(endKm)) errs.push('end km missing');
    if (!errs.length && endKm < startKm) errs.push('end km is lower than start km');
    if (errs.length) { problems.push('Row ' + row + ': ' + errs.join(', ')); continue; }

    trips.push({
      id: id, date: date, type: type, description: String(v[COL.description - 1] || ''),
      startTime: startTime, endTime: endTime, startKm: startKm, endKm: endKm,
      km: Math.round((endKm - startKm) * 10) / 10,
    });
  }
  return { trips: trips, problems: problems };
}

/** Insert or update one trip by ID. */
function upsertTrip_(t) {
  const trip = validateTrip_(t);
  const sheet = sheetOrFail_(TRIPS_SHEET);
  const now = new Date().toISOString();
  const rowIndex = findRow_(sheet, trip.id);
  const created = rowIndex > 0 ? String(sheet.getRange(rowIndex, COL.created).getValue() || now) : now;
  const row = toRow_(trip, created, now);
  if (rowIndex > 0) sheet.getRange(rowIndex, 1, 1, NCOLS).setValues([row]);
  else sheet.getRange(sheet.getLastRow() + 1, 1, 1, NCOLS).setValues([row]);

  // If this trip was deleted earlier (e.g. "Undo"), remove it from the deleted list
  const del = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DELETED_SHEET);
  if (del) { const r = findRow_(del, trip.id); if (r > 0) del.deleteRow(r); }
}

/** Bulk insert (import / first upload). Trips already present (same ID, or same date + odometer readings) are skipped. */
function upsertMany_(list) {
  const sheet = sheetOrFail_(TRIPS_SHEET);
  const existing = readTrips_().trips;
  const ids = {}, keys = {};
  existing.forEach(function (x) { ids[x.id] = true; keys[x.date + '|' + x.startKm + '|' + x.endKm] = true; });
  const now = new Date().toISOString(), rows = [];
  list.forEach(function (t) {
    const trip = validateTrip_(t);
    const key = trip.date + '|' + trip.startKm + '|' + trip.endKm;
    if (ids[trip.id] || keys[key]) return;
    ids[trip.id] = keys[key] = true;
    rows.push(toRow_(trip, now, now));
  });
  if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, NCOLS).setValues(rows);
  return rows.length;
}

function toRow_(trip, created, updated) {
  return [trip.id, trip.date, trip.type, safeText_(trip.description), trip.startTime, trip.endTime,
    trip.startKm, trip.endKm, trip.km, created, updated];
}

function deleteTrip_(id) {
  if (!id) throw err_('invalid', 'Missing trip ID.');
  const sheet = sheetOrFail_(TRIPS_SHEET);
  const r = findRow_(sheet, id);
  if (r < 0) return; // already deleted
  const values = sheet.getRange(r, 1, 1, NCOLS).getValues()[0];
  const shown = sheet.getRange(r, 1, 1, NCOLS).getDisplayValues()[0];
  values[COL.date - 1] = shown[COL.date - 1];
  values[COL.description - 1] = safeText_(values[COL.description - 1]);
  const del = getOrCreateSheet_(DELETED_SHEET, HEADERS.concat(['Deleted at']));
  del.getRange(del.getLastRow() + 1, 1, 1, NCOLS + 1).setValues([values.concat([new Date().toISOString()])]);
  sheet.deleteRow(r);
}

function sortTrips_() {
  const sheet = sheetOrFail_(TRIPS_SHEET);
  const n = sheet.getLastRow() - 1;
  if (n > 1) sheet.getRange(2, 1, n, NCOLS).sort([{ column: COL.date, ascending: true }, { column: COL.startKm, ascending: true }]);
}

function validateTrip_(t) {
  if (!t || typeof t !== 'object') throw err_('invalid', 'No trip data.');
  const id = String(t.id || '').trim();
  const date = String(t.date || '');
  const type = normType_(t.type);
  const startKm = num_(t.startKm), endKm = num_(t.endKm);
  const startTime = String(t.startTime || ''), endTime = String(t.endTime || '');
  if (!id || id.length > 64) throw err_('invalid', 'Invalid trip ID.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw err_('invalid', 'Invalid date: ' + date);
  if (!type) throw err_('invalid', 'Trip type must be Business or Private.');
  if (isNaN(startKm) || isNaN(endKm) || startKm < 0) throw err_('invalid', 'Invalid odometer readings.');
  if (endKm < startKm) throw err_('invalid', 'End km (' + endKm + ') cannot be lower than start km (' + startKm + ').');
  [startTime, endTime].forEach(function (x) { if (x && !/^([01]\d|2[0-3]):[0-5]\d$/.test(x)) throw err_('invalid', 'Invalid time: ' + x); });
  return {
    id: id, date: date, type: type, description: str_(t.description, 500), startTime: startTime, endTime: endTime,
    startKm: Math.round(startKm * 10) / 10, endKm: Math.round(endKm * 10) / 10,
    km: Math.round((endKm - startKm) * 10) / 10, // always recalculated here
  };
}

/* ---------- App sheet: small key/value store ---------- */

function getApp_(key) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(APP_SHEET);
  if (!sheet) return null;
  const r = findRow_(sheet, key);
  if (r < 0) return null;
  const raw = String(sheet.getRange(r, 2).getValue() || '');
  try { return raw ? JSON.parse(raw) : null; } catch (e) { return null; }
}

function setApp_(key, value) {
  const sheet = getOrCreateSheet_(APP_SHEET, ['Key', 'Value']);
  const text = value == null ? '' : JSON.stringify(value);
  const r = findRow_(sheet, key);
  if (r > 0) sheet.getRange(r, 2).setValue(text);
  else sheet.getRange(sheet.getLastRow() + 1, 1, 1, 2).setValues([[key, text]]);
}

/* ---------- helpers ---------- */

function findRow_(sheet, id) {
  if (!id || sheet.getLastRow() < 2) return -1;
  const cell = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).createTextFinder(id).matchEntireCell(true).findNext();
  return cell ? cell.getRow() : -1;
}

function getOrCreateSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(name);
  if (!s) s = ss.insertSheet(name);
  if (s.getLastRow() === 0) s.getRange(1, 1, 1, headers.length).setValues([headers]);
  return s;
}

function sheetOrFail_(name) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!s) throw err_('setup', 'Sheet "' + name + '" not found. Run setup() in the script editor.');
  return s;
}

function normType_(x) {
  const s = String(x || '').trim().toLowerCase();
  if (/^(business|zakelijk|b)$/.test(s)) return 'Business';
  if (/^(private|privé|prive|p)$/.test(s)) return 'Private';
  return null;
}

function normDate_(s) {
  s = String(s || '').trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return m[1] + '-' + pad_(m[2]) + '-' + pad_(m[3]);
  m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})$/);
  if (m) return m[3] + '-' + pad_(m[2]) + '-' + pad_(m[1]);
  return null;
}

function normTime_(value, shown, tz) {
  if (value instanceof Date) return Utilities.formatDate(value, tz, 'HH:mm');
  const m = String(shown || '').trim().match(/^(\d{1,2})[:.](\d{2})/);
  return m && +m[1] < 24 && +m[2] < 60 ? pad_(m[1]) + ':' + m[2] : '';
}

function num_(x) {
  if (typeof x === 'number') return x;
  const s = String(x == null ? '' : x).trim().replace(/\s/g, '').replace(',', '.');
  return /^\d+(\.\d+)?$/.test(s) ? Number(s) : NaN;
}

function str_(x, max) { return String(x == null ? '' : x).trim().slice(0, max); }
function pad_(n) { return ('0' + n).slice(-2); }
/** Stop text such as "=..." or "+..." from being treated as a formula by Sheets. */
function safeText_(s) { s = String(s == null ? '' : s); return /^[=+\-@]/.test(s) ? "'" + s : s; }
function err_(code, message) { const e = new Error(message); e.code = code; return e; }
function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
