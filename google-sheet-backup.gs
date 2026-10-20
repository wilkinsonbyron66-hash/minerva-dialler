/* Minerva Dialler - Google Sheet sync (version 3)
   1. Paste into the Google Sheet: Extensions > Apps Script (replace everything) and Save.
   2. Choose "setup" in the function list at the top and click Run once. Allow access when asked.
   3. Deploy > Manage deployments > Edit (pencil) > Version: New version > Deploy
      (first time: Deploy > New deployment > Web app, Execute as "Me", Who has access "Anyone").
   The full backup is kept as a file in your Google Drive (next to this sheet), with a daily
   snapshot file for 30 days. The sheet holds readable "Call Log" and "Leads" tabs. */
const KEY = '__KEY__';
const SCRIPT_VERSION = 3;
const KEEP_DAYS = 30;
const READABLE_EVERY_MS = 3 * 60 * 1000;  // refresh the readable tabs at most every 3 minutes

function doPost(e) {
  let res;
  try {
    const req = JSON.parse(e.postData.contents);
    if (req.key !== KEY) res = { ok: false, error: 'key' };
    else if (req.action === 'ping') res = Object.assign({ ok: true, script: SCRIPT_VERSION }, info_());
    else if (req.action === 'restore') res = restore_(req.day);
    else if (req.action === 'snapshots') res = { ok: true, list: snapList_() };
    else if (req.action === 'save') res = locked_(function () { return save_(req); });
    else res = { ok: false, error: 'action' };
    if (req.key === KEY && req.action === 'ping') { try { tidyReadable_(); } catch (x) {} }
  } catch (err) {
    res = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput('Minerva Dialler backup is running.');
}

// Run this once from the editor: it asks for permission and moves any old backup out of the sheet.
function setup() {
  folder_();
  const p = props_();
  if (!p.getProperty('v')) {
    const old = oldSheetBackup_();
    if (old) {
      file_('minerva-dialler-backup.json', true).setContent(old.json);
      p.setProperties({ v: '1', saved: String(Date.now()), leads: String(old.leads), calls: String(old.calls) });
    }
  }
  ['Backup', 'Snapshots'].forEach(function (n) { const s = ss_().getSheetByName(n); if (s && ss_().getSheets().length > 1) ss_().deleteSheet(s); });
  return 'Ready';
}

function locked_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(45000)) return { ok: false, error: 'busy' };
  try { return fn(); } finally { lock.releaseLock(); }
}

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function props_() { return PropertiesService.getScriptProperties(); }
function folder_() {
  const parents = DriveApp.getFileById(ss_().getId()).getParents();
  return parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
}
function file_(name, create) {
  const it = folder_().getFilesByName(name);
  if (it.hasNext()) return it.next();
  return create ? folder_().createFile(name, '', 'application/json') : null;
}

function info_() {
  const p = props_().getProperties();
  return { url: ss_().getUrl(), saved: Number(p.saved) || null, leads: Number(p.leads) || 0, calls: Number(p.calls) || 0, v: Number(p.v) || 0 };
}

function save_(req) {
  const st = req.state || {};
  const rows = st.rows || [];
  const leads = rows.length;
  const calls = rows.reduce(function (n, r) { return n + ((r.log || []).length); }, 0);
  const prev = info_();
  if (req.expect != null && Number(req.expect) !== prev.v) return { ok: false, error: 'stale', v: prev.v };
  if (!req.force) {
    if (leads === 0 && prev.leads > 0) return { ok: false, error: 'empty', prev: prev };
    if (prev.calls >= 20 && calls < prev.calls * 0.5) return { ok: false, error: 'shrink', prev: prev };
  }
  const now = new Date();
  const json = JSON.stringify(st);
  const ver = prev.v + 1;
  file_('minerva-dialler-backup.json', true).setContent(json);
  const day = Utilities.formatDate(now, ss_().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  const snap = file_('minerva-dialler-snapshot-' + day + '.json', true);
  snap.setContent(json);
  snap.setDescription(leads + ' leads, ' + calls + ' calls');
  props_().setProperties({ v: String(ver), saved: String(now.getTime()), leads: String(leads), calls: String(calls), dirty: '1' });
  pruneSnapshots_();
  maybeReadable_(st, false);
  return { ok: true, saved: now.getTime(), leads: leads, calls: calls, v: ver, url: ss_().getUrl() };
}

function restore_(day) {
  let json = null;
  if (day) {
    const f = file_('minerva-dialler-snapshot-' + day + '.json', false);
    if (!f) return { ok: false, error: 'none' };
    json = f.getBlob().getDataAsString();
  } else {
    const f = file_('minerva-dialler-backup.json', false);
    if (f) json = f.getBlob().getDataAsString();
    else { const old = oldSheetBackup_(); if (old) json = old.json; }
  }
  return { ok: true, state: json ? JSON.parse(json) : null, v: info_().v };
}

function snapList_() {
  const out = [];
  const it = folder_().getFiles();
  while (it.hasNext()) {
    const f = it.next(), m = f.getName().match(/^minerva-dialler-snapshot-(\d{4}-\d{2}-\d{2})\.json$/);
    if (m) out.push({ day: m[1], saved: f.getLastUpdated().toISOString(), info: f.getDescription() || '' });
  }
  out.sort(function (a, b) { return a.day < b.day ? 1 : -1; });
  return out.map(function (x) { const n = x.info.match(/(\d+) leads, (\d+) calls/) || []; return { day: x.day, saved: x.saved, leads: n[1] || '?', calls: n[2] || '?' }; });
}

function pruneSnapshots_() {
  const list = snapList_();
  list.slice(KEEP_DAYS).forEach(function (x) { const f = file_('minerva-dialler-snapshot-' + x.day + '.json', false); if (f) f.setTrashed(true); });
}

// Backups written by version 2 of this script lived inside the "Backup" sheet.
function oldSheetBackup_() {
  const b = ss_().getSheetByName('Backup');
  if (!b) return null;
  const k = Number(b.getRange('B4').getValue()) || 0;
  if (!k) return null;
  const parts = b.getRange(7, 1, k, 1).getValues().map(function (r) { return r[0]; });
  return { json: parts.join(''), leads: Number(b.getRange('B2').getValue()) || 0, calls: Number(b.getRange('B3').getValue()) || 0 };
}

function maybeReadable_(st, force) {
  const p = props_();
  const last = Number(p.getProperty('readableAt')) || 0;
  if (!force && Date.now() - last < READABLE_EVERY_MS) return;
  readable_(st);
  p.setProperties({ readableAt: String(Date.now()), dirty: '0' });
}

// Called on a quick check: if the readable tabs are behind, refresh them when nothing else is running.
function tidyReadable_() {
  const p = props_();
  if (p.getProperty('dirty') !== '1') return;
  if (Date.now() - (Number(p.getProperty('readableAt')) || 0) < READABLE_EVERY_MS) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) return;
  try {
    const f = file_('minerva-dialler-backup.json', false);
    if (f) maybeReadable_(JSON.parse(f.getBlob().getDataAsString()), true);
  } finally { lock.releaseLock(); }
}

function fit_(s, rows, cols) {
  if (s.getMaxRows() < rows) s.insertRowsAfter(s.getMaxRows(), rows - s.getMaxRows());
  if (s.getMaxColumns() < cols) s.insertColumnsAfter(s.getMaxColumns(), cols - s.getMaxColumns());
}
function sh_(name) { const s = ss_(); return s.getSheetByName(name) || s.insertSheet(name); }

// Readable copies for you and your manager. They are rebuilt automatically, so edits here are overwritten.
function readable_(st) {
  const cols = st.cols || [], rows = st.rows || [];
  const role = function (x) { const c = cols.filter(function (c) { return c.role === x; })[0]; return c ? c.n : ''; };
  const co = role('company'), ct = role('contact'), ph = role('phone');
  const v = function (r, n) { return n && r.f && r.f[n] != null ? String(r.f[n]) : ''; };
  const title = function (r) { return v(r, co) || v(r, ct) || '(blank lead)'; };
  const extra = cols.filter(function (c) { return c.show && ['company', 'contact', 'phone'].indexOf(c.role) < 0; })
    .map(function (c) { return c.n; });
  const log = [];
  rows.forEach(function (r) {
    (r.log || []).forEach(function (l) {
      log.push([new Date(l.t), title(r), v(r, ct), v(r, ph), l.o, l.note || ''].concat(extra.map(function (n) { return v(r, n); })));
    });
  });
  log.sort(function (a, b) { return b[0] - a[0]; });
  write_('Call Log', ['Date & time', 'Company', 'Contact', 'Phone', 'Outcome', 'What was discussed'].concat(extra), log, [1], [140, 220, 160, 130, 110, 420]);
  const leadRows = rows.map(function (r) {
    const lg = r.log || [], l = lg[lg.length - 1] || {};
    return cols.map(function (c) { return v(r, c.n); }).concat([
      r.done ? 'Logged' : 'To call', String(lg.length), l.o || '', l.t ? new Date(l.t) : '', r.fu ? new Date(r.fu) : '',
      lg.map(function (x) { return x.note; }).filter(Boolean).join(' | ')
    ]);
  });
  const nc = cols.length;
  write_('Leads', cols.map(function (c) { return c.n; }).concat(['Status', 'Calls', 'Last outcome', 'Last contact', 'Follow-up', 'All notes']),
    leadRows, [nc + 4, nc + 5], null);
  const s = ss_();
  const blank = s.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && s.getSheets().length > 1) s.deleteSheet(blank);
}

function write_(name, head, data, dateCols, widths) {
  const s = sh_(name);
  s.clear();
  const w = head.length, n = data.length;
  fit_(s, n + 1, w);
  s.getRange(1, 1, 1, w).setValues([head]).setFontWeight('bold').setFontColor('#ffffff').setBackground('#004679');
  s.setFrozenRows(1);
  if (n) {
    const rg = s.getRange(2, 1, n, w);
    rg.setNumberFormat('@');  // keep phone numbers like +971... as text
    dateCols.forEach(function (c) { s.getRange(2, c, n, 1).setNumberFormat('dd mmm yyyy hh:mm'); });
    rg.setValues(data.map(function (r) { r = r.slice(0, w); while (r.length < w) r.push(''); return r; }))
      .setVerticalAlignment('top');
  }
  if (widths) widths.forEach(function (px, i) { if (i < w) s.setColumnWidth(i + 1, px); });
}
