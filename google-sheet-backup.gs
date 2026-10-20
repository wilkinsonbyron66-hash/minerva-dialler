/* Minerva Dialler - Google Sheet backup
   Paste into the Google Sheet: Extensions > Apps Script (replace everything), Save,
   then Deploy > New deployment > Web app: Execute as "Me", Who has access "Anyone".
   The KEY must match the key shown in the dialler's Cloud backup panel. */
const KEY = '__KEY__';
const CHUNK = 40000;   // characters per cell (Sheets limit is 50,000)
const KEEP_DAYS = 30;  // daily snapshots kept

function doPost(e) {
  let res;
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    const req = JSON.parse(e.postData.contents);
    if (req.key !== KEY) res = { ok: false, error: 'key' };
    else if (req.action === 'ping') res = Object.assign({ ok: true }, info_());
    else if (req.action === 'save') res = save_(req);
    else if (req.action === 'restore') res = restore_(req.day);
    else if (req.action === 'snapshots') res = { ok: true, list: snapList_() };
    else res = { ok: false, error: 'action' };
  } catch (err) {
    res = { ok: false, error: String(err) };
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput('Minerva Dialler backup is running.');
}

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sh_(name) { const s = ss_(); return s.getSheetByName(name) || s.insertSheet(name); }
function fit_(s, rows, cols) {
  if (s.getMaxRows() < rows) s.insertRowsAfter(s.getMaxRows(), rows - s.getMaxRows());
  if (s.getMaxColumns() < cols) s.insertColumnsAfter(s.getMaxColumns(), cols - s.getMaxColumns());
}
function chunks_(json) { const a = []; for (let i = 0; i < json.length; i += CHUNK) a.push(json.slice(i, i + CHUNK)); return a; }

function info_() {
  const b = sh_('Backup');
  const v = b.getRange('B1:B3').getValues();
  return {
    url: ss_().getUrl(),
    saved: v[0][0] instanceof Date ? v[0][0].getTime() : null,
    leads: Number(v[1][0]) || 0,
    calls: Number(v[2][0]) || 0
  };
}

function save_(req) {
  const st = req.state || {};
  const rows = st.rows || [];
  const leads = rows.length;
  const calls = rows.reduce(function (n, r) { return n + ((r.log || []).length); }, 0);
  const prev = info_();
  // Never let an empty or much smaller copy wipe a good backup unless the user chose to overwrite.
  if (!req.force) {
    if (leads === 0 && prev.leads > 0) return { ok: false, error: 'empty', prev: prev };
    if (prev.calls >= 20 && calls < prev.calls * 0.5) return { ok: false, error: 'shrink', prev: prev };
  }
  const now = new Date();
  const parts = chunks_(JSON.stringify(st));
  const b = sh_('Backup');
  b.clear();
  b.getRange(1, 1, 5, 2).setValues([
    ['Last saved', now], ['Leads', leads], ['Calls logged', calls], ['Parts', parts.length],
    ["Do not edit this sheet. It is the dialler's full backup.", '']
  ]);
  fit_(b, 7 + parts.length, 2);
  if (parts.length) b.getRange(7, 1, parts.length, 1).setNumberFormat('@').setValues(parts.map(function (p) { return [p]; }));
  snapshot_(parts, now, leads, calls);
  readable_(st);
  const blank = ss_().getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss_().getSheets().length > 1) ss_().deleteSheet(blank);
  return { ok: true, saved: now.getTime(), leads: leads, calls: calls, url: ss_().getUrl() };
}

function snapshot_(parts, now, leads, calls) {
  const s = sh_('Snapshots');
  const day = Utilities.formatDate(now, ss_().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  if (s.getLastRow() === 0) s.appendRow(['Day', 'Saved', 'Leads', 'Calls', 'Parts', 'Data']);
  const n = s.getLastRow();
  let row = 0;
  if (n > 1) {
    const days = s.getRange(2, 1, n - 1, 1).getDisplayValues().map(function (r) { return r[0]; });
    const i = days.indexOf(day);
    if (i >= 0) row = i + 2;
  }
  if (!row) row = n + 1;
  const need = 5 + parts.length;
  fit_(s, row, need);
  if (row <= s.getLastRow()) s.getRange(row, 1, 1, s.getMaxColumns()).clearContent();
  s.getRange(row, 1).setNumberFormat('@');
  if (parts.length) s.getRange(row, 6, 1, parts.length).setNumberFormat('@');
  s.getRange(row, 1, 1, need).setValues([[day, now, leads, calls, parts.length].concat(parts)]);
  const total = s.getLastRow() - 1;
  if (total > KEEP_DAYS) s.deleteRows(2, total - KEEP_DAYS);
}

function restore_(day) {
  let parts;
  if (day) {
    const s = sh_('Snapshots');
    const n = s.getLastRow();
    if (n < 2) return { ok: false, error: 'none' };
    const days = s.getRange(2, 1, n - 1, 1).getDisplayValues().map(function (r) { return r[0]; });
    const i = days.indexOf(day);
    if (i < 0) return { ok: false, error: 'none' };
    const k = Number(s.getRange(i + 2, 5).getValue()) || 0;
    if (!k) return { ok: false, error: 'none' };
    parts = s.getRange(i + 2, 6, 1, k).getValues()[0];
  } else {
    const b = sh_('Backup');
    const k = Number(b.getRange('B4').getValue()) || 0;
    if (!k) return { ok: false, error: 'none' };
    parts = b.getRange(7, 1, k, 1).getValues().map(function (r) { return r[0]; });
  }
  return { ok: true, state: JSON.parse(parts.join('')) };
}

function snapList_() {
  const s = sh_('Snapshots');
  const n = s.getLastRow();
  if (n < 2) return [];
  return s.getRange(2, 1, n - 1, 4).getDisplayValues()
    .map(function (r) { return { day: r[0], saved: r[1], leads: r[2], calls: r[3] }; }).reverse();
}

// Readable copies for you and your manager. They are rebuilt on every backup, so edits here are overwritten.
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
  write_('Call Log', ['Date & time', 'Company', 'Contact', 'Phone', 'Outcome', 'What was discussed'].concat(extra), log, [1], [60, 220, 160, 130, 110, 420]);

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
  ['Call Log', 'Leads'].forEach(function (n, i) { s.setActiveSheet(s.getSheetByName(n)); s.moveActiveSheet(i + 1); });
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
  if (widths) widths.forEach(function (px, i) { if (i < w) s.setColumnWidth(i + 1, i === 0 ? 140 : px); });
}
