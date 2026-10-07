// BRC KAVACH Dashboard — Google Apps Script Backend v3
const SHEET_ID = '13vvSnm3KO3-rQ9iGvoFOTYSyxrvY_WSJYJKFDUDIwQI';

// Column indices (0-indexed).
// Column indices — leading columns match the sheet layout agreed with Railway:
// A  B     C        D      E     F           G    H            I     J       K             L        M     N     O           P
// SR Date  Section  Train  Loco  Train Name  Fit  LOCO OEM     UP/DN Reason  Failure Type  Station  Gear  Desc  NMS Engr    Verified By
const COL = {
  SR: 0, DATE: 1, SECTION: 2, TRAIN: 3, LOCO: 4, TRAIN_NAME: 5,
  FIT: 6, OEM_COMPANY: 7, UPDN: 8, REASON: 9, REMARK: 10, FAILURE_TYPE: 11, STATION: 12,
  GEAR: 13, DESC: 14, OEM_STAFF: 15, RLYSTAFF: 16,
  VERIFIED_DATE: 17, STATUS: 18, FLAG: 19, FLAG_NOTE: 20,
  MODE_DEG: 21, EB: 22, OP_AVAIL: 23,
  LOCO_FAULTS: 24, WRONG_OP: 25, SOC: 26, EXEC: 27,
  FOREIGN_TAG: 28, INVALID_SIG: 29, TAG_MISS: 30, THREE_TAG: 31,
  LOCO_TYPE: 32, LOCO_SHED: 33, LOCO_ZONE: 34, LOCO_MAKE: 35,
  RD_REMARK: 36, LOCO_LOG_STATUS: 37, RECT_NOTE: 38,
  LABEL: 39, MEMO_STATUS: 40, OPR_STATUS: 41, WRONG_HIGHLIGHT: 42
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('BRC KAVACH Failure Monitoring')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ── Sheet helpers ────────────────────────────────────
const MONTH_NAMES = ['JANUARY','FEBRUARY','MARCH','APRIL','MAY','JUNE',
                     'JULY','AUGUST','SEPTEMBER','OCTOBER','NOVEMBER','DECEMBER'];

const MAIN_HEADERS = [
  // A–P: Railway-agreed leading columns
  'SR No','Date','Section','Train No','Loco No','Train Name',
  'Loco Fit/Unfit','LOCO OEM','UP/DN','Reason','Remark','Failure Type','Station',
  'Gear At Fault','Failure Description','NMS Engineer','Verified By',
  // Q onwards: internal tracking columns
  'Verified Date','Status','Flag','Flag Note',
  'Mode Deg','Undue Braking','Op Avail',
  'Loco Faults','Wrong Op','SOC','Exec',
  'Foreign Tag','Invalid Sig','Tag Miss','3-Tag',
  'Loco Type','Loco Shed','Loco Zone','KAVACH Make',
  'R&D Remark','Loco Log Status','Rectification Note','Label',
  'Memo Status','OPR Status','Flagged Sub-Obs'
];

function getCurrentMonthSheetName() {
  const now = new Date();
  return MONTH_NAMES[now.getMonth()] + '-' + now.getFullYear();
}

// Sheet names that are NOT observation month sheets
const _OBS_NON_MONTH = new Set([
  'STAFF','CONFIG','LOCO_CACHE','Loco details','Train_cached',
  'RAILWAY_STAFF','LOCO_MEMO','OPR_RECORDS','TRAIN_MISSIONS',
  'GENERATION_QUEUE','RSSI_DATA'
]);

// ── Dynamic Section Map ──────────────────────────────────
// Stored in CONFIG sheet as: section_map | <JSON>
// Legacy keys obs_sheet_brc_gda / obs_sheet_brc_urn / obs_sheet_bjw_adi
// are read on first run to auto-populate the map.

function _getSectionMap() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('section_map_v2');
  if (hit) return JSON.parse(hit);

  const sh   = getOrCreateSheet('CONFIG');
  const rows = sh.getDataRange().getValues();
  let mapJson = null;
  const legacyIds = { brc_gda: '', brc_urn: '', bjw_adi: '' };

  rows.slice(1).forEach(function(r) {
    const k = String(r[0]||'').trim();
    const v = String(r[1]||'').trim();
    if (k === 'section_map') mapJson = v;
    if (k === 'obs_sheet_brc_gda') legacyIds.brc_gda = v;
    if (k === 'obs_sheet_brc_urn') legacyIds.brc_urn = v;
    if (k === 'obs_sheet_bjw_adi') legacyIds.bjw_adi = v;
  });

  let map;
  if (mapJson) {
    try { map = JSON.parse(mapJson); } catch(e) { map = null; }
  }
  if (!map) {
    map = [
      { group: 'BRC–GDA', company: 'Medha', sections: ['BRC - GDA','GDA - BRC','GDA - CYI (Toward BJW-ADI)'],
        updn: {'BRC - GDA':'DN','GDA - BRC':'UP','GDA - CYI (Toward BJW-ADI)':'UP'},
        sheetIds: legacyIds.brc_gda ? [legacyIds.brc_gda] : [] },
      { group: 'VS–URN',  company: 'Medha', sections: ['VS - URN','URN - VS','PRTN - BRC (DN)','BRC - PRTN (UP)'],
        updn: {'VS - URN':'UP','URN - VS':'DN','PRTN - BRC (DN)':'DN','BRC - PRTN (UP)':'UP'},
        sheetIds: legacyIds.brc_urn ? [legacyIds.brc_urn] : [] },
      { group: 'BJW–ADI', company: 'HBL', sections: ['BJW - ADI','ADI - BJW','ADI - BJW (Toward CYI-GDA)','BJW - ADI (From CYI-GDA)'],
        updn: {'BJW - ADI':'DN','ADI - BJW':'UP','ADI - BJW (Toward CYI-GDA)':'DN','BJW - ADI (From CYI-GDA)':'DN'},
        sheetIds: legacyIds.bjw_adi ? [legacyIds.bjw_adi] : [] }
    ];
  }
  // Back-fill updn/company for maps saved before these fields existed
  map.forEach(function(g) {
    if (!g.updn || typeof g.updn !== 'object') g.updn = {};
    if (!g.company) g.company = '';
  });
  cache.put('section_map_v2', JSON.stringify(map), 60);
  return map;
}

// Merges each group's { sectionName: 'UP'|'DN' } into one flat lookup map.
function getSectionUpDnMap() {
  const out = {};
  _getSectionMap().forEach(function(g) {
    Object.keys(g.updn || {}).forEach(function(s) { out[s] = g.updn[s]; });
  });
  return out;
}

// Merges section_map's per-group `company` field into { company: [sections...] },
// falling back to the hardcoded OEM_ALLOWED_SECTIONS for any group that hasn't
// been assigned a company yet (keeps old data working during migration).
function _getOemAllowedSectionsDynamic() {
  const out = {};
  _getSectionMap().forEach(function(g) {
    if (!g.company) return;
    if (!out[g.company]) out[g.company] = [];
    out[g.company] = out[g.company].concat(g.sections || []);
  });
  Object.keys(OEM_ALLOWED_SECTIONS).forEach(function(co) {
    if (!out[co]) out[co] = OEM_ALLOWED_SECTIONS[co];
  });
  return out;
}

function getSectionMapConfig() {
  return _getSectionMap();
}

function getSectionList() {
  return _getSectionMap().reduce(function(acc, g) { return acc.concat(g.sections); }, []);
}

function saveSectionMapConfig(data) {
  try {
    const auth = validateRailwayStaffPin(data.staffName, data.pin);
    if (!auth.valid) return { success: false, error: 'Invalid staff name or PIN' };
    const map = data.map;
    if (!Array.isArray(map)) return { success: false, error: 'Invalid data format' };
    for (var i = 0; i < map.length; i++) {
      const g = map[i];
      if (!g.group || !Array.isArray(g.sections) || !Array.isArray(g.sheetIds))
        return { success: false, error: 'Group "' + (g.group||'?') + '" is missing required fields' };
      if (!g.updn || typeof g.updn !== 'object') g.updn = {};
      if (!g.company) g.company = '';
    }
    const sh   = getOrCreateSheet('CONFIG');
    const rows = sh.getDataRange().getValues();
    let found  = false;
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]||'').trim() === 'section_map') {
        sh.getRange(i + 1, 2).setValue(JSON.stringify(map));
        found = true; break;
      }
    }
    if (!found) sh.appendRow(['section_map', JSON.stringify(map)]);
    CacheService.getScriptCache().remove('section_map_v2');
    SpreadsheetApp.flush();
    _bumpStatsCache();
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// Returns SpreadsheetApp for a section (uses first sheetId in matching group)
function _getObsSS(section) {
  const secStr = String(section||'').trim();
  const map = _getSectionMap();
  for (var i = 0; i < map.length; i++) {
    if (map[i].sections.indexOf(secStr) >= 0 && map[i].sheetIds.length)
      return SpreadsheetApp.openById(map[i].sheetIds[0]);
  }
  return SpreadsheetApp.openById(SHEET_ID);
}

// Returns array of unique SpreadsheetApp across all groups (all sheetIds)
function _getAllObsSS() {
  const map  = _getSectionMap();
  const seen = new Set(), list = [];
  map.forEach(function(g) {
    g.sheetIds.forEach(function(id) {
      if (id && !seen.has(id)) {
        seen.add(id);
        try { list.push(SpreadsheetApp.openById(id)); } catch(e) {}
      }
    });
  });
  if (!list.length) list.push(SpreadsheetApp.openById(SHEET_ID));
  return list;
}

// Gets or creates current-month sheet in the correct obs spreadsheet
function getObsSheet(section) {
  const ss   = _getObsSS(section);
  const name = getCurrentMonthSheetName();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(MAIN_HEADERS);
    sh.getRange(1, 1, 1, MAIN_HEADERS.length).setFontWeight('bold').setBackground('#1A2540').setFontColor('#7B9EFF');
    sh.setFrozenRows(1);
    _polishObsSheetLayout(sh);
  }
  return sh;
}

// One-time cosmetic layout applied when a month sheet is first created (and
// re-applied retroactively by beautifyAllObsSheets()): sensible column widths,
// freeze header row + SR/Date columns, a taller/larger header row, and a
// uniform readable font for the data rows below it.
function _polishObsSheetLayout(sh) {
  try {
    // Header row — bold, bigger, colored, frozen (independent of the body font below)
    sh.getRange(1, 1, 1, MAIN_HEADERS.length)
      .setFontWeight('bold').setBackground('#1A2540').setFontColor('#7B9EFF')
      .setFontFamily('Calibri').setFontSize(11).setVerticalAlignment('middle').setHorizontalAlignment('center');
    sh.setRowHeight(1, 34);
    sh.setFrozenRows(1);
    sh.setFrozenColumns(2);

    // Body rows — smaller, plain font
    const lastRow = sh.getMaxRows();
    if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, MAIN_HEADERS.length).setFontFamily('Calibri').setFontSize(10);

    const widths = {
      [COL.SR]: 55, [COL.DATE]: 85, [COL.SECTION]: 130, [COL.TRAIN]: 75, [COL.LOCO]: 75,
      [COL.TRAIN_NAME]: 110, [COL.FIT]: 75, [COL.OEM_COMPANY]: 90, [COL.UPDN]: 60,
      [COL.REASON]: 140, [COL.FAILURE_TYPE]: 130, [COL.STATION]: 110, [COL.GEAR]: 120,
      [COL.DESC]: 260, [COL.OEM_STAFF]: 120, [COL.RLYSTAFF]: 120, [COL.FLAG_NOTE]: 180,
      [COL.RD_REMARK]: 180, [COL.RECT_NOTE]: 180
    };
    Object.keys(widths).forEach(function(c){ sh.setColumnWidth(Number(c) + 1, widths[c]); });
  } catch(e) { /* best-effort */ }
}

// Returns all obs month sheets across all configured obs spreadsheets
function _getAllObsMonthSheets() {
  const sheets = [];
  _getAllObsSS().forEach(ss => {
    ss.getSheets().forEach(sh => {
      if (!_OBS_NON_MONTH.has(sh.getName()) && /^[A-Z]+-\d{4}$/.test(sh.getName()))
        sheets.push(sh);
    });
  });
  return sheets;
}

// Searches all obs sheets for an SR number; returns {sh, rowIndex} or null
function _findObsRowBySr(srNo, section) {
  const srStr  = String(srNo).trim();
  const secStr = section ? String(section).trim().toLowerCase() : '';
  for (const ss of _getAllObsSS()) {
    for (const sh of ss.getSheets()) {
      if (_OBS_NON_MONTH.has(sh.getName()) || !/^[A-Z]+-\d{4}$/.test(sh.getName())) continue;
      const rows = sh.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][COL.SR]).trim() !== srStr) continue;
        // If section provided, must match — prevents wrong-sheet hits when SR numbers collide
        if (secStr && String(rows[i][COL.SECTION] || '').trim().toLowerCase() !== secStr) continue;
        return { sh, rowIndex: i + 1 };
      }
    }
  }
  // Fallback: match by SR only (handles old calls without section)
  if (secStr) return _findObsRowBySr(srNo, '');
  return null;
}

// ── ICMS Sync (KAVACH portal cross-check) ────────────────────────────
// The KAVACH failure-report portal (10.3.2.55, internal RailNet only) can't be
// reached by this Apps Script backend directly (private IP, not internet-routable),
// so ICMS cases arrive via doPost from a local script (kavach_sync.py) that logs
// into KAVACH itself and posts its CSV export here, or via manual paste in the
// Admin Panel. Either path lands in syncIcmsCases().

function _findObsRowsByTrainDate(trainNo, dateStr) {
  var out = [];
  trainNo = String(trainNo || '').trim();
  if (!trainNo || !dateStr) return out;
  var target = new Date(String(dateStr) + 'T00:00:00');
  if (isNaN(target)) return out;
  var targetKey = Utilities.formatDate(target, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  _getAllObsSS().forEach(function(ss) {
    ss.getSheets().forEach(function(sh) {
      if (_OBS_NON_MONTH.has(sh.getName()) || !/^[A-Z]+-\d{4}$/.test(sh.getName())) return;
      var rows = sh.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        var r = rows[i];
        if (String(r[COL.TRAIN] || '').trim() !== trainNo) continue;
        var rd = r[COL.DATE] instanceof Date ? r[COL.DATE] : new Date(String(r[COL.DATE] || ''));
        if (isNaN(rd)) continue;
        if (Utilities.formatDate(rd, Session.getScriptTimeZone(), 'yyyy-MM-dd') !== targetKey) continue;
        out.push({ sh: sh, rowIndex: i + 1, sr: r[COL.SR], section: r[COL.SECTION] });
      }
    });
  });
  return out;
}

// Additive, mirrors addObsFlag()'s combine-onto-existing-flags pattern
function _icmsAddFlag(sh, row, note) {
  var curFlag = String(sh.getRange(row, COL.FLAG + 1).getValue() || '').trim();
  var flags = (curFlag && curFlag !== 'None') ? curFlag.split(',').map(function(s){ return s.trim(); }).filter(Boolean) : [];
  if (flags.indexOf('ICMS Flagged') < 0) flags.push('ICMS Flagged');
  sh.getRange(row, COL.FLAG + 1).setValue(flags.join(', '));
  if (note) {
    var stamp = '[ICMS Sync — ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yy HH:mm') + '] ' + note;
    var curNote = String(sh.getRange(row, COL.FLAG_NOTE + 1).getValue() || '').trim();
    sh.getRange(row, COL.FLAG_NOTE + 1).setValue(curNote ? curNote + '\n' + stamp : stamp);
  }
  _applyObsRowHighlight(sh, row);
}

// Train/Loco numbers aren't in clean CSV columns on the KAVACH side — they only
// show up inside free-text remark/detail fields (e.g. "Detention of train 22939...",
// "LOCO -39220,WAP-7", "Loco ID 37830"), so we regex them out of the combined text.
function _extractTrainNo(text) {
  var m = String(text || '').match(/train\s*(?:no\.?)?\s*[:\-]?\s*(\d{3,6})/i);
  return m ? m[1] : '';
}
function _extractLocoNo(text) {
  var m = String(text || '').match(/loco\s*(?:no\.?|id)?\s*[:\-]?\s*(\d{3,6})/i);
  return m ? m[1] : '';
}

// Parses a KAVACH failure_report.php CSV export (RFC4180-style; Utilities.parseCsv
// handles the embedded quoted/multi-line fields correctly).
function _parseIcmsCsv(csvText) {
  if (!csvText) return [];
  var table = Utilities.parseCsv(csvText);
  if (!table.length) return [];
  var idx = {};
  table[0].forEach(function(h, i) { idx[String(h || '').trim()] = i; });
  function col(row, name) { return idx[name] !== undefined ? String(row[idx[name]] || '').trim() : ''; }

  var out = [];
  for (var r = 1; r < table.length; r++) {
    var row = table[r];
    if (!row || !row.length || !col(row, 'Sr No')) continue;
    var remark  = col(row, 'PARENTSHED (Remark)');
    var details = col(row, 'Failure Details');
    var cause   = col(row, 'Cause');
    var combined = remark + ' ' + details + ' ' + cause;
    out.push({
      srNo: col(row, 'Sr No'),
      failureDate: col(row, 'Failure Date'),
      failureTime: col(row, 'Failure Time'),
      stationCode: col(row, 'Station Code'),
      stationName: col(row, 'Station Name'),
      section: col(row, 'Section Name'),
      primaryFault: col(row, 'Primary Fault'),
      gearAtFault: col(row, 'Gear At Fault'),
      cause: cause,
      failureDetails: details,
      trainNo: _extractTrainNo(combined),
      locoNo: col(row, 'Loco No') || _extractLocoNo(combined)
    });
  }
  return out;
}

function _getIcmsSyncKey() {
  var sh = getOrCreateSheet('CONFIG');
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === 'icms_sync_key') return String(rows[i][1] || '').trim();
  }
  return '';
}

// Admin Panel: view (or lazily create) the shared secret kavach_sync.py must send
function getIcmsSyncKey(pin) {
  if (!validateAdmin(pin)) return { success: false, error: 'Wrong PIN' };
  var key = _getIcmsSyncKey();
  if (key) return { success: true, key: key };
  return regenerateIcmsSyncKey(pin);
}

function regenerateIcmsSyncKey(pin) {
  if (!validateAdmin(pin)) return { success: false, error: 'Wrong PIN' };
  var key = Utilities.getUuid().replace(/-/g, '');
  var sh = getOrCreateSheet('CONFIG');
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === 'icms_sync_key') {
      sh.getRange(i + 1, 2).setValue(key);
      return { success: true, key: key };
    }
  }
  sh.appendRow(['icms_sync_key', key]);
  return { success: true, key: key };
}

// payload: { key, entries: [{ division:'BRC'|'BRC-N', csv:'<raw CSV export text>' }, ...] }
// Upserts every case into ICMS_CASES (keyed by division+srNo so re-syncing the
// same case updates it in place instead of duplicating), and — on a Train+Date
// match against the observation sheets — additively sets the "ICMS Flagged" flag,
// the same combinable flag the Journey Summary / exclude system already treats
// like R&D Flagged / Loco Log Flagged.
function syncIcmsCases(payload) {
  try {
    var key = String((payload && payload.key) || '').trim();
    var savedKey = _getIcmsSyncKey();
    if (!savedKey || key !== savedKey) return { success: false, error: 'Invalid sync key' };

    var entries = (payload && Array.isArray(payload.entries)) ? payload.entries : [];
    var allCases = [];
    entries.forEach(function(e) {
      var division = String(e.division || '').trim();
      _parseIcmsCsv(String(e.csv || '')).forEach(function(c) {
        c.division = division;
        allCases.push(c);
      });
    });

    var sh = getOrCreateSheet('ICMS_CASES');
    var existing = sh.getDataRange().getValues();
    var keyToRow = {};
    for (var i = 1; i < existing.length; i++) {
      keyToRow[String(existing[i][0]) + '|' + String(existing[i][1])] = i + 1;
    }

    var matched = 0, unmatched = 0, created = 0, updated = 0;
    var now = new Date();

    allCases.forEach(function(c) {
      var found = _findObsRowsByTrainDate(c.trainNo, c.failureDate);
      var matchStatus, matchedRef;
      if (found.length) {
        matched++;
        matchStatus = 'Matched';
        matchedRef = found.map(function(f) { return f.section + ' SR#' + f.sr; }).join('; ');
        found.forEach(function(f) {
          _icmsAddFlag(f.sh, f.rowIndex, 'KAVACH ' + c.division + ' Sr#' + c.srNo + ' — ' + (c.primaryFault || c.cause || 'ICMS case'));
        });
      } else {
        unmatched++;
        matchStatus = 'Unmatched';
        matchedRef = '';
      }

      var rowKey = c.division + '|' + c.srNo;
      var rowData = [c.division, c.srNo, c.failureDate, c.failureTime, c.stationCode, c.stationName,
                      c.section, c.trainNo, c.locoNo, c.primaryFault, c.gearAtFault, c.cause,
                      c.failureDetails, matchStatus, matchedRef, now];
      if (keyToRow[rowKey]) {
        sh.getRange(keyToRow[rowKey], 1, 1, rowData.length).setValues([rowData]);
        updated++;
      } else {
        sh.appendRow(rowData);
        created++;
      }
    });

    SpreadsheetApp.flush();
    return { success: true, total: allCases.length, matched: matched, unmatched: unmatched, created: created, updated: updated };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// In-dashboard manual fallback (Admin PIN instead of the external sync key) —
// used by the ICMS tab's "paste CSV" box when kavach_sync.py isn't available.
function syncIcmsCasesManual(pin, division, csvText) {
  if (!validateAdmin(pin)) return { success: false, error: 'Wrong PIN' };
  return syncIcmsCases({ key: _getIcmsSyncKey(), entries: [{ division: division, csv: csvText }] });
}

function getIcmsCases() {
  try {
    var sh = getOrCreateSheet('ICMS_CASES');
    var rows = sh.getDataRange().getValues();
    var fields = ['division','srNo','failureDate','failureTime','stationCode','stationName','section',
                  'trainNo','locoNo','primaryFault','gearAtFault','cause','failureDetails','matchStatus','matchedRef','syncedAt'];
    var out = [];
    for (var i = 1; i < rows.length; i++) {
      var obj = {};
      fields.forEach(function(f, ci) { obj[f] = rows[i][ci]; });
      out.push(obj);
    }
    out.sort(function(a, b) { return new Date(b.failureDate) - new Date(a.failureDate); });
    return { success: true, cases: out };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// External entry point for kavach_sync.py — POST { action:'icmsSync', key, entries:[...] }
function doPost(e) {
  try {
    var body = {};
    try { body = JSON.parse(e.postData.contents); } catch (parseErr) {}
    var result;
    if (body.action === 'icmsSync') {
      result = syncIcmsCases(body);
    } else if (body.action === 'hblLocoSync') {
      result = registerHblLocoEntries(body);
    } else {
      result = { success: false, error: 'Unknown action' };
    }
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Admin: get current obs sheet IDs
function getObsSheetConfig() {
  return { success: true, ids: _getObsSheetIds() };
}

// Admin: connect + test a single obs sheet (save ID, create month tab, dummy write/delete)
function connectObsSheet(sheetKey, sheetId) {
  try {
    sheetId = String(sheetId || '').trim();
    if (!sheetId) return { success: false, error: 'No Sheet ID provided.' };

    // 1. Open the spreadsheet — validates access & ID correctness
    let ss;
    try { ss = SpreadsheetApp.openById(sheetId); }
    catch(e) { return { success: false, error: 'Cannot open sheet — verify ID and that it is shared with this script\'s account. (' + e.message + ')' }; }

    // 2. Save the ID to CONFIG
    const cfgKeyMap = { brc_gda: 'obs_sheet_brc_gda', brc_urn: 'obs_sheet_brc_urn', bjw_adi: 'obs_sheet_bjw_adi' };
    const cfgKey = cfgKeyMap[sheetKey];
    if (!cfgKey) return { success: false, error: 'Unknown sheet key: ' + sheetKey };
    const cfgSh   = getOrCreateSheet('CONFIG');
    const cfgRows = cfgSh.getDataRange().getValues();
    const cfgKeys = cfgRows.map(r => String(r[0]||'').trim());
    const cfgIdx  = cfgKeys.indexOf(cfgKey);
    if (cfgIdx >= 0) cfgSh.getRange(cfgIdx + 1, 2).setValue(sheetId);
    else cfgSh.appendRow([cfgKey, sheetId]);
    CacheService.getScriptCache().remove('obs_sheet_ids');

    // 3. Create or verify current month tab
    const tabName = getCurrentMonthSheetName();
    let sh = ss.getSheetByName(tabName);
    let tabCreated = false;
    if (!sh) {
      sh = ss.insertSheet(tabName);
      sh.appendRow(MAIN_HEADERS);
      sh.getRange(1, 1, 1, MAIN_HEADERS.length).setFontWeight('bold').setBackground('#1A2540').setFontColor('#7B9EFF');
      sh.setFrozenRows(1);
      tabCreated = true;
    } else {
      // Extend headers if sheet predates new columns
      const lastCol = sh.getLastColumn();
      if (lastCol < MAIN_HEADERS.length) {
        const existing = lastCol > 0 ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
        for (let i = lastCol; i < MAIN_HEADERS.length; i++) {
          if (!existing[i]) sh.getRange(1, i + 1).setValue(MAIN_HEADERS[i]);
        }
      }
    }

    // 4. Dummy write + delete to confirm write access
    const dummy = Array(MAIN_HEADERS.length).fill('');
    dummy[0] = '_TEST_' + Date.now();
    sh.appendRow(dummy);
    sh.deleteRow(sh.getLastRow());

    return { success: true, tabName: tabName, ssName: ss.getName(), tabCreated: tabCreated };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// Admin: save obs sheet IDs
function saveObsSheetConfig(data) {
  try {
    const sh   = getOrCreateSheet('CONFIG');
    const rows = sh.getDataRange().getValues();
    const keys = rows.map(r => String(r[0]||'').trim());
    const map  = {
      obs_sheet_brc_gda: String(data.brc_gda||'').trim(),
      obs_sheet_brc_urn: String(data.brc_urn||'').trim(),
      obs_sheet_bjw_adi: String(data.bjw_adi||'').trim()
    };
    Object.entries(map).forEach(([k, v]) => {
      const idx = keys.indexOf(k);
      if (idx >= 0) sh.getRange(idx + 1, 2).setValue(v);
      else sh.appendRow([k, v]);
    });
    CacheService.getScriptCache().remove('obs_sheet_ids');
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ── UI Layout ─────────────────────────────────────────
function getUILayout() {
  try {
    const sh   = getOrCreateSheet('CONFIG');
    const rows = sh.getDataRange().getValues();
    let tabOrder = [], tabHidden = [], cardOrder = [], cardHidden = [];
    let sidebarTabs = [];
    rows.forEach(function(r) {
      const k = String(r[0]||'').trim();
      const v = String(r[1]||'').trim();
      if (k === 'UI_TAB_ORDER')    tabOrder    = v ? v.split(',') : [];
      if (k === 'UI_TAB_HIDDEN')   tabHidden   = v ? v.split(',') : [];
      if (k === 'UI_CARD_ORDER')   cardOrder   = v ? v.split(',') : [];
      if (k === 'UI_CARD_HIDDEN')  cardHidden  = v ? v.split(',') : [];
      if (k === 'UI_SIDEBAR_TABS') sidebarTabs = v ? v.split(',') : [];
    });
    return { success: true, tabOrder, tabHidden, cardOrder, cardHidden, sidebarTabs };
  } catch(e) { return { success: true, tabOrder:[], tabHidden:[], cardOrder:[], cardHidden:[], sidebarTabs:[] }; }
}

function saveUILayout(data) {
  try {
    const sh   = getOrCreateSheet('CONFIG');
    const rows = sh.getDataRange().getValues();
    const keys = rows.map(function(r){ return String(r[0]||'').trim(); });
    const map  = {
      UI_TAB_ORDER:    (data.tabOrder    || []).join(','),
      UI_TAB_HIDDEN:   (data.tabHidden   || []).join(','),
      UI_CARD_ORDER:   (data.cardOrder   || []).join(','),
      UI_CARD_HIDDEN:  (data.cardHidden  || []).join(','),
      UI_SIDEBAR_TABS: (data.sidebarTabs || []).join(',')
    };
    Object.entries(map).forEach(function([k, v]) {
      const idx = keys.indexOf(k);
      if (idx >= 0) sh.getRange(idx + 1, 2).setValue(v);
      else sh.appendRow([k, v]);
    });
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function getMainSheet() {
  const ss   = SpreadsheetApp.openById(SHEET_ID);
  const name = getCurrentMonthSheetName();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(MAIN_HEADERS);
    sh.getRange(1, 1, 1, MAIN_HEADERS.length).setFontWeight('bold').setBackground('#1A2540').setFontColor('#7B9EFF');
    sh.setFrozenRows(1);
  }
  return sh;
}

function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (name === 'STAFF') {
      sh.appendRow(['ID', 'Name', 'Company', 'Added Date', 'Active']);
      sh.getRange(1, 1, 1, 5).setFontWeight('bold');
    }
    if (name === 'RAILWAY_STAFF') {
      sh.appendRow(['ID', 'Name', 'Designation', 'PIN', 'Added Date', 'Active']);
      sh.getRange(1, 1, 1, 6).setFontWeight('bold');
    }
    if (name === 'CONFIG') {
      sh.appendRow(['Key', 'Value']);
      sh.appendRow(['admin_pin', '1234']);
      sh.appendRow(['dashboard_title', 'BRC KAVACH Failure Monitoring']);
      sh.appendRow(['default_section', 'ALL']);
      sh.appendRow(['show_stats', 'true']);
      sh.appendRow(['show_recent', 'true']);
      sh.appendRow(['oem_companies', JSON.stringify(['HBL', 'Medha'])]);
      sh.appendRow(['sections', JSON.stringify(['BRC-GDA','GDA-BRC','CYI-GDA','GDA-CYI'])]);
      sh.appendRow(['directions', JSON.stringify(['UP','DN'])]);
      sh.appendRow(['loco_statuses', JSON.stringify(['FIT','UNFIT'])]);
    }
    if (name === 'ICMS_CASES') {
      sh.appendRow(['Division','Sr No','Failure Date','Failure Time','Station Code','Station Name','Section','Train No','Loco No','Primary Fault','Gear At Fault','Cause','Failure Details','Match Status','Matched Ref','Synced At']);
      sh.getRange(1, 1, 1, 16).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

// ── Railway Staff (individual PIN per staff) ──────────
function getRailwayStaffList() {
  const sh = getOrCreateSheet('RAILWAY_STAFF');
  const last = sh.getLastRow();
  if (last < 2) return [];
  const rows = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  return rows
    .filter(r => r[0] && r[5] !== false && String(r[5]).toUpperCase() !== 'FALSE')
    .map(r => ({ id: String(r[0]), name: String(r[1]), designation: String(r[2] || ''), addedDate: String(r[4] || '') }));
  // PIN (r[3]) never exposed to frontend
}

function addRailwayStaff(data) {
  try {
    if (!data.name || !data.pin) return { success: false, error: 'Name and PIN required' };
    const sh = getOrCreateSheet('RAILWAY_STAFF');
    const id = Date.now().toString();
    sh.appendRow([id, data.name, data.designation || '',
      String(data.pin),
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yy'), true]);
    return { success: true, id };
  } catch (e) { return { success: false, error: e.message }; }
}

function getRailwayStaffPin(id) {
  try {
    const sh = getOrCreateSheet('RAILWAY_STAFF');
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id) && rows[i][5] !== false && rows[i][5] !== 'FALSE') {
        return { success: true, pin: String(rows[i][3]) };
      }
    }
    return { success: false, error: 'Staff not found' };
  } catch (e) { return { success: false, error: e.message }; }
}

function updateRailwayStaff(data) {
  try {
    const sh = getOrCreateSheet('RAILWAY_STAFF');
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        if (data.name)        sh.getRange(i + 1, 2).setValue(data.name);
        if (data.designation !== undefined) sh.getRange(i + 1, 3).setValue(data.designation);
        if (data.pin)         sh.getRange(i + 1, 4).setValue(String(data.pin));
        return { success: true };
      }
    }
    return { success: false, error: 'Staff not found' };
  } catch (e) { return { success: false, error: e.message }; }
}

function removeRailwayStaff(id) {
  try {
    const sh = getOrCreateSheet('RAILWAY_STAFF');
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] == id) { sh.getRange(i + 1, 6).setValue(false); break; }
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

function validateRailwayStaffPin(name, pin) {
  const sh = getOrCreateSheet('RAILWAY_STAFF');
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0] || r[5] === false || r[5] === 'FALSE') continue;
    if (String(r[1]).trim().toLowerCase() === String(name).trim().toLowerCase()
        && String(r[3]).trim() === String(pin).trim()) {
      return { valid: true, name: r[1], designation: r[2] };
    }
  }
  return { valid: false };
}

// ── Admin Auth ───────────────────────────────────────
function validateAdmin(pin) {
  const sh = getOrCreateSheet('CONFIG');
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === 'admin_pin') return rows[i][1] == pin;
  }
  return pin === '1234';
}

function changeAdminPin(oldPin, newPin) {
  if (!validateAdmin(oldPin)) return { success: false, error: 'Wrong current PIN' };
  const sh = getOrCreateSheet('CONFIG');
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === 'admin_pin') {
      sh.getRange(i + 1, 2).setValue(newPin);
      return { success: true };
    }
  }
  sh.appendRow(['admin_pin', newPin]);
  return { success: true };
}

// ── Staff Management ─────────────────────────────────
function getStaffList() {
  const sh = getOrCreateSheet('STAFF');
  const last = sh.getLastRow();
  if (last < 2) return [];
  const rows = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  return rows
    .filter(r => r[0] && r[4] !== false && String(r[4]).toUpperCase() !== 'FALSE')
    .map(r => ({ id: String(r[0]), name: String(r[1]), company: String(r[2]), addedDate: String(r[3] || '') }));
}


function addStaff(data) {
  try {
    const sh = getOrCreateSheet('STAFF');
    const id = Date.now().toString();
    sh.appendRow([id, data.name, data.company,
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yy'), true]);
    return { success: true, id };
  } catch (e) { return { success: false, error: e.message }; }
}

function updateStaff(data) {
  try {
    const sh = getOrCreateSheet('STAFF');
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        if (data.name)    sh.getRange(i + 1, 2).setValue(data.name);
        if (data.company) sh.getRange(i + 1, 3).setValue(data.company);
        return { success: true };
      }
    }
    return { success: false, error: 'Staff not found' };
  } catch (e) { return { success: false, error: e.message }; }
}

function removeStaff(id) {
  try {
    const sh = getOrCreateSheet('STAFF');
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] == id) { sh.getRange(i + 1, 5).setValue(false); break; }
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

// ── Config ───────────────────────────────────────────
function getAdminConfig() {
  const sh = getOrCreateSheet('CONFIG');
  const rows = sh.getDataRange().getValues();
  const cfg = {};
  rows.slice(1).forEach(r => { if (r[0]) cfg[r[0]] = r[1]; });
  delete cfg.admin_pin; // never expose PIN
  return cfg;
}

function updateAdminConfig(updates) {
  try {
    const sh = getOrCreateSheet('CONFIG');
    const rows = sh.getDataRange().getValues();
    const keys = rows.map(r => r[0]);
    Object.entries(updates).forEach(([k, v]) => {
      if (k === 'admin_pin') return; // can't update PIN here
      const idx = keys.indexOf(k);
      if (idx >= 0) sh.getRange(idx + 1, 2).setValue(v);
      else sh.appendRow([k, v]);
    });
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

// Get config lists (oem_companies, sections, directions, loco_statuses)
function getConfigLists() {
  try {
    const sh = getOrCreateSheet('CONFIG');
    const rows = sh.getDataRange().getValues();
    const listKeys = ['oem_companies','sections','directions','loco_statuses'];
    const out = {};
    rows.slice(1).forEach(r => {
      const key = String(r[0]||'').trim();
      if (listKeys.includes(key)) {
        try { out[key] = JSON.parse(r[1]); } catch(e) { out[key] = []; }
      }
    });
    listKeys.forEach(k => { if (!out[k]) out[k] = []; });
    return { success: true, ...out };
  } catch(e) { return { success: false, oem_companies:[], sections:[], directions:[], loco_statuses:[] }; }
}

function updateConfigList(key, items) {
  try {
    const allowed = ['oem_companies','sections','directions','loco_statuses'];
    if (!allowed.includes(key)) return { success: false, error: 'Invalid list key' };
    const sh = getOrCreateSheet('CONFIG');
    const rows = sh.getDataRange().getValues();
    const val = JSON.stringify(items);
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === key) { sh.getRange(i + 1, 2).setValue(val); return { success: true }; }
    }
    sh.appendRow([key, val]);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ── Ensure observation sheet has all headers ─────────
function ensureMainSheetHeaders(sheet) {
  if (!sheet) sheet = getMainSheet();
  const lastCol = sheet.getLastColumn();
  if (lastCol >= MAIN_HEADERS.length) return; // already full
  // Extend headers if sheet was created before new columns were added
  const headerRow = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  for (let i = lastCol; i < MAIN_HEADERS.length; i++) {
    if (!headerRow[i]) sheet.getRange(1, i + 1).setValue(MAIN_HEADERS[i]);
  }
}

// ── Submit Observation ───────────────────────────────
// data.isNil=true  → single NIL row
// data.failures[]  → one row per failure point (same train/loco/oem)
// OEM company → sections they are authorised to submit for
const OEM_ALLOWED_SECTIONS = {
  'HBL':   ['BJW - ADI', 'ADI - BJW', 'ADI - BJW (Toward CYI-GDA)', 'BJW - ADI (From CYI-GDA)'],
  'Medha': ['VS - URN',  'URN - VS',  'BRC - GDA', 'GDA - BRC', 'GDA - CYI (Toward BJW-ADI)', 'BRC - PRTN (UP)', 'PRTN - BRC (DN)']
};

// Finds the sheet row (1-based) after which a NEW row for `dateStr` should
// land, so month sheets stay date-ordered even when someone back-fills an
// older date after later ones are already entered — instead of it always
// landing at the physical bottom of the sheet. Walks upward from the bottom
// (robust even against a sheet that already has some out-of-order rows from
// before this fix existed — see fixUnsortedDateOrder() for repairing those):
// finds the last row whose date is <= the target date and returns it, so the
// new row goes right after it, clustered with any other same-date rows in
// their original entry order. Returns sheet.getLastRow() itself for the
// common case (today's date, nothing later already exists) — the caller can
// then just appendRow() instead of doing a real mid-sheet insert.
function _findDateOrderedInsertRow(sheet, dateStr) {
  const target = new Date(String(dateStr || '') + 'T00:00:00');
  const lastRow = sheet.getLastRow();
  if (isNaN(target)) return lastRow; // can't reason about an invalid date — just append
  if (lastRow < 2) return 1; // only the header exists — insert right after it
  const targetTime = target.getTime();
  const dateVals = sheet.getRange(2, COL.DATE + 1, lastRow - 1, 1).getValues();
  for (let i = dateVals.length - 1; i >= 0; i--) {
    const rd = dateVals[i][0];
    const rowDate = rd instanceof Date ? rd : new Date(String(rd || ''));
    if (!isNaN(rowDate) && rowDate.getTime() <= targetTime) return 2 + i;
  }
  return 1; // every existing row is dated AFTER the target — insert right after the header
}

// Writes `row` into the sheet at its date-ordered position (appending if that
// position is the current last row, inserting otherwise) and returns the
// 1-based row index it actually landed on, for _formatObsRow to target.
function _insertObsRowInDateOrder(sheet, row, dateStr) {
  const insertAfter = _findDateOrderedInsertRow(sheet, dateStr);
  if (insertAfter >= sheet.getLastRow()) {
    sheet.appendRow(row);
    return sheet.getLastRow();
  }
  sheet.insertRowAfter(insertAfter);
  const targetRow = insertAfter + 1;
  sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
  return targetRow;
}

function submitObservation(data) {
  try {
    // ── No future-dated observations ────────────────────────
    // The frontend already caps the date picker at today, but that's
    // bypassable (manual typing in some browsers, or calling this function
    // directly) — enforce it server-side too.
    if (data.date && String(data.date) > Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')) {
      return { success: false, error: 'Date cannot be in the future' };
    }
    // ── Section eligibility check ──────────────────────────
    const allowedForOEM = _getOemAllowedSectionsDynamic()[String(data.oemCompany||'').trim()];
    if (allowedForOEM && !allowedForOEM.includes(String(data.section||'').trim())) {
      return { success: false, error: (data.oemCompany||'This OEM') + ' staff are not authorised to submit observations for section: ' + data.section };
    }
    // ── Everything below reads the sheet's current row count (to assign the
    // new row's SR No) and/or its existing rows (to detect the same journey
    // already having a row) — this MUST be serialized. Without a lock, two
    // submissions close together in time (e.g. catching up on several
    // forgotten entries back-to-back) can both read the sheet before either
    // has appended, and end up assigning the SAME SR No to two different
    // rows — SR No is this app's unique row identifier, used everywhere
    // (verify/reject/flag/label all look a row up by it), so a collision
    // here silently breaks whichever action later targets that SR.
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {

    const sheet = getObsSheet(data.section);
    ensureMainSheetHeaders(sheet);

    function buildBaseRow(srNo) {
      const r = new Array(MAIN_HEADERS.length).fill('');
      r[COL.SR]          = srNo;
      r[COL.DATE]        = data.date;
      r[COL.TRAIN]       = data.trainNo;
      r[COL.UPDN]        = data.upDn;
      r[COL.LOCO]        = data.locoNo;
      r[COL.FIT]         = data.fitUnfit || 'FIT';
      r[COL.SECTION]     = data.section;
      // "LOCO OEM" is the loco's own registered KAVACH vendor (looked up from
      // its loco number, same source as LOCO_MAKE below) — NOT the submitting
      // staff's company, which just mirrors the section since staff are
      // section-locked to their own company already. Falls back to the
      // submitting company only if the loco isn't in the Loco Details sheet,
      // so this cell is never left blank.
      r[COL.OEM_COMPANY] = data.locoMake || data.oemCompany;
      r[COL.OEM_STAFF]   = data.oemStaff;
      r[COL.STATUS]      = 'Pending';
      r[COL.LOCO_TYPE]   = data.locoType || '';
      r[COL.LOCO_SHED]   = data.locoShed || '';
      r[COL.LOCO_ZONE]   = data.locoZone  || '';
      r[COL.LOCO_MAKE]   = data.locoMake  || '';
      r[COL.TRAIN_NAME]  = data.trainName || '';
      r[COL.LABEL]       = data.label || '';
      r[COL.MEMO_STATUS] = data.memoStatus || '';
      r[COL.OPR_STATUS]  = data.oprStatus  || '';
      // Free-text remark — independent of Observation Type. In NIL mode it's the
      // reason for no fault; in Fault Observed mode it's just a note attached to
      // the submission and is NOT itself counted as an observation.
      r[COL.REMARK]      = data.nilRemark  || '';
      return r;
    }

    // ── NIL report ──────────────────────────────────
    if (data.isNil) {
      const srNo = sheet.getLastRow();
      const row  = buildBaseRow(srNo);
      row[COL.GEAR]        = 'NIL';
      row[COL.REASON]      = 'NIL';
      row[COL.FAILURE_TYPE]= 'NIL';
      row[COL.STATION]     = 'NIL';
      row[COL.DESC]        = 'NIL - No Fault';
      row[COL.MODE_DEG]    = 'NIL';
      row[COL.EB]          = 'NIL';
      row[COL.FLAG]        = 'None';
      const nilRowIdx = _insertObsRowInDateOrder(sheet, row, data.date);
      _formatObsRow(sheet, nilRowIdx);
      _bumpStatsCache();
      return { success: true, count: 1, srNos: [srNo] };
    }

    // ── Multiple failure points → ONE row per journey ──────────────────────
    // All observations of a single train journey are aggregated into one row.
    // Multi-value cells hold numbered lists ("1. …\n2. …") like the official record sheet.
    const failures = data.failures || [];
    if (!failures.length) return { success: false, error: 'No failure points provided' };

    // Numbered list when >1 item, plain value when single
    const numList = arr => {
      const a = arr.filter(v => String(v).trim() !== '');
      if (!a.length) return '';
      return a.length === 1 ? a[0] : a.map((v,i) => (i+1) + '. ' + v).join('\n');
    };

    // If the same journey (train/loco/date/section/up-dn) already has a row,
    // any failure point identical to one already recorded there — same gear,
    // reason, station, and description — is a duplicate re-entry and must be
    // skipped, not appended as a second copy. Unique new failure points still
    // get appended onto the existing row rather than creating a separate one.
    const existing = _findExistingJourneyRow(sheet, data);
    const seenSignatures = new Set();
    let existingParsed = null;
    if (existing) {
      existingParsed = _readJourneyArrays(sheet, existing.rowIndex);
      existingParsed.gears.forEach((g, idx) => {
        seenSignatures.add(_failureSignature(g, existingParsed.reasons[idx], existingParsed.stations[idx], existingParsed.descs[idx]));
      });
    }

    const uniqueFailures = [], duplicates = [];
    failures.forEach(f => {
      const sig = _failureSignature(f.gearAtFault, f.failureReason, f.station, f.description);
      if (seenSignatures.has(sig)) {
        duplicates.push({ gear: f.gearAtFault||'', station: f.station||'', desc: String(f.description||'').slice(0,100) });
      } else {
        seenSignatures.add(sig); // also catches duplicates within this same batch
        uniqueFailures.push(f);
      }
    });

    if (!uniqueFailures.length) {
      return { success: false, error: 'All submitted observation(s) are already registered for this journey', duplicates: duplicates };
    }

    const gears=[], reasons=[], types=[], stations=[], descs=[], mcs=[], ebs=[], flags=[], flagNotes=[];
    uniqueFailures.forEach(f => {
      gears.push(f.gearAtFault || '');
      reasons.push(f.failureReason || '');
      types.push(f.subType || '');
      stations.push(f.station || '');
      descs.push(f.description || '');
      if (f.category === 'Mode Change'  && f.subType) mcs.push(f.subType);
      if (f.category === 'Undue Braking' && f.subType) ebs.push(f.subType);
      const fl = f.flag || 'None';
      if (fl !== 'None' && flags.indexOf(fl) < 0) flags.push(fl);
      if (fl !== 'None' && f.flagNote) flagNotes.push(f.flagNote);
    });

    if (existing) {
      // Append the unique new failures onto the existing journey row.
      const mergedGears    = existingParsed.gears.concat(gears);
      const mergedReasons  = existingParsed.reasons.concat(reasons);
      const mergedTypes    = existingParsed.types.concat(types);
      const mergedStations = existingParsed.stations.concat(stations);
      const mergedDescs    = existingParsed.descs.concat(descs);
      const mergedMcs      = existingParsed.mcs.concat(mcs);
      const mergedEbs      = existingParsed.ebs.concat(ebs);
      const mergedFlagNotes= existingParsed.flagNotes.concat(flagNotes);
      const mergedFlags    = existingParsed.flags.slice();
      flags.forEach(fl => { if (mergedFlags.indexOf(fl) < 0) mergedFlags.push(fl); });

      const rIdx = existing.rowIndex;
      sheet.getRange(rIdx, COL.GEAR+1).setValue(numList(mergedGears));
      sheet.getRange(rIdx, COL.REASON+1).setValue(numList(mergedReasons));
      sheet.getRange(rIdx, COL.FAILURE_TYPE+1).setValue(numList(mergedTypes));
      sheet.getRange(rIdx, COL.STATION+1).setValue(numList(mergedStations));
      sheet.getRange(rIdx, COL.DESC+1).setValue(numList(mergedDescs));
      sheet.getRange(rIdx, COL.MODE_DEG+1).setValue(numList(mergedMcs));
      sheet.getRange(rIdx, COL.EB+1).setValue(numList(mergedEbs));
      sheet.getRange(rIdx, COL.FLAG+1).setValue(mergedFlags.length ? mergedFlags.join(', ') : 'None');
      sheet.getRange(rIdx, COL.FLAG_NOTE+1).setValue(numList(mergedFlagNotes));
      if (data.nilRemark) {
        const curRemark = String(sheet.getRange(rIdx, COL.REMARK+1).getValue() || '').trim();
        sheet.getRange(rIdx, COL.REMARK+1).setValue(curRemark ? curRemark + '\n' + data.nilRemark : data.nilRemark);
      }
      // New unverified content was added — the journey needs re-verification.
      sheet.getRange(rIdx, COL.STATUS+1).setValue('Pending');
      sheet.getRange(rIdx, COL.RLYSTAFF+1).setValue('');
      sheet.getRange(rIdx, COL.VERIFIED_DATE+1).setValue('');
      _formatObsRow(sheet, rIdx);
      _bumpStatsCache();

      return {
        success: true, count: uniqueFailures.length, srNos: [existing.srNo],
        appendedToExisting: true, duplicates: duplicates
      };
    }

    const srNo = sheet.getLastRow();
    const row  = buildBaseRow(srNo);
    row[COL.GEAR]        = numList(gears);
    row[COL.REASON]      = numList(reasons);
    row[COL.FAILURE_TYPE]= numList(types);
    row[COL.STATION]     = numList(stations);
    row[COL.DESC]        = numList(descs);
    row[COL.MODE_DEG]    = numList(mcs);
    row[COL.EB]          = numList(ebs);
    row[COL.FLAG]        = flags.length ? flags.join(', ') : 'None';
    row[COL.FLAG_NOTE]   = numList(flagNotes);
    const newRowIdx = _insertObsRowInDateOrder(sheet, row, data.date);
    _formatObsRow(sheet, newRowIdx);
    _bumpStatsCache();

    return { success: true, count: uniqueFailures.length, srNos: [srNo], duplicates: duplicates };

    } finally { lock.releaseLock(); }
  } catch (e) { return { success: false, error: e.message }; }
}

// ── Data Integrity: Duplicate SR Numbers ─────────────────────────────
// Before the LockService fix above, two submissions close together in time
// could both read the same "next SR" before either had appended, giving two
// different rows the same SR No. This pair (scan + fix) is the Admin Panel
// tool to find and repair any that already happened.

// Read-only: finds every SR No used by more than one row in the same sheet.
function scanDuplicateSrNumbers() {
  try {
    const tz = Session.getScriptTimeZone();
    const duplicates = [];
    _getAllObsMonthSheets().forEach(function(sh) {
      const rows = sh.getDataRange().getValues();
      const bySr = {};
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r[COL.SR] && !r[COL.DATE]) continue;
        const srStr = String(r[COL.SR] || '').trim();
        if (!srStr) continue;
        const dv = r[COL.DATE];
        const dateStr = dv instanceof Date ? Utilities.formatDate(dv, tz, 'dd/MM/yy') : String(dv || '');
        (bySr[srStr] = bySr[srStr] || []).push({
          rowIndex: i + 1, date: dateStr,
          trainNo: String(r[COL.TRAIN] || ''), locoNo: String(r[COL.LOCO] || ''), section: String(r[COL.SECTION] || '')
        });
      }
      Object.keys(bySr).forEach(function(srStr) {
        if (bySr[srStr].length > 1) {
          duplicates.push({ sheetName: sh.getName(), srNo: srStr, rows: bySr[srStr] });
        }
      });
    });
    return { success: true, duplicates: duplicates };
  } catch(e) { return { success: false, error: e.message, duplicates: [] }; }
}

// Fixes what scanDuplicateSrNumbers() found: for each SR No shared by more
// than one row, the earliest-created row (lowest sheet row index) KEEPS that
// SR unchanged; every later duplicate row is given a fresh SR No, starting
// right after that sheet's current highest SR. Only the SR cell is ever
// touched — no other column, no row order, no deletions. Locked the same way
// as submitObservation() so it can't race a concurrent submission.
function fixDuplicateSrNumbers() {
  try {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      let fixed = 0;
      const details = [];
      _getAllObsMonthSheets().forEach(function(sh) {
        const rows = sh.getDataRange().getValues();
        const bySr = {};
        let maxSr = 0;
        for (let i = 1; i < rows.length; i++) {
          const r = rows[i];
          if (!r[COL.SR] && !r[COL.DATE]) continue;
          const srNum = Number(r[COL.SR]);
          if (!isNaN(srNum) && srNum > maxSr) maxSr = srNum;
          const srStr = String(r[COL.SR] || '').trim();
          if (!srStr) continue;
          (bySr[srStr] = bySr[srStr] || []).push(i + 1);
        }
        let nextSr = maxSr + 1;
        Object.keys(bySr).forEach(function(srStr) {
          const rowIdxs = bySr[srStr];
          if (rowIdxs.length <= 1) return;
          for (let k = 1; k < rowIdxs.length; k++) {
            const rIdx = rowIdxs[k];
            sh.getRange(rIdx, COL.SR + 1).setValue(nextSr);
            details.push({ sheet: sh.getName(), row: rIdx, oldSr: srStr, newSr: nextSr });
            nextSr++;
            fixed++;
          }
        });
      });
      _bumpStatsCache();
      return { success: true, fixed: fixed, details: details };
    } finally { lock.releaseLock(); }
  } catch(e) { return { success: false, error: e.message }; }
}

// ── Data Integrity: Out-of-Date-Order Rows ───────────────────────────
// submitObservation() now inserts new rows at their date-ordered position
// (see _findDateOrderedInsertRow above) instead of always appending at the
// sheet's physical bottom — but that only prevents NEW disorder. Rows
// entered before that fix (a back-filled older date that landed after a
// later one) are still sitting out of order. This pair repairs those.

// Read-only: for each month sheet, stable-sorts a COPY of its rows by date
// (ties keep original relative order) and reports how many rows would
// actually move if that sort were applied for real.
function scanUnsortedDateOrder() {
  try {
    const issues = [];
    _getAllObsMonthSheets().forEach(function(sh) {
      const rows = sh.getDataRange().getValues();
      const entries = [];
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r[COL.SR] && !r[COL.DATE]) continue;
        const dv = r[COL.DATE];
        const d = dv instanceof Date ? dv : new Date(String(dv || ''));
        entries.push({ origIdx: entries.length, time: isNaN(d) ? Infinity : d.getTime() });
      }
      if (entries.length < 2) return;
      const sorted = entries.slice().sort(function(a, b) {
        return (a.time - b.time) || (a.origIdx - b.origIdx);
      });
      let moved = 0;
      for (let i = 0; i < entries.length; i++) { if (sorted[i].origIdx !== entries[i].origIdx) moved++; }
      if (moved > 0) issues.push({ sheetName: sh.getName(), rowsOutOfOrder: moved, totalRows: entries.length });
    });
    return { success: true, issues: issues };
  } catch(e) { return { success: false, error: e.message, issues: [] }; }
}

// Fixes what scanUnsortedDateOrder() found: stable-sorts every month sheet's
// data rows by Date ascending (ties keep their original relative order —
// same-day entries are never shuffled against each other). Each row moves as
// one complete unit — every column travels with it, nothing is edited
// in-place, nothing is deleted or duplicated. SR numbers travel WITH their
// row unchanged (SR is an entry-order ID, not a position — this never
// renumbers anything, see fixDuplicateSrNumbers above for that). Locked the
// same way as submitObservation() so it can't race a concurrent submission.
function fixUnsortedDateOrder() {
  try {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      let sheetsFixed = 0, rowsMoved = 0;
      _getAllObsMonthSheets().forEach(function(sh) {
        const rows = sh.getDataRange().getValues();
        if (rows.length < 3) return; // header + fewer than 2 data rows — nothing to reorder
        const dataRows = rows.slice(1);
        const entries = dataRows.map(function(r, i) {
          const dv = r[COL.DATE];
          const d = dv instanceof Date ? dv : new Date(String(dv || ''));
          return { row: r, origIdx: i, time: isNaN(d) ? Infinity : d.getTime() };
        });
        const sorted = entries.slice().sort(function(a, b) {
          return (a.time - b.time) || (a.origIdx - b.origIdx);
        });
        let moved = 0;
        for (let i = 0; i < entries.length; i++) { if (sorted[i].origIdx !== entries[i].origIdx) moved++; }
        if (moved === 0) return;
        const newDataRows = sorted.map(function(e) { return e.row; });
        // Width must match what getDataRange() actually returned (rows[0].length,
        // the header row) — not MAIN_HEADERS.length, which can be WIDER than a
        // sheet whose trailing columns happen to be empty across every row,
        // and setValues() throws on any width mismatch.
        sh.getRange(2, 1, newDataRows.length, rows[0].length).setValues(newDataRows);
        sheetsFixed++;
        rowsMoved += moved;
      });
      _bumpStatsCache();
      return { success: true, sheetsFixed: sheetsFixed, rowsMoved: rowsMoved };
    } finally { lock.releaseLock(); }
  } catch(e) { return { success: false, error: e.message }; }
}

// A journey (same train + loco + date + section + up/dn direction) that already
// has a row in this month's sheet — re-submitting for it should append new
// failure points to that row instead of creating a duplicate journey row.
// Returns null if no match (NIL rows are never matched — they carry no failures).
function _findExistingJourneyRow(sheet, data) {
  const trainNo = String(data.trainNo||'').trim();
  const locoNo  = String(data.locoNo||'').trim();
  const upDn    = String(data.upDn||'').trim();
  const dateWanted = String(data.date||'').trim();
  if (!trainNo || !dateWanted) return null;
  const tz = Session.getScriptTimeZone();
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    const r = rows[i];
    if (!r[COL.SR]) continue;
    if (String(r[COL.GEAR]).trim() === 'NIL') continue;
    if (String(r[COL.TRAIN]).trim() !== trainNo) continue;
    if (String(r[COL.LOCO]).trim() !== locoNo) continue;
    if (upDn && String(r[COL.UPDN]).trim() !== upDn) continue;
    const rd = r[COL.DATE];
    const rdStr = rd instanceof Date ? Utilities.formatDate(rd, tz, 'yyyy-MM-dd') : String(rd||'').trim();
    if (rdStr !== dateWanted) continue;
    return { rowIndex: i + 1, srNo: r[COL.SR] };
  }
  return null;
}

// Reverses numList()'s encoding ("1. x\n2. y" for multiple, plain "x" for one) back
// into an array — empty string/'NIL' → []. Used to reconstruct an existing journey
// row's failure points so new submissions can be de-duplicated and appended.
function _splitFailureList(text) {
  text = String(text||'').trim();
  if (!text || text === 'NIL') return [];
  const lines = text.split('\n');
  if (lines.length > 1 && /^\d+\.\s/.test(lines[0])) {
    return lines.map(l => l.replace(/^\d+\.\s*/, '').trim());
  }
  return [text];
}

function _readJourneyArrays(sheet, rowIndex) {
  const r = sheet.getRange(rowIndex, 1, 1, MAIN_HEADERS.length).getValues()[0];
  const flagStr = String(r[COL.FLAG]||'').trim();
  return {
    srNo: r[COL.SR],
    gears:    _splitFailureList(r[COL.GEAR]),
    reasons:  _splitFailureList(r[COL.REASON]),
    types:    _splitFailureList(r[COL.FAILURE_TYPE]),
    stations: _splitFailureList(r[COL.STATION]),
    descs:    _splitFailureList(r[COL.DESC]),
    mcs:      _splitFailureList(r[COL.MODE_DEG]),
    ebs:      _splitFailureList(r[COL.EB]),
    flags:    (flagStr && flagStr !== 'None') ? flagStr.split(',').map(s=>s.trim()).filter(Boolean) : [],
    flagNotes:_splitFailureList(r[COL.FLAG_NOTE])
  };
}

// Duplicate-detection key for a single failure point — same gear at fault,
// failure reason, station, and description text = the same observation.
function _failureSignature(gear, reason, station, desc) {
  const norm = s => String(s||'').trim().toLowerCase().replace(/\s+/g, ' ');
  return [norm(gear), norm(reason), norm(station), norm(desc)].join('|');
}

// ── Row formatting: borders, wrap, alignment (applied to every new obs row) ──
function _formatObsRow(sheet, rowIdx) {
  try {
    const nCols = MAIN_HEADERS.length;
    const rng = sheet.getRange(rowIdx, 1, 1, nCols);
    rng.setVerticalAlignment('middle')
       .setWrap(true)
       .setHorizontalAlignment('center')
       .setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);
    // Long-text columns read better left-aligned
    [COL.DESC, COL.FLAG_NOTE, COL.RD_REMARK, COL.RECT_NOTE].forEach(function(c) {
      sheet.getRange(rowIdx, c + 1).setHorizontalAlignment('left');
    });
    _applyObsRowHighlight(sheet, rowIdx);
  } catch(e) { /* formatting is best-effort — never block a submission */ }
}

// ── Flag / Status highlight ──────────────────────────────
// A row flagged R&D Required / Loco Log Required, or sent back for
// Rectification, gets a bright background so anyone opening the raw backend
// sheet directly — dashboard up or not — can spot it at a glance. Overrides
// the day-banding color only while one of these applies; call this again
// after any change to the FLAG or STATUS columns so it stays in sync (falls
// back to the normal day-banding once the row is no longer flagged).
function _applyObsRowHighlight(sheet, rowIdx) {
  try {
    const status = String(sheet.getRange(rowIdx, COL.STATUS + 1).getValue() || '').trim();
    const flag   = String(sheet.getRange(rowIdx, COL.FLAG   + 1).getValue() || '').trim();
    let color = null;
    if (status === 'Rectification')                   color = '#FF8A80'; // coral — sent back, under analysis
    else if (flag.indexOf('R&D Required') >= 0)        color = '#C3B1FF'; // lavender — R&D log required
    else if (flag.indexOf('Loco Log Required') >= 0)   color = '#4DD9C8'; // teal — loco log required
    if (color) {
      sheet.getRange(rowIdx, 1, 1, MAIN_HEADERS.length).setBackground(color);
    } else {
      _applyDateBanding(sheet, rowIdx, sheet.getRange(rowIdx, COL.DATE + 1).getValue());
    }
  } catch(e) { /* best-effort — never block the calling write */ }
}

// ── Per-day color banding ────────────────────────────────
// Every distinct calendar day gets its own row background, alternating between
// two soft tints so adjacent days are visually distinct at a glance.
const _DATE_BAND_COLORS = ['#EAF0FF', '#FFFFFF'];

function _dateBandKey(dateVal) {
  if (dateVal instanceof Date) return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(dateVal || '');
}

// Stateless by design: derives the band color from the row physically ABOVE
// this one in the sheet, not from any persisted "last date/color seen" —
// that used to be tracked in CONFIG, but a persisted "last call" state can
// only ever be right when every row is appended at the true bottom in order.
// It goes stale the moment a row is manually typed into Sheets directly (no
// Apps Script call happens at all), or — since submitObservation() now
// inserts a back-filled date's row next to its own date's other rows instead
// of always at the bottom (see _findDateOrderedInsertRow) — the moment a row
// is inserted anywhere but the physical end. Reading the actual neighbor
// row's actual current color is correct in both cases, with no state to
// desync. beautifyAllObsSheets() uses the same "did the date change from the
// row before it" rule for its from-scratch bulk pass.
function _applyDateBanding(sheet, rowIdx, dateVal) {
  try {
    const dateStr = _dateBandKey(dateVal);
    let color = _DATE_BAND_COLORS[0]; // first data row (or no usable neighbor) always starts here
    if (rowIdx > 2) {
      const aboveColor = sheet.getRange(rowIdx - 1, 1).getBackground();
      const aboveDateStr = _dateBandKey(sheet.getRange(rowIdx - 1, COL.DATE + 1).getValue());
      if (aboveColor === _DATE_BAND_COLORS[0] || aboveColor === _DATE_BAND_COLORS[1]) {
        color = (aboveDateStr === dateStr) ? aboveColor
              : (aboveColor === _DATE_BAND_COLORS[0] ? _DATE_BAND_COLORS[1] : _DATE_BAND_COLORS[0]);
      }
      // else: row above isn't itself a recognized band color (e.g. it's an
      // unformatted manual row, or carries a flag-highlight color instead) —
      // can't infer a continuation from it, so start the sequence fresh here.
    }
    sheet.getRange(rowIdx, 1, 1, MAIN_HEADERS.length).setBackground(color);
  } catch(e) { /* best-effort */ }
}

// Moves the "Remark" column (if present) to sit right after "Reason", matching
// COL.REMARK's position — older sheets created before this reorder still have
// it as their last column. Safe to re-run: no-ops once already in place.
// Finds columns by header text rather than trusting existing positions, since
// that's the one thing guaranteed still correct on an unmigrated old sheet.
function _migrateRemarkColumnPosition(sheet) {
  try {
    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) return;
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const remarkIdx = headers.indexOf('Remark');  // 0-indexed
    const reasonIdx = headers.indexOf('Reason');  // 0-indexed
    if (remarkIdx < 0 || reasonIdx < 0 || remarkIdx === reasonIdx + 1) return;
    const colRange = sheet.getRange(1, remarkIdx + 1, sheet.getMaxRows(), 1);
    sheet.moveColumns(colRange, reasonIdx + 2); // 1-indexed "insert before" position, pre-move coordinates
  } catch(e) { /* best-effort */ }
}

// "LOCO OEM" used to store the submitting staff's company; it now stores the
// loco's own registered KAVACH vendor (see buildBaseRow in submitObservation).
// Existing rows already have that value sitting in LOCO_MAKE — copy it over
// wherever present so old rows read correctly too, without touching rows
// where the loco was never found in the Loco Details sheet (LOCO_MAKE blank).
function _backfillLocoOemColumn(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    const n = lastRow - 1;
    const makes = sheet.getRange(2, COL.LOCO_MAKE   + 1, n, 1).getValues();
    const oems  = sheet.getRange(2, COL.OEM_COMPANY + 1, n, 1).getValues();
    if (!makes.some(function(r) { return String(r[0] || '').trim(); })) return; // nothing to backfill
    const merged = makes.map(function(r, i) {
      const make = String(r[0] || '').trim();
      return [make || oems[i][0]]; // keep the existing value for rows with no known loco make
    });
    sheet.getRange(2, COL.OEM_COMPANY + 1, n, 1).setValues(merged);
  } catch(e) { /* best-effort */ }
}

// The "Braking Type (if applicable)" picker was removed from the form —
// its FSB column is now dead weight nobody reads. Finds it by header text
// (not position) and deletes the column outright. No-ops once already gone.
function _migrateRemoveFsbColumn(sheet) {
  try {
    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) return;
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const fsbIdx = headers.indexOf('FSB'); // 0-indexed
    if (fsbIdx < 0) return;
    sheet.deleteColumn(fsbIdx + 1);
  } catch(e) { /* best-effort */ }
}

// Retroactively re-applies day-banding + layout polish across ALL existing
// observation month sheets, in every configured obs spreadsheet. Admin-triggered
// (safe to re-run any time — recomputes from scratch, doesn't depend on prior state).
function beautifyAllObsSheets() {
  try {
    let sheetsTouched = 0, rowsTouched = 0;
    _getAllObsMonthSheets().forEach(function(sh) {
      _migrateRemoveFsbColumn(sh);
      _migrateRemarkColumnPosition(sh);
      _backfillLocoOemColumn(sh);
      // Rename the legacy 'EB' header (and fix the earlier 'Undue Breaking'
      // spelling mistake) — the column holds the full Undue Braking
      // classification (incl. FSB events), so 'EB' was misleading and
      // 'Breaking' was a typo for 'Braking'. Located by header text, not
      // position, so unmigrated sheets are safe.
      try {
        const hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
        const ebIdx = hdrs.indexOf('EB') >= 0 ? hdrs.indexOf('EB') : hdrs.indexOf('Undue Breaking');
        if (ebIdx >= 0) sh.getRange(1, ebIdx + 1).setValue('Undue Braking');
      } catch(e) { /* best-effort */ }
      _polishObsSheetLayout(sh);
      const lastRow = sh.getLastRow();
      if (lastRow < 2) { sheetsTouched++; return; }
      const dates    = sh.getRange(2, COL.DATE   + 1, lastRow - 1, 1).getValues();
      const statuses = sh.getRange(2, COL.STATUS + 1, lastRow - 1, 1).getValues();
      const flags    = sh.getRange(2, COL.FLAG   + 1, lastRow - 1, 1).getValues();
      let color = _DATE_BAND_COLORS[0], lastDateStr = null;
      const colors = dates.map(function(r, i) {
        const d = _dateBandKey(r[0]);
        if (lastDateStr !== null && d !== lastDateStr) {
          color = (color === _DATE_BAND_COLORS[0]) ? _DATE_BAND_COLORS[1] : _DATE_BAND_COLORS[0];
        }
        lastDateStr = d;
        // Flagged / rectification rows keep their alert color instead of the day-band tint
        const st = String(statuses[i][0] || '').trim();
        const fl = String(flags[i][0] || '').trim();
        let rowColor = color;
        if (st === 'Rectification')                 rowColor = '#FF8A80';
        else if (fl.indexOf('R&D Required') >= 0)    rowColor = '#C3B1FF';
        else if (fl.indexOf('Loco Log Required') >= 0) rowColor = '#4DD9C8';
        return new Array(MAIN_HEADERS.length).fill(rowColor);
      });
      if (colors.length) sh.getRange(2, 1, colors.length, MAIN_HEADERS.length).setBackgrounds(colors);
      // No state to persist — _applyDateBanding (used for every subsequent
      // single-row submission) now reads the actual row above it instead of
      // any stored "last date/color", so it stays correct on its own from
      // here regardless of where in the sheet the next row lands.
      sheetsTouched++; rowsTouched += colors.length;
    });
    SpreadsheetApp.flush();
    return { success: true, sheets: sheetsTouched, rows: rowsTouched };
  } catch (e) { return { success: false, error: e.message }; }
}

// ── Quick Register ────────────────────────────────────
// Backend sheet: QUICK_REG  cols: ID|Date|Time|TrainNo|LocoNo|Section|Status|RegisteredBy|Notes|TrainName|LocoType|LocoMake|UpDn|FitUnfit
function _getQuickRegSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName('QUICK_REG');
  if (!sh) {
    sh = ss.insertSheet('QUICK_REG');
    sh.getRange(1,1,1,14).setValues([['ID','Date','Time','Train No','Loco No','Section','Status','Registered By','Notes','Train Name','Loco Type','Loco Make','UP/DN','Fit/Unfit']]);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,14).setFontWeight('bold').setBackground('#1A2540').setFontColor('#7B9EFF');
  }
  return sh;
}

function addQuickReg(data) {
  try {
    const sh  = _getQuickRegSheet();
    const id  = Utilities.getUuid();
    const tz  = Session.getScriptTimeZone();
    const now = new Date();
    sh.appendRow([
      id,
      Utilities.formatDate(now, tz, 'dd/MM/yy'),
      Utilities.formatDate(now, tz, 'HH:mm'),
      String(data.trainNo      ||'').trim(),
      String(data.locoNo       ||'').trim(),
      String(data.section      ||'').trim(),
      'Pending',
      String(data.registeredBy ||'').trim(),
      String(data.notes        ||'').trim(),
      String(data.trainName    ||'').trim(),
      String(data.locoType     ||'').trim(),
      String(data.locoMake     ||'').trim(),
      String(data.upDn         ||'').trim(),
      String(data.fitUnfit     ||'').trim()
    ]);
    return { success: true, id };
  } catch(e) { return { success: false, error: e.message }; }
}

function getQuickRegs(section) {
  try {
    const sh   = _getQuickRegSheet();
    const tz   = Session.getScriptTimeZone();
    const rows = sh.getDataRange().getValues();
    if (rows.length <= 1) return { success: true, rows: [], pendingCount: 0 };
    const now  = Date.now();
    const all  = rows.slice(1).filter(r => r[0] && String(r[6]).trim() === 'Pending');
    const filtered = all
      .filter(r => !section || section === 'ALL' || String(r[5]).trim() === String(section).trim())
      .map(r => {
        // Safely stringify every cell — Date objects from getValues() must be converted
        const _s = v => v instanceof Date ? Utilities.formatDate(v, tz, 'dd/MM/yy') : String(v||'');
        const _t = v => v instanceof Date ? Utilities.formatDate(v, tz, 'HH:mm')    : String(v||'');
        const dateStr = _s(r[1]);
        const timeStr = _t(r[2]);
        let elapsed = '';
        try {
          const [dd,mm,yy] = dateStr.split('/');
          const [hh,mi]    = timeStr.split(':');
          const d = new Date(2000+Number(yy), Number(mm)-1, Number(dd), Number(hh), Number(mi));
          const diffMin = Math.floor((now - d.getTime()) / 60000);
          if (diffMin < 1)    elapsed = 'just now';
          else if (diffMin < 60)   elapsed = diffMin + ' min ago';
          else if (diffMin < 1440) elapsed = Math.floor(diffMin/60) + 'h ' + (diffMin%60) + 'm ago';
          else elapsed = Math.floor(diffMin/1440) + 'd ago';
        } catch(e2) {}
        return {
          id:           String(r[0] ||''),
          date:         dateStr,
          time:         timeStr,
          trainNo:      String(r[3] ||''),
          locoNo:       String(r[4] ||''),
          section:      String(r[5] ||''),
          status:       String(r[6] ||''),
          registeredBy: String(r[7] ||''),
          notes:        String(r[8] ||''),
          trainName:    String(r[9] ||''),
          locoType:     String(r[10]||''),
          locoMake:     String(r[11]||''),
          upDn:         String(r[12]||''),
          fitUnfit:     String(r[13]||''),
          elapsed
        };
      })
      .reverse();
    return { success: true, rows: filtered, pendingCount: all.length };
  } catch(e) { return { success: false, rows: [], pendingCount: 0, error: e.message + ' | ' + e.stack }; }
}

function removeQuickReg(id) {
  try {
    const sh   = _getQuickRegSheet();
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === id) { sh.deleteRow(i + 1); return { success: true }; }
    }
    return { success: false, error: 'Not found' };
  } catch(e) { return { success: false, error: e.message }; }
}

function getQuickRegPendingCount() {
  try {
    const sh   = _getQuickRegSheet();
    const rows = sh.getDataRange().getValues();
    if (rows.length <= 1) return 0;
    return rows.slice(1).filter(r => r[0] && r[6] === 'Pending').length;
  } catch(e) { return 0; }
}

// ── HBL NMS live loco auto-registration ───────────────
// hbl_nms_sync.ps1 logs into the HBL NMS portal (10.35.251.23, internal RailNet
// only, same reachability wall as KAVACH) on a dedicated account, polls its live
// loco feed, and POSTs snapshots here. Each snapshot becomes a normal QUICK_REG
// "Pending" entry -- exactly the same entry point as a staff member manually
// registering a loco -- so staff complete it the same way once the journey ends;
// only the initial "a loco entered the network" step is automated.

function _getHblNmsSyncKey() {
  var sh = getOrCreateSheet('CONFIG');
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === 'hbl_nms_sync_key') return String(rows[i][1] || '').trim();
  }
  return '';
}

function getHblNmsSyncKey(pin) {
  if (!validateAdmin(pin)) return { success: false, error: 'Wrong PIN' };
  var key = _getHblNmsSyncKey();
  if (key) return { success: true, key: key };
  return regenerateHblNmsSyncKey(pin);
}

function regenerateHblNmsSyncKey(pin) {
  if (!validateAdmin(pin)) return { success: false, error: 'Wrong PIN' };
  var key = Utilities.getUuid().replace(/-/g, '');
  var sh = getOrCreateSheet('CONFIG');
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === 'hbl_nms_sync_key') {
      sh.getRange(i + 1, 2).setValue(key);
      return { success: true, key: key };
    }
  }
  sh.appendRow(['hbl_nms_sync_key', key]);
  return { success: true, key: key };
}

// A loco can sit inside the same station's live view for several poll cycles
// (station dwell, slow approach, etc.) -- without a dedup window every poll
// would spam a fresh QUICK_REG row for it. Keyed on locoId+stationId so the
// SAME loco reappearing at a DIFFERENT station (the next leg of its journey)
// still registers as a new entry.
var HBL_NMS_DEDUP_WINDOW_MIN = 240;

function registerHblLocoEntries(payload) {
  try {
    var key = String((payload && payload.key) || '').trim();
    var savedKey = _getHblNmsSyncKey();
    if (!savedKey || key !== savedKey) return { success: false, error: 'Invalid sync key' };

    var entries = (payload && Array.isArray(payload.entries)) ? payload.entries : [];
    if (!entries.length) return { success: true, total: 0, registered: 0, skipped: 0 };

    var sh = _getQuickRegSheet();
    var rows = sh.getDataRange().getValues();
    var tz = Session.getScriptTimeZone();
    var now = new Date();

    // Build a recent-registration lookup: "locoNo|section" -> most recent Date
    var recent = {};
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (String(r[7] || '').indexOf('HBL NMS') < 0) continue; // only our own auto-registered rows gate dedup
      var loco = String(r[4] || '').trim();
      var sec  = String(r[5] || '').trim();
      if (!loco) continue;
      var dateStr = r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'dd/MM/yy') : String(r[1] || '');
      var timeStr = r[2] instanceof Date ? Utilities.formatDate(r[2], tz, 'HH:mm')    : String(r[2] || '');
      var parts = dateStr.split('/');
      var tparts = timeStr.split(':');
      if (parts.length !== 3 || tparts.length !== 2) continue;
      var rowDate = new Date(2000 + Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]), Number(tparts[0]), Number(tparts[1]));
      var dedupKey = loco + '|' + sec;
      if (!recent[dedupKey] || rowDate > recent[dedupKey]) recent[dedupKey] = rowDate;
    }

    var registered = 0, skipped = 0;

    entries.forEach(function(e) {
      var locoNo  = String(e.locoNo  || '').trim();
      var section = String(e.station || e.section || '').trim();
      if (!locoNo) { skipped++; return; }

      var dedupKey = locoNo + '|' + section;
      var last = recent[dedupKey];
      if (last && (now - last) / 60000 < HBL_NMS_DEDUP_WINDOW_MIN) { skipped++; return; }

      var res = addQuickReg({
        trainNo:      e.trainNo   || '',
        locoNo:       locoNo,
        section:      section,
        registeredBy: 'HBL NMS (Auto-Sync)',
        notes:        'Live loco entry from HBL NMS' + (e.direction ? (' -- Direction: ' + e.direction) : '') + (e.absLoc ? (', AbsLoc: ' + e.absLoc) : ''),
        trainName:    e.trainName || '',
        locoType:     e.locoType  || '',
        locoMake:     e.locoMake  || '',
        upDn:         e.direction || ''
      });
      if (res.success) {
        recent[dedupKey] = now; // avoid re-registering the same loco+station again within this same call
        registered++;
      } else {
        skipped++;
      }
    });

    return { success: true, total: entries.length, registered: registered, skipped: skipped };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ── Admin password gate (for AI generation) ─────────
function verifyAdminPassword(pwd) {
  try {
    const stored = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
    if (!stored) return { ok: false, msg: 'ADMIN_PASSWORD not set in Script Properties.' };
    return { ok: String(pwd) === String(stored), msg: 'Wrong password.' };
  } catch(e) { return { ok: false, msg: e.message }; }
}

// ── Memo / OPR Pending ──────────────────────────────
function getMemoOprPendingObs() {
  try {
    const tz = Session.getScriptTimeZone();
    const memoPending = [], oprPending = [];
    _getAllObsMonthSheets().forEach(sh => {
      const rows = sh.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r[COL.SR] && !r[COL.DATE]) continue;
        const rd = r[COL.DATE];
        const dateStr = rd instanceof Date ? Utilities.formatDate(rd, tz, 'dd/MM/yy') : String(rd||'');
        const base = {
          srNo: String(r[COL.SR]||''), date: dateStr,
          trainNo: r[COL.TRAIN]||'', trainName: r[COL.TRAIN_NAME]||'',
          locoNo: r[COL.LOCO]||'', section: r[COL.SECTION]||'',
          gear: r[COL.GEAR]||'', fitUnfit: r[COL.FIT]||'',
          status: r[COL.STATUS]||'Pending', label: r[COL.LABEL]||''
        };
        if (String(r[COL.MEMO_STATUS]||'').trim() === 'Pending') memoPending.push({...base, memoStatus:'Pending'});
        if (String(r[COL.OPR_STATUS] ||'').trim() === 'Pending') oprPending.push({...base, oprStatus:'Pending'});
      }
    });
    memoPending.sort((a,b)=>Number(b.srNo)-Number(a.srNo));
    oprPending.sort((a,b)=>Number(b.srNo)-Number(a.srNo));
    return { success: true, memoPending, oprPending };
  } catch(e) { return { success: false, memoPending:[], oprPending:[], error:e.message }; }
}

function updateMemoStatus(srNo, status) {
  return _updateObsColumn(srNo, COL.MEMO_STATUS, status);
}

function updateOprStatus(srNo, status) {
  return _updateObsColumn(srNo, COL.OPR_STATUS, status);
}

function _updateObsColumn(srNo, colIdx, value) {
  try {
    const found = _findObsRowBySr(srNo);
    if (!found) return { success: false, error: 'SR not found' };
    found.sh.getRange(found.rowIndex, colIdx + 1).setValue(value);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ── Flag Wrong Sub-Observation(s) ─────────────────────
// data: { srNo, section, flags: {"1":"red", "3":"amber"}, staffName, pin }
// flags maps a 1-based line number (within the numbered Failure Description
// list) to a color key. Persisted as JSON in COL.WRONG_HIGHLIGHT, and also
// painted directly onto the DESC cell text via rich-text font color so it's
// visible when the backend sheet is opened directly — not just in the app.
const WRONG_FLAG_COLORS = { red: '#D32F2F', amber: '#E8920A', blue: '#1565C0', purple: '#7B5CE8' };

function setObsWrongFlags(data) {
  try {
    const auth = validateRailwayStaffPin(data.staffName, data.pin);
    if (!auth.valid) return { success: false, error: 'Invalid staff name or PIN' };
    const found = _findObsRowBySr(data.srNo, data.section);
    if (!found) return { success: false, error: 'Entry not found' };
    const { sh, rowIndex: row } = found;
    const flags = (data.flags && typeof data.flags === 'object') ? data.flags : {};
    sh.getRange(row, COL.WRONG_HIGHLIGHT + 1).setValue(Object.keys(flags).length ? JSON.stringify(flags) : '');
    _applyDescRichTextHighlight(sh, row, flags);
    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

// Colors the numbered lines of the DESC cell that have a flag, via per-run
// font color/bold (cell backgrounds can't vary by character in Sheets).
// Passing an empty flags object resets the cell back to plain text.
function _applyDescRichTextHighlight(sheet, rowIdx, flags) {
  try {
    const cell = sheet.getRange(rowIdx, COL.DESC + 1);
    const text = String(cell.getValue() || '');
    if (!Object.keys(flags).length || !text) { cell.setValue(text); return; }
    const lines = text.split('\n');
    const builder = SpreadsheetApp.newRichTextValue().setText(text);
    let offset = 0;
    lines.forEach(function(line, i) {
      const lineNo = String(i + 1);
      const start = offset, end = offset + line.length;
      const colorKey = flags[lineNo];
      if (colorKey && WRONG_FLAG_COLORS[colorKey] && end > start) {
        builder.setTextStyle(start, end, SpreadsheetApp.newTextStyle()
          .setForegroundColor(WRONG_FLAG_COLORS[colorKey]).setBold(true).build());
      }
      offset = end + 1; // +1 for the '\n' consumed between lines
    });
    cell.setRichTextValue(builder.build());
  } catch(e) { /* best-effort — never block the calling save */ }
}

// ── Verify Observation ───────────────────────────────
// data: { srNo, staffName, pin }
function verifyObservation(data) {
  try {
    const auth = validateRailwayStaffPin(data.staffName, data.pin);
    if (!auth.valid) return { success: false, error: 'Invalid staff name or PIN' };
    const found = _findObsRowBySr(data.srNo, data.section);
    if (!found) return { success: false, error: 'Entry not found' };
    const { sh, rowIndex: row } = found;
    sh.getRange(row, COL.STATUS + 1).setValue('Verified');
    sh.getRange(row, COL.RLYSTAFF + 1).setValue(auth.name + (auth.designation ? ' ('+auth.designation+')' : ''));
    sh.getRange(row, COL.VERIFIED_DATE + 1).setValue(
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yy HH:mm')
    );
    _applyObsRowHighlight(sh, row);
    SpreadsheetApp.flush();
    _bumpStatsCache();
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

// ── Reject Observation → Rectification ───────────────
// data: { srNo, staffName, pin, rectNote }
function rejectObservation(data) {
  try {
    const auth = validateRailwayStaffPin(data.staffName, data.pin);
    if (!auth.valid) return { success: false, error: 'Invalid staff name or PIN' };
    if (!data.rectNote || !String(data.rectNote).trim()) return { success: false, error: 'Rectification note required' };
    const found = _findObsRowBySr(data.srNo, data.section);
    if (!found) return { success: false, error: 'Entry not found' };
    const { sh, rowIndex: row } = found;
    sh.getRange(row, COL.STATUS + 1).setValue('Rectification');
    sh.getRange(row, COL.RLYSTAFF + 1).setValue(auth.name + (auth.designation ? ' ('+auth.designation+')' : ''));
    sh.getRange(row, COL.VERIFIED_DATE + 1).setValue(
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yy HH:mm')
    );
    sh.getRange(row, COL.RECT_NOTE + 1).setValue(String(data.rectNote).trim());
    _applyObsRowHighlight(sh, row);
    SpreadsheetApp.flush();
    _bumpStatsCache();
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

// ── Resolve Rectification → Verified (skips Pending queue) ───────────────
// data: { srNo, section, staffName, pin }
function resolveRectification(data) {
  try {
    const auth = validateRailwayStaffPin(data.staffName, data.pin);
    if (!auth.valid) return { success: false, error: 'Invalid staff name or PIN' };
    const found = _findObsRowBySr(data.srNo, data.section);
    if (!found) return { success: false, error: 'Entry not found' };
    const { sh, rowIndex: row } = found;
    if (String(sh.getRange(row, COL.STATUS + 1).getValue()).trim() !== 'Rectification')
      return { success: false, error: 'Entry is not in Rectification' };
    sh.getRange(row, COL.STATUS + 1).setValue('Verified');
    sh.getRange(row, COL.RLYSTAFF + 1).setValue(auth.name + (auth.designation ? ' ('+auth.designation+')' : ''));
    sh.getRange(row, COL.VERIFIED_DATE + 1).setValue(
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yy HH:mm')
    );
    // Mark the rectification note as resolved (keep the history)
    const curNote = String(sh.getRange(row, COL.RECT_NOTE + 1).getValue() || '');
    sh.getRange(row, COL.RECT_NOTE + 1).setValue(
      curNote + '\n[Resolved] ' + auth.name + ' — ' +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yy HH:mm')
    );
    _applyObsRowHighlight(sh, row);
    SpreadsheetApp.flush();
    _bumpStatsCache();
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

// ── Delete Observation (admin action, PIN protected) ────────────────
// data: { srNo, section, staffName, pin }
function deleteObservation(data) {
  try {
    const auth = validateRailwayStaffPin(data.staffName, data.pin);
    if (!auth.valid) return { success: false, error: 'Invalid staff name or PIN' };
    const found = _findObsRowBySr(data.srNo, data.section);
    if (!found) return { success: false, error: 'Entry not found' };
    found.sh.deleteRow(found.rowIndex);
    SpreadsheetApp.flush();
    _bumpStatsCache();
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

// ── Add R&D Remark ────────────────────────────────────
// data: { srNo, staffName, pin, remark }
function addRdRemark(data) {
  try {
    const auth = validateRailwayStaffPin(data.staffName, data.pin);
    if (!auth.valid) return { success: false, error: 'Invalid staff name or PIN' };
    if (!data.remark || !String(data.remark).trim()) return { success: false, error: 'Remark required' };
    const found = _findObsRowBySr(data.srNo);
    if (!found) return { success: false, error: 'Entry not found' };
    found.sh.getRange(found.rowIndex, COL.RD_REMARK + 1).setValue(
      String(data.remark).trim() + ' — ' + auth.name +
      ' (' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yy HH:mm') + ')'
    );
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

// ── Close Loco Log ────────────────────────────────────
// data: { srNo, staffName, pin }
function closeLocoLog(data) {
  try {
    const auth = validateRailwayStaffPin(data.staffName, data.pin);
    if (!auth.valid) return { success: false, error: 'Invalid staff name or PIN' };
    const found = _findObsRowBySr(data.srNo);
    if (!found) return { success: false, error: 'Entry not found' };
    found.sh.getRange(found.rowIndex, COL.LOCO_LOG_STATUS + 1).setValue(
      'Closed — ' + auth.name +
      ' (' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yy HH:mm') + ')'
    );
    // Replace only the Loco Log token — preserve a combined R&D flag if present
    var curFlagL = String(found.sh.getRange(found.rowIndex, COL.FLAG + 1).getValue() || '').trim();
    if (curFlagL.indexOf('Loco Log Required') >= 0) {
      found.sh.getRange(found.rowIndex, COL.FLAG + 1).setValue(curFlagL.replace('Loco Log Required', 'Loco Log Closed'));
    }
    _applyObsRowHighlight(found.sh, found.rowIndex);
    SpreadsheetApp.flush();
    _bumpStatsCache();
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

// ── Close R&D Case (remark satisfied) ────────────────
function closeRdCase(data) {
  try {
    const auth = validateRailwayStaffPin(data.staffName, data.pin);
    if (!auth.valid) return { success: false, error: 'Invalid staff name or PIN' };
    const found = _findObsRowBySr(data.srNo, data.section);
    if (!found) return { success: false, error: 'Entry not found' };
    // Replace only the R&D token — preserve a combined Loco Log flag if present
    var curFlag = String(found.sh.getRange(found.rowIndex, COL.FLAG + 1).getValue() || '').trim();
    var newFlag = curFlag.indexOf('R&D Required') >= 0
      ? curFlag.replace('R&D Required', 'R&D Closed')
      : (curFlag && curFlag !== 'None' ? curFlag + ', R&D Closed' : 'R&D Closed');
    found.sh.getRange(found.rowIndex, COL.FLAG + 1).setValue(newFlag);
    found.sh.getRange(found.rowIndex, COL.RD_REMARK + 1).setValue(
      String(found.sh.getRange(found.rowIndex, COL.RD_REMARK + 1).getValue() || '') +
      (data.closingNote ? '\n[Closed] ' + data.closingNote : '') +
      ' — ' + auth.name +
      ' (' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yy HH:mm') + ')'
    );
    _applyObsRowHighlight(found.sh, found.rowIndex);
    SpreadsheetApp.flush();
    _bumpStatsCache();
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ── Get Rectification Observations ───────────────────
function getRectificationObs() {
  try {
    const rows = _getObservationsByFilter(r => r[COL.STATUS] === 'Rectification');
    return rows;
  } catch(e) {
    return { _error: e.message };
  }
}

// ── Update Flag ──────────────────────────────────────
function updateFlag(data) {
  try {
    const found = _findObsRowBySr(data.srNo);
    if (!found) return { success: false, error: 'Entry not found' };
    found.sh.getRange(found.rowIndex, COL.FLAG + 1).setValue(data.flag);
    found.sh.getRange(found.rowIndex, COL.FLAG_NOTE + 1).setValue(data.flagNote || '');
    _applyObsRowHighlight(found.sh, found.rowIndex);
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

// ── Flag an entry into R&D / Loco Log Queue from Pending/Rectification ──
// data: { srNo, section, targets: ['rd','loco'], note, staffName, pin }
// Additive — merges onto any existing flags rather than overwriting, and the
// note is appended with a timestamp so a history of who flagged what for
// which sub-observation survives across multiple flags on the same row. The
// FLAG column drives whole-row queue membership (that's the only granularity
// the data model supports); `note` is where a specific sub-observation gets
// called out when only part of a multi-observation row needs attention.
function addObsFlag(data) {
  try {
    const auth = validateRailwayStaffPin(data.staffName, data.pin);
    if (!auth.valid) return { success: false, error: 'Invalid staff name or PIN' };
    const targets = Array.isArray(data.targets) ? data.targets : [];
    if (!targets.length) return { success: false, error: 'Select R&D, Loco Log, or both' };
    const found = _findObsRowBySr(data.srNo, data.section);
    if (!found) return { success: false, error: 'Entry not found' };
    const { sh, rowIndex: row } = found;

    const curFlag = String(sh.getRange(row, COL.FLAG + 1).getValue() || '').trim();
    const flags = (curFlag && curFlag !== 'None') ? curFlag.split(',').map(function(s){ return s.trim(); }).filter(Boolean) : [];
    if (targets.indexOf('rd')   >= 0 && flags.indexOf('R&D Required')       < 0) flags.push('R&D Required');
    if (targets.indexOf('loco') >= 0 && flags.indexOf('Loco Log Required')  < 0) flags.push('Loco Log Required');
    sh.getRange(row, COL.FLAG + 1).setValue(flags.length ? flags.join(', ') : 'None');

    if (data.note && String(data.note).trim()) {
      const stamp = '[' + auth.name + ' — ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yy HH:mm') + '] ' + String(data.note).trim();
      const curNote = String(sh.getRange(row, COL.FLAG_NOTE + 1).getValue() || '').trim();
      sh.getRange(row, COL.FLAG_NOTE + 1).setValue(curNote ? curNote + '\n' + stamp : stamp);
    }
    _applyObsRowHighlight(sh, row);
    SpreadsheetApp.flush();
    _bumpStatsCache();
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

// ── Get Pending ──────────────────────────────────────
function getPendingObservations() {
  return _getObservationsByFilter(r => r[COL.STATUS] === 'Pending');
}

// ── OEM Loco Analytics ──────────────────────────────
// Returns per-OEM unique loco count + loco list for a date range.
// sections: optional array of section names to restrict to (multi-select filter).
// Empty/omitted = all sections. Each OEM's result now also breaks down loco counts
// per section, so you can see how many of each OEM's locos ran in which section
// during the date range, not just an overall total.
function getOemLocoAnalytics(dateFrom, dateTo, sections) {
  try {
    const tz = Session.getScriptTimeZone();
    const wantSections = Array.isArray(sections) && sections.length ? new Set(sections) : null;
    const oemMap = {};

    _getAllObsMonthSheets().forEach(sh => {
      const rows = sh.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r[COL.SR] && !r[COL.DATE]) continue;
        if (r[COL.GEAR] === 'NIL') continue; // skip NIL rows

        // Date filter
        const rd = r[COL.DATE];
        const dStr = rd instanceof Date ? Utilities.formatDate(rd, tz, 'yyyy-MM-dd') : '';
        if (dateFrom && dStr < dateFrom) continue;
        if (dateTo   && dStr > dateTo)   continue;

        const section = String(r[COL.SECTION] || '').trim() || 'Unknown';
        if (wantSections && !wantSections.has(section)) continue;

        const oem  = String(r[COL.LOCO_MAKE] || '').trim() || 'Unknown';
        const loco = String(r[COL.LOCO]      || '').trim();
        if (!oemMap[oem]) oemMap[oem] = { locos: {}, obs: 0, bySection: {} };
        if (loco) oemMap[oem].locos[loco] = true;
        oemMap[oem].obs++;
        if (!oemMap[oem].bySection[section]) oemMap[oem].bySection[section] = { locos: {}, obs: 0 };
        if (loco) oemMap[oem].bySection[section].locos[loco] = true;
        oemMap[oem].bySection[section].obs++;
      }
    });

    const result = Object.entries(oemMap).map(([oem, d]) => ({
      oem,
      locoCount: Object.keys(d.locos).length,
      locos: Object.keys(d.locos).sort(),
      obs: d.obs,
      bySection: Object.entries(d.bySection).map(([section, sd]) => ({
        section, locoCount: Object.keys(sd.locos).length, obs: sd.obs
      })).sort((a, b) => b.locoCount - a.locoCount)
    })).sort((a, b) => b.locoCount - a.locoCount);

    return { success: true, data: result, dateFrom, dateTo };
  } catch(e) { return { success: false, error: e.message, data: [] }; }
}

// ── Get Flagged ──────────────────────────────────────
function getFlaggedObservations(flagType) {
  try {
    const qCfg = getQueueReasonConfig();
    const isRd   = flagType === 'R&D Required';
    const isLoco = flagType === 'Loco Log Required';
    const reasons = isRd   ? (qCfg.rdReasons   || []).map(function(s){ return String(s).trim().toLowerCase(); })
                  : isLoco ? (qCfg.locoReasons || []).map(function(s){ return String(s).trim().toLowerCase(); })
                  : [];
    const results = [];
    _getAllObsMonthSheets().forEach(function(sheet) {
      const rows = sheet.getDataRange().getValues();
      if (rows.length <= 1) return;
      for (var i = 1; i < rows.length; i++) {
        var r = rows[i];
        if (!r[COL.SR] && !r[COL.DATE]) continue;
        var flagVal  = String(r[COL.FLAG]   || '').trim();
        var reasonVal = String(r[COL.REASON] || '').trim().toLowerCase();
        // Combined flags possible ("R&D Required, Loco Log Required") — check per-type close marker
        var closedMarker = flagType.replace('Required', 'Closed');
        if (flagVal.indexOf(closedMarker) >= 0) continue;
        var matched = flagVal.indexOf(flagType) >= 0 ||
                      (reasons.length && reasons.some(function(x){ return reasonVal.indexOf(x) >= 0; }));
        if (!matched) continue;
        var rawDate = r[COL.DATE];
        var dateStr = rawDate instanceof Date
          ? Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'dd/MM/yy')
          : String(rawDate || '');
        results.push({
          srNo         : String(r[COL.SR]           || ''),
          date         : dateStr,
          trainNo      : String(r[COL.TRAIN]         || ''),
          trainName    : String(r[COL.TRAIN_NAME]    || ''),
          locoNo       : String(r[COL.LOCO]          || ''),
          locoType     : String(r[COL.LOCO_TYPE]     || ''),
          section      : String(r[COL.SECTION]       || ''),
          station      : String(r[COL.STATION]       || ''),
          upDn         : String(r[COL.UPDN]          || ''),
          fitUnfit     : String(r[COL.FIT]           || ''),
          gear         : String(r[COL.GEAR]          || ''),
          reason       : String(r[COL.REASON]        || ''),
          category     : _obsCategory(r),
          modeDeg      : String(r[COL.MODE_DEG]      || ''),
          eb           : String(r[COL.EB]            || ''),
          subType      : String(r[COL.MODE_DEG] || r[COL.EB] || '-'),
          desc         : String(r[COL.DESC]          || ''),
          status       : String(r[COL.STATUS]        || 'Pending'),
          oemCompany   : String(r[COL.OEM_COMPANY]   || '-'),
          oemStaff     : String(r[COL.OEM_STAFF]     || '-'),
          flag         : String(r[COL.FLAG]          || 'None'),
          flagNote     : String(r[COL.FLAG_NOTE]     || ''),
          verifiedBy   : String(r[COL.RLYSTAFF]      || ''),
          verifiedDate : String(r[COL.VERIFIED_DATE] || ''),
          rdRemark     : String(r[COL.RD_REMARK]     || ''),
          locoLogStatus: String(r[COL.LOCO_LOG_STATUS]|| ''),
          rectNote     : String(r[COL.RECT_NOTE]     || ''),
          label        : String(r[COL.LABEL]         || '')
        });
      }
    });
    return results.reverse();
  } catch(e) { return [{ _filterError: e.message }]; }
}

// ── Debug: R&D queue diagnosis ───────────────────────
function debugRdQueue() {
  try {
    const sheets = _getAllObsMonthSheets();
    const info = { sheetCount: sheets.length, sheets: [], flagMatches: [], reasonMatches: [], error: null };
    const qCfg = getQueueReasonConfig();
    info.rdReasons = qCfg.rdReasons || [];
    sheets.forEach(function(sh) {
      const rows = sh.getDataRange().getValues();
      info.sheets.push({ name: sh.getName(), rows: rows.length - 1 });
      for (var i = 1; i < rows.length; i++) {
        var r = rows[i];
        if (!r[COL.SR] && !r[COL.DATE]) continue;
        var flagVal = String(r[COL.FLAG] || '');
        var reasonVal = String(r[COL.REASON] || '');
        if (flagVal === 'R&D Required') {
          info.flagMatches.push({ sr: r[COL.SR], section: r[COL.SECTION], flag: flagVal, status: r[COL.STATUS] });
        }
        if (info.rdReasons.some(function(x){ return reasonVal.trim().toLowerCase() === String(x).trim().toLowerCase(); })) {
          info.reasonMatches.push({ sr: r[COL.SR], section: r[COL.SECTION], reason: reasonVal });
        }
      }
    });
    return info;
  } catch(e) { return { error: e.message }; }
}

// ── Shared row mapper ────────────────────────────────
function _getObservationsByFilter(filterFn) {
  try {
    const results = [];
    _getAllObsMonthSheets().forEach(sheet => {
      const rows = sheet.getDataRange().getValues();
      if (rows.length <= 1) return;
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r[COL.SR] && !r[COL.DATE]) continue;
        if (filterFn(r)) {
        const rawDate = r[COL.DATE];
        const dateStr = rawDate instanceof Date
          ? Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'dd/MM/yy')
          : String(rawDate || '');
        results.push({
          rowIndex     : i + 1,
          srNo         : String(r[COL.SR] || ''),
          date         : dateStr,
          trainNo      : r[COL.TRAIN],
          trainName    : r[COL.TRAIN_NAME] || '',
          locoNo       : r[COL.LOCO],
          locoType     : r[COL.LOCO_TYPE] || '',
          section      : r[COL.SECTION],
          station      : r[COL.STATION],
          upDn         : r[COL.UPDN] || '',
          fitUnfit     : r[COL.FIT]  || '',
          gear         : r[COL.GEAR],
          reason       : r[COL.REASON],
          category     : _obsCategory(r),
          modeDeg      : r[COL.MODE_DEG] || '',
          eb           : r[COL.EB]  || '',
          subType      : r[COL.MODE_DEG] || r[COL.EB] || '-',
          desc         : r[COL.DESC],
          status       : r[COL.STATUS] || 'Pending',
          oemCompany   : r[COL.OEM_COMPANY] || '-',
          oemStaff     : r[COL.OEM_STAFF] || '-',
          flag         : r[COL.FLAG] || 'None',
          flagNote     : r[COL.FLAG_NOTE] || '',
          verifiedBy   : String(r[COL.RLYSTAFF] || ''),
          verifiedDate : r[COL.VERIFIED_DATE] instanceof Date
            ? Utilities.formatDate(r[COL.VERIFIED_DATE], Session.getScriptTimeZone(), 'dd/MM/yy HH:mm')
            : String(r[COL.VERIFIED_DATE] || ''),
          rdRemark     : String(r[COL.RD_REMARK] || ''),
          locoLogStatus: String(r[COL.LOCO_LOG_STATUS] || ''),
          rectNote     : String(r[COL.RECT_NOTE] || ''),
          label        : String(r[COL.LABEL] || ''),
          remark       : String(r[COL.REMARK] || ''),
          wrongFlags   : String(r[COL.WRONG_HIGHLIGHT] || '')
        });
        }
      }
    });
    return results.reverse();
  } catch (e) { return [{ _filterError: e.message, _filterStack: String(e.stack||'').slice(0,400) }]; }
}

// ── Analytics ────────────────────────────────────────
function getAnalytics() {
  try {
    const section={}, gear={}, status={Pending:0,Verified:0}, reasons={}, monthly={};

    // Corridor groups (BRC-GDA / VS-URN / BJW-ADI) — always present in the output
    // even with zero data, with UP/DN split from the row's own direction column.
    const sectionMap = _getSectionMap();
    const rawToGroup = {};
    sectionMap.forEach(function(g){ g.sections.forEach(function(s){ rawToGroup[s] = g.group; }); });
    const groupTally = {};
    sectionMap.forEach(function(g){ groupTally[g.group] = {total:0, up:0, dn:0}; });

    _getAllObsMonthSheets().forEach(sheet => {
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r[COL.SR] && !r[COL.DATE]) continue;
        const sec  = String(r[COL.SECTION]||'Unknown').trim();
        const st   = String(r[COL.STATUS]||'Pending').trim();
        const dt   = String(r[COL.DATE]  ||'').trim();
        const updn = String(r[COL.UPDN]||'').trim().toUpperCase();
        section[sec] = (section[sec]||0)+1;
        status[st]   = (status[st]||0)+1;
        const mkey = dt.slice(3,8);
        if (mkey) monthly[mkey] = (monthly[mkey]||0)+1;

        const grp = rawToGroup[sec];
        if (grp) {
          groupTally[grp].total++;
          if (updn === 'UP') groupTally[grp].up++;
          else if (updn === 'DN') groupTally[grp].dn++;
        }

        // Gear/Reason cells hold numbered lists ("1. X\n2. Y") on multi-failure
        // journeys — tally individual tokens (strips NIL) instead of counting the
        // whole raw cell text as one bogus category.
        _rptSplitVals(r[COL.GEAR]).forEach(function(g){ gear[g] = (gear[g]||0)+1; });
        _rptSplitVals(r[COL.REASON]).forEach(function(rs){ reasons[rs] = (reasons[rs]||0)+1; });
      }
    });

    // Admin-configured category visibility for the dashboard charts
    const filt = _getDashboardChartFilters();
    const gearFiltered = {};
    Object.keys(gear).forEach(function(k){ if (filt.gearExclude.indexOf(k) < 0) gearFiltered[k] = gear[k]; });
    const reasonEntries = Object.entries(reasons).filter(function(e){ return filt.reasonExclude.indexOf(e[0]) < 0; });

    const topReasons = reasonEntries.sort((a,b)=>b[1]-a[1]).slice(0,5).map(([r,c])=>({reason:r,count:c}));
    const monthlyArr = Object.entries(monthly).sort((a,b)=>a[0].localeCompare(b[0])).slice(-4).map(([m,c])=>({month:m,count:c}));
    const sectionGroups = sectionMap.map(function(g){
      return { group: g.group, total: groupTally[g.group].total, up: groupTally[g.group].up, dn: groupTally[g.group].dn };
    });

    return {
      section, sectionGroups, gear: gearFiltered, gearAll: gear,
      status, reasons: topReasons, reasonsAll: reasons, monthly: monthlyArr
    };
  } catch(e) { return { section:{}, sectionGroups:[], gear:{}, gearAll:{}, status:{}, reasons:[], reasonsAll:{}, monthly:[] }; }
}

// ── Dashboard Chart Filters (admin-controlled category visibility) ──────
// Hides chosen Gear/Reason categories (e.g. "NIL") from the dashboard's
// Gear Fault Distribution / Top Failure Reasons charts without touching data.
function _getDashboardChartFilters() {
  try {
    const sh = getOrCreateSheet('CONFIG');
    const rows = sh.getDataRange().getValues();
    let gearExclude = ['NIL'], reasonExclude = ['NIL'];
    rows.slice(1).forEach(function(r){
      const k = String(r[0]||'').trim();
      if (k === 'dashboard_gear_exclude')   { try { gearExclude   = JSON.parse(r[1]); } catch(e){} }
      if (k === 'dashboard_reason_exclude') { try { reasonExclude = JSON.parse(r[1]); } catch(e){} }
    });
    return { gearExclude: gearExclude, reasonExclude: reasonExclude };
  } catch(e) { return { gearExclude:['NIL'], reasonExclude:['NIL'] }; }
}

function getDashboardChartFilters() {
  const f = _getDashboardChartFilters();
  return { success:true, gearExclude:f.gearExclude, reasonExclude:f.reasonExclude };
}

function saveDashboardChartFilters(data) {
  try {
    const sh = getOrCreateSheet('CONFIG');
    const rows = sh.getDataRange().getValues();
    const _set = function(key, val) {
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]||'').trim() === key) { sh.getRange(i+1,2).setValue(JSON.stringify(val)); return; }
      }
      sh.appendRow([key, JSON.stringify(val)]);
    };
    _set('dashboard_gear_exclude',   data.gearExclude||[]);
    _set('dashboard_reason_exclude', data.reasonExclude||[]);
    return { success:true };
  } catch(e) { return { success:false, error:e.message }; }
}

// ── Queue Reason Config ──────────────────────────────
function getQueueReasonConfig() {
  try {
    const sh   = getOrCreateSheet('CONFIG');
    const rows = sh.getDataRange().getValues();
    let rdReasons = [], locoReasons = [];
    rows.slice(1).forEach(r => {
      const key = String(r[0]||'').trim();
      const val = String(r[1]||'').trim();
      if (!key || !val) return;
      if (key === 'rd_track_reasons')   try { rdReasons   = JSON.parse(val); } catch(e) {}
      if (key === 'loco_track_reasons') try { locoReasons = JSON.parse(val); } catch(e) {}
    });
    return { success: true, rdReasons, locoReasons };
  } catch(e) { return { success: false, rdReasons: [], locoReasons: [] }; }
}

function saveQueueReasonConfig(rdReasons, locoReasons) {
  try {
    const sh   = getOrCreateSheet('CONFIG');
    const rows = sh.getDataRange().getValues();
    const _set = function(key, val) {
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]||'').trim() === key) {
          sh.getRange(i + 1, 2).setValue(JSON.stringify(val));
          rows[i][1] = JSON.stringify(val);
          return;
        }
      }
      sh.appendRow([key, JSON.stringify(val)]);
      rows.push([key, JSON.stringify(val)]);
    };
    _set('rd_track_reasons',   rdReasons   || []);
    _set('loco_track_reasons', locoReasons || []);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// Derive display category from a row — 'NIL' text in the MODE_DEG/EB cells must NOT count
function _obsCategory(r) {
  const md = String(r[COL.MODE_DEG]||'').trim();
  const eb = String(r[COL.EB]||'').trim();
  if (md && md !== 'NIL') return 'Mode Change';
  if (eb && eb !== 'NIL') return 'Undue Braking';
  return String(r[COL.GEAR]||'').trim() === 'NIL' ? 'NIL' : '-';
}

// True only for a real ICMS label — "Non-ICMS" must NOT match
function _isIcmsLabel(v) {
  return String(v||'').toUpperCase().split(',').some(function(l){
    l = l.trim();
    return l.indexOf('ICMS') >= 0 && l.indexOf('NON') < 0;
  });
}

// ── Stats ────────────────────────────────────────────
// Drop the cached stats snapshot after any data mutation so the
// actor (and everyone polling) sees the change immediately.
function _bumpStatsCache() {
  try { CacheService.getScriptCache().remove('stats_v1__'); } catch(e) {}
}

function getStats(opts) {
  try {
    // 20s shared cache — many users poll every 30s; serve them the same snapshot
    const cacheKey = 'stats_v1_' + ((opts&&opts.dateFrom)||'') + '_' + ((opts&&opts.dateTo)||'');
    const cache = CacheService.getScriptCache();
    const hit = cache.get(cacheKey);
    if (hit) { try { return JSON.parse(hit); } catch(e) {} }
    let total=0, modeChange=0, undueEB=0, todayCount=0, icmsCount=0, rdCount=0, locoCount=0, rectification=0, pendingCount=0;
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yy').slice(0,5);
    const tz = Session.getScriptTimeZone();
    const dateFrom = (opts && opts.dateFrom) ? opts.dateFrom : '';
    const dateTo   = (opts && opts.dateTo)   ? opts.dateTo   : '';
    // Total Entries / Undue Braking / Undue Mode Degradation are compared
    // directly against Monthly Report ("TOTAL JOURNEYS" and the "…Analysis"
    // tables), which default to the CURRENT MONTH — not all-time. When no
    // explicit date filter has been applied here, scope those three counters
    // to the current month to match; Today/ICMS/queues stay all-time (or
    // respect an explicit filter) since they have no such counterpart.
    // Note: Monthly Report shows one "TOTAL JOURNEYS" number PER SECTION TAB
    // (BRC-GDA / VS-URN / BJW-ADI); this card is the sum across all of them,
    // so it matches the sum of all three tabs' totals, not any single tab.
    let mcFrom = dateFrom, mcTo = dateTo;
    if (!dateFrom && !dateTo) {
      const now = new Date();
      mcFrom = Utilities.formatDate(new Date(now.getFullYear(), now.getMonth(), 1), tz, 'yyyy-MM-dd');
      mcTo   = Utilities.formatDate(new Date(now.getFullYear(), now.getMonth() + 1, 0), tz, 'yyyy-MM-dd');
    }
    // Load configured reasons for queue cards
    const qCfg = getQueueReasonConfig();
    const rdReasons   = (qCfg.rdReasons   || []).map(function(s){ return String(s).trim().toLowerCase(); });
    const locoReasons = (qCfg.locoReasons || []).map(function(s){ return String(s).trim().toLowerCase(); });
    _getAllObsMonthSheets().forEach(sheet => {
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r[0] && !r[1]) continue;
        const rdVal = r[COL.DATE];
        const dStr  = rdVal instanceof Date ? Utilities.formatDate(rdVal, tz, 'yyyy-MM-dd') : '';
        if (dateFrom || dateTo) {
          if (dateFrom && dStr < dateFrom) continue;
          if (dateTo   && dStr > dateTo)   continue;
        }
        // Count individual sub-observations, not just rows that have at least
        // one — a single journey row can carry several bundled Mode Degradation/
        // Undue Braking events ("1. FS-SR\n2. FS-LS"), and Monthly Report tallies
        // each one separately. Using the same split here keeps this card in
        // sync with the "Undue Braking/Mode Degradation Analysis" totals there.
        const inDefaultRange = (!mcFrom || dStr >= mcFrom) && (!mcTo || dStr <= mcTo);
        if (inDefaultRange) {
          total++;
          modeChange += _rptSplitVals(r[COL.MODE_DEG]).length;
          undueEB    += _rptSplitVals(r[COL.EB]).length;
        }
        if (String(r[COL.DATE]).includes(todayStr)) todayCount++;
        if (_isIcmsLabel(r[COL.LABEL])) icmsCount++;
        if (r[COL.STATUS] === 'Rectification') rectification++;
        if (!r[COL.STATUS] || r[COL.STATUS] === 'Pending') pendingCount++;
        const rsn = String(r[COL.REASON]||'').trim().toLowerCase();
        const flagV = String(r[COL.FLAG]||'').trim();
        const isRd   = flagV.indexOf('R&D Closed') < 0 &&
                       (flagV.indexOf('R&D Required') >= 0      || (rdReasons.length   && rdReasons.some(function(x){   return rsn.indexOf(x)>=0; })));
        const isLoco = flagV.indexOf('Loco Log Closed') < 0 &&
                       (flagV.indexOf('Loco Log Required') >= 0 || (locoReasons.length && locoReasons.some(function(x){ return rsn.indexOf(x)>=0; })));
        if (isRd)   rdCount++;
        if (isLoco) locoCount++;
      }
    });
    const out = { total, modeChange, undueEB, today: todayCount, icmsCount, rdCount, locoCount, rectification, pendingCount };
    cache.put(cacheKey, JSON.stringify(out), 20);
    return out;
  } catch (e) { return { total:0, modeChange:0, undueEB:0, today:0, icmsCount:0, rdCount:0, locoCount:0, rectification:0, pendingCount:0 }; }
}

// ── Get Filtered Entries (for stat card drilldown) ───
function getFilteredEntries(opts) {
  const filter   = (typeof opts === 'string') ? opts : (opts && opts.filter) || 'all';
  const dateFrom = (typeof opts === 'object' && opts) ? (opts.dateFrom || '') : '';
  const dateTo   = (typeof opts === 'object' && opts) ? (opts.dateTo   || '') : '';
  try {
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yy').slice(0,5);
    const results  = [];
    _getAllObsMonthSheets().forEach(sheet => {
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r[COL.SR] && !r[COL.DATE]) continue;
        if (dateFrom || dateTo) {
          const rd2 = r[COL.DATE];
          const dStr2 = rd2 instanceof Date ? Utilities.formatDate(rd2, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
          if (dateFrom && dStr2 < dateFrom) continue;
          if (dateTo   && dStr2 > dateTo)   continue;
        }
        let include = false;
        if (filter === 'all')        include = true;
        if (filter === 'today')      include = String(r[COL.DATE]).includes(todayStr);
        if (filter === 'modeChange') include = !!(r[COL.MODE_DEG] && r[COL.MODE_DEG] !== 'NIL' && r[COL.MODE_DEG] !== '');
        if (filter === 'undueEB')    include = !!(r[COL.EB]       && r[COL.EB]       !== 'NIL' && r[COL.EB] !== '');
        if (filter === 'icms')       include = _isIcmsLabel(r[COL.LABEL]);
        if (filter === 'rdReason' || filter === 'locoReason') {
          const qc  = getQueueReasonConfig();
          const reasons = filter === 'rdReason' ? (qc.rdReasons||[]) : (qc.locoReasons||[]);
          const rsn = String(r[COL.REASON]||'').trim().toLowerCase();
          if (reasons.length) include = reasons.some(function(x){ return rsn === String(x).trim().toLowerCase(); });
          else include = r[COL.FLAG] === (filter === 'rdReason' ? 'R&D Required' : 'Loco Log Required');
        }
        if (include) {
          const rd = r[COL.DATE];
          results.push({
            srNo      : String(r[COL.SR] || ''),
            date      : rd instanceof Date ? Utilities.formatDate(rd, Session.getScriptTimeZone(), 'dd/MM/yy') : String(rd || ''),
            trainNo   : r[COL.TRAIN],
            locoNo    : r[COL.LOCO],
            section   : r[COL.SECTION],
            station   : r[COL.STATION],
            gear      : r[COL.GEAR]  || '—',
            reason    : r[COL.REASON] || '—',
            category  : _obsCategory(r) === '-' ? '—' : _obsCategory(r),
            subType   : r[COL.MODE_DEG] || r[COL.EB] || '—',
            desc      : r[COL.DESC]  || '',
            status    : r[COL.STATUS] || 'Pending',
            oemCompany: r[COL.OEM_COMPANY] || '—',
            oemStaff  : r[COL.OEM_STAFF]   || '—',
            flag      : r[COL.FLAG]  || 'None',
            remark    : r[COL.REMARK] || '',
            wrongFlags: r[COL.WRONG_HIGHLIGHT] || '',
          });
        }
      }
    });
    _attachHblRemarks(results);
    return results.reverse();
  } catch (e) { return []; }
}

// Parses "dd.mm.yy", "dd/mm/yyyy", etc. (and real Date objects) into a
// comparable "yyyy-MM-dd" key — dates in the observation sheet and the HBL
// workbook aren't guaranteed to share the same on-screen format.
function _normDateKey(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const m = String(v || '').trim().match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})$/);
  if (!m) return '';
  const d = m[1].padStart(2, '0'), mo = m[2].padStart(2, '0');
  const y = m[3].length === 2 ? '20' + m[3] : m[3];
  return y + '-' + mo + '-' + d;
}

// ══════════════════════════════════════════════════════
// HISTORICAL OBSERVATION BULK IMPORTER
// ══════════════════════════════════════════════════════

// "2026-03-01" -> "MARCH-2026"
function _histMonthYearSheetName(dateKey) {
  const parts = String(dateKey || '').split('-'); // yyyy-mm-dd
  if (parts.length !== 3) return null;
  const y = +parts[0], m = +parts[1] - 1;
  if (!y || m < 0 || m > 11) return null;
  return MONTH_NAMES[m] + '-' + y;
}

// sectionGroupLabel is the exact `.group` string from _getSectionMap(),
// e.g. "VS–URN" (en dash) — never a hand-typed key.
function _histResolveDestSpreadsheet(sectionGroupLabel) {
  const map = _getSectionMap();
  const grp = map.find(function(g) { return g.group === sectionGroupLabel; });
  if (grp && grp.sheetIds.length) return SpreadsheetApp.openById(grp.sheetIds[0]);
  return SpreadsheetApp.openById(SHEET_ID);
}

// Normalized identity key for duplicate detection: same day + train + loco.
function _histRowKey(dateKey, trainNo, locoNo) {
  return dateKey + '|' + String(trainNo || '').trim() + '|' + String(locoNo || '').trim();
}

// Exposes the configured section groups (label + their section list) so the
// frontend's picker and section-validation warning never drift from what
// Section Map admin actually has configured.
function getHistSectionGroups() {
  try {
    const map = _getSectionMap();
    return { success: true, groups: map.map(function(g) { return { label: g.group, sections: g.sections, hasSheet: !!(g.sheetIds && g.sheetIds.length) }; }) };
  } catch (e) { return { success: false, error: e.message }; }
}

// rows: [{ date, trainNo, locoNo }]. Returns which of them already exist
// in their own resolved target sheet, by (date, trainNo, locoNo).
function checkHistoricalImportDuplicates(sectionGroupLabel, rows) {
  try {
    if (!Array.isArray(rows) || !rows.length) return { success: true, duplicates: [] };
    const grp = _getSectionMap().find(function(g) { return g.group === sectionGroupLabel; });
    if (!grp || !grp.sheetIds.length) return { success: false, error: 'Section group "' + sectionGroupLabel + '" has no linked spreadsheet configured — check Admin Panel > Observation Sheet IDs.' };
    const ss = SpreadsheetApp.openById(grp.sheetIds[0]);
    const sheetCache = {}; // sheetName -> Set of existing keys -> srNo
    const duplicates = [];
    rows.forEach(function(r) {
      const dateKey = _normDateKey(r.date);
      if (!dateKey) return;
      const sheetName = _histMonthYearSheetName(dateKey);
      if (!sheetName) return;
      if (!sheetCache[sheetName]) {
        const map = {};
        const sh = ss.getSheetByName(sheetName);
        if (sh) {
          const data = sh.getDataRange().getValues();
          for (let i = 1; i < data.length; i++) {
            const rd = data[i][COL.DATE];
            const dk = rd instanceof Date ? Utilities.formatDate(rd, Session.getScriptTimeZone(), 'yyyy-MM-dd') : _normDateKey(rd);
            if (!dk) continue;
            const key = _histRowKey(dk, data[i][COL.TRAIN], data[i][COL.LOCO]);
            map[key] = data[i][COL.SR];
          }
        }
        sheetCache[sheetName] = map;
      }
      const key = _histRowKey(dateKey, r.trainNo, r.locoNo);
      if (sheetCache[sheetName].hasOwnProperty(key)) {
        duplicates.push({ key: key, existingSrNo: sheetCache[sheetName][key], existingSheet: sheetName });
      }
    });
    return { success: true, duplicates: duplicates };
  } catch (e) { return { success: false, error: e.message }; }
}

function bulkImportHistoricalObservations(sectionGroupLabel, rows, opts) {
  try {
    if (!Array.isArray(rows) || !rows.length) return { success: false, error: 'No rows to import' };
    opts = opts || {};
    const skipSet = new Set(opts.skipKeys || []);
    const grp = _getSectionMap().find(function(g) { return g.group === sectionGroupLabel; });
    if (!grp || !grp.sheetIds.length) return { success: false, error: 'Section group "' + sectionGroupLabel + '" has no linked spreadsheet configured — check Admin Panel > Observation Sheet IDs.' };
    const ss = SpreadsheetApp.openById(grp.sheetIds[0]);
    const tz = Session.getScriptTimeZone();

    // Group rows by their own resolved target sheet name.
    const bySheet = {}; // sheetName -> array of row objects (each still carries _dateKey)
    let skippedBadDate = 0;
    let skippedDuplicate = 0;
    rows.forEach(function(r) {
      const dateKey = _normDateKey(r.date);
      if (!dateKey) { skippedBadDate++; return; }
      const key = _histRowKey(dateKey, r.trainNo, r.locoNo);
      if (skipSet.has(key)) { skippedDuplicate++; return; }
      const sheetName = _histMonthYearSheetName(dateKey);
      if (!sheetName) { skippedBadDate++; return; }
      (bySheet[sheetName] = bySheet[sheetName] || []).push(Object.assign({ _dateKey: dateKey }, r));
    });

    let imported = 0;
    const sheetsTouched = [];
    Object.keys(bySheet).forEach(function(sheetName) {
      let sh = ss.getSheetByName(sheetName);
      if (!sh) {
        sh = ss.insertSheet(sheetName);
        sh.appendRow(MAIN_HEADERS);
        sh.getRange(1, 1, 1, MAIN_HEADERS.length).setFontWeight('bold').setBackground('#1A2540').setFontColor('#7B9EFF');
        sh.setFrozenRows(1);
        _polishObsSheetLayout(sh);
      }
      const lastRow = sh.getLastRow();
      let nextSr = 1;
      if (lastRow > 1) {
        const srCol = sh.getRange(2, COL.SR + 1, lastRow - 1, 1).getValues();
        srCol.forEach(function(v) { const n = +v[0]; if (n > nextSr - 1) nextSr = n + 1; });
      }

      const block = bySheet[sheetName].map(function(r) {
        let locoType = r.locoType || '', locoShed = r.locoShed || '', locoZone = r.locoZone || '', locoMake = r.locoMake || '';
        const needsBackfill = opts.backfillLocoDetails && r.locoNo &&
          [locoType, locoShed, locoZone, locoMake].some(function(v) { return !v || /^not found$/i.test(v); });
        if (needsBackfill) {
          const info = fetchLocoInfo(r.locoNo);
          if (info && info.success) {
            if (!locoType || /^not found$/i.test(locoType)) locoType = info.locoType || locoType;
            if (!locoShed || /^not found$/i.test(locoShed)) locoShed = info.shed || locoShed;
            if (!locoZone || /^not found$/i.test(locoZone)) locoZone = info.zone || locoZone;
            if (!locoMake || /^not found$/i.test(locoMake)) locoMake = info.make || locoMake;
          }
        }
        const row = new Array(MAIN_HEADERS.length).fill('');
        row[COL.SR]           = nextSr++;
        row[COL.DATE]         = Utilities.formatDate(new Date(r._dateKey + 'T00:00:00'), tz, 'dd/MM/yy');
        row[COL.SECTION]      = r.section || '';
        row[COL.TRAIN]        = r.trainNo || '';
        row[COL.LOCO]         = r.locoNo || '';
        row[COL.TRAIN_NAME]   = r.trainName || '';
        row[COL.FIT]          = r.fitUnfit || '';
        row[COL.OEM_COMPANY]  = r.oemCompany || '';
        row[COL.UPDN]         = r.upDn || '';
        row[COL.REASON]       = r.reason || '';
        row[COL.REMARK]       = r.remark || '';
        row[COL.FAILURE_TYPE] = r.failureType || '';
        row[COL.STATION]      = r.station || '';
        row[COL.GEAR]         = r.gearAtFault || '';
        row[COL.DESC]         = r.desc || '';
        row[COL.OEM_STAFF]    = r.nmsEngineer || '';
        row[COL.RLYSTAFF]     = r.verifiedBy || '';
        row[COL.STATUS]       = 'Pending';
        row[COL.FLAG]         = 'None';
        row[COL.MODE_DEG]     = r.modeDeg || '';
        row[COL.EB]           = r.undueBraking || '';
        row[COL.OP_AVAIL]     = r.opAvail || '';
        row[COL.LOCO_FAULTS]  = r.locoFaults || '';
        row[COL.WRONG_OP]     = r.wrongOp || '';
        row[COL.SOC]          = r.soc || '';
        row[COL.EXEC]         = r.exec || '';
        row[COL.FOREIGN_TAG]  = r.foreignTag || '';
        row[COL.INVALID_SIG]  = r.invalidSig || '';
        row[COL.TAG_MISS]     = r.tagMiss || '';
        row[COL.THREE_TAG]    = r.threeTag || '';
        row[COL.LOCO_TYPE]    = locoType;
        row[COL.LOCO_SHED]    = locoShed;
        row[COL.LOCO_ZONE]    = locoZone;
        row[COL.LOCO_MAKE]    = locoMake;
        imported++;
        return row;
      });

      if (block.length) {
        sh.getRange(sh.getLastRow() + 1, 1, block.length, MAIN_HEADERS.length).setValues(block);
        sheetsTouched.push(sheetName);
      }
    });

    return { success: true, imported: imported, skippedBadDate: skippedBadDate, skippedDuplicate: skippedDuplicate, sheetsTouched: sheetsTouched };
  } catch (e) { return { success: false, error: e.message }; }
}

// Maps a destField key (from _HIST_DEST_FIELDS) to the COL constant name it
// writes. Deliberately excludes date/trainNo/locoNo — those three form the
// match identity in bulkCorrectHistoricalObservations below and are never
// themselves correctable (a row is found BY them, not written by this path).
const _HIST_FIELD_TO_COL = {
  section: 'SECTION', trainName: 'TRAIN_NAME', fitUnfit: 'FIT', oemCompany: 'OEM_COMPANY',
  upDn: 'UPDN', reason: 'REASON', remark: 'REMARK', failureType: 'FAILURE_TYPE',
  station: 'STATION', gearAtFault: 'GEAR', desc: 'DESC', nmsEngineer: 'OEM_STAFF',
  verifiedBy: 'RLYSTAFF', modeDeg: 'MODE_DEG', undueBraking: 'EB', opAvail: 'OP_AVAIL',
  locoFaults: 'LOCO_FAULTS', wrongOp: 'WRONG_OP', soc: 'SOC', exec: 'EXEC',
  foreignTag: 'FOREIGN_TAG', invalidSig: 'INVALID_SIG', tagMiss: 'TAG_MISS', threeTag: 'THREE_TAG',
  locoType: 'LOCO_TYPE', locoShed: 'LOCO_SHED', locoZone: 'LOCO_ZONE', locoMake: 'LOCO_MAKE'
};

// Correction pass for already-imported rows: overwrites ONLY the columns
// named in fieldsToWrite, on the row matching each incoming row's own
// (date, trainNo, locoNo) identity — never by row order (duplicate-skips
// during the original import can shift order between the source file and
// what actually landed in the sheet). Never inserts a new row: a source row
// with no match in the destination sheet is simply skipped and counted in
// `unmatched`. SR No, Status, Flag, and every column not in fieldsToWrite
// are left byte-for-byte untouched.
function bulkCorrectHistoricalObservations(sectionGroupLabel, rows, fieldsToWrite) {
  try {
    if (!Array.isArray(rows) || !rows.length) return { success: false, error: 'No rows to process' };
    const cols = (fieldsToWrite || []).filter(function(f) { return _HIST_FIELD_TO_COL.hasOwnProperty(f); });
    if (!cols.length) return { success: false, error: 'No correctable fields selected' };
    const grp = _getSectionMap().find(function(g) { return g.group === sectionGroupLabel; });
    if (!grp || !grp.sheetIds.length) return { success: false, error: 'Section group "' + sectionGroupLabel + '" has no linked spreadsheet configured — check Admin Panel > Observation Sheet IDs.' };
    const ss = SpreadsheetApp.openById(grp.sheetIds[0]);
    const tz = Session.getScriptTimeZone();

    const bySheet = {};
    let skippedBadDate = 0;
    rows.forEach(function(r) {
      const dateKey = _normDateKey(r.date);
      if (!dateKey) { skippedBadDate++; return; }
      const sheetName = _histMonthYearSheetName(dateKey);
      if (!sheetName) { skippedBadDate++; return; }
      (bySheet[sheetName] = bySheet[sheetName] || []).push(Object.assign({ _dateKey: dateKey }, r));
    });

    let corrected = 0, unmatched = 0;
    const sheetsTouched = [];
    Object.keys(bySheet).forEach(function(sheetName) {
      const sh = ss.getSheetByName(sheetName);
      if (!sh) { unmatched += bySheet[sheetName].length; return; }
      const data = sh.getDataRange().getValues();

      const rowIdxByKey = {};
      for (let i = 1; i < data.length; i++) {
        const rd = data[i][COL.DATE];
        const dk = rd instanceof Date ? Utilities.formatDate(rd, tz, 'yyyy-MM-dd') : _normDateKey(rd);
        if (!dk) continue;
        const k = _histRowKey(dk, data[i][COL.TRAIN], data[i][COL.LOCO]);
        (rowIdxByKey[k] = rowIdxByKey[k] || []).push(i);
      }

      let sheetChanged = false;
      bySheet[sheetName].forEach(function(r) {
        const idxList = rowIdxByKey[_histRowKey(r._dateKey, r.trainNo, r.locoNo)];
        if (!idxList) { unmatched++; return; }
        idxList.forEach(function(idx) {
          cols.forEach(function(fieldKey) {
            data[idx][COL[_HIST_FIELD_TO_COL[fieldKey]]] = r[fieldKey] || '';
          });
          corrected++;
        });
        sheetChanged = true;
      });

      if (sheetChanged) {
        cols.forEach(function(fieldKey) {
          const colIdx = COL[_HIST_FIELD_TO_COL[fieldKey]];
          const colValues = [];
          for (let i = 1; i < data.length; i++) colValues.push([data[i][colIdx]]);
          sh.getRange(2, colIdx + 1, colValues.length, 1).setValues(colValues);
        });
        sheetsTouched.push(sheetName);
      }
    });

    return { success: true, corrected: corrected, unmatched: unmatched, skippedBadDate: skippedBadDate, sheetsTouched: sheetsTouched };
  } catch (e) { return { success: false, error: e.message }; }
}

// ── Update Observation (edit → back to Pending) ──────
function updateObservation(data) {
  try {
    if (data.date && String(data.date) > Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')) {
      return { success: false, error: 'Date cannot be in the future' };
    }
    const found = _findObsRowBySr(data.srNo);
    if (!found) return { success: false, error: 'Entry not found' };
    let { sh: sheet, rowIndex: row } = found;

    // Date is edited via a calendar picker as "yyyy-mm-dd". Month sheets are
    // named e.g. "JULY-2026", so a date change that crosses into a different
    // month/year has to physically move the row — otherwise every report that
    // scans month sheets by name would never see it under its new date.
    if (data.date) {
      const parts = String(data.date).split('-'); // yyyy-mm-dd
      if (parts.length === 3) {
        const y = +parts[0], m = +parts[1] - 1;
        const targetSheetName = MONTH_NAMES[m] + '-' + y;
        const newDateStr = Utilities.formatDate(new Date(y, m, +parts[2]), Session.getScriptTimeZone(), 'dd/MM/yy');
        if (targetSheetName !== sheet.getName()) {
          const ss = sheet.getParent();
          let destSheet = ss.getSheetByName(targetSheetName);
          if (!destSheet) {
            destSheet = ss.insertSheet(targetSheetName);
            destSheet.appendRow(MAIN_HEADERS);
            destSheet.getRange(1, 1, 1, MAIN_HEADERS.length).setFontWeight('bold').setBackground('#1A2540').setFontColor('#7B9EFF');
            destSheet.setFrozenRows(1);
            _polishObsSheetLayout(destSheet);
          }
          const rowValues = sheet.getRange(row, 1, 1, MAIN_HEADERS.length).getValues()[0];
          destSheet.appendRow(rowValues);
          const destRow = destSheet.getLastRow();
          destSheet.getRange(destRow, COL.DATE + 1).setValue(newDateStr);
          sheet.deleteRow(row);
          sheet = destSheet;
          row = destRow;
        } else {
          sheet.getRange(row, COL.DATE + 1).setValue(newDateStr);
        }
      }
    }

    if (data.trainNo  !== undefined) sheet.getRange(row, COL.TRAIN    + 1).setValue(data.trainNo);
    if (data.locoNo   !== undefined) sheet.getRange(row, COL.LOCO     + 1).setValue(data.locoNo);
    if (data.section  !== undefined) sheet.getRange(row, COL.SECTION  + 1).setValue(data.section);
    if (data.station  !== undefined) sheet.getRange(row, COL.STATION  + 1).setValue(data.station);
    if (data.upDn     !== undefined) sheet.getRange(row, COL.UPDN     + 1).setValue(data.upDn);
    if (data.fitUnfit !== undefined) sheet.getRange(row, COL.FIT      + 1).setValue(data.fitUnfit);
    if (data.gear     !== undefined) sheet.getRange(row, COL.GEAR     + 1).setValue(data.gear);
    if (data.reason   !== undefined) sheet.getRange(row, COL.REASON   + 1).setValue(data.reason);
    if (data.modeDeg  !== undefined) sheet.getRange(row, COL.MODE_DEG + 1).setValue(data.modeDeg);
    if (data.eb       !== undefined) sheet.getRange(row, COL.EB       + 1).setValue(data.eb);
    if (data.desc     !== undefined) sheet.getRange(row, COL.DESC     + 1).setValue(data.desc);
    if (data.label    !== undefined) sheet.getRange(row, COL.LABEL    + 1).setValue(data.label);
    sheet.getRange(row, COL.STATUS        + 1).setValue('Pending');
    sheet.getRange(row, COL.RLYSTAFF      + 1).setValue('');
    sheet.getRange(row, COL.VERIFIED_DATE + 1).setValue('');
    sheet.getRange(row, COL.RECT_NOTE     + 1).setValue('');
    // An edit implies the flagged wrong sub-observation(s) were addressed —
    // clear the flags (the DESC rewrite above already reset any rich-text
    // color, since .setValue() drops prior text-run formatting).
    sheet.getRange(row, COL.WRONG_HIGHLIGHT + 1).setValue('');
    _applyObsRowHighlight(sheet, row);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ── Shared Drive folder ───────────────────────────────
function _getBrcFolder() {
  const folders = DriveApp.getFoldersByName('BRC_KAVACH_DOCS');
  return folders.hasNext() ? folders.next() : DriveApp.createFolder('BRC_KAVACH_DOCS');
}

function _uploadFileToDrive(fileData, fileName, mimeType, subfolder) {
  const parent = _getBrcFolder();
  let folder;
  const subs = parent.getFoldersByName(subfolder);
  folder = subs.hasNext() ? subs.next() : parent.createFolder(subfolder);
  const blob = Utilities.newBlob(Utilities.base64Decode(fileData), mimeType || 'application/octet-stream', fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { fileId: file.getId(), fileUrl: file.getUrl(), fileName: file.getName() };
}

// ── Loco MEMO ────────────────────────────────────────
// Sheet cols: ID | Memo Type | Details | Date | File Name | File ID | File URL | Linked SRs (JSON)
function _getMemoSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName('LOCO_MEMO');
  if (!sh) {
    sh = ss.insertSheet('LOCO_MEMO');
    sh.getRange(1, 1, 1, 8).setValues([['ID','Memo Type','Details','Date','File Name','File ID','File URL','Linked SRs']]);
    sh.setFrozenRows(1);
  } else {
    // migrate: if old 6-col sheet, extend headers
    const lastCol = sh.getLastColumn();
    if (lastCol < 8) {
      const newHeaders = ['File Name','File ID','File URL','Linked SRs'].slice(-(8 - lastCol));
      sh.getRange(1, lastCol + 1, 1, newHeaders.length).setValues([newHeaders]);
    }
  }
  return sh;
}

// ── Labels (ICMS / Non-ICMS / custom) ────────────────
const DEFAULT_LABELS = ['ICMS', 'Non-ICMS'];

function getLabels() {
  try {
    const sh = getOrCreateSheet('CONFIG');
    const rows = sh.getDataRange().getValues();
    for (const r of rows.slice(1)) {
      if (String(r[0]).trim() === 'obs_labels') {
        try { return { success: true, labels: JSON.parse(r[1]) }; } catch(e) {}
      }
    }
    return { success: true, labels: DEFAULT_LABELS };
  } catch(e) { return { success: true, labels: DEFAULT_LABELS }; }
}

function saveLabels(labels) {
  try {
    const sh = getOrCreateSheet('CONFIG');
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === 'obs_labels') {
        sh.getRange(i + 1, 2).setValue(JSON.stringify(labels));
        return { success: true };
      }
    }
    sh.appendRow(['obs_labels', JSON.stringify(labels)]);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ── NIL/Remark Presets — admin-curated common remarks, offered as a
// one-click dropdown next to the free-text Remark field to cut down on
// retyping the same recurring notes (e.g. "BRCY Offline due to EI
// Remodeling"). Storage mirrors Labels exactly.
const DEFAULT_NIL_REMARK_PRESETS = ['BRCY Offline due to EI Remodeling'];

function getNilRemarkPresets() {
  try {
    const sh = getOrCreateSheet('CONFIG');
    const rows = sh.getDataRange().getValues();
    for (const r of rows.slice(1)) {
      if (String(r[0]).trim() === 'nil_remark_presets') {
        try { return { success: true, presets: JSON.parse(r[1]) }; } catch(e) {}
      }
    }
    return { success: true, presets: DEFAULT_NIL_REMARK_PRESETS };
  } catch(e) { return { success: true, presets: DEFAULT_NIL_REMARK_PRESETS }; }
}

function saveNilRemarkPresets(presets) {
  try {
    const sh = getOrCreateSheet('CONFIG');
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === 'nil_remark_presets') {
        sh.getRange(i + 1, 2).setValue(JSON.stringify(presets));
        return { success: true };
      }
    }
    sh.appendRow(['nil_remark_presets', JSON.stringify(presets)]);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function updateObservationLabel(data) {
  try {
    const found = _findObsRowBySr(data.srNo);
    if (!found) return { success: false, error: 'SR not found' };
    found.sh.getRange(found.rowIndex, COL.LABEL + 1).setValue(data.label || '');
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// sections: optional array of exact section names (e.g. from a group's
// .sections list) to restrict to — empty/undefined means no filter (all).
function getObservationsByLabel(label, sections) {
  try {
    const results = [];
    const tz = Session.getScriptTimeZone();
    const sectionSet = (Array.isArray(sections) && sections.length)
      ? new Set(sections.map(s => String(s || '').trim())) : null;
    _getAllObsMonthSheets().forEach(sh => {
      const name = sh.getName();
      const rows = sh.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r[COL.SR] && !r[COL.DATE]) continue;
        if (sectionSet && !sectionSet.has(String(r[COL.SECTION] || '').trim())) continue;
        const rowLabel = String(r[COL.LABEL] || '').trim();
        const rowLabels = rowLabel ? rowLabel.split(',').map(l=>l.trim()) : [];
        if (label === '__ALL__' || rowLabels.includes(label) || (label === '__UNLABELLED__' && !rowLabel)) {
          const rawDate = r[COL.DATE];
          const dateStr = rawDate instanceof Date ? Utilities.formatDate(rawDate, tz, 'dd/MM/yy') : String(rawDate||'');
          results.push({
            srNo: String(r[COL.SR]||''), date: dateStr,
            trainNo: r[COL.TRAIN]||'', trainName: r[COL.TRAIN_NAME]||'',
            locoNo: r[COL.LOCO]||'', section: r[COL.SECTION]||'',
            gear: r[COL.GEAR]||'', status: r[COL.STATUS]||'Pending',
            category: _obsCategory(r),
            subType: r[COL.MODE_DEG]||r[COL.EB]||'-',
            label: rowLabel, sheetName: name
          });
        }
      }
    });
    results.sort((a,b)=>Number(b.srNo)-Number(a.srNo));
    return { success: true, rows: results };
  } catch(e) { return { success: false, error: e.message, rows: [] }; }
}

// sections: optional array of exact section names — empty/undefined means
// no filter (all sections), same contract as getObservationsByLabel above.
function getLabelCounts(sections) {
  try {
    const counts = {};
    let total = 0, unlabelled = 0;
    const sectionSet = (Array.isArray(sections) && sections.length)
      ? new Set(sections.map(s => String(s || '').trim())) : null;
    _getAllObsMonthSheets().forEach(sh => {
      const rows = sh.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r[COL.SR] && !r[COL.DATE]) continue;
        if (sectionSet && !sectionSet.has(String(r[COL.SECTION] || '').trim())) continue;
        total++;
        const lbl = String(r[COL.LABEL]||'').trim();
        if (lbl) { lbl.split(',').map(l=>l.trim()).filter(Boolean).forEach(l=>{ counts[l]=(counts[l]||0)+1; }); }
        else unlabelled++;
      }
    });
    return { success: true, counts, total, unlabelled };
  } catch(e) { return { success: false, counts: {}, total: 0, unlabelled: 0 }; }
}

// ── MEMO & OPR Summary ────────────────────────────────
function getMemoOprSummary() {
  try {
    // Get memos
    const memoSh = _getMemoSheet();
    const memoRows = memoSh.getDataRange().getValues();
    const memos = memoRows.slice(1).filter(r=>r[0]).map(r=>{
      let linkedObs = []; try { linkedObs = JSON.parse(r[7]||'[]'); } catch(e){}
      return {
        id: r[0], memoType: r[1]||'', details: r[2]||'',
        date: r[3] instanceof Date ? Utilities.formatDate(r[3],Session.getScriptTimeZone(),'dd/MM/yy') : String(r[3]||''),
        fileUrl: r[6]||'', fileName: r[4]||'', linkedObs
      };
    });

    // Get OPR records
    const oprRows = (()=>{ try { const s=SpreadsheetApp.openById(SHEET_ID).getSheetByName('OPR_RECORDS'); return s?s.getDataRange().getValues():[[]] } catch(e){return [[]]} })();
    const oprs = oprRows.slice(1).filter(r=>r[0]).map(r=>{
      let linkedObs=[]; try { linkedObs=JSON.parse(r[2]||'[]'); } catch(e){}
      return {
        id: r[0], issue: r[1]||'', fileUrl: r[5]||'', fileName: r[3]||'',
        date: r[6] instanceof Date ? Utilities.formatDate(r[6],Session.getScriptTimeZone(),'dd/MM/yy') : String(r[6]||''),
        linkedObs
      };
    });

    return { success: true, memos, oprs };
  } catch(e) { return { success: false, memos: [], oprs: [], error: e.message }; }
}

const DEFAULT_MEMO_CATEGORIES = ['Mode Degradation', 'RFID Issue'];

function getMemoCategories() {
  try {
    const sh = getOrCreateSheet('CONFIG');
    const rows = sh.getDataRange().getValues();
    for (const r of rows.slice(1)) {
      if (String(r[0]).trim() === 'memo_categories') {
        try { return { success: true, categories: JSON.parse(r[1]) }; } catch(e) {}
      }
    }
    return { success: true, categories: DEFAULT_MEMO_CATEGORIES };
  } catch(e) { return { success: true, categories: DEFAULT_MEMO_CATEGORIES }; }
}

function saveMemoCategories(categories) {
  try {
    const sh = getOrCreateSheet('CONFIG');
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === 'memo_categories') {
        sh.getRange(i + 1, 2).setValue(JSON.stringify(categories));
        return { success: true };
      }
    }
    sh.appendRow(['memo_categories', JSON.stringify(categories)]);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function addLocoMemo(data) {
  // data: {memoType, details, date, linkedSrNos[], linkedObs[], fileData, fileName, mimeType}
  try {
    let fileId = '', fileUrl = '', fileName = '';
    if (data.fileData) {
      const r = _uploadFileToDrive(data.fileData, data.fileName, data.mimeType, 'LOCO_MEMO');
      fileId = r.fileId; fileUrl = r.fileUrl; fileName = r.fileName;
    }
    const sh = _getMemoSheet();
    const id = Utilities.getUuid();
    const dateStr = data.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yy');
    const linkedSrs = JSON.stringify(data.linkedSrNos || []);
    sh.appendRow([id, data.memoType || '', data.details || '', dateStr, fileName, fileId, fileUrl, linkedSrs]);
    return { success: true, id };
  } catch(e) { return { success: false, error: e.message }; }
}

function getLocoMemos() {
  try {
    const sh = _getMemoSheet();
    const rows = sh.getDataRange().getValues();
    if (rows.length <= 1) return [];
    return rows.slice(1).map(r => {
      let linkedObs = [];
      try { linkedObs = JSON.parse(r[7] || '[]'); } catch(e) {}
      // handle old 6-col rows: col indices may not exist
      return {
        id: r[0], memoType: r[1] || '', details: r[2] || '',
        date: r[3] instanceof Date
          ? Utilities.formatDate(r[3], Session.getScriptTimeZone(), 'dd/MM/yy')
          : String(r[3] || ''),
        fileName: r[4] || '', fileId: r[5] || '', fileUrl: r[6] || '',
        linkedObs
      };
    }).reverse();
  } catch(e) { return []; }
}

function deleteLocoMemo(data) {
  // data: {id} or string id (backwards compat)
  const id = (typeof data === 'object') ? data.id : data;
  const delFile = (typeof data === 'object') ? data.deleteFile : false;
  try {
    const sh = _getMemoSheet();
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === id) {
        if (delFile && rows[i][5]) {
          try { DriveApp.getFileById(rows[i][5]).setTrashed(true); } catch(e) {}
        }
        sh.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Not found' };
  } catch(e) { return { success: false, error: e.message }; }
}

// ── OPR (One Page Report) ────────────────────────────
// Sheet cols: ID | Issue | Linked SRs (JSON) | File Name | File ID | File URL | Added Date
function _getOprSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName('OPR_RECORDS');
  if (!sh) {
    sh = ss.insertSheet('OPR_RECORDS');
    sh.getRange(1, 1, 1, 7).setValues([['ID','Issue','Linked SRs','File Name','File ID','File URL','Added Date']]);
    sh.setFrozenRows(1);
  } else {
    // migrate old 9-col format if needed
    const lastCol = sh.getLastColumn();
    if (lastCol < 7) {
      const add = ['File Name','File ID','File URL','Added Date'].slice(-(7 - lastCol));
      sh.getRange(1, lastCol + 1, 1, add.length).setValues([add]);
    }
  }
  return sh;
}

function addOprRecord(data) {
  // data: {issue, linkedSrNos[], linkedObs[], fileData, fileName, mimeType}
  try {
    let fileId = '', fileUrl = '', fileName = '';
    if (data.fileData) {
      const r = _uploadFileToDrive(data.fileData, data.fileName, data.mimeType, 'OPR');
      fileId = r.fileId; fileUrl = r.fileUrl; fileName = r.fileName;
    }
    const sh = _getOprSheet();
    const id = Utilities.getUuid();
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yy');
    const linkedSrs = JSON.stringify(data.linkedSrNos || []);
    sh.appendRow([id, data.issue || '', linkedSrs, fileName, fileId, fileUrl, dateStr]);
    return { success: true, id, fileUrl, fileName };
  } catch(e) { return { success: false, error: e.message }; }
}

function getOprRecords() {
  try {
    const sh = _getOprSheet();
    const rows = sh.getDataRange().getValues();
    if (rows.length <= 1) return [];
    return rows.slice(1).map(r => {
      let linkedObs = [];
      try { linkedObs = JSON.parse(r[2] || '[]'); } catch(e) {}
      return {
        id: r[0], issue: r[1] || '', linkedObs,
        fileName: r[3] || '', fileId: r[4] || '', fileUrl: r[5] || '',
        date: r[6] instanceof Date
          ? Utilities.formatDate(r[6], Session.getScriptTimeZone(), 'dd/MM/yy')
          : String(r[6] || '')
      };
    }).reverse();
  } catch(e) { return []; }
}

function deleteOprRecord(data) {
  const id = (typeof data === 'object') ? data.id : data;
  const delFile = (typeof data === 'object') ? data.deleteFile : false;
  try {
    const sh = _getOprSheet();
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === id) {
        if (delFile && rows[i][4]) {
          try { DriveApp.getFileById(rows[i][4]).setTrashed(true); } catch(e) {}
        }
        sh.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Not found' };
  } catch(e) { return { success: false, error: e.message }; }
}

// ── PPT (manual attach — nav-rail quick add, parallel to MEMO/OPR) ───
// Same shape as OPR_RECORDS: ID | Issue | Linked SRs | File Name | File ID | File URL | Added Date
function _getPptRecordsSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName('PPT_RECORDS');
  if (!sh) {
    sh = ss.insertSheet('PPT_RECORDS');
    sh.getRange(1, 1, 1, 7).setValues([['ID','Issue','Linked SRs','File Name','File ID','File URL','Added Date']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function addPptRecord(data) {
  // data: {issue, linkedSrNos[], linkedObs[], fileData, fileName, mimeType}
  try {
    let fileId = '', fileUrl = '', fileName = '';
    if (data.fileData) {
      const r = _uploadFileToDrive(data.fileData, data.fileName, data.mimeType, 'PPT');
      fileId = r.fileId; fileUrl = r.fileUrl; fileName = r.fileName;
    }
    const sh = _getPptRecordsSheet();
    const id = Utilities.getUuid();
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yy');
    const linkedSrs = JSON.stringify(data.linkedSrNos || []);
    sh.appendRow([id, data.issue || '', linkedSrs, fileName, fileId, fileUrl, dateStr]);
    return { success: true, id, fileUrl, fileName };
  } catch(e) { return { success: false, error: e.message }; }
}

function getPptRecords() {
  try {
    const sh = _getPptRecordsSheet();
    const rows = sh.getDataRange().getValues();
    if (rows.length <= 1) return [];
    return rows.slice(1).map(r => {
      let linkedObs = [];
      try { linkedObs = JSON.parse(r[2] || '[]'); } catch(e) {}
      return {
        id: r[0], issue: r[1] || '', linkedObs,
        fileName: r[3] || '', fileId: r[4] || '', fileUrl: r[5] || '',
        date: r[6] instanceof Date
          ? Utilities.formatDate(r[6], Session.getScriptTimeZone(), 'dd/MM/yy')
          : String(r[6] || '')
      };
    }).reverse();
  } catch(e) { return []; }
}

function deletePptRecord(data) {
  const id = (typeof data === 'object') ? data.id : data;
  const delFile = (typeof data === 'object') ? data.deleteFile : false;
  try {
    const sh = _getPptRecordsSheet();
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === id) {
        if (delFile && rows[i][4]) {
          try { DriveApp.getFileById(rows[i][4]).setTrashed(true); } catch(e) {}
        }
        sh.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'Not found' };
  } catch(e) { return { success: false, error: e.message }; }
}

// ── Observation Picker (for MEMO / OPR linking) ───────
function getObservationsForPicker(sheetName) {
  try {
    // Find the named sheet across all obs spreadsheets
    const targetName = sheetName || getCurrentMonthSheetName();
    let sh = null;
    for (const ss of _getAllObsSS()) {
      sh = ss.getSheetByName(targetName);
      if (sh) break;
    }
    if (!sh) return [];
    const rows = sh.getDataRange().getValues();
    if (rows.length <= 1) return [];
    return rows.slice(1).filter(r => r[COL.SR] && r[COL.DATE]).map(r => {
      const rd = r[COL.DATE];
      return {
        srNo    : String(r[COL.SR]),
        date    : rd instanceof Date ? Utilities.formatDate(rd, Session.getScriptTimeZone(), 'dd/MM/yy') : String(rd || ''),
        trainNo : String(r[COL.TRAIN] || ''),
        locoNo  : String(r[COL.LOCO]  || ''),
        section : String(r[COL.SECTION] || ''),
        gear    : String(r[COL.GEAR]   || ''),
        status  : String(r[COL.STATUS] || 'Pending'),
        category: r[COL.MODE_DEG] ? 'MC' : r[COL.EB] ? 'EB' : '—'
      };
    }).reverse();
  } catch(e) { return []; }
}

// ── Recent Entries ───────────────────────────────────
// Most recent N observations across ALL configured obs spreadsheets — the old
// version only read the default SHEET_ID's current-month tab via getMainSheet(),
// which missed everything logged in the section-specific spreadsheets and left
// the Recent Entries panel looking empty.
function getRecentEntries() {
  try {
    const tz = Session.getScriptTimeZone();
    const rows = [];
    _getAllObsMonthSheets().forEach(function(sheet) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const r = data[i];
        if (!r[COL.SR] && !r[COL.DATE]) continue;
        if (String(r[COL.GEAR]).trim() === 'NIL') continue;
        const rd = r[COL.DATE];
        rows.push({
          srNo      : String(r[COL.SR] || ''),
          date      : rd instanceof Date ? Utilities.formatDate(rd, tz, 'dd/MM/yy') : String(rd || ''),
          _sortDate : rd instanceof Date ? rd.getTime() : 0,
          trainNo   : r[COL.TRAIN],
          section   : r[COL.SECTION],
          category  : _obsCategory(r) === 'Undue Braking' ? 'Undue EB' : _obsCategory(r),
          gear      : r[COL.GEAR] || '-',
          status    : r[COL.STATUS] || 'Pending',
          flag      : r[COL.FLAG] || 'None'
        });
      }
    });
    rows.sort(function(a,b){ return (b._sortDate - a._sortDate) || (Number(b.srNo)-Number(a.srNo)); });
    rows.forEach(function(r){ delete r._sortDate; });
    return rows.slice(0, 8);
  } catch (e) { return []; }
}

// ── Loco Details Sheet ───────────────────────────────
// "Loco details" sheet columns (0-indexed): S.No | Loco No | Loco Type | Zone | Shed/PU | Make
const LOCO_COL = { SNO: 0, NO: 1, TYPE: 2, ZONE: 3, SHED: 4, MAKE: 5 };

function getLocoDetailsSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName('Loco details');
  if (!sh) {
    sh = ss.insertSheet('Loco details');
    sh.appendRow(['S.No', 'Loco No', 'Loco Type', 'Zone', 'Shed/PU', 'Make']);
    sh.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  return sh;
}

function fetchLocoInfo(locoNo) {
  if (!locoNo) return { success: false, error: 'No loco number' };
  locoNo = String(locoNo).trim();

  const sh   = getLocoDetailsSheet();
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][LOCO_COL.NO]).trim() === locoNo) {
      return {
        success:  true,
        locoNo:   locoNo,
        locoType: String(rows[i][LOCO_COL.TYPE] || '').trim(),
        zone:     String(rows[i][LOCO_COL.ZONE] || '').trim(),
        shed:     String(rows[i][LOCO_COL.SHED] || '').trim(),
        make:     String(rows[i][LOCO_COL.MAKE] || '').trim()
      };
    }
  }
  return { success: false, notFound: true, error: 'Loco ' + locoNo + ' not found in Loco details sheet' };
}

// Add new loco entry to "Loco details" sheet
function addLocoInfo(data) {
  try {
    if (!data.locoNo) return { success: false, error: 'Loco No required' };
    const locoNo = String(data.locoNo).trim();
    const sh     = getLocoDetailsSheet();
    const rows   = sh.getDataRange().getValues();

    // Prevent duplicates
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][LOCO_COL.NO]).trim() === locoNo) {
        // Update existing row
        sh.getRange(i + 1, 1, 1, 6).setValues([[
          rows[i][LOCO_COL.SNO], locoNo,
          data.locoType || '', data.zone || '', data.shed || '', data.make || ''
        ]]);
        return { success: true, updated: true };
      }
    }

    // Append new row — S.No = last data row number
    const sNo = rows.length; // header is row 1, so next S.No = rows.length
    sh.appendRow([sNo, locoNo, data.locoType || '', data.zone || '', data.shed || '', data.make || '']);
    return { success: true, updated: false };
  } catch (e) { return { success: false, error: e.message }; }
}

// Bulk-import from Admin Panel → Loco Details Manager. Unlike addLocoInfo()
// (which does a full overwrite for the single-loco "fix this one" flow),
// this only FILLS BLANK fields on locos that already exist — it never
// overwrites data that's already saved — and appends locos not yet on the
// sheet. Batches the whole sheet read/write since imports can run into the
// thousands of rows.
function bulkImportLocoDetails(rows) {
  try {
    if (!Array.isArray(rows)) return { success: false, error: 'Invalid data' };
    const sh   = getLocoDetailsSheet();
    const data = sh.getDataRange().getValues(); // row 0 = header
    const rowIdxByLoco = {};
    for (let i = 1; i < data.length; i++) {
      const no = String(data[i][LOCO_COL.NO] || '').trim();
      if (no && rowIdxByLoco[no] === undefined) rowIdxByLoco[no] = i;
    }
    const newByLoco = {}; // locoNo -> merged fields, for locos not yet on the sheet
    const touchedExisting = new Set();
    const changedExisting = new Set();
    let skipped = 0;
    rows.forEach(function(r) {
      const locoNo = String(r.locoNo || '').trim();
      if (!locoNo) { skipped++; return; }
      const incoming = {
        TYPE: String(r.locoType || '').trim(),
        ZONE: String(r.zone || '').trim(),
        SHED: String(r.shed || '').trim(),
        MAKE: String(r.make || '').trim()
      };
      const rowIdx = rowIdxByLoco[locoNo];
      if (rowIdx !== undefined) {
        touchedExisting.add(rowIdx);
        ['TYPE', 'ZONE', 'SHED', 'MAKE'].forEach(function(key) {
          const col = LOCO_COL[key];
          if (!String(data[rowIdx][col] || '').trim() && incoming[key]) {
            data[rowIdx][col] = incoming[key];
            changedExisting.add(rowIdx);
          }
        });
      } else {
        const acc = newByLoco[locoNo] || (newByLoco[locoNo] = { TYPE: '', ZONE: '', SHED: '', MAKE: '' });
        ['TYPE', 'ZONE', 'SHED', 'MAKE'].forEach(function(key) { if (!acc[key] && incoming[key]) acc[key] = incoming[key]; });
      }
    });

    if (changedExisting.size && data.length > 1) {
      const block = data.slice(1).map(r => [r[LOCO_COL.TYPE], r[LOCO_COL.ZONE], r[LOCO_COL.SHED], r[LOCO_COL.MAKE]]);
      sh.getRange(2, LOCO_COL.TYPE + 1, block.length, 4).setValues(block);
    }

    let nextSNo = data.length; // header is row 1, so next S.No = data.length
    const toAppend = Object.keys(newByLoco).map(function(locoNo) {
      const f = newByLoco[locoNo];
      return [nextSNo++, locoNo, f.TYPE, f.ZONE, f.SHED, f.MAKE];
    });
    if (toAppend.length) sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, 6).setValues(toAppend);

    return {
      success: true,
      added: toAppend.length,
      updated: changedExisting.size,
      unchanged: touchedExisting.size - changedExisting.size,
      skipped: skipped
    };
  } catch (e) { return { success: false, error: e.message }; }
}

// ── Train Name Lookup ────────────────────────────────
// Train_cached columns (0-indexed): TrainNo | TrainName | From | To | CachedAt | BrcRoute | CachedSection
const TRAIN_CACHE_COL = { NO: 0, NAME: 1, FROM: 2, TO: 3, CACHED_AT: 4, BRC_ROUTE: 5, CACHED_SECTION: 6 };

function getTrainCacheSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName('Train_cached');
  if (!sh) {
    sh = ss.insertSheet('Train_cached');
    sh.appendRow(['TrainNo', 'TrainName', 'From', 'To', 'CachedAt', 'BrcRoute', 'CachedSection']);
    sh.getRange(1, 1, 1, 7).setFontWeight('bold');
  } else {
    // Migrate old 6-col sheet — add CachedSection header if missing
    if (sh.getLastColumn() < 7) sh.getRange(1, 7).setValue('CachedSection');
  }
  return sh;
}

// Append a confirmed section to the train's cached section list (max 2 unique values).
// A train passes through 2 sections — we learn them one observation at a time.
// Once both are known, the dropdown restricts to only those two.
function updateTrainSection(trainNo, section) {
  try {
    trainNo = String(trainNo || '').trim();
    section = String(section || '').trim();
    if (!trainNo || !section) return { success: false };
    const sh   = getTrainCacheSheet();
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][TRAIN_CACHE_COL.NO]).trim() === trainNo) {
        const existing = String(rows[i][TRAIN_CACHE_COL.CACHED_SECTION] || '').trim();
        const parts = existing ? existing.split(',').map(s => s.trim()).filter(Boolean) : [];
        if (!parts.includes(section)) {
          parts.push(section);
          // Cap at 2 — a train has at most 2 BRC-division sections
          const updated = parts.slice(0, 2).join(',');
          sh.getRange(i + 1, TRAIN_CACHE_COL.CACHED_SECTION + 1).setValue(updated);
        }
        return { success: true };
      }
    }
    return { success: false };
  } catch(e) { return { success: false }; }
}

function fetchTrainInfo(trainNo) {
  if (!trainNo) return { success: false, error: 'No train number' };
  trainNo = String(trainNo).trim();

  // 1. Check Train_cached sheet first
  const sh   = getTrainCacheSheet();
  const rows = sh.getDataRange().getValues();
  let cacheRowIdx = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][TRAIN_CACHE_COL.NO]).trim() === trainNo) {
      const cachedName = String(rows[i][TRAIN_CACHE_COL.NAME] || '').trim();
      if (cachedName) {
        return {
          success:       true,
          cached:        true,
          trainNo:       trainNo,
          trainName:     cachedName,
          from:          String(rows[i][TRAIN_CACHE_COL.FROM]           || '').trim(),
          to:            String(rows[i][TRAIN_CACHE_COL.TO]             || '').trim(),
          brcRoute:      [], // no longer derived — section is cached directly
          cachedSections: String(rows[i][TRAIN_CACHE_COL.CACHED_SECTION] || '').trim().split(',').map(s=>s.trim()).filter(Boolean)
        };
      }
      cacheRowIdx = i + 1;
      break;
    }
  }

  // 2. Wikipedia full-text search (action=query&list=search)
  //    This searches article BODIES, so it finds trains whose articles are titled just the
  //    name ("Avantika Superfast Express") but whose body mentions the train number.
  const TRAIN_KW = ['express','mail','rajdhani','shatabdi','superfast','intercity',
                    'passenger','duronto','humsafar','tejas','vande','garib','sampark',
                    'special','link','fast'];
  const _wikiFullSearch = function(query) {
    try {
      const url  = 'https://en.wikipedia.org/w/api.php?action=query&list=search&srsearch='
                 + encodeURIComponent(query) + '&srlimit=5&format=json';
      const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, deadline: 12 });
      const data = JSON.parse(resp.getContentText());
      return (data && data.query && data.query.search) ? data.query.search : [];
    } catch(e) { return []; }
  };

  const _pickResult = function(results) {
    for (let i = 0; i < results.length; i++) {
      const title   = results[i].title   || '';
      const snippet = results[i].snippet || '';
      const tl = title.toLowerCase();
      if (!TRAIN_KW.some(function(k){ return tl.indexOf(k) >= 0; })) continue;
      if (snippet.replace(/<[^>]+>/g, '').indexOf(trainNo) === -1) continue;
      return title.replace(/^\d+[\/\d]*\s+/, '').trim().replace(/^[–\-—:]\s*/, '');
    }
    return null;
  };

  // Pass A: "<trainNo> express indian railway"
  let trainName = _pickResult(_wikiFullSearch(trainNo + ' express indian railway'));
  // Pass B (fallback): bare number — catches trains like 12861 whose snippet
  //                    only mentions the number in a secondary context
  if (!trainName) trainName = _pickResult(_wikiFullSearch(trainNo));

  if (trainName) {
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yy HH:mm');
    if (cacheRowIdx > 0) sh.getRange(cacheRowIdx, 1, 1, 5).setValues([[trainNo, trainName, '', '', now]]);
    else sh.appendRow([trainNo, trainName, '', '', now, '', '']);
    return { success: true, cached: false, trainNo: trainNo, trainName: trainName,
             from: '', to: '', brcRoute: [], cachedSections: [] };
  }

  // 3. Not found anywhere — user types manually; saved to cache on submit
  return { success: false, notInCache: true, error: '' };
}

// ── Station Manager ───────────────────────────────────
// Sheet: STATIONS  cols: Name|Code|KM|SortOrder|Section|Type|BlockFromPrevKM|Active
const ST_COL = { NAME:0, CODE:1, KM:2, SORT:3, SECTION:4, TYPE:5, BLOCK_KM:6, ACTIVE:7 };
const ST_HEADERS = ['StationName','Code','KM','SortOrder','Section','Type','BlockFromPrev_KM','Active'];

function _getStationSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName('STATIONS');
  if (!sh) {
    sh = ss.insertSheet('STATIONS');
    sh.appendRow(ST_HEADERS);
    sh.getRange(1,1,1,ST_HEADERS.length).setFontWeight('bold').setBackground('#1A2540').setFontColor('#7B9EFF');
    sh.setFrozenRows(1);
    // Seed with VS-URN section data (ST=Surat → BRC=Vadodara/Baroda)
    // Note: BRC = Vadodara (Baroda), BH = Bharuch
    const seed = [
      ['Surat',          'ST',   266.78,  1, 'VS-URN', 'Jn/Spl',  0.00, true],
      ['Utran',          'URN',  270.31,  2, 'VS-URN', 'Spl',      3.53, true],
      ['Kosad',          'KSE',  273.61,  3, 'VS-URN', 'Spl',      3.30, true],
      ['Gothangam',      'GTX',  277.16,  4, 'VS-URN', 'Spl',      3.55, true],
      ['Sayan',          'SYN',  280.25,  5, 'VS-URN', 'Spl',      3.09, true],
      ['Kim',            'KIM',  290.11,  6, 'VS-URN', 'Spl',      9.86, true],
      ['Kosamba Jn',     'KSB',  297.89,  7, 'VS-URN', 'Jn/Spl',  7.78, true],
      ['Hathuran',       'HAT',  302.22,  8, 'VS-URN', 'D',        4.33, true],
      ['Panoli',         'PAO',  306.08,  9, 'VS-URN', 'Spl',      3.86, true],
      ['Ankleshwar Jn',  'AKV',  316.32, 10, 'VS-URN', 'Jn/Spl', 10.24, true],
      ['Bharuch Jn',     'BH',   325.56, 11, 'VS-URN', 'Jn/Spl',  9.24, true],
      ['Chavaj',         'CVJ',  0,      12, 'VS-URN', 'Spl',      0,    true],
      ['Nabipur',        'NIU',  0,      13, 'VS-URN', 'Spl',      0,    true],
      ['Varedia',        'VRE',  0,      14, 'VS-URN', 'Spl',      0,    true],
      ['Palej',          'PLJ',  0,      15, 'VS-URN', 'Spl',      0,    true],
      ['Lakodra',        'LKD',  0,      16, 'VS-URN', 'Spl',      0,    true],
      ['Miyagam Karjan', 'MYG',  0,      17, 'VS-URN', 'Jn/Spl',  0,    true],
      ['Kashipura',      'KSPR', 0,      18, 'VS-URN', 'Spl',      0,    true],
      ['Itola',          'ITA',  0,      19, 'VS-URN', 'Spl',      0,    true],
      ['Varnama',        'VRM',  0,      20, 'VS-URN', 'Spl',      0,    true],
      ['Makarpura',      'MPR',  0,      21, 'VS-URN', 'Spl',      0,    true],
      ['Vishwamitri',    'VS',   0,      22, 'VS-URN', 'Spl',      0,    true],
      ['Vadodara Jn',    'BRC',  0,      23, 'VS-URN', 'Jn/Spl',  0,    true],
    ];
    if (seed.length) sh.getRange(2,1,seed.length,ST_HEADERS.length).setValues(seed);
  }
  return sh;
}

function getStations(section) {
  try {
    const sh   = _getStationSheet();
    const rows = sh.getDataRange().getValues();
    let result = rows.slice(1)
      .map((r,i) => ({
        rowIdx:   i + 2,
        name:     String(r[ST_COL.NAME]    ||'').trim(),
        code:     String(r[ST_COL.CODE]    ||'').trim(),
        km:       Number(r[ST_COL.KM]      ||0),
        sort:     Number(r[ST_COL.SORT]    ||0),
        section:  String(r[ST_COL.SECTION] ||'').trim(),
        type:     String(r[ST_COL.TYPE]    ||'').trim(),
        blockKm:  Number(r[ST_COL.BLOCK_KM]||0),
        active:   r[ST_COL.ACTIVE] === true || String(r[ST_COL.ACTIVE]).toLowerCase() === 'true'
      }))
      .filter(r => r.name && (!section || section === 'ALL' || r.section === section));
    result.sort((a,b) => a.sort - b.sort || a.km - b.km);
    return { success: true, stations: result };
  } catch(e) { return { success: false, error: e.message }; }
}

function saveStation(data) {
  try {
    const sh   = _getStationSheet();
    const rows = sh.getDataRange().getValues();
    const row  = [
      String(data.name    ||'').trim(),
      String(data.code    ||'').trim().toUpperCase(),
      Number(data.km      ||0),
      Number(data.sort    ||0),
      String(data.section ||'VS-URN').trim(),
      String(data.type    ||'Spl').trim(),
      Number(data.blockKm ||0),
      data.active !== false
    ];
    if (data.rowIdx && data.rowIdx > 1) {
      sh.getRange(data.rowIdx, 1, 1, ST_HEADERS.length).setValues([row]);
    } else {
      sh.appendRow(row);
    }
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function deleteStation(rowIdx) {
  try {
    if (!rowIdx || rowIdx < 2) return { success: false, error: 'Invalid row' };
    _getStationSheet().deleteRow(rowIdx);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// Admin: fix legacy Section labels in STATIONS sheet — stations entered
// under the old "BRC-ST" name (pre-rename) now show up correctly under VS-URN.
function migrateStationSectionLabels() {
  try {
    const sh   = _getStationSheet();
    const rows = sh.getDataRange().getValues();
    const legacyMatches = ['BRC-ST', 'BRC ST', 'BRCST', 'BRC_ST'];
    let fixed = 0;
    for (let i = 1; i < rows.length; i++) {
      const cur = String(rows[i][ST_COL.SECTION] || '').trim().toUpperCase();
      if (legacyMatches.indexOf(cur) >= 0) {
        sh.getRange(i + 1, ST_COL.SECTION + 1).setValue('VS-URN');
        fixed++;
      }
    }
    SpreadsheetApp.flush();
    return { success: true, fixed: fixed };
  } catch(e) { return { success: false, error: e.message }; }
}

// Admin: wipe and reseed STATIONS sheet from the built-in list
function reseedStations() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const existing = ss.getSheetByName('STATIONS');
    if (existing) ss.deleteSheet(existing);
    _getStationSheet(); // recreates with seed data
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ── Bulk Train Prefetch ───────────────────────────────
// Client sends batches of train numbers; this fetches Wikipedia for each
// and caches in Train_cached. Returns per-number results.
function bulkPrefetchBatch(trainNos) {
  try {
    if (!Array.isArray(trainNos)) return { success: false, error: 'trainNos must be array' };
    const results = trainNos.map(function(num) {
      try {
        const res = fetchTrainInfo(String(num).trim());
        return { num: String(num), ok: res.success, name: res.trainName || '', cached: !!res.cached };
      } catch(e) {
        return { num: String(num), ok: false, name: '', err: e.message };
      }
    });
    return { success: true, results: results };
  } catch(e) { return { success: false, error: e.message }; }
}

// ── Custom Options (gear + reasons) ──────────────────
// Stored in CONFIG sheet as JSON strings:
//   custom_gears          → JSON array of {val, subtitle, color}
//   custom_reason_<GEAR>  → JSON array of {val, color}

// ── Gear / Reason Classification — single source of truth ────────────
// Previously the 5 "built-in" gears (LOCO KAVACH, STATION KAVACH, …) were
// hardcoded in the frontend and only admin-added extras were editable.
// Now everything lives in one CONFIG key ('gear_reason_map'), auto-seeded
// from these same defaults on first run — so the originals can be renamed,
// have reasons added/removed, or be deleted entirely, same as custom ones.
const DEFAULT_GEAR_REASONS = [
  { val:'LOCO KAVACH', subtitle:'Loco Radio, RFID Reader', color:'amber', reasons:[
    {val:'Mode degradation due to faulty loco radio', color:'amber'},
    {val:'False SPAD due to loco radio issue', color:'amber'},
    {val:'EB due to foreign RFID tag', color:'amber'},
    {val:'Mode degradation due to RF issue (Loco)', color:'amber'}
  ]},
  { val:'STATION KAVACH', subtitle:'Signal Bob, Kavach Relay, Logic, Radio', color:'lavender', reasons:[
    {val:'Mode degradation due to STCAS/RIU card or connectivity issues', color:'lavender'},
    {val:'Mode degradation due to STCAS RF Issue', color:'lavender'},
    {val:'Mode degradation at HBL-Medha boundary (BRC-BJW)', color:'lavender'},
    {val:'EB due to relay failure', color:'lavender'}
  ]},
  { val:'REMOTE INTERFACE UNIT', subtitle:'RIU Issue', color:'teal', reasons:[
    {val:'Mode degradation due to STCAS/RIU card or connectivity issues', color:'teal'},
    {val:'RIU card failure', color:'teal'}
  ]},
  { val:'STATION TO STATION COMM', subtitle:'S-S Communication Issue', color:'teal', reasons:[
    {val:'Mode degradation at HBL-Medha boundary (BRC-BJW)', color:'teal'},
    {val:'S-S communication failure', color:'teal'}
  ]},
  { val:'LP Incorrect Working', subtitle:'Wrong Operation by Loco Pilot', color:'coral', reasons:[
    {val:'LP Incorrect Working', color:'coral'}
  ]}
];

function _getGearReasonMap() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('gear_reason_map_v1');
  if (hit) { try { return JSON.parse(hit); } catch(e) {} }
  const sh   = getOrCreateSheet('CONFIG');
  const rows = sh.getDataRange().getValues();
  let mapJson = null, legacyGears = null;
  const legacyReasons = {};
  rows.slice(1).forEach(r => {
    const key = String(r[0]||'').trim(), val = String(r[1]||'').trim();
    if (key === 'gear_reason_map') mapJson = val;
    if (key === 'custom_gears') { try { legacyGears = JSON.parse(val); } catch(e) {} }
    if (key.indexOf('custom_reason_') === 0) {
      const g = key.slice('custom_reason_'.length);
      try { legacyReasons[g] = JSON.parse(val); } catch(e) {}
    }
  });
  let map;
  if (mapJson) { try { map = JSON.parse(mapJson); } catch(e) { map = null; } }
  if (!map) {
    // First run: seed from built-in defaults, then fold in anything that
    // was previously saved under the old custom_gears / custom_reason_* keys.
    map = JSON.parse(JSON.stringify(DEFAULT_GEAR_REASONS));
    if (legacyGears) legacyGears.forEach(g => {
      if (!map.find(m => m.val === g.val)) map.push({ val:g.val, subtitle:g.subtitle||'', color:g.color||'mint', reasons:[] });
    });
    Object.keys(legacyReasons).forEach(g => {
      let entry = map.find(m => m.val === g);
      if (!entry) { entry = { val:g, subtitle:'', color:'mint', reasons:[] }; map.push(entry); }
      legacyReasons[g].forEach(r => { if (!entry.reasons.find(x => x.val === r.val)) entry.reasons.push(r); });
    });
    _saveGearReasonMap(map);
  }
  cache.put('gear_reason_map_v1', JSON.stringify(map), 60);
  return map;
}

function _saveGearReasonMap(map) {
  const sh   = getOrCreateSheet('CONFIG');
  const rows = sh.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]||'').trim() === 'gear_reason_map') {
      sh.getRange(i+1, 2).setValue(JSON.stringify(map)); found = true; break;
    }
  }
  if (!found) sh.appendRow(['gear_reason_map', JSON.stringify(map)]);
  CacheService.getScriptCache().remove('gear_reason_map_v1');
}

function getCustomOptions() {
  try {
    const map = _getGearReasonMap();
    const gears = map.map(g => ({ val:g.val, subtitle:g.subtitle||'', color:g.color||'mint' }));
    const reasons = {};
    map.forEach(g => { reasons[g.val] = g.reasons || []; });
    return { success: true, gears: gears, reasons: reasons };
  } catch(e) { return { success: false, gears: [], reasons: {} }; }
}

function addCustomGear(val, subtitle, color) {
  try {
    val = (val || '').trim().toUpperCase();
    if (!val) return { success: false, error: 'Name required' };
    const map = _getGearReasonMap();
    if (map.find(g => g.val === val)) return { success: false, error: 'Gear already exists' };
    map.push({ val: val, subtitle: subtitle||'', color: color||'mint', reasons: [] });
    _saveGearReasonMap(map);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function addCustomReason(gear, val, color) {
  try {
    gear = (gear || '').trim();
    val  = (val  || '').trim();
    if (!gear || !val) return { success: false, error: 'Gear and reason required' };
    const map = _getGearReasonMap();
    let entry = map.find(g => g.val === gear);
    if (!entry) { entry = { val:gear, subtitle:'', color:'mint', reasons:[] }; map.push(entry); }
    if (!entry.reasons) entry.reasons = [];
    if (entry.reasons.find(r => r.val === val)) return { success: false, error: 'Reason already exists' };
    entry.reasons.push({ val: val, color: color || 'mint' });
    _saveGearReasonMap(map);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function removeCustomGear(val) {
  try {
    val = (val || '').trim();
    const map = _getGearReasonMap();
    const updated = map.filter(g => g.val !== val);
    if (updated.length === map.length) return { success: false, error: 'Gear not found' };
    _saveGearReasonMap(updated);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function removeCustomReason(gear, val) {
  try {
    gear = (gear || '').trim();
    val  = (val  || '').trim();
    const map = _getGearReasonMap();
    const entry = map.find(g => g.val === gear);
    if (!entry) return { success: false, error: 'Gear not found' };
    entry.reasons = (entry.reasons || []).filter(r => r.val !== val);
    _saveGearReasonMap(map);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// Rename a reason within a gear (works for defaults and custom alike).
// Note: renaming does NOT retroactively change the Reason column text on already-submitted rows.
function renameReason(gear, oldVal, newVal) {
  try {
    gear   = (gear   || '').trim();
    oldVal = (oldVal || '').trim();
    newVal = (newVal || '').trim();
    if (!gear || !oldVal || !newVal) return { success: false, error: 'Name required' };
    const map = _getGearReasonMap();
    const entry = map.find(g => g.val === gear);
    if (!entry) return { success: false, error: 'Gear not found' };
    const r = (entry.reasons || []).find(x => x.val === oldVal);
    if (!r) return { success: false, error: 'Reason not found' };
    if (newVal !== oldVal && entry.reasons.find(x => x.val === newVal)) return { success: false, error: 'A reason with that name already exists' };
    r.val = newVal;
    _saveGearReasonMap(map);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// Rename a gear and/or edit its subtitle/color (works for defaults and custom alike).
// Note: renaming does NOT retroactively change the Gear column text on already-submitted rows.
function renameGear(oldVal, newVal, subtitle, color) {
  try {
    oldVal = (oldVal || '').trim();
    newVal = (newVal || '').trim();
    if (!oldVal || !newVal) return { success: false, error: 'Name required' };
    const map = _getGearReasonMap();
    const entry = map.find(g => g.val === oldVal);
    if (!entry) return { success: false, error: 'Gear not found' };
    if (newVal !== oldVal && map.find(g => g.val === newVal)) return { success: false, error: 'A gear with that name already exists' };
    entry.val = newVal;
    if (subtitle !== undefined) entry.subtitle = subtitle;
    if (color) entry.color = color;
    _saveGearReasonMap(map);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// One-time admin migration (2026-07): adds the new gear categories /
// reasons supplied by Railway on top of the existing set — merge only,
// nothing existing is removed or overwritten (except STATION KAVACH's
// subtitle, which is updated to carry the SYN subtotal note as requested).
function applyGearReasonUpdate_202607() {
  try {
    const map = _getGearReasonMap();
    function ensureGear(name, subtitle, color) {
      let g = map.find(x => x.val === name);
      if (!g) { g = { val: name, subtitle: subtitle || '', color: color || 'mint', reasons: [] }; map.push(g); }
      if (!g.reasons) g.reasons = [];
      return g;
    }
    function ensureReason(g, val, color) {
      if (!g.reasons.find(r => r.val === val)) g.reasons.push({ val: val, color: color || g.color || 'mint' });
    }

    const loco = ensureGear('LOCO KAVACH', 'Loco Radio, RFID Reader', 'amber');
    ['Loco RF','Foreign Tag','Odo error','Consecutive tag miss','System failure'].forEach(r => ensureReason(loco, r, 'amber'));

    const interlock = ensureGear('INTERLOCKING ISSUE', 'Relay / Signal Interlocking', 'lavender');
    ['Relay pickup fault','Signal bob during cascading'].forEach(r => ensureReason(interlock, r, 'lavender'));

    ensureGear('KAVACH RELAY ISSUE', 'Kavach Relay Fault', 'coral');

    const lpw = ensureGear('LP WORKING', 'Loco Pilot Acknowledgement / Override', 'coral');
    ensureReason(lpw, 'Late Ack of mode degradation + OVRD not performed properly', 'coral');

    // STATION KAVACH already exists — merge in the 3 new reasons and
    // update its subtitle to carry the SYN subtotal note (per admin instruction).
    const stnKavach = ensureGear('STATION KAVACH', 'Signal Bob, Kavach Relay, Logic, Radio', 'lavender');
    stnKavach.subtitle = 'Signal Bob, Kavach Relay, Logic, Radio — incl. SYN subtotal of 79';
    ['Rear-end collision','Head-on collision','Low Station RF'].forEach(r => ensureReason(stnKavach, r, 'lavender'));

    _saveGearReasonMap(map);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ── Fault Code Lookup (VCC / VIC) ─────────────────────
// Sheet: FAULT_CODES  cols: ID | Category | FaultName | Hex | Dec | Notes
const FC_COL = { ID:0, CATEGORY:1, NAME:2, HEX:3, DEC:4, NOTES:5 };
const FC_HEADERS = ['ID','Category','FaultName','Hex','Dec','Notes'];

function _getFaultCodeSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName('FAULT_CODES');
  if (!sh) {
    sh = ss.insertSheet('FAULT_CODES');
    sh.appendRow(FC_HEADERS);
    sh.getRange(1,1,1,FC_HEADERS.length).setFontWeight('bold').setBackground('#1A2540').setFontColor('#7B9EFF');
    sh.setFrozenRows(1);
  }
  return sh;
}

function getFaultCodes() {
  try {
    const sh   = _getFaultCodeSheet();
    const rows = sh.getDataRange().getValues();
    const result = rows.slice(1)
      .map((r,i) => ({
        id: String(r[FC_COL.ID]||''), rowIdx: i+2,
        category: String(r[FC_COL.CATEGORY]||'').trim(),
        faultName: String(r[FC_COL.NAME]||'').trim(),
        hex: String(r[FC_COL.HEX]||'').trim(),
        dec: String(r[FC_COL.DEC]||'').trim(),
        notes: String(r[FC_COL.NOTES]||'').trim()
      }))
      .filter(r => r.faultName || r.hex || r.dec);
    return { success: true, codes: result };
  } catch(e) { return { success: false, error: e.message, codes: [] }; }
}

// data: { id (optional — update if present), category, faultName, hex, dec, notes }
function saveFaultCode(data) {
  try {
    const category  = String(data.category||'').trim().toUpperCase();
    const faultName = String(data.faultName||'').trim();
    if (!category || !faultName) return { success:false, error:'Category and fault name are required' };
    const hex   = String(data.hex||'').trim().toUpperCase();
    const dec   = String(data.dec||'').trim();
    const notes = String(data.notes||'').trim();
    const sh = _getFaultCodeSheet();
    if (data.id) {
      const rows = sh.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][FC_COL.ID]) === String(data.id)) {
          sh.getRange(i+1, 1, 1, FC_HEADERS.length).setValues([[data.id, category, faultName, hex, dec, notes]]);
          return { success:true, id:data.id };
        }
      }
      return { success:false, error:'Entry not found' };
    }
    const id = 'FC_' + Date.now() + '_' + Math.floor(Math.random()*1000);
    sh.appendRow([id, category, faultName, hex, dec, notes]);
    return { success:true, id:id };
  } catch(e) { return { success:false, error:e.message }; }
}

function deleteFaultCode(id) {
  try {
    const sh   = _getFaultCodeSheet();
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][FC_COL.ID]) === String(id)) { sh.deleteRow(i+1); return { success:true }; }
    }
    return { success:false, error:'Entry not found' };
  } catch(e) { return { success:false, error:e.message }; }
}

// Bulk import from a parsed Excel (VIC/VCC fault-code sheets).
// rows: [{category, faultName, hex, dec, notes}], mode: 'replace' | 'merge'
function bulkImportFaultCodes(rows, mode) {
  try {
    if (!Array.isArray(rows)) return { success:false, error:'Invalid data' };
    const sh = _getFaultCodeSheet();
    let existing = [];
    if (mode === 'replace') {
      const lastRow = sh.getLastRow();
      if (lastRow > 1) sh.getRange(2, 1, lastRow-1, FC_HEADERS.length).clearContent();
    } else {
      existing = sh.getDataRange().getValues().slice(1).map(r => ({
        category: String(r[FC_COL.CATEGORY]||'').trim().toUpperCase(),
        faultName: String(r[FC_COL.NAME]||'').trim(),
        hex: String(r[FC_COL.HEX]||'').trim().toUpperCase()
      }));
    }
    const seen = new Set(existing.map(e => e.category+'|'+e.faultName+'|'+e.hex));
    const toAppend = [];
    let skipped = 0;
    rows.forEach(function(r) {
      const category  = String(r.category||'').trim().toUpperCase();
      const faultName = String(r.faultName||'').trim();
      const hex       = String(r.hex||'').trim().toUpperCase();
      const dec       = String(r.dec||'').trim();
      const notes     = String(r.notes||'').trim();
      if (!category || !faultName) return;
      const key = category+'|'+faultName+'|'+hex;
      if (seen.has(key)) { skipped++; return; }
      seen.add(key);
      toAppend.push(['FC_'+Date.now()+'_'+Math.floor(Math.random()*100000)+toAppend.length, category, faultName, hex, dec, notes]);
    });
    if (toAppend.length) sh.getRange(sh.getLastRow()+1, 1, toAppend.length, FC_HEADERS.length).setValues(toAppend);
    return { success:true, imported: toAppend.length, skipped: skipped };
  } catch(e) { return { success:false, error:e.message }; }
}

// ── Sub-type options (mc / eb / nil) ──────────────────
const DEFAULT_SUBTYPES = {
  mc: [
    'Mode Change due to RF issue from Loco',
    'STCAS Failure',
    'Low Station RF (KIM-SYN)',
    'Foreign Tag',
    'No Track Profile Failure (HBL Handover)'
  ],
  eb: [
    'EB during Staff Working',
    'EB due to Loco RF Issue (SPAD done)',
    'EB due to Relay Failure',
    'LP Late Ack of Degraded Mode'
  ],
  nil: [
    'Single Tag Miss',
    'Both Tag Miss',
    'RF Issue'
  ]
};

// Sub-type items can be either a plain string (legacy, no sub-options) or
// { val, children } (adds a second-level pick under that option). Normalizing
// on every read means older saved lists keep working without a migration step.
function _normalizeSubTypeList(arr) {
  return (arr || []).map(function(x) {
    return (typeof x === 'string') ? { val: x, children: [] } : { val: x.val, children: x.children || [] };
  });
}

function getSubTypes() {
  try {
    const sh   = getOrCreateSheet('CONFIG');
    const rows = sh.getDataRange().getValues();
    let mc = null, eb = null, nil = null;
    rows.slice(1).forEach(r => {
      const key = String(r[0] || '').trim();
      const val = String(r[1] || '').trim();
      if (key === 'subtypes_mc')  { try { mc  = JSON.parse(val); } catch(e){} }
      if (key === 'subtypes_eb')  { try { eb  = JSON.parse(val); } catch(e){} }
      if (key === 'subtypes_nil') { try { nil = JSON.parse(val); } catch(e){} }
    });
    return {
      success: true,
      mc:  _normalizeSubTypeList(mc  || DEFAULT_SUBTYPES.mc),
      eb:  _normalizeSubTypeList(eb  || DEFAULT_SUBTYPES.eb),
      nil: _normalizeSubTypeList(nil || DEFAULT_SUBTYPES.nil)
    };
  } catch(e) {
    return {
      success: true,
      mc: _normalizeSubTypeList(DEFAULT_SUBTYPES.mc),
      eb: _normalizeSubTypeList(DEFAULT_SUBTYPES.eb),
      nil: _normalizeSubTypeList(DEFAULT_SUBTYPES.nil)
    };
  }
}

function saveSubTypes(data) {
  try {
    const sh   = getOrCreateSheet('CONFIG');
    const rows = sh.getDataRange().getValues();
    const keys = ['subtypes_mc', 'subtypes_eb', 'subtypes_nil'];
    const vals = { subtypes_mc: JSON.stringify(data.mc), subtypes_eb: JSON.stringify(data.eb), subtypes_nil: JSON.stringify(data.nil || []) };
    const found = { subtypes_mc: false, subtypes_eb: false, subtypes_nil: false };
    rows.slice(1).forEach((r, i) => {
      const key = String(r[0] || '').trim();
      if (keys.includes(key)) { sh.getRange(i + 2, 2).setValue(vals[key]); found[key] = true; }
    });
    keys.forEach(k => { if (!found[k]) sh.appendRow([k, vals[k]]); });
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ── Monthly Report — per section group ───────────────
// Splits a multi-value cell ("1. FS - SR\n2. FSB" or "A, B") into clean tokens.
// Commas inside parentheses do NOT split — "LOCO RADIO ISSUE (FS-SR, FS-LS)"
// is one sub-categorized value, not two broken halves.
function _rptSplitVals(v) {
  return String(v || '')
    .split(/\n|,(?![^()]*\))/)
    .map(function(s){ return s.replace(/^\s*\d+[\.\)]\s*/, '').trim(); })
    .filter(function(s){ return s && s !== 'NIL'; });
}

// Shared "special" (Journey Summary) row-matching predicate — used by both
// getReportDrillRows (drill-down / linked-category discovery) and
// getReportExcludeSubset (partial-exclude counting), so the two can never
// disagree on what counts as e.g. a "NIL — No Fault Journey".
// typeWantedLower must already be trimmed + lowercased.
function _rptSpecialMatch(r, typeWantedLower) {
  var gearRaw = String(r[COL.GEAR] || '').trim();
  var hasMc = _rptSplitVals(r[COL.MODE_DEG]).length > 0;
  var hasEb = _rptSplitVals(r[COL.EB]).length > 0;
  var flagV = String(r[COL.FLAG] || '').trim();
  if (typeWantedLower.indexOf('nil') === 0)           return gearRaw === 'NIL';
  if (typeWantedLower.indexOf('no mode') === 0)       return gearRaw !== 'NIL' && !hasMc && !hasEb;
  if (typeWantedLower.indexOf('r&d') === 0)           return flagV.indexOf('R&D Required') >= 0;
  if (typeWantedLower.indexOf('loco log') === 0)      return flagV.indexOf('Loco Log Required') >= 0;
  if (typeWantedLower.indexOf('icms') === 0)          return flagV.indexOf('ICMS Flagged') >= 0;
  return false;
}

function getMonthlyReport() {
  try {
    const ids = _getObsSheetIds();
    // Build a group list: only configured (non-empty) IDs, deduplicated
    const groupDefs = [
      { key: 'brc_gda', label: 'BRC-GDA',  sections: 'BRC-GDA / GDA-BRC / GDA-CYI (Toward BJW-ADI)',       id: ids.brc_gda },
      { key: 'brc_urn', label: 'VS-URN',  sections: 'VS-URN / URN-VS',                                       id: ids.brc_urn },
      { key: 'bjw_adi', label: 'BJW-ADI',  sections: 'BJW-ADI / ADI-BJW / BJW-ADI (Toward CYI-GDA)',          id: ids.bjw_adi }
    ];
    const seenIds = new Set();
    const groups  = groupDefs.filter(g => {
      if (!g.id) { g.id = SHEET_ID; } // fallback: backend sheet
      if (seenIds.has(g.id)) return false;
      seenIds.add(g.id);
      return true;
    });

    function _buildReport(ss) {
      const monthSh = ss.getSheets().filter(sh =>
        !_OBS_NON_MONTH.has(sh.getName()) && /^[A-Z]+-\d{4}$/.test(sh.getName())
      );
      if (!monthSh.length) return null;
      monthSh.sort((a, b) => {
        const [ma, ya] = a.getName().split('-'), [mb, yb] = b.getName().split('-');
        return new Date(Number(ya), MONTH_NAMES.indexOf(ma)) - new Date(Number(yb), MONTH_NAMES.indexOf(mb));
      });
      const months = monthSh.map(s => {
        const [m, y] = s.getName().split('-');
        return m.charAt(0) + m.slice(1).toLowerCase() + ' ' + y;
      });
      const mc = {}, eb = {}, gear = {}, reason = {};
      const mcTotal = {}, ebTotal = {}, gearTotal = {};
      months.forEach(m => { mcTotal[m]=0; ebTotal[m]=0; gearTotal[m]=0; });

      monthSh.forEach((sh, idx) => {
        const ml   = months[idx];
        const rows = sh.getDataRange().getValues();
        for (let i = 1; i < rows.length; i++) {
          const r = rows[i];
          if (!r[COL.SR] && !r[COL.DATE]) continue;
          _rptSplitVals(r[COL.MODE_DEG]).forEach(t=>{mc[t]=mc[t]||{};mc[t][ml]=(mc[t][ml]||0)+1;mcTotal[ml]++;});
          _rptSplitVals(r[COL.EB]).forEach(t=>{eb[t]=eb[t]||{};eb[t][ml]=(eb[t][ml]||0)+1;ebTotal[ml]++;});
          _rptSplitVals(r[COL.GEAR]).forEach(t=>{gear[t]=gear[t]||{};gear[t][ml]=(gear[t][ml]||0)+1;gearTotal[ml]++;});
          const gRaw = String(r[COL.GEAR]||'').trim();
          if (gRaw === 'NIL') { gear['NIL']=gear['NIL']||{};gear['NIL'][ml]=(gear['NIL'][ml]||0)+1;gearTotal[ml]++; }
          _rptSplitVals(r[COL.REASON]).forEach(t=>{reason[t]=reason[t]||{};reason[t][ml]=(reason[t][ml]||0)+1;});
        }
      });
      const toRows = obj => Object.keys(obj).map(type=>({type,counts:months.map(m=>obj[type][m]||0)}));
      return {
        months,
        modeChange:      toRows(mc),   modeChangeTotal:    months.map(m=>mcTotal[m]||0),
        undueBraking:    toRows(eb),   undueBrakingTotal:  months.map(m=>ebTotal[m]||0),
        gearAtFault:     toRows(gear), gearAtFaultTotal:   months.map(m=>gearTotal[m]||0),
        reasons:         toRows(reason)
      };
    }

    const result = { success: true, groups: [] };
    groups.forEach(g => {
      const ss   = SpreadsheetApp.openById(g.id);
      const data = _buildReport(ss);
      result.groups.push({ key: g.key, label: g.label, sections: g.sections, data });
    });
    return result;
  } catch (e) { return { success: false, error: e.message }; }
}

// ── New report function: weekly (≤2 months) or monthly (>2 months) ──────────
function getReportData(params) {
  try {
    var fromStr = (params && params.fromDate) || '';
    var toStr   = (params && params.toDate)   || '';
    var now = new Date();
    var fromDate = fromStr ? new Date(fromStr + 'T00:00:00')
                           : new Date(now.getFullYear(), now.getMonth(), 1);
    var toDate   = toStr   ? new Date(toStr   + 'T23:59:59')
                           : new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59);

    var monthList = _rptMonthsInRange(fromDate, toDate);
    var viewMode  = monthList.length <= 2 ? 'weekly' : 'monthly';

    // Build group list from dynamic section map
    var map = _getSectionMap();
    var groups = [], seen = new Set();
    map.forEach(function(g) {
      g.sheetIds.forEach(function(id) {
        if (!id || seen.has(id)) return;
        seen.add(id);
        var key = g.group.toLowerCase().replace(/[^a-z0-9]/g, '_');
        groups.push({ key: key, label: g.group, sections: g.sections.join(' / '), sheetId: id });
      });
    });
    if (!groups.length) groups.push({ key: 'main', label: 'All Sections', sections: 'All', sheetId: SHEET_ID });

    var result = { success: true, viewMode: viewMode, groups: [] };
    groups.forEach(function(grp) {
      try {
        var ss   = SpreadsheetApp.openById(grp.sheetId);
        var data = viewMode === 'weekly'
          ? _rptBuildWeekly(ss, fromDate, toDate, monthList)
          : _rptBuildMonthly(ss, fromDate, toDate, monthList);
        result.groups.push({ key: grp.key, label: grp.label, sections: grp.sections, data: data });
      } catch(e) {
        result.groups.push({ key: grp.key, label: grp.label, sections: grp.sections, data: null, error: e.message });
      }
    });
    return result;
  } catch(e) { return { success: false, error: e.message }; }
}

function _rptMonthsInRange(fromDate, toDate) {
  var months = [];
  var d = new Date(fromDate.getFullYear(), fromDate.getMonth(), 1);
  while (d <= toDate) {
    months.push({ year: d.getFullYear(), month: d.getMonth() });
    d.setMonth(d.getMonth() + 1);
  }
  return months;
}

function _rptWeekSpans(year, month, fromDate, toDate) {
  var monthStart = new Date(year, month, 1);
  var monthEnd   = new Date(year, month + 1, 0, 23, 59, 59);
  var clampFrom  = fromDate > monthStart ? fromDate : monthStart;
  var clampTo    = toDate   < monthEnd   ? toDate   : monthEnd;
  function fmt(d) { return (d.getMonth()+1) + '/' + d.getDate(); }
  var spans = [];
  var cur = new Date(clampFrom.getFullYear(), clampFrom.getMonth(), clampFrom.getDate());
  var dow = cur.getDay(); // 0=Sun
  cur.setDate(cur.getDate() - (dow === 0 ? 6 : dow - 1)); // back to Monday
  if (cur < clampFrom) cur = new Date(clampFrom.getFullYear(), clampFrom.getMonth(), clampFrom.getDate());
  var wn = 1;
  while (cur <= clampTo) {
    var wStart = new Date(cur);
    var wEnd   = new Date(cur); wEnd.setDate(wEnd.getDate() + 6); wEnd.setHours(23,59,59);
    var effEnd = wEnd < clampTo ? wEnd : new Date(clampTo);
    spans.push({ label: 'W'+wn+'\n'+fmt(wStart)+'–'+fmt(effEnd), start: new Date(wStart), end: new Date(effEnd) });
    wn++;
    cur = new Date(effEnd); cur.setDate(cur.getDate() + 1); cur.setHours(0,0,0);
  }
  return spans;
}

function _rptBuildWeekly(ss, fromDate, toDate, monthList) {
  var MN = MONTH_NAMES;
  var allSpans = [], columnGroups = [];
  monthList.forEach(function(m) {
    var spans = _rptWeekSpans(m.year, m.month, fromDate, toDate);
    var ml = MN[m.month].charAt(0) + MN[m.month].slice(1).toLowerCase() + ' ' + m.year;
    columnGroups.push({ label: ml, count: spans.length });
    spans.forEach(function(s) { allSpans.push(s); });
  });
  if (!allSpans.length) return null;

  var mc = {}, eb = {}, gear = {}, reason = {};
  var nCol = allSpans.length;
  var zero = function(){ return allSpans.map(function(){return 0;}); };
  var mcTot=zero(), ebTot=zero(), gTot=zero();
  var spNil=zero(), spNmb=zero(), spRd=zero(), spLoco=zero(), spIcms=zero(), spJourneys=zero();

  monthList.forEach(function(m) {
    var sh = ss.getSheetByName(MN[m.month] + '-' + m.year);
    if (!sh) return;
    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (!r[COL.SR] && !r[COL.DATE]) continue;
      var dv = r[COL.DATE];
      var rd = dv instanceof Date ? dv : new Date(String(dv));
      if (isNaN(rd)) continue;
      var si = -1;
      for (var j = 0; j < allSpans.length; j++) {
        if (rd >= allSpans[j].start && rd <= allSpans[j].end) { si = j; break; }
      }
      if (si < 0) continue;
      spJourneys[si]++;
      var gearRaw = String(r[COL.GEAR]||'').trim();
      var mcToks = _rptSplitVals(r[COL.MODE_DEG]);
      var ebToks = _rptSplitVals(r[COL.EB]);
      var flagV  = String(r[COL.FLAG]||'').trim();
      if (gearRaw === 'NIL') spNil[si]++;                       // NIL journey — no fault, kept out of gear table
      else if (!mcToks.length && !ebToks.length) spNmb[si]++;   // fault but no mode change / brake event
      if (flagV.indexOf('R&D Required') >= 0)      spRd[si]++;
      if (flagV.indexOf('Loco Log Required') >= 0) spLoco[si]++;
      if (flagV.indexOf('ICMS Flagged') >= 0)      spIcms[si]++;
      mcToks.forEach(function(t){ mc[t]=mc[t]||{}; mc[t][si]=(mc[t][si]||0)+1; mcTot[si]++; });
      ebToks.forEach(function(t){ eb[t]=eb[t]||{}; eb[t][si]=(eb[t][si]||0)+1; ebTot[si]++; });
      _rptSplitVals(r[COL.GEAR]).forEach(function(t){ gear[t]=gear[t]||{}; gear[t][si]=(gear[t][si]||0)+1; gTot[si]++; });
      _rptSplitVals(r[COL.REASON]).forEach(function(t){ reason[t]=reason[t]||{}; reason[t][si]=(reason[t][si]||0)+1; });
    }
  });
  function toRows(obj) {
    return Object.keys(obj).map(function(t){ return { type:t, counts:allSpans.map(function(_,i){ return obj[t][i]||0; }) }; });
  }
  return {
    columns: allSpans.map(function(s){ return s.label; }),
    columnGroups: columnGroups,
    modeChange: toRows(mc),    modeChangeTotal:   mcTot,
    undueBraking: toRows(eb),  undueBrakingTotal: ebTot,
    gearAtFault: toRows(gear), gearAtFaultTotal:  gTot,
    reasons: toRows(reason),
    special: [
      { type:'NIL — No Fault Journey',        counts: spNil  },
      { type:'No Mode / Brake Event',         counts: spNmb  },
      { type:'R&D Flagged (Cause Unknown)',   counts: spRd   },
      { type:'Loco Log Flagged',              counts: spLoco },
      { type:'ICMS Flagged',                  counts: spIcms }
    ],
    journeyTotals: spJourneys
  };
}

function _rptBuildMonthly(ss, fromDate, toDate, monthList) {
  var MN = MONTH_NAMES;
  var months = monthList.map(function(m){ return MN[m.month].charAt(0)+MN[m.month].slice(1).toLowerCase()+' '+m.year; });
  var mc = {}, eb = {}, gear = {}, reason = {};
  var zero = function(){ return months.map(function(){return 0;}); };
  var mcTot=zero(), ebTot=zero(), gTot=zero();
  var spNil=zero(), spNmb=zero(), spRd=zero(), spLoco=zero(), spIcms=zero(), spJourneys=zero();

  monthList.forEach(function(m, idx) {
    var sh = ss.getSheetByName(MN[m.month] + '-' + m.year);
    if (!sh) return;
    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      if (!r[COL.SR] && !r[COL.DATE]) continue;
      var dv = r[COL.DATE];
      var rd = dv instanceof Date ? dv : new Date(String(dv));
      if (!isNaN(rd) && (rd < fromDate || rd > toDate)) continue;
      spJourneys[idx]++;
      var gearRaw = String(r[COL.GEAR]||'').trim();
      var mcToks = _rptSplitVals(r[COL.MODE_DEG]);
      var ebToks = _rptSplitVals(r[COL.EB]);
      var flagV  = String(r[COL.FLAG]||'').trim();
      if (gearRaw === 'NIL') spNil[idx]++;
      else if (!mcToks.length && !ebToks.length) spNmb[idx]++;
      if (flagV.indexOf('R&D Required') >= 0)      spRd[idx]++;
      if (flagV.indexOf('Loco Log Required') >= 0) spLoco[idx]++;
      if (flagV.indexOf('ICMS Flagged') >= 0)      spIcms[idx]++;
      mcToks.forEach(function(t){ mc[t]=mc[t]||{}; mc[t][idx]=(mc[t][idx]||0)+1; mcTot[idx]++; });
      ebToks.forEach(function(t){ eb[t]=eb[t]||{}; eb[t][idx]=(eb[t][idx]||0)+1; ebTot[idx]++; });
      _rptSplitVals(r[COL.GEAR]).forEach(function(t){ gear[t]=gear[t]||{}; gear[t][idx]=(gear[t][idx]||0)+1; gTot[idx]++; });
      _rptSplitVals(r[COL.REASON]).forEach(function(t){ reason[t]=reason[t]||{}; reason[t][idx]=(reason[t][idx]||0)+1; });
    }
  });
  function toRows(obj) {
    return Object.keys(obj).map(function(t){ return { type:t, counts:months.map(function(_,i){ return obj[t][i]||0; }) }; });
  }
  return {
    columns: months,
    columnGroups: null,
    modeChange: toRows(mc),    modeChangeTotal:   mcTot,
    undueBraking: toRows(eb),  undueBrakingTotal: ebTot,
    gearAtFault: toRows(gear), gearAtFaultTotal:  gTot,
    reasons: toRows(reason),
    special: [
      { type:'NIL — No Fault Journey',        counts: spNil  },
      { type:'No Mode / Brake Event',         counts: spNmb  },
      { type:'R&D Flagged (Cause Unknown)',   counts: spRd   },
      { type:'Loco Log Flagged',              counts: spLoco },
      { type:'ICMS Flagged',                  counts: spIcms }
    ],
    journeyTotals: spJourneys
  };
}

// ── Report drill-down: observations behind one report row ─────────────
// params: { groupKey, kind: 'gear'|'reason'|'eb'|'mc', type, fromDate, toDate }
function getReportDrillRows(params) {
  try {
    var fromDate = new Date((params.fromDate||'1970-01-01') + 'T00:00:00');
    var toDate   = new Date((params.toDate  ||'2099-12-31') + 'T23:59:59');
    var typeWanted = String(params.type||'').trim().toLowerCase();
    var kind = String(params.kind||'gear');

    // Resolve the sheet for this group (same mapping as getReportData)
    var map = _getSectionMap();
    var sheetId = null, seen = new Set();
    map.forEach(function(g) {
      g.sheetIds.forEach(function(id) {
        if (!id || seen.has(id)) return;
        seen.add(id);
        var key = g.group.toLowerCase().replace(/[^a-z0-9]/g, '_');
        if (key === params.groupKey) sheetId = id;
      });
    });
    if (!sheetId) sheetId = SHEET_ID;
    var ss = SpreadsheetApp.openById(sheetId);

    var colFor = { gear: COL.GEAR, reason: COL.REASON, eb: COL.EB, mc: COL.MODE_DEG };
    var col = colFor[kind] !== undefined ? colFor[kind] : COL.GEAR;
    // Gear/Reason/Station/Description are written 1-per-failure-point in the same
    // order (see submitObservation), so they're reliably index-aligned — drilling
    // into one specific Gear or Reason value can isolate that exact sub-observation's
    // station/description instead of showing the whole journey's combined text.
    // Mode Degradation / Undue Braking are NOT aligned this way (only pushed for
    // failures that actually had that category), so those stay whole-row best-effort.
    var isAligned = (kind === 'gear' || kind === 'reason');

    var out = [];
    var monthList = _rptMonthsInRange(fromDate, toDate);
    monthList.forEach(function(m) {
      var sh = ss.getSheetByName(MONTH_NAMES[m.month] + '-' + m.year);
      if (!sh) return;
      var rows = sh.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        var r = rows[i];
        if (!r[COL.SR] && !r[COL.DATE]) continue;
        var dv = r[COL.DATE];
        var rd = dv instanceof Date ? dv : new Date(String(dv));
        if (!isNaN(rd) && (rd < fromDate || rd > toDate)) continue;
        var dateStr = rd instanceof Date && !isNaN(rd) ? Utilities.formatDate(rd, Session.getScriptTimeZone(), 'dd/MM/yy') : String(dv||'');
        var baseRow = {
          srNo: String(r[COL.SR]||''), date: dateStr, trainNo: String(r[COL.TRAIN]||''),
          locoNo: String(r[COL.LOCO]||''), section: String(r[COL.SECTION]||''),
          status: String(r[COL.STATUS]||'Pending'), flag: String(r[COL.FLAG]||'None'),
          // Journey Summary link: lets the report's Exclude dialog offer
          // "No Mode / Brake Event" as a linked category (same predicate the
          // report table itself uses, so counts always agree).
          special: _rptSpecialMatch(r, 'no mode') ? ['No Mode / Brake Event'] : []
        };

        if (kind === 'special') {
          if (!_rptSpecialMatch(r, typeWanted)) continue;
          out.push(Object.assign({}, baseRow, {
            station: String(r[COL.STATION]||''), desc: String(r[COL.DESC]||''),
            gear: _rptSplitVals(r[COL.GEAR]), reason: _rptSplitVals(r[COL.REASON]),
            modeDeg: _rptSplitVals(r[COL.MODE_DEG]), eb: _rptSplitVals(r[COL.EB])
          }));
          continue;
        }

        if (isAligned) {
          var alignedTokens = _splitFailureList(r[col]).map(function(t){ return t.toLowerCase(); });
          var idx = alignedTokens.indexOf(typeWanted);
          if (idx < 0) continue;
          var stationsArr = _splitFailureList(r[COL.STATION]);
          var descsArr    = _splitFailureList(r[COL.DESC]);
          var gearsArr    = _splitFailureList(r[COL.GEAR]);
          var reasonsArr  = _splitFailureList(r[COL.REASON]);
          out.push(Object.assign({}, baseRow, {
            station: (stationsArr[idx] || ''),
            desc: (descsArr[idx] || ''),
            // The matched dimension's own value is this drill's "type" already —
            // only surface the OTHER dimension's exact linked value as a cross-ref.
            gear: kind==='gear' ? [] : (gearsArr[idx] ? [gearsArr[idx]] : []),
            reason: kind==='reason' ? [] : (reasonsArr[idx] ? [reasonsArr[idx]] : []),
            // Not index-aligned — best-effort, may include other sub-observations' values
            modeDeg: _rptSplitVals(r[COL.MODE_DEG]), eb: _rptSplitVals(r[COL.EB])
          }));
          continue;
        }

        // mc/eb kind — not index-aligned, whole-row best-effort as before
        var tokens = _rptSplitVals(r[col]).map(function(t){ return t.toLowerCase(); });
        if (tokens.indexOf(typeWanted) < 0) continue;
        out.push(Object.assign({}, baseRow, {
          station: String(r[COL.STATION]||'').slice(0,60), desc: String(r[COL.DESC]||'').slice(0,220),
          gear: _rptSplitVals(r[COL.GEAR]), reason: _rptSplitVals(r[COL.REASON]),
          modeDeg: _rptSplitVals(r[COL.MODE_DEG]), eb: _rptSplitVals(r[COL.EB])
        }));
      }
    });
    return { success: true, rows: out };
  } catch(e) { return { success: false, error: e.message, rows: [] }; }
}

// ── Report partial-exclude: per-column count of {kind:type} observations
// that are ALSO linked to one of condVals in dimension condKind — e.g. "how
// many LOCO KAVACH (gear) observations per week are also tagged Both Tag
// Miss (reason)". This is what a report row's count gets reduced by when the
// user chooses "Exclude only checked cases" instead of excluding the whole
// row. Reuses the exact same week/month bucketing as getReportData /
// _rptBuildWeekly / _rptBuildMonthly (same fromDate/toDate → same viewMode,
// same _rptWeekSpans/_rptMonthsInRange) so the subtracted counts land in
// exactly the columns the report is already showing — no re-derivation, no
// risk of the two disagreeing.
// params: { groupKey, kind, type, condKind, condVals, fromDate, toDate }
function getReportExcludeSubset(params) {
  try {
    var fromStr = (params && params.fromDate) || '';
    var toStr   = (params && params.toDate)   || '';
    var now = new Date();
    var fromDate = fromStr ? new Date(fromStr + 'T00:00:00')
                           : new Date(now.getFullYear(), now.getMonth(), 1);
    var toDate   = toStr   ? new Date(toStr   + 'T23:59:59')
                           : new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59);
    var monthList = _rptMonthsInRange(fromDate, toDate);
    var viewMode  = monthList.length <= 2 ? 'weekly' : 'monthly';

    var map = _getSectionMap();
    var sheetId = null, seen = new Set();
    map.forEach(function(g) {
      g.sheetIds.forEach(function(id) {
        if (!id || seen.has(id)) return;
        seen.add(id);
        var key = g.group.toLowerCase().replace(/[^a-z0-9]/g, '_');
        if (key === params.groupKey) sheetId = id;
      });
    });
    if (!sheetId) sheetId = SHEET_ID;
    var ss = SpreadsheetApp.openById(sheetId);

    var colFor = { gear: COL.GEAR, reason: COL.REASON, eb: COL.EB, mc: COL.MODE_DEG };
    var kindCol = colFor[params.kind] !== undefined ? colFor[params.kind] : COL.GEAR;
    var condCol = colFor[params.condKind] !== undefined ? colFor[params.condKind] : COL.REASON;
    var typeWanted = String(params.type || '').trim().toLowerCase();
    var condSet = new Set((params.condVals || []).map(function(v) { return String(v || '').trim().toLowerCase(); }));

    var rowMatches = function(r) {
      if (params.kind === 'special') {
        if (!_rptSpecialMatch(r, typeWanted)) return false;
      } else {
        var kindToks = _rptSplitVals(r[kindCol]).map(function(t) { return t.toLowerCase(); });
        if (kindToks.indexOf(typeWanted) < 0) return false;
      }
      if (params.condKind === 'special') {
        return Array.from(condSet).some(function(v) { return _rptSpecialMatch(r, v); });
      }
      var condToks = _rptSplitVals(r[condCol]).map(function(t) { return t.toLowerCase(); });
      return condToks.some(function(t) { return condSet.has(t); });
    };

    var counts;
    if (viewMode === 'weekly') {
      var allSpans = [];
      monthList.forEach(function(m) {
        _rptWeekSpans(m.year, m.month, fromDate, toDate).forEach(function(s) { allSpans.push(s); });
      });
      counts = allSpans.map(function() { return 0; });
      monthList.forEach(function(m) {
        var sh = ss.getSheetByName(MONTH_NAMES[m.month] + '-' + m.year);
        if (!sh) return;
        var rows = sh.getDataRange().getValues();
        for (var i = 1; i < rows.length; i++) {
          var r = rows[i];
          if (!r[COL.SR] && !r[COL.DATE]) continue;
          var dv = r[COL.DATE];
          var rd = dv instanceof Date ? dv : new Date(String(dv));
          if (isNaN(rd)) continue;
          var si = -1;
          for (var j = 0; j < allSpans.length; j++) {
            if (rd >= allSpans[j].start && rd <= allSpans[j].end) { si = j; break; }
          }
          if (si < 0) continue;
          if (rowMatches(r)) counts[si]++;
        }
      });
    } else {
      counts = monthList.map(function() { return 0; });
      monthList.forEach(function(m, idx) {
        var sh = ss.getSheetByName(MONTH_NAMES[m.month] + '-' + m.year);
        if (!sh) return;
        var rows = sh.getDataRange().getValues();
        for (var i = 1; i < rows.length; i++) {
          var r = rows[i];
          if (!r[COL.SR] && !r[COL.DATE]) continue;
          var dv = r[COL.DATE];
          var rd = dv instanceof Date ? dv : new Date(String(dv));
          if (!isNaN(rd) && (rd < fromDate || rd > toDate)) continue;
          if (rowMatches(r)) counts[idx]++;
        }
      });
    }
    return { success: true, counts: counts };
  } catch(e) { return { success: false, error: e.message, counts: [] }; }
}

// Returns list of all available month sheet names
function getAvailableMonths() {
  try {
    const seen = new Set();
    _getAllObsMonthSheets().forEach(sh => seen.add(sh.getName()));
    return [...seen].reverse();
  } catch(e) { return []; }
}

// ── Train Mission & Observation Files ────────────────
// TRAIN_MISSIONS sheet cols (v2):
//   ID | Train No | SR No | File Type | File Name | File ID | File URL | Uploaded At
// FILE_TYPE: TRAIN_MISSION | LOCO_RSSI | LOCO_LOGS | RF_COMM
const TM = { ID:0, TRAIN:1, SR:2, TYPE:3, NAME:4, FID:5, FURL:6, AT:7 };

function _getMissionSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName('TRAIN_MISSIONS');
  if (!sh) {
    sh = ss.insertSheet('TRAIN_MISSIONS');
    sh.getRange(1,1,1,8).setValues([['ID','Train No','SR No','File Type','File Name','File ID','File URL','Uploaded At']]);
    sh.setFrozenRows(1);
    return sh;
  }
  // Migrate old 6-col sheet to 8-col (insert SR No + File Type after Train No)
  if (sh.getLastColumn() < 8) {
    const lastRow = sh.getLastRow();
    // Read all existing data
    const data = lastRow > 1 ? sh.getRange(2,1,lastRow-1,6).getValues() : [];
    // Rewrite header
    sh.getRange(1,1,1,8).setValues([['ID','Train No','SR No','File Type','File Name','File ID','File URL','Uploaded At']]);
    // Rewrite rows (shift cols 3-6 → cols 5-8, insert blanks at 3-4)
    for (let i = 0; i < data.length; i++) {
      const r = data[i]; // [ID, TrainNo, FileName, FileID, FileURL, UploadedAt]
      sh.getRange(i+2,1,1,8).setValues([[r[0],r[1],'','TRAIN_MISSION',r[2],r[3],r[4],r[5]]]);
    }
  }
  return sh;
}

// data = {trainNo, srNo, fileType, fileName, jsonStr}
// fileType: TRAIN_MISSION | LOCO_RSSI | LOCO_LOGS | RF_COMM
function uploadTrainMission(data) {
  try {
    const fileType = (data.fileType || 'TRAIN_MISSION').toString().trim();
    const trainNo  = (data.trainNo  || '').toString().trim();
    const srNo     = (data.srNo     || '').toString().trim();

    const parent = _getBrcFolder();
    const subs   = parent.getFoldersByName('TRAIN_MISSIONS');
    const folder = subs.hasNext() ? subs.next() : parent.createFolder('TRAIN_MISSIONS');
    const safeName = data.fileName.replace(/[^a-zA-Z0-9._\-]/g,'_');
    const blob   = Utilities.newBlob(data.jsonStr,'application/json', safeName+'.json');
    const file   = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const id  = 'TM_' + Date.now();
    const now = Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd/MM/yyyy HH:mm');
    const sh  = _getMissionSheet();
    sh.appendRow([id, trainNo, srNo, fileType, data.fileName, file.getId(), file.getUrl(), now]);

    return { success:true, id, fileId:file.getId(), fileUrl:file.getUrl() };
  } catch(e) { return { success:false, error:e.message }; }
}

// Stores an LM Analyzer output workbook (xlsx binary, base64) attached to an observation
// data = {trainNo, srNo, fileName, base64}
function uploadLmOutput(data) {
  try {
    const parent = _getBrcFolder();
    const subs   = parent.getFoldersByName('TRAIN_MISSIONS');
    const folder = subs.hasNext() ? subs.next() : parent.createFolder('TRAIN_MISSIONS');
    const safeName = String(data.fileName||'LM_Output.xlsx').replace(/[^a-zA-Z0-9._\- ]/g,'_');
    const blob = Utilities.newBlob(
      Utilities.base64Decode(data.base64),
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      safeName
    );
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const id  = 'LM_' + Date.now();
    const now = Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd/MM/yyyy HH:mm');
    const sh  = _getMissionSheet();
    sh.appendRow([id, String(data.trainNo||'').trim(), String(data.srNo||'').trim(), 'LM_OUTPUT', safeName, file.getId(), file.getUrl(), now]);

    return { success:true, id, fileId:file.getId(), fileUrl:file.getUrl() };
  } catch(e) { return { success:false, error:e.message }; }
}

// Returns TRAIN_MISSION files for a given train number
function getTrainMissions(trainNo) {
  try {
    const sh   = _getMissionSheet();
    const rows = sh.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      const rType = (r[TM.TYPE]||'TRAIN_MISSION').toString().trim();
      if ((r[TM.TRAIN]||'').toString().trim() === trainNo.toString().trim()
          && rType === 'TRAIN_MISSION') {
        result.push({id:r[TM.ID],trainNo:r[TM.TRAIN],srNo:r[TM.SR],fileType:rType,
                     fileName:r[TM.NAME],fileId:r[TM.FID],fileUrl:r[TM.FURL],uploadedAt:r[TM.AT]});
      }
    }
    return { success:true, missions:result };
  } catch(e) { return { success:false, error:e.message, missions:[] }; }
}

// Returns EVERY file linked to a train number, regardless of type — used by
// the Log Observation "Saved Files" chip row so RF Comm / RSSI / Loco Log
// uploads are viewable side-by-side with the Train Mission XLS for that train.
function getTrainAllFiles(trainNo) {
  try {
    const sh   = _getMissionSheet();
    const rows = sh.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if ((r[TM.TRAIN]||'').toString().trim() === trainNo.toString().trim()) {
        result.push({id:r[TM.ID],trainNo:r[TM.TRAIN],srNo:r[TM.SR],
                     fileType:(r[TM.TYPE]||'TRAIN_MISSION').toString().trim(),
                     fileName:r[TM.NAME],fileId:r[TM.FID],fileUrl:r[TM.FURL],uploadedAt:r[TM.AT]});
      }
    }
    return { success:true, files: result };
  } catch(e) { return { success:false, error:e.message, files:[] }; }
}

// Returns all non-mission files linked to a specific SR observation
function getObsFiles(srNo) {
  try {
    const sh   = _getMissionSheet();
    const rows = sh.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if ((r[TM.SR]||'').toString().trim() === srNo.toString().trim()) {
        result.push({id:r[TM.ID],trainNo:r[TM.TRAIN],srNo:r[TM.SR],
                     fileType:(r[TM.TYPE]||'TRAIN_MISSION').toString().trim(),
                     fileName:r[TM.NAME],fileId:r[TM.FID],fileUrl:r[TM.FURL],uploadedAt:r[TM.AT]});
      }
    }
    return { success:true, files:result };
  } catch(e) { return { success:false, error:e.message, files:[] }; }
}

// Fetches the JSON data of a stored mission/obs file from Drive
function getTrainMissionData(fileId) {
  try {
    const file   = DriveApp.getFileById(fileId);
    const jsonStr = file.getBlob().getDataAsString('UTF-8');
    const data   = JSON.parse(jsonStr);
    return { success:true, headers:data.headers, rows:data.rows };
  } catch(e) { return { success:false, error:e.message }; }
}

// Overwrite the JSON content of a stored Excel file (for in-dashboard editing)
// data = {fileId, jsonStr}  — creates replacement file, updates sheet record
function updateExcelFile(data) {
  try {
    const oldFile = DriveApp.getFileById(data.fileId);
    const folder  = oldFile.getParents().next();
    const blob    = Utilities.newBlob(data.jsonStr,'application/json', oldFile.getName());
    const newFile = folder.createFile(blob);
    newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Update sheet record
    const sh   = _getMissionSheet();
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if ((rows[i][TM.FID]||'').toString() === data.fileId) {
        sh.getRange(i+1, TM.FID+1).setValue(newFile.getId());
        sh.getRange(i+1, TM.FURL+1).setValue(newFile.getUrl());
        break;
      }
    }
    oldFile.setTrashed(true);
    return { success:true, newFileId:newFile.getId(), newFileUrl:newFile.getUrl() };
  } catch(e) { return { success:false, error:e.message }; }
}

// Deletes any file record by ID (works for all file types)
function deleteTrainMission(id) {
  try {
    const sh   = _getMissionSheet();
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if ((rows[i][TM.ID]||'').toString() === id.toString()) {
        try { DriveApp.getFileById(rows[i][TM.FID]).setTrashed(true); } catch(e) {}
        sh.deleteRow(i + 1);
        return { success:true };
      }
    }
    return { success:false, error:'Record not found' };
  } catch(e) { return { success:false, error:e.message }; }
}

// ── Get raw file content from Drive (for RSSI JSON blobs) ──
function getFileContent(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const content = file.getBlob().getDataAsString('UTF-8');
    return { success:true, content, name:file.getName() };
  } catch(e) { return { success:false, error:e.toString() }; }
}

// ══════════════════════════════════════════════════════
// DOCUMENT EDITOR — read / write Google Doc as HTML
// ══════════════════════════════════════════════════════

function getGeneratedDocContent(fileId) {
  try {
    const doc  = DocumentApp.openById(fileId);
    const body = doc.getBody();
    const parts = [];
    let listBuf = []; // buffered list items

    function flushList() {
      if (!listBuf.length) return;
      // detect ordered vs unordered from glyph type stored in first item
      const tag = listBuf[0].ordered ? 'ol' : 'ul';
      parts.push('<' + tag + '>' + listBuf.map(i => '<li>' + i.html + '</li>').join('') + '</' + tag + '>');
      listBuf = [];
    }

    for (let i = 0; i < body.getNumChildren(); i++) {
      const child = body.getChild(i);
      const ctype = child.getType();

      if (ctype === DocumentApp.ElementType.PARAGRAPH) {
        flushList();
        const para    = child.asParagraph();
        const heading = para.getHeading();
        const inner   = _docParaToHtml(para);
        if      (heading === DocumentApp.ParagraphHeading.HEADING1) parts.push('<h1>' + inner + '</h1>');
        else if (heading === DocumentApp.ParagraphHeading.HEADING2) parts.push('<h2>' + inner + '</h2>');
        else if (heading === DocumentApp.ParagraphHeading.HEADING3) parts.push('<h3>' + inner + '</h3>');
        else if (heading === DocumentApp.ParagraphHeading.HEADING4) parts.push('<h4>' + inner + '</h4>');
        else if (inner === '' || inner === '&nbsp;')                parts.push('<p><br></p>');
        else                                                        parts.push('<p>' + inner + '</p>');

      } else if (ctype === DocumentApp.ElementType.LIST_ITEM) {
        const item = child.asListItem();
        const g    = item.getGlyphType();
        const ordered = (g === DocumentApp.GlyphType.DECIMAL ||
                         g === DocumentApp.GlyphType.LATIN_LOWER ||
                         g === DocumentApp.GlyphType.LATIN_UPPER ||
                         g === DocumentApp.GlyphType.ROMAN_LOWER ||
                         g === DocumentApp.GlyphType.ROMAN_UPPER);
        listBuf.push({ html: _docParaToHtml(item), ordered });

      } else if (ctype === DocumentApp.ElementType.HORIZONTAL_RULE ||
                 ctype === DocumentApp.ElementType.PAGE_BREAK) {
        flushList();
        parts.push('<hr>');

      } else if (ctype === DocumentApp.ElementType.TABLE) {
        flushList();
        const table = child.asTable();
        let tHtml = '<table border="1" style="border-collapse:collapse;width:100%">';
        for (let r = 0; r < table.getNumRows(); r++) {
          tHtml += '<tr>';
          const row = table.getRow(r);
          for (let c = 0; c < row.getNumCells(); c++) {
            tHtml += '<td style="padding:6px 8px">' + (row.getCell(c).getText()||'&nbsp;') + '</td>';
          }
          tHtml += '</tr>';
        }
        tHtml += '</table>';
        parts.push(tHtml);
      }
    }
    flushList();
    return { success: true, html: parts.join('\n'), title: doc.getName() };
  } catch(e) { return { success: false, error: e.message }; }
}

function _docParaToHtml(para) {
  let html = '';
  for (let i = 0; i < para.getNumChildren(); i++) {
    if (para.getChild(i).getType() !== DocumentApp.ElementType.TEXT) continue;
    const te      = para.getChild(i).asText();
    const text    = te.getText();
    if (!text) continue;
    const indices = te.getTextAttributeIndices();
    const runs    = [...indices, text.length];
    for (let k = 0; k < runs.length - 1; k++) {
      const s   = runs[k], e = runs[k+1];
      let chunk = text.slice(s, e).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
      if (te.isUnderline(s)) chunk = '<u>' + chunk + '</u>';
      if (te.isItalic(s))    chunk = '<em>' + chunk + '</em>';
      if (te.isBold(s))      chunk = '<strong>' + chunk + '</strong>';
      html += chunk;
    }
  }
  return html;
}

function updateGeneratedDocContent(data) {
  try {
    const doc  = DocumentApp.openById(data.fileId);
    const body = doc.getBody();
    body.clear();

    // Parse the HTML into lines and rebuild the Google Doc
    const html = (data.htmlContent || '')
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<hr\s*\/?>/gi, '\n§HR§\n')
      .replace(/<\/?(ul|ol)[^>]*>/gi, '\n')
      .replace(/<h1[^>]*>([\s\S]*?)<\/h1>/gi, '\n§H1§$1§END§\n')
      .replace(/<h2[^>]*>([\s\S]*?)<\/h2>/gi, '\n§H2§$1§END§\n')
      .replace(/<h3[^>]*>([\s\S]*?)<\/h3>/gi, '\n§H3§$1§END§\n')
      .replace(/<h4[^>]*>([\s\S]*?)<\/h4>/gi, '\n§H4§$1§END§\n')
      .replace(/<li[^>]*>([\s\S]*?)<\/li>/gi, '\n§LI§$1§END§\n')
      .replace(/<\/?(p|div)[^>]*>/gi, '\n')
      .replace(/<[^>]+>/g, '')
      .replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&nbsp;/g,' ');

    const hStyle = DocumentApp.ParagraphHeading;
    html.split('\n').forEach(raw => {
      const line = raw.trim();
      if (!line) return;
      if (line === '§HR§')  { body.appendHorizontalRule(); return; }
      const h1 = line.match(/^§H1§([\s\S]*?)§END§$/); if (h1) { body.appendParagraph(h1[1]).setHeading(hStyle.HEADING1); return; }
      const h2 = line.match(/^§H2§([\s\S]*?)§END§$/); if (h2) { body.appendParagraph(h2[1]).setHeading(hStyle.HEADING2); return; }
      const h3 = line.match(/^§H3§([\s\S]*?)§END§$/); if (h3) { body.appendParagraph(h3[1]).setHeading(hStyle.HEADING3); return; }
      const h4 = line.match(/^§H4§([\s\S]*?)§END§$/); if (h4) { body.appendParagraph(h4[1]).setHeading(hStyle.HEADING4); return; }
      const li = line.match(/^§LI§([\s\S]*?)§END§$/); if (li) { body.appendListItem(li[1]); return; }
      body.appendParagraph(line);
    });

    doc.saveAndClose();
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ══════════════════════════════════════════════════════
// AI DOCUMENT GENERATION QUEUE
// Sheet: GENERATION_QUEUE
// Cols: ID|Type|Status|SR Numbers|Requested By|Requested At|Started At|Completed At|File ID|File URL|Error|Notes
// ══════════════════════════════════════════════════════
const GQ = { ID:0,TYPE:1,STATUS:2,SR_NOS:3,REQ_BY:4,REQ_AT:5,STARTED:6,COMPLETED:7,FILE_ID:8,FILE_URL:9,ERROR:10,NOTES:11,DRAFT:12,ATTACHMENTS:13 };

function _getQueueSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName('GENERATION_QUEUE');
  if (!sh) {
    sh = ss.insertSheet('GENERATION_QUEUE');
    sh.getRange(1,1,1,14).setValues([[
      'ID','Type','Status','SR Refs (JSON: srNo+section)','Requested By','Requested At',
      'Started At','Completed At','File ID','File URL','Error','Notes','Draft Content','Attachments (JSON)'
    ]]);
    sh.setFrozenRows(1);
  }
  // Migrate older 13-col sheet — add Attachments column
  if (sh.getLastColumn() < 14) {
    sh.getRange(1,14).setValue('Attachments (JSON)');
  }
  return sh;
}

// data.srRefs: [{srNo, section}] — section-qualified so cross-section SR
// collisions can't pull in the wrong observation (SR numbers are only
// unique per-section-sheet, not globally).
function addToGenerationQueue(data) {
  try {
    const refs = Array.isArray(data.srRefs) && data.srRefs.length
      ? data.srRefs
      : (data.srNos||[]).map(function(s){ return { srNo:String(s), section:'' }; }); // legacy fallback
    if (!refs.length) return { success:false, error:'Select at least one observation' };
    const sh  = _getQueueSheet();
    const id  = 'TGQ_' + Date.now();
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
    sh.appendRow([id, data.type||'MEMO', 'Queued',
      JSON.stringify(refs), data.requestedBy||'User',
      now,'','','','','', data.notes||'', '', JSON.stringify(data.attachments||[])]);
    // fire processing immediately — will no-op if another task is running
    _tryProcessQueue();
    return { success:true, taskId:id };
  } catch(e) { return { success:false, error:e.message }; }
}

// TEMPORARY DIAGNOSTIC — run this manually from the Apps Script editor
// (select debugGenQueue in the function dropdown, click Run, then View > Logs).
// Safe to delete once the queue-display issue is confirmed fixed.
function debugGenQueue() {
  const sh = _getQueueSheet();
  const raw = sh.getDataRange().getValues();
  Logger.log('--- RAW SHEET (headers + rows) ---');
  Logger.log(JSON.stringify(raw));
  Logger.log('--- Row count (excluding header): ' + (raw.length - 1));
  Logger.log('--- getGenerationQueue() result ---');
  const result = getGenerationQueue();
  Logger.log(JSON.stringify(result));
  Logger.log('--- Result length: ' + result.length);
  return result;
}

// Sheets silently reinterprets plain text that "looks like" a date/time
// (e.g. the "dd/MM/yyyy HH:mm:ss" strings this queue writes) as a real Date
// cell. Reading it back via getValues() then hands us a raw JS Date object —
// which google.script.run can fail to serialize across the client bridge,
// collapsing the ENTIRE response to null instead of throwing a catchable error.
// Every field leaving this function must be a plain string/primitive.
function _cellToStr(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  return String(v == null ? '' : v);
}

function getGenerationQueue() {
  try {
    const rows = _getQueueSheet().getDataRange().getValues();
    const out = [];
    rows.slice(1).reverse().forEach(r => {
      // One malformed legacy row (e.g. SR_NOS saved as a bare number/string
      // before the section-qualified refs format existed) must never wipe
      // out the whole queue — isolate each row's transform.
      try {
        const refs = _parseJsonArray(r[GQ.SR_NOS]);
        out.push({
          id          : _cellToStr(r[GQ.ID]),
          type        : _cellToStr(r[GQ.TYPE]),
          status      : _cellToStr(r[GQ.STATUS]),
          srRefs      : refs,
          srNos       : refs.map(function(x){ return (x && typeof x==='object') ? x.srNo : x; }), // display convenience
          requestedBy : _cellToStr(r[GQ.REQ_BY]),
          requestedAt : _cellToStr(r[GQ.REQ_AT]),
          startedAt   : _cellToStr(r[GQ.STARTED]),
          completedAt : _cellToStr(r[GQ.COMPLETED]),
          fileId      : _cellToStr(r[GQ.FILE_ID]),
          fileUrl     : _cellToStr(r[GQ.FILE_URL]),
          error       : _cellToStr(r[GQ.ERROR]),
          // Truncated for the list view — some notes have an entire pasted reference
          // doc (many KB), and shipping that for every row on every 5s poll can bloat
          // the google.script.run payload enough to fail silently. Full text isn't
          // needed here; the draft review modal fetches full content separately.
          notes       : _cellToStr(r[GQ.NOTES]).slice(0, 300),
          notesTruncated: _cellToStr(r[GQ.NOTES]).length > 300,
          attachments : _parseJsonArray(r[GQ.ATTACHMENTS]),
          hasDraft    : !!(r[GQ.DRAFT] && _cellToStr(r[GQ.DRAFT]).length > 10)
        });
      } catch(rowErr) {
        out.push({ id:_cellToStr(r[GQ.ID])||('row_err_'+out.length), type:_cellToStr(r[GQ.TYPE]), status:'Error',
          srRefs:[], srNos:[], requestedBy:_cellToStr(r[GQ.REQ_BY]), requestedAt:_cellToStr(r[GQ.REQ_AT]),
          error:'Corrupt row data: '+rowErr.message, notes:'', attachments:[], hasDraft:false });
      }
    });
    return out;
  } catch(e) { return []; }
}

// Like _parseJson, but guarantees an array is always returned — legacy rows
// may hold a bare number/string/object instead of a JSON array.
function _parseJsonArray(str) {
  const v = _parseJson(str, []);
  return Array.isArray(v) ? v : [];
}

function getDraftContent(taskId) {
  try {
    const sh = _getQueueSheet();
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][GQ.ID] === taskId) {
        return { success: true, draft: String(rows[i][GQ.DRAFT] || ''), status: rows[i][GQ.STATUS] };
      }
    }
    return { success: false, error: 'Task not found' };
  } catch(e) { return { success: false, error: e.message }; }
}

function finalizeGeneration(taskId, resolutions, labels) {
  try {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sh   = _getQueueSheet();
      const rows = sh.getDataRange().getValues();
      let taskRow = -1;
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][GQ.ID] === taskId && rows[i][GQ.STATUS] === 'Draft Ready') { taskRow = i + 1; break; }
      }
      if (taskRow < 0) return { success: false, error: 'Task not found or not in Draft Ready state' };

      const r      = rows[taskRow - 1];
      const type   = r[GQ.TYPE];
      const refs   = _parseJsonArray(r[GQ.SR_NOS]);
      const srNos  = refs.map(function(x){ return typeof x==='object' ? x.srNo : x; });
      const draft  = String(r[GQ.DRAFT] || '');
      const notes  = String(r[GQ.NOTES] || '');
      const attachments = _parseJsonArray(r[GQ.ATTACHMENTS]);
      const tz     = Session.getScriptTimeZone();
      const now    = () => Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');

      sh.getRange(taskRow, GQ.STATUS+1).setValue('Finalizing');

      // DAILY produces a plain WhatsApp-ready text message, not a Google Doc —
      // stored back into the DRAFT cell (no Drive file, no Reports registry entry).
      if (type === 'DAILY') {
        const finalText = _generateDailyFinal(draft, resolutions || '', notes);
        sh.getRange(taskRow, GQ.DRAFT+1).setValue(finalText);
        sh.getRange(taskRow, GQ.STATUS+1).setValue('Completed');
        sh.getRange(taskRow, GQ.COMPLETED+1).setValue(now());
        return { success: true, finalText: finalText };
      }

      const final  = _generateFinalFromDraft(draft, resolutions || '', type, notes);
      const saved  = _saveGeneratedDoc(final, type, taskId, srNos, attachments);

      sh.getRange(taskRow, GQ.STATUS+1).setValue('Completed');
      sh.getRange(taskRow, GQ.COMPLETED+1).setValue(now());
      sh.getRange(taskRow, GQ.FILE_ID+1).setValue(saved.fileId);
      sh.getRange(taskRow, GQ.FILE_URL+1).setValue(saved.fileUrl);

      // Register in the Reports registry — OPR gets the draft/signed/PPT lifecycle,
      // MEMO is considered final immediately (no signing step requested for MEMO).
      _createReportEntry({
        taskId: taskId, type: type, srRefs: refs,
        draftFileId: saved.fileId, draftFileUrl: saved.fileUrl,
        labels: Array.isArray(labels) ? labels : []
      });

      return { success: true, fileId: saved.fileId, fileUrl: saved.fileUrl };
    } finally { lock.releaseLock(); }
  } catch(e) { return { success: false, error: e.message }; }
}

// Client calls this every few seconds to drive the queue forward
function checkAndProcessQueue() {
  try { _tryProcessQueue(); return { success:true }; }
  catch(e) { return { success:false, error:e.message }; }
}

function _tryProcessQueue() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(300)) return; // another execution holds the lock
  try {
    const sh   = _getQueueSheet();
    const rows = sh.getDataRange().getValues();
    // abort if any task is already In Progress or Finalizing
    for (let i=1;i<rows.length;i++) {
      if (rows[i][GQ.STATUS]==='In Progress' || rows[i][GQ.STATUS]==='Finalizing') return;
    }
    // find oldest Queued (rows are appended chronologically — scan from end)
    let taskRow = -1;
    for (let i=rows.length-1;i>=1;i--) { if (rows[i][GQ.STATUS]==='Queued'){taskRow=i+1;break;} }
    if (taskRow<0) return;

    const tz  = Session.getScriptTimeZone();
    const now = () => Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');

    // Mark In Progress
    sh.getRange(taskRow, GQ.STATUS+1).setValue('In Progress');
    sh.getRange(taskRow, GQ.STARTED+1).setValue(now());

    const r     = rows[taskRow-1];
    const id    = r[GQ.ID];
    const type  = r[GQ.TYPE];
    const refs  = _parseJsonArray(r[GQ.SR_NOS]);
    const notes = r[GQ.NOTES]||'';
    const attachments = _parseJsonArray(r[GQ.ATTACHMENTS]);

    try {
      const obsData = _getObsDataForRefs(refs);
      const draft   = type==='MEMO' ? _generateMemoDraft(obsData,notes,attachments)
                    : type==='DAILY' ? _generateDailyReportDraft(obsData,notes,attachments)
                    : _generateOprDraft(obsData,notes,attachments);
      // Save draft and mark as Draft Ready — user must review before final doc is saved
      sh.getRange(taskRow, GQ.DRAFT+1).setValue(draft);
      sh.getRange(taskRow, GQ.STATUS+1).setValue('Draft Ready');
    } catch(err) {
      sh.getRange(taskRow, GQ.STATUS+1).setValue('Error');
      sh.getRange(taskRow, GQ.ERROR+1).setValue(err.message);
    }
  } finally { lock.releaseLock(); }
}

function _parseJson(str, def) {
  try { return JSON.parse(str||'null') || def; } catch(e) { return def; }
}

// A single train journey row can aggregate multiple failure points as a numbered
// list ("1. foo\n2. bar\n3. baz") in the Gear/Reason/Desc/etc columns. Removes the
// given 1-based indices and renumbers what's left. Single-value (non-numbered) text
// passes through untouched unless index 1 is excluded (which clears it).
function _filterNumberedList(text, excludeIdx) {
  text = String(text||'');
  if (!excludeIdx || !excludeIdx.length) return text;
  const lines = text.split('\n');
  const isNumbered = lines.length > 1 && /^\d+\.\s/.test(lines[0]);
  if (!isNumbered) return excludeIdx.indexOf(1) >= 0 ? '' : text;
  const kept = lines.filter(function(line, i){ return excludeIdx.indexOf(i+1) < 0; });
  if (!kept.length) return '';
  if (kept.length === 1) return kept[0].replace(/^\d+\.\s/, '');
  return kept.map(function(line, i){ return (i+1) + '. ' + line.replace(/^\d+\.\s/, ''); }).join('\n');
}

// refs: [{srNo, section, exclude?:[1,2,...]}] — section-qualified match (falls back to
// SR-only match when section is blank, for legacy queue rows created before this fix).
// `exclude` lets the user drop specific numbered sub-observations out of a multi-failure
// train journey row without dropping the whole row.
function _getObsDataForRefs(refs) {
  const obs = [];
  const wanted = refs.map(function(r){
    return { srNo:String(r.srNo||r).trim(), section:String(r.section||'').trim().toLowerCase(), exclude: r.exclude||[] };
  });
  _getAllObsMonthSheets().forEach(sh=>{
    const rows = sh.getDataRange().getValues();
    for (let i=1;i<rows.length;i++) {
      const rowSr  = String(rows[i][COL.SR]).trim();
      const rowSec = String(rows[i][COL.SECTION]||'').trim().toLowerCase();
      const w = wanted.find(function(w){ return w.srNo === rowSr && (!w.section || w.section === rowSec); });
      if (!w) continue;
      const r=rows[i];
      const ex = w.exclude;
      obs.push({
        srNo:r[COL.SR], date:r[COL.DATE], trainNo:r[COL.TRAIN], locoNo:r[COL.LOCO],
        section:r[COL.SECTION], station:_filterNumberedList(r[COL.STATION], ex),
        gear:_filterNumberedList(r[COL.GEAR], ex),
        reason:_filterNumberedList(r[COL.REASON], ex),
        subType:_filterNumberedList(r[COL.MODE_DEG], ex) || _filterNumberedList(r[COL.EB], ex) || '',
        desc:_filterNumberedList(r[COL.DESC], ex), oemStaff:r[COL.OEM_STAFF]||'', status:r[COL.STATUS],
        label:r[COL.LABEL]||'', locoType:r[COL.LOCO_TYPE]||''
      });
    }
  });
  return obs;
}

// Full, un-truncated numbered sub-items for a single observation row — used by the
// picker's "exclude sub-items" panel when a train journey has multiple failure points.
function getObsFullDetail(srNo, section) {
  try {
    const srStr = String(srNo).trim();
    const secStr = String(section||'').trim().toLowerCase();
    for (const sh of _getAllObsMonthSheets()) {
      const rows = sh.getDataRange().getValues();
      for (let i=1;i<rows.length;i++) {
        if (String(rows[i][COL.SR]).trim() !== srStr) continue;
        if (secStr && String(rows[i][COL.SECTION]||'').trim().toLowerCase() !== secStr) continue;
        const r = rows[i];
        const splitNum = function(text){
          text = String(text||'');
          const lines = text.split('\n');
          return (lines.length > 1 && /^\d+\.\s/.test(lines[0]))
            ? lines.map(function(l){ return l.replace(/^\d+\.\s/, ''); })
            : (text ? [text] : []);
        };
        const stations = splitNum(r[COL.STATION]), descs = splitNum(r[COL.DESC]);
        const n = Math.max(stations.length, descs.length);
        const items = [];
        for (let k=0;k<n;k++) items.push({ idx:k+1, station:stations[k]||'', desc:descs[k]||'' });
        return { success:true, items: items };
      }
    }
    return { success:false, error:'Observation not found', items:[] };
  } catch(e) { return { success:false, error:e.message, items:[] }; }
}

// Lean search for the Generate-Document observation picker (all sections)
function searchObservationsForGen(query, filters) {
  try {
    query = String(query||'').trim().toLowerCase();
    filters = filters || {};
    const fSection = String(filters.section||'').trim();
    const fFrom = filters.dateFrom ? new Date(filters.dateFrom+'T00:00:00') : null;
    const fTo   = filters.dateTo   ? new Date(filters.dateTo+'T23:59:59')   : null;
    if (!query && !fSection && !fFrom && !fTo) return { success:true, rows:[] };
    const out = [];
    _getAllObsMonthSheets().forEach(function(sh){
      const rows = sh.getDataRange().getValues();
      for (let i=1;i<rows.length;i++) {
        const r = rows[i];
        if (!r[COL.SR] && !r[COL.DATE]) continue;
        const srNo = String(r[COL.SR]||'');
        const section = String(r[COL.SECTION]||'');
        const trainNo = String(r[COL.TRAIN]||'');
        const desc = String(r[COL.DESC]||'');
        const rd = r[COL.DATE];
        if (fSection && section !== fSection) continue;
        if ((fFrom || fTo) && rd instanceof Date) {
          if (fFrom && rd < fFrom) continue;
          if (fTo && rd > fTo) continue;
        } else if (fFrom || fTo) {
          continue; // date filter set but row has no parseable date
        }
        if (query) {
          const hay = (srNo+' '+section+' '+trainNo+' '+r[COL.LOCO]+' '+desc).toLowerCase();
          if (hay.indexOf(query) < 0) continue;
        }
        out.push({
          srNo:srNo, section:section, trainNo:trainNo, locoNo:String(r[COL.LOCO]||''),
          date: rd instanceof Date ? Utilities.formatDate(rd, Session.getScriptTimeZone(), 'dd/MM/yy') : String(rd||''),
          _sortDate: rd instanceof Date ? rd.getTime() : 0,
          desc: desc.slice(0,140), gear:String(r[COL.GEAR]||''),
          hasMulti: /^\d+\.\s/.test(desc.trimStart()) && desc.indexOf('\n2. ')>=0
        });
        if (out.length >= 300) break;
      }
    });
    out.sort(function(a,b){ return b._sortDate - a._sortDate; });
    out.forEach(function(o){ delete o._sortDate; });
    return { success:true, rows: out.slice(0,80) };
  } catch(e) { return { success:false, error:e.message, rows:[] }; }
}

// Distinct section values across all obs sheets, cached briefly to keep the picker fast
function getObsSectionList() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('OBS_SECTION_LIST');
    if (cached) return { success:true, sections: JSON.parse(cached) };
    const set = {};
    _getAllObsMonthSheets().forEach(function(sh){
      const rows = sh.getDataRange().getValues();
      for (let i=1;i<rows.length;i++) {
        const s = String(rows[i][COL.SECTION]||'').trim();
        if (s) set[s] = true;
      }
    });
    const sections = Object.keys(set).sort();
    cache.put('OBS_SECTION_LIST', JSON.stringify(sections), 300);
    return { success:true, sections: sections };
  } catch(e) { return { success:false, error:e.message, sections:[] }; }
}

function _buildObsText(obsData) {
  return obsData.map((o,i)=>`
Observation ${i+1} (SR #${o.srNo}):
  Date        : ${o.date}
  Train No    : ${o.trainNo}   Loco No   : ${o.locoNo} (${o.locoType})
  Section     : ${o.section}   Station   : ${o.station}
  Gear        : ${o.gear}
  Failure     : ${o.subType||'-'}
  Reason      : ${o.reason||'-'}
  Description : ${o.desc||'-'}
  OEM Staff   : ${o.oemStaff||'-'}
  Label       : ${o.label||'-'}`).join('\n\n---\n');
}

// TEMPORARY DIAGNOSTIC — select this function in the dropdown next to ▶ Run and
// run it directly, then View > Logs. Confirms whether Script Properties actually
// persisted the API keys, independent of the web app / dashboard.
// Safe to delete once the DAILY_API_KEY / PPT_API_KEY issue is resolved.
function debugCheckApiKeys() {
  const props = PropertiesService.getScriptProperties();
  const keys = ['MEMO_API_KEY','OPR_API_KEY','DAILY_API_KEY','PPT_API_KEY'];
  keys.forEach(function(k) {
    const v = props.getProperty(k);
    Logger.log(k + ': ' + (v ? ('SET (' + v.length + ' chars)') : 'MISSING'));
  });
  Logger.log('--- All Script Property keys currently stored: ---');
  Logger.log(JSON.stringify(props.getKeys()));
}

// TEMPORARY SETUP HELPER — the Project Settings > Script Properties web panel has
// been unreliable at persisting newly-added rows (DAILY_API_KEY vanishing after
// clicking Run is that panel losing an unsaved/uncommitted row, not anything in
// this script). This bypasses that panel entirely with a direct, guaranteed write.
// 1. Paste your real keys into the two lines below (replace the placeholder text).
// 2. Select setDailyAndPptApiKeys in the dropdown next to ▶ Run, click Run.
// 3. Run debugCheckApiKeys() afterward to confirm both show as SET.
// 4. Delete this whole function once confirmed — don't leave real key values sitting in source.
function setDailyAndPptApiKeys() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('DAILY_API_KEY', 'PASTE_YOUR_DAILY_KEY_HERE');
  props.setProperty('PPT_API_KEY',   'PASTE_YOUR_PPT_KEY_HERE');
  Logger.log('Done. Now run debugCheckApiKeys() to confirm.');
}

// Project-memory instructions, editable from Admin Panel. Falls back to the
// built-in default if the admin hasn't set one yet. DAILY falls back to the
// full KAVACH_Failure_Report_Rules.md reference doc so it works out of the box.
function _getProjectMemoryPrompt(type) {
  const props = PropertiesService.getScriptProperties();
  const key = type==='MEMO' ? 'MEMO_SYSTEM_PROMPT' : type==='OPR' ? 'OPR_SYSTEM_PROMPT'
            : type==='DAILY' ? 'DAILY_SYSTEM_PROMPT' : 'PPT_SYSTEM_PROMPT';
  const stored = props.getProperty(key);
  if (stored) return stored;
  return type==='DAILY' ? _DAILY_DEFAULT_MEMORY : '';
}

// Seeded default Project Memory for Daily Reports — the full house-style rules doc
// (KAVACH_Failure_Report_Rules.md, BRC Division). Editable from Admin Panel; this is
// only the fallback used until the admin saves an edited version.
const _DAILY_DEFAULT_MEMORY =
`# KAVACH Failure Report — Generation Rules & Format Reference
## (For API Integration — BRC Division, Western Railway)

## 1. OUTPUT FORMAT

### Header
🗓️ KAVACH FAILURE REPORT — DD.MM.YYYY
🔷 [SECTION] | Total Trains: X

### Summary Block
▸ Summary
- Undue FSB: XX Nos
- Undue EB: XX Nos
- Undue NSB: XX Nos
- Undue Mode Change: XX Nos
- FSB, EB, NSB are counted separately.
- NSB (Normal Service Brake) must always appear just below Undue EB.
- Undue Mode Change = 1 per train per cause type (not per event).
- Speed Control FSBs are EXCLUDED from summary unless explicitly asked.
- ICMS Cases with FSB/NSB ARE included in summary count.

### Observations Block
▸ Observations

📌 *Category Name (XX Nos)*
1. In Train XXXXX (DIR)/Loco XXXXX, at [STN], [observation] due to [cause].

### Station Radio Issue Block (when applicable)
📌 *Station Radio Issue (XX Trains | XX Events)*
1. ...

▸ Station-wise Count
• STN1    — XX Nos
• STN2    — XX Nos
─────────────────────
  Total   — XX Events

## 2. SECTION NAMES
Always use in this order (never reverse): BRC-URN, BJW-ADI, BRC-GDA

## 3. COUNTING RULES
- Total Trains = All rows in data (UP+DN combined)
- Undue FSB = Count per train where FSB is a genuine undue brake event
- Undue EB = Count per train where EB is applied
- Undue NSB = Count per train where NSB is applied
- Undue Mode Change = 1 per train per cause type
- Station Radio Issue heading count = per train; table count = per station occurrence/event
- Both numbers shown as (XX Trains | XX Events)

## 4. EXCLUSIONS (ALWAYS APPLY)
- BRCY offline due to EI Remodeling — EXCLUDE entirely
- BJD STCAS offline — EXCLUDE entirely, no Note
- Speed Control (FSB/NB/NSB) — EXCLUDE by default unless explicitly asked or ICMS
- OVRD as per rule / Signal Override as per rule — EXCLUDE
- Tag Miss (standard cases) — EXCLUDE unless systemic site issue or explicitly asked
- Rear End SOS — INCLUDE unless user says remove
- CYI STCAS offline — EXCLUDE
- NIL rows — EXCLUDE
- Unidentified observations where "No issues found from station side, loco log needed" — REMOVE
- Under Analysis with no identified cause — REMOVE unless EB/FSB involved

## 5. OBSERVATION CATEGORIES (use these exact category names)
SPAD Detection; Rear End Collision; Foreign RFID; Foreign RFID — Loco Reader Issue; S2S Communication Issue;
S2S Offline — [STN1]-[STN2]; Signal/Interlocking Issue; Signal/Interlocking Issue — Signal Bobbing;
Signal/Interlocking Issue — S2S/RRI; Signal/Interlocking Issue — SOC; Signal/Interlocking Issue — TPR Drop;
Signal/Interlocking Issue — C46TPR Drop; Signal/Interlocking Issue — Relay [X] Contact Trouble;
Station Tower Radio Offline; Station Radio Issue; Station Offline; Station Offline — IPS Work / OFC Work / Fault Rectification;
Station Handover Issue; Loco Radio Issue; Loco Side Issue; Loco Issue — Consecutive Tag Miss;
Loco Issue — Continuous Tag Miss; Direction Invalid — RFID Reader Issue; Direction Invalid — SLIP;
Track Profile Issue; LTCAS Issue; LTCAS System Failure; KAVACH System Failure; KAVACH Isolation;
On Board Odometer Error; FSB — Late ACK by LP; FSB — No/Late ACK by LP; FSB — Under Analysis;
FSB — Sudden MA Fluctuation; FSB — LP Operation; FSB — Loco Near Danger Signal; FSB — EOA;
FSB & EB — Late/No ACK by LP; NSB — Speed Control; LP Operation; STCAS Issue; STCAS Issue — SR Authority;
STCAS Issue — CYI Offline; Rollback Detected; ICMS Case — FSB due to PSR; ICMS Case — NSB due to PSR;
ICMS Case — Relay [X] Contact Trouble; ICMS Case — FSB due to No ACK by LP; Adjustment Tag Miss — BH, AKV & PAO;
Adjustment Tag Miss — BH & AKV; Both Tag Miss — Multiple Stations; Signal Override; Under Analysis

## 6. OBSERVATION FORMAT RULES
Standard line: "In Train XXXXX (DIR)/Loco XXXXX, at [STN], [observation] due to [cause]."
Always use "due to" in every observation.
OEM Branding: add (HBL) or (Medha) or (KERNEX) after loco number when specified in data.
CCB notation: if loco is FIT (CCB), mention (CCB) after loco number when relevant.
Station abbreviations: use standard station codes from data as-is (URN, BH, AKV, PAO, KIM, SYN, GTX, KSE, etc.)

## 7. SPECIFIC OBSERVATION RULES
- Track Profile Issue: include wrong sequence notation e.g. 4-2-2-4 vs station sent 4-4-4
- Rear End Collision: include TIN numbers and distance between locos when available — "distance between locos XXX m after deducting ahead train length (XXX m)"
- Station Radio Issue (RSSI): below -70dB = station radio issue (NOT loco fault), do not include in loco fault observations. Single Radio Down pattern: ~50% reception at all stations in tight band = one radio failed.
- Signal Bobbing: include bobbing duration (ms) and signal combination causing invalid aspect; use G-Y for signal combination (not DECR-HECR)
- C46TPR Drop (WAGO): label "Signal/Interlocking Issue — C46TPR Drop (WAGO Cable Fault)"; include SOC location, loco trip details
- S2S / RRI Issues: "station sent RED aspect through RRI while loco travelling towards [signal]; as per R&D team analysis"
- Adjustment Tag Miss: R-649 (missed after R-651) at BH; R-345 (missed after R-347) at AKV; R-213 (missed after R-215) at PAO — label "Adjustment Tag Miss" NOT "Both Tag Miss"; group all three stations in one heading when applicable
- ICMS Cases: only mark as ICMS when explicitly confirmed/identified; label "ICMS Case — [description]"; include in FSB summary count
- NSB: always include in summary block below Undue EB; label observations "NSB — Speed Control"
- LP Operation: STANDBY by LP with no KAVACH issue = "LP Operation"; note "Nothing wrong occurred but LP did STANDBY at [STN]"
- OFC / NMS Data Not Available: add note at bottom "*Note: NMS data not generated due to [STN] station offline from OFC cut.*"
- Direction Invalid: due to RFID reader = "Direction Invalid — RFID Reader Issue"; due to wheel slip = "Direction Invalid — SLIP"; due to OnBoardOdoError = "On Board Odometer Error"

## 8. SPECIAL NOTES FORMAT
"*Note: [note text]*"
Common notes: OFC cut — "NMS data not generated due to NIU station offline from OFC cut."; Data not available — "NMS data not available for [STN]-[STN] section due to cable cut."; Data pump pending — "NMS data not available for Train Nos. X–Y; data pump pending from Stationary KAVACH post NMS issue rectification."

## 9. SECTION-SPECIFIC SUMMARY FORMAT
BRC-URN / BJW-ADI: Undue FSB, Undue EB, Undue NSB, Undue Mode Change
BRC-GDA: Undue Brake Application (instead of Undue FSB), Undue NSB, Undue EB, Undue Mode Change

## 10. KNOWN SYSTEMIC ISSUES (standing exclusions)
BRCY offline — EI Remodeling: EXCLUDE all trains with only this reason.
BJD STCAS offline: EXCLUDE, no Note.
CYI STCAS offline — EI interfacing: include only if mode degradation observed.
Adjustment tags at BH/AKV/PAO: include only when explicitly asked or systemic.
GER-BJD S2S/RRI issue: include as "Signal/Interlocking Issue — S2S/RRI".
PLJ station radio failure: include as "Station Radio Issue — PLJ Station Radios Failure".
GDA-KIZ S2S offline: include as "S2S Offline — GDA-KIZ".

## 11. DIAGNOSTIC PATTERNS
All stations ~50% reception in tight band → Single Radio Completely Failed.
Non-uniform varying reception below 95% → Intermittent Radio Fault.
Mode FS-SR, cause = station provided Profile ID but loco didn't process → Loco Side Issue.
Mode = Failure, EB, EmergencyStatus = No Emergency, Abs Loc freezes → LTCAS Failure.
Signal crossed on green, R-tags not read, Abs Loc not updated → TRIP/SPAD from packet loss.
Loco read tag on different track → Foreign RFID (correct KAVACH safety behavior).
LDoubtOver/Under crossing 60 → SR mode — Under Analysis or Track Profile.
FSB = BrakeStatus, EmergencyStatus = Unusual Stoppage in Block Section, AckInfo = NoAck past 7 sec → FSB — Late ACK by LP.
RSSI below -70dB → Station Radio Issue.
Frame offset increasing, station sending but loco not receiving → LTCAS Issue / Loco Radio Issue.
3+ consecutive tag miss → Loco Issue — Consecutive Tag Miss.
LP crossed signal without selecting override and waiting 2 min → SPAD Detection.
MA fluctuation (e.g. 1035→1760→752) → FSB — Sudden MA Fluctuation.

## 12. OUTPUT REQUIREMENTS
1. Maintain category order: EB categories first → FSB categories → NSB → Mode Change categories.
2. FSB and EB observations always listed first within category.
3. Never include Speed Control, OVRD as per rule, Signal Override as per rule, standard Tag Miss, or Rear End SOS (unless causing EB/FSB) as observations by default.
4. Combine same-cause trains in one observation point where logical.
5. Station-wise count table mandatory for Station Radio Issue.
6. Total Trains = all rows in input data (UP + DN combined).
7. NIL output is valid: if no genuine failures, write "▸ Observations\\nNIL".
8. Loco radio issue and loco communication issue are the same category — always use "Loco Radio Issue".

## 13. ICMS vs NON-ICMS DISTINCTION
Only mark as ICMS when explicitly confirmed by user or clearly identified from context. Do NOT auto-classify Speed Control FSBs as ICMS. When marked ICMS, use heading "ICMS Case — [description]". ICMS cases are included in FSB/NSB summary count.

## 14. EXAMPLE OUTPUT
🗓️ KAVACH FAILURE REPORT — DD.MM.YYYY
🔷 BRC-URN | Total Trains: 42

▸ Summary
- Undue FSB: 02 Nos
- Undue EB: 01 Nos
- Undue NSB: 00 Nos
- Undue Mode Change: 05 Nos

▸ Observations

📌 *SPAD Detection (01 Nos)*
1. In Train 22908 (UP)/Loco 30507, at ADI, SPAD detected; EB applied and mode degraded OS-TRIP; LP crossed signal S178 which was already at Red.

📌 *FSB — Late ACK by LP (01 Nos)*
1. In Train 12480 (DN)/Loco 37515, at VS, FSB applied due to Late ACK by LP for FS-SR mode change.

📌 *Loco Radio Issue (02 Nos)*
1. In Train 12953 (DN)/Loco 37602, at KIM & KSB, mode degraded FS-LS-SR at both stations due to momentary packet loss from both loco radios.
2. In Train 19015 (DN)/Loco 30606, at URN-KSE-GTX, mode degraded FS-SR due to loco communication issue.

📌 *Station Radio Issue (03 Trains | 03 Events)*
1. In Train 12218 (UP)/Loco 37215, at SYN, mode degraded FS-LS due to Station radio issue.
2. In Train 12912 (UP)/Loco 37503, at SYN, mode degraded FS-SR due to Station radio issue.
3. In Train 22717 (UP)/Loco 39248, at SYN, mode degraded FS-SR due to Station radio issue.

▸ Station-wise Count
• SYN     — 03 Nos
─────────────────────
  Total   — 03 Events

## 15. KEY TERMINOLOGY STANDARDS
Contact issue → "contact trouble" (NOT "poor contact")
KAVACH Brake types → FSB, EB, NSB (not "brakes applied")
Profile ID wrong sequence → "wrong sequence (X-Y-Z) while station sent A-B-C"
Signal bobbing → "bobbed for XXX ms"
LRP not received → "LRP not received by station"
Station sending OS → "station sending OS packets"
Loco not receiving → "loco unable to receive"
MA fluctuation → "MA fluctuation (X→Y→Z)"
Distance between locos → "distance between locos XXX m"
Tag to tower distance → "tag to radio tower distance XXXX m"`;

// Structural formatting contract for the Daily Report — output must be plain text
// ready to paste directly into WhatsApp (single-tap copy from the dashboard), NOT a
// Google Doc. Different from _DOC_FORMATTING_RULES, which targets DocumentApp markup.
const _DAILY_FORMATTING_RULES =
  'OUTPUT FORMATTING CONTRACT — this text is copied verbatim into WhatsApp, not rendered as a document:\n' +
  '- Plain text only. NO markdown code fences (no ``` blocks), no # headings, no ** double-asterisk bold.\n' +
  '- Use a SINGLE asterisk for bold (WhatsApp bold syntax), exactly as shown in the Project Memory examples, e.g. *Category Name*.\n' +
  '- Keep the emoji/bullet structure (🗓️ 🔷 ▸ 📌 • ─) and line breaks exactly as specified in the Project Memory.\n' +
  '- Output ONLY the final message text — no preamble, no "Here is the report", no trailing commentary.';

// Structural markup contract the doc-builder (_saveGeneratedDoc) can actually render.
// This is separate from the admin's Project Memory (which governs CONTENT/wording rules) —
// this governs the TECHNICAL syntax needed so tables/subject/signatures render as real
// Word-style elements instead of being flattened into plain text.
const _DOC_FORMATTING_RULES =
  'OUTPUT FORMATTING CONTRACT (follow exactly — this output is parsed by code, not read by a human as raw text):\n' +
  '- For ANY tabular data (fault summary, mode degradation, tag miss, RF communication, etc.) use standard ' +
  'GitHub-flavored Markdown tables: a header row "| Col1 | Col2 |", a separator row "|---|---|", then data rows. ' +
  'Do not use any other table format.\n' +
  '- Wrap the document Subject line exactly as: §SUBJECT§Your subject text here§END§\n' +
  '- If (and only if) signatory names/designations are known from the Project Memory or context, end the document ' +
  'with a 3-column signature block wrapped exactly as:\n' +
  '§SIGNATURES§\n' +
  'Name One (Designation One)\n' +
  'Name Two (Designation Two)\n' +
  'Name Three (Designation Three)\n' +
  '§END§\n' +
  '(omit this block entirely if signatories are not known — never invent names)\n' +
  '- Use #, ##, ### for section headings and **text** for bold. Do not use any other markup.';

function _attachmentsNote(attachments) {
  if (!attachments || !attachments.length) return '';
  return '\n\nAttached reference files (' + attachments.length + '): ' +
    attachments.map(function(a){ return a.fileName; }).join(', ') +
    '. These are attached as an appendix to the final document for the reviewer\'s reference — ' +
    'describe/cross-check against them where the filename suggests relevance (e.g. NMS screenshot, loco log).';
}

function _generateMemoDraft(obsData, notes, attachments) {
  const props  = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('MEMO_API_KEY');
  if (!apiKey) throw new Error('MEMO_API_KEY not set in Script Properties');
  const custom = _getProjectMemoryPrompt('MEMO');
  const system = (custom ? custom + '\n\n' : '') +
    'You are a senior KAVACH railway signalling engineer drafting a Loco MEMO. ' +
    'Structure it with: (1) Header with MEMO No & Date, (2) Loco & Train Details, (3) Nature of Fault, ' +
    '(4) Detailed Description, (5) Technical Analysis, (6) Recommended Action & Corrective Measures. ' +
    'Use Indian Railways terminology. Be precise and factual. ' +
    'If multiple observations are provided, cover each as its own numbered item within the relevant sections. ' +
    'IMPORTANT: If you notice any data discrepancy or missing information that needs clarification ' +
    '(e.g. NMS screenshot may differ from observation record), insert a conflict marker exactly like this:\n' +
    '⚠ CONFLICT: [describe the discrepancy clearly — what the sheet says vs what may differ] — Please confirm which is correct.\n' +
    'Insert these markers inline at the relevant point in the document. The user will review and resolve them before the final document is saved.\n\n' +
    _DOC_FORMATTING_RULES;
  const user = `Generate a DRAFT Loco MEMO for the following KAVACH system observation(s):\n\n${_buildObsText(obsData)}\n\nAdditional Context: ${notes||'None'}${_attachmentsNote(attachments)}`;
  return _callClaudeAPI(apiKey, system, user);
}

function _generateOprDraft(obsData, notes, attachments) {
  const props  = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('OPR_API_KEY');
  if (!apiKey) throw new Error('OPR_API_KEY not set in Script Properties');
  const custom = _getProjectMemoryPrompt('OPR');
  const system = (custom ? custom + '\n\n' : '') +
    'You are a senior KAVACH railway signalling engineer drafting a One Page Report (OPR). ' +
    'Structure it with: (1) Executive Summary, (2) Incident Details (date/loco/location), ' +
    '(3) Technical Findings & Root Cause Analysis, (4) Impact Assessment, ' +
    '(5) Corrective Actions Taken, (6) Recommendations & Preventive Measures. ' +
    'If multiple observations are provided, summarize each briefly within the relevant sections — keep the overall report concise. ' +
    'Keep it to one page — concise, clear, and professional. Use Indian Railways format. ' +
    'IMPORTANT: If you notice any data discrepancy or missing information that needs clarification ' +
    '(e.g. NMS screenshot may differ from observation record), insert a conflict marker exactly like this:\n' +
    '⚠ CONFLICT: [describe the discrepancy clearly — what the sheet says vs what may differ] — Please confirm which is correct.\n' +
    'Insert these markers inline at the relevant point in the document. The user will review and resolve them before the final document is saved.\n\n' +
    _DOC_FORMATTING_RULES;
  const user = `Generate a DRAFT OPR for the following KAVACH system observation(s):\n\n${_buildObsText(obsData)}\n\nAdditional Context: ${notes||'None'}${_attachmentsNote(attachments)}`;
  return _callClaudeAPI(apiKey, system, user);
}

function _generateDailyReportDraft(obsData, notes, attachments) {
  const props  = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('DAILY_API_KEY') || props.getProperty('MEMO_API_KEY') || props.getProperty('OPR_API_KEY');
  if (!apiKey) throw new Error('No API key configured (DAILY_API_KEY, MEMO_API_KEY, or OPR_API_KEY)');
  const custom = _getProjectMemoryPrompt('DAILY');
  const system = custom + '\n\n' +
    'You are a senior KAVACH railway signalling engineer compiling today\'s failure observation report. ' +
    'Follow the Project Memory rules above exactly — category names, exclusions, counting rules, summary format, ' +
    'and section-specific summary wording all matter and must be followed precisely. ' +
    'This is a DRAFT for the user to review — if you are uncertain how to classify an observation or the data ' +
    'looks incomplete, insert a conflict marker exactly like this:\n' +
    '⚠ CONFLICT: [describe the uncertainty] — Please confirm.\n' +
    'The user will review and can remove/edit observations before the final message is produced.\n\n' +
    _DAILY_FORMATTING_RULES;
  const user = `Compile the daily KAVACH failure report from the following observation(s):\n\n${_buildObsText(obsData)}\n\nAdditional Context: ${notes||'None'}${_attachmentsNote(attachments)}`;
  return _callClaudeAPI(apiKey, system, user);
}

function _generateDailyFinal(draft, edits, notes) {
  const props  = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('DAILY_API_KEY') || props.getProperty('MEMO_API_KEY') || props.getProperty('OPR_API_KEY');
  if (!apiKey) throw new Error('No API key configured (DAILY_API_KEY, MEMO_API_KEY, or OPR_API_KEY)');
  const custom = _getProjectMemoryPrompt('DAILY');
  const system = custom + '\n\n' +
    'You are a senior KAVACH railway signalling engineer finalizing today\'s failure observation report. ' +
    'You will receive a draft report and the user\'s edit instructions (which observations to remove, correct, ' +
    'reword, or reclassify). Apply those instructions exactly, remove any remaining ⚠ CONFLICT markers, ' +
    'recompute the Summary counts to match the final observation list, and produce the clean final message.\n\n' +
    _DAILY_FORMATTING_RULES;
  const user =
    `DRAFT REPORT:\n${draft}\n\n` +
    `USER EDITS / REMOVALS:\n${edits || 'None — finalize as-is, just remove any ⚠ CONFLICT markers.'}\n\n` +
    `Additional Notes: ${notes || 'None'}\n\n` +
    'Now produce the final WhatsApp-ready message with all edits applied.';
  return _callClaudeAPI(apiKey, system, user);
}

function _generateFinalFromDraft(draft, resolutions, type, notes) {
  const props  = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty(type === 'MEMO' ? 'MEMO_API_KEY' : 'OPR_API_KEY');
  if (!apiKey) throw new Error((type==='MEMO'?'MEMO':'OPR') + '_API_KEY not set in Script Properties');
  const custom = _getProjectMemoryPrompt(type);
  const system = (custom ? custom + '\n\n' : '') +
    'You are a senior KAVACH railway signalling engineer. ' +
    'You will receive a draft document that may contain ⚠ CONFLICT markers where data discrepancies were found. ' +
    'The user has provided resolutions for those conflicts. ' +
    'Your task: rewrite the document as a clean, final, professional version — ' +
    'incorporate the user\'s conflict resolutions, remove all ⚠ CONFLICT markers, and produce a polished final ' +
    (type === 'MEMO' ?
      'Loco MEMO using Indian Railways format.' :
      'One Page Report (OPR) using Indian Railways format — keep it to one page.') +
    '\n\n' + _DOC_FORMATTING_RULES;
  const user =
    `DRAFT DOCUMENT:\n${draft}\n\n` +
    `USER CONFLICT RESOLUTIONS:\n${resolutions || 'No conflicts were flagged — clean up and finalize as-is.'}\n\n` +
    `Additional Notes: ${notes || 'None'}\n\n` +
    'Now produce the final clean document with all conflicts resolved.';
  return _callClaudeAPI(apiKey, system, user);
}

function _generateMemoDoc(obsData, notes) {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('MEMO_API_KEY');
  if (!apiKey) throw new Error('MEMO_API_KEY not set in Script Properties');
  const sysProp = props.getProperty('MEMO_SYSTEM_PROMPT');
  const system  = sysProp ||
    'You are a senior KAVACH railway signalling engineer. Generate a formal, professional Loco MEMO document. ' +
    'Structure it with: (1) Header with MEMO No & Date, (2) Loco & Train Details, (3) Nature of Fault, ' +
    '(4) Detailed Description, (5) Technical Analysis, (6) Recommended Action & Corrective Measures. ' +
    'Use Indian Railways terminology. Be precise and factual.';
  const user = `Generate a Loco MEMO for the following KAVACH system observation(s):\n\n${_buildObsText(obsData)}\n\nAdditional Context: ${notes||'None'}`;
  return _callClaudeAPI(apiKey, system, user);
}

function _generateOprDoc(obsData, notes) {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('OPR_API_KEY');
  if (!apiKey) throw new Error('OPR_API_KEY not set in Script Properties');
  const sysProp = props.getProperty('OPR_SYSTEM_PROMPT');
  const system  = sysProp ||
    'You are a senior KAVACH railway signalling engineer. Generate a formal One Page Report (OPR). ' +
    'Structure it with: (1) Executive Summary, (2) Incident Details (date/loco/location), ' +
    '(3) Technical Findings & Root Cause Analysis, (4) Impact Assessment, ' +
    '(5) Corrective Actions Taken, (6) Recommendations & Preventive Measures. ' +
    'Keep it to one page — concise, clear, and professional. Use Indian Railways format.';
  const user = `Generate an OPR for the following KAVACH system observation(s):\n\n${_buildObsText(obsData)}\n\nAdditional Context: ${notes||'None'}`;
  return _callClaudeAPI(apiKey, system, user);
}

function _callClaudeAPI(apiKey, systemPrompt, userPrompt) {
  const payload = {
    model  : 'claude-sonnet-4-6',
    max_tokens : 4096,
    system : [{ type:'text', text:systemPrompt, cache_control:{type:'ephemeral'} }],
    messages   : [{ role:'user', content:userPrompt }]
  };
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method : 'POST',
    headers: {
      'x-api-key'         : apiKey,
      'anthropic-version' : '2023-06-01',
      'content-type'      : 'application/json',
      'anthropic-beta'    : 'prompt-caching-2024-07-31'
    },
    payload           : JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const body = JSON.parse(res.getContentText());
  if (code!==200) throw new Error('Claude API ' + code + ': ' + (body.error&&body.error.message||JSON.stringify(body)));
  return body.content[0].text;
}

// Parses the AI output (markdown tables, §SUBJECT§, §SIGNATURES§, #/##/### headings, **bold**)
// and appends real Google Docs elements (tables, styled paragraphs) into `body`.
function _renderMarkupIntoBody(body, content, hStyle, FONT, FSIZE) {
  const lines = content.split('\n');
  let i = 0;
  while (i < lines.length) {
    const raw = lines[i];
    const t = raw.trimStart();

    // §SUBJECT§...§END§ (may span the rest of the line or be inline)
    const subjMatch = t.match(/^§SUBJECT§(.*?)§END§\s*$/);
    if (subjMatch) {
      body.appendParagraph(subjMatch[1].trim()).setBold(true).setUnderline(true).setFontFamily(FONT).setFontSize(FSIZE);
      i++; continue;
    }

    // §SIGNATURES§ ... §END§ block -> borderless 3-col table
    if (t === '§SIGNATURES§') {
      const sigLines = [];
      i++;
      while (i < lines.length && lines[i].trim() !== '§END§') { sigLines.push(lines[i].trim()); i++; }
      i++; // skip §END§
      _appendSignatureBlock(body, sigLines.filter(function(l){return l;}), FONT, FSIZE);
      continue;
    }

    // Markdown table: header row, separator row, then data rows
    if (t.startsWith('|') && i + 1 < lines.length && /^\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?\s*$/.test(lines[i+1].trim())) {
      const tableLines = [t];
      let j = i + 2; // skip header + separator
      while (j < lines.length && lines[j].trim().startsWith('|')) { tableLines.push(lines[j].trim()); j++; }
      const headerCells = _splitMdRow(t);
      const dataRows = tableLines.slice(1).map(_splitMdRow);
      _appendStyledTable(body, headerCells, dataRows, FONT, FSIZE);
      i = j;
      continue;
    }

    if (t.startsWith('### '))      body.appendParagraph(t.slice(4)).setHeading(hStyle.HEADING3).setFontFamily(FONT);
    else if (t.startsWith('## '))  body.appendParagraph(t.slice(3)).setHeading(hStyle.HEADING2).setFontFamily(FONT);
    else if (t.startsWith('# '))   body.appendParagraph(t.slice(2)).setHeading(hStyle.HEADING1).setFontFamily(FONT);
    else if (t.startsWith('**') && t.endsWith('**') && t.length > 3) body.appendParagraph(t.replace(/\*\*/g,'')).setBold(true).setFontFamily(FONT).setFontSize(FSIZE);
    else body.appendParagraph(raw).setFontFamily(FONT).setFontSize(FSIZE);
    i++;
  }
}

function _splitMdRow(line) {
  let s = line.trim();
  if (s.startsWith('|')) s = s.slice(1);
  if (s.endsWith('|')) s = s.slice(0, -1);
  return s.split('|').map(function(c){ return c.trim().replace(/\*\*/g,''); });
}

// Plain black-and-white table: black 1pt borders, white background, bold header row (no fill color)
function _appendStyledTable(body, headerCells, dataRows, FONT, FSIZE) {
  const allRows = [headerCells].concat(dataRows);
  const table = body.appendTable(allRows);
  table.setBorderColor('#000000').setBorderWidth(1);
  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    for (let c = 0; c < row.getNumCells(); c++) {
      const cell = row.getCell(c);
      cell.setBackgroundColor('#FFFFFF');
      const txt = cell.editAsText();
      txt.setFontFamily(FONT).setFontSize(FSIZE);
      if (r === 0) txt.setBold(true);
    }
  }
}

// Borderless 3-column signature row (each line = "Name (Designation)")
function _appendSignatureBlock(body, sigLines, FONT, FSIZE) {
  if (!sigLines.length) return;
  while (sigLines.length < 3) sigLines.push('');
  const table = body.appendTable([sigLines.slice(0, 3)]);
  table.setBorderWidth(0);
  const row = table.getRow(0);
  for (let c = 0; c < row.getNumCells(); c++) {
    const cell = row.getCell(c);
    const txt = cell.editAsText();
    txt.setFontFamily(FONT).setFontSize(FSIZE).setBold(true);
    cell.setPaddingTop(24);
  }
}

// Root "GENERATED" folder inside BRC_KAVACH_DOCS that holds all AI-generated /
// uploaded Memo, OPR, and PPT files — organized Year > Month > Day underneath.
function _getGeneratedRootFolder() {
  const parent = _getBrcFolder();
  const subs = parent.getFoldersByName('GENERATED');
  return subs.hasNext() ? subs.next() : parent.createFolder('GENERATED');
}

// Returns (creating if needed) the Year/MonthName/DD subfolder under `rootFolder`
// for the given date — e.g. GENERATED/2026/July/11. Keeps generated reports
// browsable in Kavach Docs in a properly structured, chronological layout.
function _getDateOrganizedFolder(rootFolder, dateObj) {
  const tz = Session.getScriptTimeZone();
  const year  = Utilities.formatDate(dateObj, tz, 'yyyy');
  const month = MONTH_NAMES[dateObj.getMonth()].charAt(0) + MONTH_NAMES[dateObj.getMonth()].slice(1).toLowerCase();
  const day   = Utilities.formatDate(dateObj, tz, 'dd');
  function getOrMake(parentFolder, name) {
    const it = parentFolder.getFoldersByName(name);
    return it.hasNext() ? it.next() : parentFolder.createFolder(name);
  }
  return getOrMake(getOrMake(getOrMake(rootFolder, year), month), day);
}

function _saveGeneratedDoc(content, type, taskId, srNos, attachments) {
  const folder = _getDateOrganizedFolder(_getGeneratedRootFolder(), new Date());
  const title  = (type==='MEMO'?'LOCO_MEMO':'OPR') + '_SR' + srNos.join('-') + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
  const doc    = DocumentApp.create(title);
  const body   = doc.getBody();
  const hStyle = DocumentApp.ParagraphHeading;
  const FONT   = 'Calibri', FSIZE = 9;

  // Twips->points (1pt = 20 twips): top/bottom 720 twips = 36pt, left/right 1080 twips = 54pt
  body.setMarginTop(36).setMarginBottom(36).setMarginLeft(54).setMarginRight(54);

  body.appendParagraph(type==='MEMO'?'LOCO MEMO — BRC KAVACH':'ONE PAGE REPORT (OPR) — BRC KAVACH').setHeading(hStyle.HEADING1);
  body.appendParagraph('Generated: '+new Date().toLocaleString()+'  |  SR(s): '+srNos.join(', ')).setItalic(true).setFontFamily(FONT).setFontSize(FSIZE);
  body.appendHorizontalRule();

  _renderMarkupIntoBody(body, content, hStyle, FONT, FSIZE);

  // Attachments appendix — images are embedded inline, other file types get a linked reference
  if (attachments && attachments.length) {
    body.appendHorizontalRule();
    body.appendParagraph('Attachments').setHeading(hStyle.HEADING2);
    attachments.forEach(function(a) {
      try {
        const af = DriveApp.getFileById(a.fileId);
        const mime = af.getMimeType();
        if (mime.indexOf('image/') === 0) {
          body.appendParagraph(a.fileName).setBold(true);
          const blob = af.getBlob();
          const img = body.appendImage(blob);
          const maxW = 420;
          if (img.getWidth() > maxW) { const ratio = maxW/img.getWidth(); img.setWidth(maxW); img.setHeight(Math.round(img.getHeight()*ratio)); }
        } else {
          const p = body.appendParagraph('📎 ' + a.fileName);
          p.setLinkUrl(af.getUrl());
        }
      } catch(e) { body.appendParagraph('⚠ Could not embed attachment: ' + a.fileName); }
    });
  }

  doc.saveAndClose();

  const f = DriveApp.getFileById(doc.getId());
  folder.addFile(f);
  try { DriveApp.getRootFolder().removeFile(f); } catch(e){}
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { fileId:doc.getId(), fileUrl:f.getUrl() };
}

// Uploads a reference attachment (NMS screenshot, loco log screenshot, etc.)
// used by the Generate Document dialog. Returns a reference the client
// accumulates and submits with the generation task.
function uploadGenAttachment(fileName, base64, mimeType) {
  try {
    const parent = _getBrcFolder();
    const subs   = parent.getFoldersByName('GEN_ATTACHMENTS');
    const folder = subs.hasNext() ? subs.next() : parent.createFolder('GEN_ATTACHMENTS');
    const safeName = String(fileName||'attachment').replace(/[^a-zA-Z0-9._\- ]/g,'_');
    const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType||'application/octet-stream', safeName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success:true, fileId:file.getId(), fileName:safeName, url:file.getUrl(), mimeType:mimeType||'' };
  } catch(e) { return { success:false, error:e.message }; }
}

// ── Admin: Project Memory (system-prompt instructions for MEMO/OPR/PPT) ──
function getProjectMemoryPrompts() {
  const props = PropertiesService.getScriptProperties();
  return {
    success: true,
    memo:  props.getProperty('MEMO_SYSTEM_PROMPT')  || '',
    opr:   props.getProperty('OPR_SYSTEM_PROMPT')   || '',
    ppt:   props.getProperty('PPT_SYSTEM_PROMPT')   || '',
    daily: props.getProperty('DAILY_SYSTEM_PROMPT') || _DAILY_DEFAULT_MEMORY
  };
}

function saveProjectMemoryPrompt(data) {
  try {
    const type = String(data.type||'').toUpperCase();
    const key = type==='MEMO' ? 'MEMO_SYSTEM_PROMPT' : type==='OPR' ? 'OPR_SYSTEM_PROMPT'
              : type==='PPT' ? 'PPT_SYSTEM_PROMPT' : type==='DAILY' ? 'DAILY_SYSTEM_PROMPT' : null;
    if (!key) return { success:false, error:'Invalid type' };
    PropertiesService.getScriptProperties().setProperty(key, String(data.text||''));
    return { success:true };
  } catch(e) { return { success:false, error:e.message }; }
}

// ══════════════════════════════════════════════════════
// REPORTS REGISTRY — OPR draft→signed lifecycle + PPT (gated on signed OPR)
// ══════════════════════════════════════════════════════
const RPT = { ID:0, TASK_ID:1, TYPE:2, SR_REFS:3, DRAFT_FILE_ID:4, DRAFT_FILE_URL:5,
  SIGNED_FILE_ID:6, SIGNED_FILE_URL:7, PPT_FILE_ID:8, PPT_FILE_URL:9, STATUS:10, CREATED_AT:11, LABELS:12 };

function _getReportsSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName('REPORTS');
  if (!sh) {
    sh = ss.insertSheet('REPORTS');
    sh.getRange(1,1,1,13).setValues([[
      'ID','TaskId','Type','SR Refs (JSON)','Draft File ID','Draft File URL',
      'Signed File ID','Signed File URL','PPT File ID','PPT File URL','Status','Created At','Labels (JSON)'
    ]]);
    sh.setFrozenRows(1);
  }
  // Migrate older 12-col sheet — add Labels column
  if (sh.getLastColumn() < 13) sh.getRange(1,13).setValue('Labels (JSON)');
  return sh;
}

// type MEMO → Status 'Final' immediately (no signing step for Memo).
// type OPR  → Status 'Draft' until a signed copy is uploaded.
// labels: category tags (e.g. "Radio Issue", "RFID Reader Issue") settable at
// generation time or later via updateReportLabels — used to filter Kavach Docs.
function _createReportEntry(data) {
  const sh = _getReportsSheet();
  const id = 'RPT_' + Date.now();
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  const status = data.type === 'MEMO' ? 'Final' : 'Draft';
  sh.appendRow([id, data.taskId, data.type, JSON.stringify(data.srRefs||[]),
    data.draftFileId||'', data.draftFileUrl||'', '', '', '', '', status, now, JSON.stringify(data.labels||[])]);
  return id;
}

function getReportsRegistry() {
  try {
    const rows = _getReportsSheet().getDataRange().getValues();
    const out = [];
    rows.slice(1).reverse().forEach(r => {
      try {
        out.push({
          id: _cellToStr(r[RPT.ID]), taskId: _cellToStr(r[RPT.TASK_ID]), type: _cellToStr(r[RPT.TYPE]),
          srRefs: _parseJsonArray(r[RPT.SR_REFS]),
          draftFileId: _cellToStr(r[RPT.DRAFT_FILE_ID]), draftFileUrl: _cellToStr(r[RPT.DRAFT_FILE_URL]),
          signedFileId: _cellToStr(r[RPT.SIGNED_FILE_ID]), signedFileUrl: _cellToStr(r[RPT.SIGNED_FILE_URL]),
          pptFileId: _cellToStr(r[RPT.PPT_FILE_ID]), pptFileUrl: _cellToStr(r[RPT.PPT_FILE_URL]),
          status: _cellToStr(r[RPT.STATUS]), createdAt: _cellToStr(r[RPT.CREATED_AT]),
          labels: _parseJsonArray(r[RPT.LABELS])
        });
      } catch(rowErr) { /* skip corrupt row rather than breaking the whole list */ }
    });
    return out;
  } catch(e) { return []; }
}

// Sets (replaces) the label tags on an existing report — used by the Reports
// Registry UI so labels can be added/edited any time, not just at creation.
function updateReportLabels(reportId, labels) {
  try {
    const sh = _getReportsSheet();
    const rows = sh.getDataRange().getValues();
    for (let i=1;i<rows.length;i++) {
      if (rows[i][RPT.ID] === reportId) {
        sh.getRange(i+1, RPT.LABELS+1).setValue(JSON.stringify(Array.isArray(labels)?labels:[]));
        return { success:true };
      }
    }
    return { success:false, error:'Report not found' };
  } catch(e) { return { success:false, error:e.message }; }
}

// Suggested label chips for tagging generated documents (Kavach Docs filter set).
// Admin-editable via CONFIG key 'doc_report_labels'; seeded with common KAVACH
// failure categories on first use.
const DEFAULT_DOC_LABELS = ['Radio Issue','RFID Reader Issue','Signal/Interlocking Issue',
  'Station Issue','Loco Issue','Tag Miss','Mode Degradation','KAVACH System Failure'];

function getDocLabelOptions() {
  try {
    const sh = getOrCreateSheet('CONFIG');
    const rows = sh.getDataRange().getValues();
    for (const r of rows.slice(1)) {
      if (String(r[0]).trim() === 'doc_report_labels') {
        try { return { success:true, labels: JSON.parse(r[1]) }; } catch(e) {}
      }
    }
    return { success:true, labels: DEFAULT_DOC_LABELS };
  } catch(e) { return { success:true, labels: DEFAULT_DOC_LABELS }; }
}

function saveDocLabelOptions(labels) {
  try {
    const sh = getOrCreateSheet('CONFIG');
    const rows = sh.getDataRange().getValues();
    for (let i=1;i<rows.length;i++) {
      if (String(rows[i][0]).trim() === 'doc_report_labels') {
        sh.getRange(i+1,2).setValue(JSON.stringify(labels));
        return { success:true };
      }
    }
    sh.appendRow(['doc_report_labels', JSON.stringify(labels)]);
    return { success:true };
  } catch(e) { return { success:false, error:e.message }; }
}

// Upload the signed hard copy of an OPR (only OPR-type reports use this).
function uploadSignedOpr(reportId, fileName, base64, mimeType) {
  try {
    const sh = _getReportsSheet();
    const rows = sh.getDataRange().getValues();
    for (let i=1;i<rows.length;i++) {
      if (rows[i][RPT.ID] === reportId) {
        if (rows[i][RPT.TYPE] !== 'OPR') return { success:false, error:'Only OPR reports accept a signed copy' };
        const folder = _getDateOrganizedFolder(_getGeneratedRootFolder(), new Date());
        const safeName = String(fileName||'OPR_signed.pdf').replace(/[^a-zA-Z0-9._\- ]/g,'_');
        const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType||'application/pdf', safeName);
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        sh.getRange(i+1, RPT.SIGNED_FILE_ID+1).setValue(file.getId());
        sh.getRange(i+1, RPT.SIGNED_FILE_URL+1).setValue(file.getUrl());
        sh.getRange(i+1, RPT.STATUS+1).setValue('Signed');
        SpreadsheetApp.flush();
        return { success:true, fileId:file.getId(), fileUrl:file.getUrl() };
      }
    }
    return { success:false, error:'Report not found' };
  } catch(e) { return { success:false, error:e.message }; }
}

// Upload a PPT (user-made) for a report — only allowed once the OPR is Signed.
function uploadPpt(reportId, fileName, base64, mimeType) {
  try {
    const sh = _getReportsSheet();
    const rows = sh.getDataRange().getValues();
    for (let i=1;i<rows.length;i++) {
      if (rows[i][RPT.ID] === reportId) {
        if (rows[i][RPT.STATUS] !== 'Signed' && rows[i][RPT.STATUS] !== 'Complete') {
          return { success:false, error:'PPT can only be attached after the OPR signed copy is uploaded' };
        }
        const folder = _getDateOrganizedFolder(_getGeneratedRootFolder(), new Date());
        const safeName = String(fileName||'OPR_slides.pptx').replace(/[^a-zA-Z0-9._\- ]/g,'_');
        const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType||'application/vnd.openxmlformats-officedocument.presentationml.presentation', safeName);
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        sh.getRange(i+1, RPT.PPT_FILE_ID+1).setValue(file.getId());
        sh.getRange(i+1, RPT.PPT_FILE_URL+1).setValue(file.getUrl());
        sh.getRange(i+1, RPT.STATUS+1).setValue('Complete');
        SpreadsheetApp.flush();
        return { success:true, fileId:file.getId(), fileUrl:file.getUrl() };
      }
    }
    return { success:false, error:'Report not found' };
  } catch(e) { return { success:false, error:e.message }; }
}

// AI-generates a PPT (Google Slides deck) from the signed OPR's content —
// only allowed once the OPR is Signed. Uses the PPT project-memory prompt.
function generatePptFromReport(reportId, notes) {
  try {
    const sh = _getReportsSheet();
    const rows = sh.getDataRange().getValues();
    let row = -1, refs = [];
    for (let i=1;i<rows.length;i++) {
      if (rows[i][RPT.ID] === reportId) {
        if (rows[i][RPT.STATUS] !== 'Signed' && rows[i][RPT.STATUS] !== 'Complete') {
          return { success:false, error:'PPT can only be generated after the OPR signed copy is uploaded' };
        }
        row = i+1; refs = _parseJson(rows[i][RPT.SR_REFS], []);
        break;
      }
    }
    if (row < 0) return { success:false, error:'Report not found' };

    const props = PropertiesService.getScriptProperties();
    const apiKey = props.getProperty('PPT_API_KEY') || props.getProperty('OPR_API_KEY') || props.getProperty('MEMO_API_KEY');
    if (!apiKey) throw new Error('No API key configured (PPT_API_KEY, OPR_API_KEY, or MEMO_API_KEY)');

    const obsData = _getObsDataForRefs(refs);
    const custom = _getProjectMemoryPrompt('PPT');
    const system = (custom ? custom + '\n\n' : '') +
      'You are a senior KAVACH railway signalling engineer preparing presentation slides from an approved OPR. ' +
      'Produce 5-8 slides. For EACH slide, output exactly in this format (repeat per slide):\n' +
      'SLIDE: <title>\n- <bullet 1>\n- <bullet 2>\n- <bullet 3 (optional)>\n\n' +
      'Keep bullets short and presentation-ready (not full sentences). Cover: title/overview, incident details, ' +
      'technical findings, corrective actions, recommendations.';
    const user = `Generate presentation slide content based on this approved OPR data:\n\n${_buildObsText(obsData)}\n\nAdditional Notes: ${notes||'None'}`;
    const raw = _callClaudeAPI(apiKey, system, user);

    // Parse "SLIDE: title\n- bullet..." blocks
    const slides = [];
    raw.split(/\n(?=SLIDE:)/).forEach(function(block) {
      const m = block.match(/^SLIDE:\s*(.+)/);
      if (!m) return;
      const title = m[1].trim();
      const bullets = block.split('\n').slice(1).map(function(l){ return l.replace(/^-\s*/,'').trim(); }).filter(Boolean);
      slides.push({ title:title, bullets:bullets });
    });
    if (!slides.length) throw new Error('AI did not return any slides — try again or check the PPT project memory prompt');

    const presentation = SlidesApp.create('OPR_SLIDES_' + reportId + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm'));
    const pres = presentation;
    const defaultSlide = pres.getSlides()[0]; // every new presentation starts with exactly one blank slide
    slides.forEach(function(sl) {
      const slide = pres.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
      slide.getShapes().forEach(function(shape) {
        if (shape.getPlaceholderType() === SlidesApp.PlaceholderType.TITLE || shape.getPlaceholderType() === SlidesApp.PlaceholderType.CENTERED_TITLE) {
          shape.getText().setText(sl.title);
        } else if (shape.getPlaceholderType() === SlidesApp.PlaceholderType.BODY) {
          shape.getText().setText(sl.bullets.join('\n'));
        }
      });
    });
    defaultSlide.remove(); // drop the original blank slide now that real ones exist
    presentation.saveAndClose();

    const f = DriveApp.getFileById(presentation.getId());
    const folder = _getDateOrganizedFolder(_getGeneratedRootFolder(), new Date());
    folder.addFile(f);
    try { DriveApp.getRootFolder().removeFile(f); } catch(e) {}
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    sh.getRange(row, RPT.PPT_FILE_ID+1).setValue(f.getId());
    sh.getRange(row, RPT.PPT_FILE_URL+1).setValue(f.getUrl());
    sh.getRange(row, RPT.STATUS+1).setValue('Complete');
    SpreadsheetApp.flush();
    return { success:true, fileId:f.getId(), fileUrl:f.getUrl() };
  } catch(e) { return { success:false, error:e.message }; }
}

// ── Kavach Docs: dynamic root folders (adds Generated Reports alongside the fixed section folders) ──
// Merged roots: the two fixed section folders + the dynamically-created
// GENERATED folder (MEMO/OPR/PPT output). Used by the tree, breadcrumb, and
// search so all three treat "Generated Reports" as a first-class root too.
function _getAllDocRoots() {
  try {
    const roots = Object.assign({}, DOC_FOLDERS);
    roots['Generated Reports'] = _getGeneratedRootFolder().getId();
    return roots;
  } catch(e) { return DOC_FOLDERS; }
}

function getDocRootFolders() {
  try {
    return { success:true, roots: _getAllDocRoots() };
  } catch(e) { return { success:false, error:e.message, roots: DOC_FOLDERS }; }
}

function cancelQueueTask(taskId) {
  try {
    const sh=_getQueueSheet(); const rows=sh.getDataRange().getValues();
    for (let i=1;i<rows.length;i++) {
      if (rows[i][GQ.ID]===taskId && rows[i][GQ.STATUS]==='Queued') {
        sh.getRange(i+1,GQ.STATUS+1).setValue('Cancelled'); return {success:true};
      }
    }
    return {success:false,error:'Not found or not Queued'};
  } catch(e){return{success:false,error:e.message};}
}

function retryQueueTask(taskId) {
  try {
    const sh=_getQueueSheet(); const rows=sh.getDataRange().getValues();
    for (let i=1;i<rows.length;i++) {
      if (rows[i][GQ.ID]===taskId && rows[i][GQ.STATUS]==='Error') {
        sh.getRange(i+1,GQ.STATUS+1).setValue('Queued');
        sh.getRange(i+1,GQ.ERROR+1).setValue('');
        sh.getRange(i+1,GQ.STARTED+1).setValue('');
        _tryProcessQueue();
        return {success:true};
      }
    }
    return {success:false,error:'Not found or not in Error state'};
  } catch(e){return{success:false,error:e.message};}
}

// ══════════════════════════════════════════════════════════════
// CUSTOM ANALYTICS / DEEP ANALYTICS — cross-filter comparison engine
// (read-only — scans month sheets, no writes)
// params: { dateFrom, dateTo,
//   filters: { sections[], labels[], gears[], reasons[], categories[],
//              oemCompanies[], locoNumbers[], stations[], statuses[],
//              flags[], upDn[], excludeNil:bool },
//   groupBy:   gear|reason|section|label|category|oem|loco|station|status|flag|upDn|locomake|locotype|locoshed|month,
//   compareBy: (same set, optional),
//   locoMatrix: true/false
// }
// ══════════════════════════════════════════════════════════════
function getCustomAnalytics(params) {
  try {
    var dateFrom  = (params && params.dateFrom)  || '';
    var dateTo    = (params && params.dateTo)    || '';
    var filters   = (params && params.filters)   || {};
    var groupBy   = String((params && params.groupBy) || 'gear').trim().toLowerCase();
    var compareBy = String((params && params.compareBy) || '').trim().toLowerCase();
    var locoMatrix = params && params.locoMatrix;
    var tz = Session.getScriptTimeZone();

    // ── Collect matching rows ──────────────────────────
    var rows = [];
    _getAllObsMonthSheets().forEach(function(sh) {
      var data = sh.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var r = data[i];
        if (!r[COL.SR] && !r[COL.DATE]) continue;

        // Date filter
        var dv = r[COL.DATE];
        var dStr = '';
        var monthKey = '';
        if (dv instanceof Date) {
          dStr = Utilities.formatDate(dv, tz, 'yyyy-MM-dd');
          monthKey = MONTH_NAMES[dv.getMonth()].charAt(0) + MONTH_NAMES[dv.getMonth()].slice(1).toLowerCase() + ' ' + dv.getFullYear();
        } else {
          dStr = String(dv || '');
        }
        if (dateFrom && dStr < dateFrom) continue;
        if (dateTo && dStr > dateTo) continue;

        // Parse row fields
        var gearRaw    = String(r[COL.GEAR] || '').trim();
        var mdRaw      = String(r[COL.MODE_DEG] || '').trim();
        var ebRaw      = String(r[COL.EB] || '').trim();
        var labelRaw   = String(r[COL.LABEL] || '').trim();
        var reasonRaw  = String(r[COL.REASON] || '').trim();
        var section    = String(r[COL.SECTION] || '').trim();
        var station    = String(r[COL.STATION] || '').trim();
        var oem        = String(r[COL.OEM_COMPANY] || '').trim();
        var loco       = String(r[COL.LOCO] || '').trim();
        var locoType   = String(r[COL.LOCO_TYPE] || '').trim();
        var locoMake   = String(r[COL.LOCO_MAKE] || '').trim();
        var locoShed   = String(r[COL.LOCO_SHED] || '').trim();
        var trainNo    = String(r[COL.TRAIN] || '').trim();
        var status     = String(r[COL.STATUS] || 'Pending').trim();
        var flag       = String(r[COL.FLAG] || 'None').trim();
        var upDn       = String(r[COL.UPDN] || '').trim();
        var fitUnfit   = String(r[COL.FIT] || '').trim();

        // Category
        var category = 'Other';
        if (gearRaw === 'NIL') category = 'NIL';
        else if (mdRaw && mdRaw !== 'NIL') category = 'Mode Change';
        else if (ebRaw && ebRaw !== 'NIL') category = 'Undue Braking';

        // Labels array
        var rowLabels = labelRaw ? labelRaw.split(',').map(function(l){return l.trim();}).filter(Boolean) : [];

        // ── Apply filters ───────────────────────────────
        if (filters.sections && filters.sections.length) {
          if (filters.sections.indexOf(section) < 0) continue;
        }
        if (filters.labels && filters.labels.length) {
          if (!filters.labels.some(function(fl){ return rowLabels.indexOf(fl) >= 0; })) continue;
        }
        if (filters.gears && filters.gears.length) {
          var rg = _rptSplitVals(gearRaw);
          if (!filters.gears.some(function(fg){ return rg.indexOf(fg) >= 0; })) continue;
        }
        if (filters.reasons && filters.reasons.length) {
          var rr = _rptSplitVals(reasonRaw);
          if (!filters.reasons.some(function(fr){ return rr.indexOf(fr) >= 0; })) continue;
        }
        if (filters.categories && filters.categories.length) {
          if (filters.categories.indexOf(category) < 0) continue;
        }
        if (filters.oemCompanies && filters.oemCompanies.length) {
          if (filters.oemCompanies.indexOf(oem) < 0) continue;
        }
        if (filters.locoNumbers && filters.locoNumbers.length) {
          if (filters.locoNumbers.indexOf(loco) < 0) continue;
        }
        if (filters.stations && filters.stations.length) {
          if (filters.stations.indexOf(station) < 0) continue;
        }
        if (filters.statuses && filters.statuses.length) {
          if (filters.statuses.indexOf(status) < 0) continue;
        }
        if (filters.flags && filters.flags.length) {
          if (filters.flags.indexOf(flag) < 0) continue;
        }
        if (filters.upDn && filters.upDn.length) {
          if (filters.upDn.indexOf(upDn) < 0) continue;
        }
        if (filters.excludeNil && gearRaw === 'NIL') continue;
        if (filters.fitUnfit && filters.fitUnfit.length) {
          if (filters.fitUnfit.indexOf(fitUnfit) < 0) continue;
        }

        rows.push({
          srNo: String(r[COL.SR]||''), date: dStr, monthKey: monthKey,
          section: section, station: station, gear: gearRaw, reason: reasonRaw,
          category: category, modeDeg: mdRaw, eb: ebRaw,
          oemCompany: oem, locoNo: loco, locoType: locoType,
          locoMake: locoMake, locoShed: locoShed, trainNo: trainNo,
          status: status, flag: flag, label: labelRaw, upDn: upDn,
          fitUnfit: fitUnfit,
          desc: String(r[COL.DESC]||'').trim()
        });
      }
    });

    // ── Filter metadata (distinct values for UI dropdowns) ──
    var filterMeta = { sections:{}, labels:{}, gears:{}, reasons:{}, categories:{},
                       oemCompanies:{}, locoNumbers:{}, stations:{}, statuses:{}, flags:{}, upDn:{},
                       locoMakes:{}, locoTypes:{}, locoSheds:{} };
    rows.forEach(function(row) {
      if (row.section)    filterMeta.sections[row.section] = true;
      if (row.gear && row.gear !== 'NIL') _rptSplitVals(row.gear).forEach(function(g){ filterMeta.gears[g] = true; });
      if (row.reason)     _rptSplitVals(row.reason).forEach(function(r){ filterMeta.reasons[r] = true; });
      if (row.category && row.category !== 'NIL') filterMeta.categories[row.category] = true;
      if (row.oemCompany) filterMeta.oemCompanies[row.oemCompany] = true;
      if (row.locoNo)     filterMeta.locoNumbers[row.locoNo] = true;
      if (row.station)    _rptSplitVals(row.station).forEach(function(s){ filterMeta.stations[s] = true; });
      if (row.status)     filterMeta.statuses[row.status] = true;
      if (row.flag && row.flag !== 'None') filterMeta.flags[row.flag] = true;
      if (row.upDn)       filterMeta.upDn[row.upDn] = true;
      if (row.locoMake)   filterMeta.locoMakes[row.locoMake] = true;
      if (row.locoType)   filterMeta.locoTypes[row.locoType] = true;
      if (row.locoShed)   filterMeta.locoSheds[row.locoShed] = true;
      if (row.label) {
        row.label.split(',').map(function(l){return l.trim();}).filter(Boolean).forEach(function(l){ filterMeta.labels[l] = true; });
      }
    });
    // Convert to sorted arrays
    Object.keys(filterMeta).forEach(function(k){ filterMeta[k] = Object.keys(filterMeta[k]).sort(); });

    // ── Summary stats ─────────────────────────────────
    var nilCount = 0, mcCount = 0, ebCount = 0, otherCount = 0;
    var verifiedCount = 0, pendingCount = 0, rectCount = 0;
    var uniqueLocos = {};
    rows.forEach(function(row) {
      if (row.category === 'NIL') nilCount++;
      else if (row.category === 'Mode Change') mcCount++;
      else if (row.category === 'Undue Braking') ebCount++;
      else otherCount++;
      if (row.status === 'Verified') verifiedCount++;
      else if (row.status === 'Rectification') rectCount++;
      else pendingCount++;
      if (row.locoNo) uniqueLocos[row.locoNo] = true;
    });

    var summary = {
      total: rows.length, nilCount: nilCount, mcCount: mcCount, ebCount: ebCount,
      otherCount: otherCount, verifiedCount: verifiedCount, pendingCount: pendingCount,
      rectCount: rectCount, uniqueLocos: Object.keys(uniqueLocos).length
    };

    // ── Helper: extract group key from a row ────────────
    function groupKey(row, dim) {
      switch(dim) {
        case 'gear':      return row.gear === 'NIL' ? 'NIL' : _rptSplitVals(row.gear).join(', ');
        case 'reason':    return _rptSplitVals(row.reason).join(', ');
        case 'section':   return row.section || 'Unknown';
        case 'label':     return row.label || 'Unlabelled';
        case 'category':  return row.category;
        case 'oem':       return row.oemCompany || 'Unknown';
        case 'loco':      return row.locoNo || 'Unknown';
        case 'station':   return _rptSplitVals(row.station).join(', ');
        case 'status':    return row.status;
        case 'flag':      return row.flag;
        case 'updn':      return row.upDn || 'N/A';
        case 'locomake':  return row.locoMake || 'Unknown';
        case 'locotype':  return row.locoType || 'Unknown';
        case 'locoshed':  return row.locoShed || 'Unknown';
        case 'month':     return row.monthKey || 'Unknown';
        default:          return row.gear || 'Unknown';
      }
    }

    // ── Loco Matrix mode ──────────────────────────────
    if (locoMatrix) {
      // Build loco × failure-type matrix
      var locoMap = {};
      var failureTypesSet = {};
      rows.forEach(function(row) {
        if (row.gear === 'NIL') return; // skip NIL for loco matrix
        var lk = row.locoNo || 'Unknown';
        if (!locoMap[lk]) {
          locoMap[lk] = { locoNo: lk, locoType: row.locoType, locoMake: row.locoMake,
                          locoShed: row.locoShed, failures: {}, total: 0,
                          mcCount: 0, ebCount: 0, obsDates: [] };
        }
        var entry = locoMap[lk];
        entry.total++;
        entry.obsDates.push(row.date);
        if (row.category === 'Mode Change') entry.mcCount++;
        if (row.category === 'Undue Braking') entry.ebCount++;

        // Individual gear types
        _rptSplitVals(row.gear).forEach(function(g) {
          entry.failures[g] = (entry.failures[g] || 0) + 1;
          failureTypesSet[g] = true;
        });
        // Individual reasons
        _rptSplitVals(row.reason).forEach(function(r) {
          entry.failures[r] = (entry.failures[r] || 0) + 1;
          failureTypesSet[r] = true;
        });
        // MC / EB subtypes
        _rptSplitVals(row.modeDeg).forEach(function(m) {
          if (m && m !== 'NIL') { entry.failures['MC: ' + m] = (entry.failures['MC: ' + m] || 0) + 1; failureTypesSet['MC: ' + m] = true; }
        });
        _rptSplitVals(row.eb).forEach(function(e) {
          if (e && e !== 'NIL') { entry.failures['EB: ' + e] = (entry.failures['EB: ' + e] || 0) + 1; failureTypesSet['EB: ' + e] = true; }
        });
      });

      var columns = Object.keys(failureTypesSet).sort();
      var matrixRows = Object.values(locoMap).sort(function(a,b){ return b.total - a.total; });

      return {
        success: true, summary: summary, filterMeta: filterMeta,
        locoMatrix: { columns: columns, rows: matrixRows },
        groups: [], timeline: []
      };
    }

    // ── Standard group + compare mode ──────────────────
    // Each row can have multi-value cells (gear, reason, station).
    // For groupBy gear/reason/station, we explode into individual tokens.
    var explodeDims = {'gear':true,'reason':true,'station':true};

    var groupMap = {};  // groupKey → { label, total, compareValues: { compareKey: count } }

    rows.forEach(function(row) {
      var gKeys;
      if (explodeDims[groupBy]) {
        gKeys = _rptSplitVals(groupKey(row, groupBy));
        if (!gKeys.length) gKeys = ['—'];
      } else {
        var gk = groupKey(row, groupBy);
        gKeys = gk ? gk.split(',').map(function(s){return s.trim();}).filter(Boolean) : ['—'];
      }

      gKeys.forEach(function(gk) {
        if (!gk) gk = '—';
        if (!groupMap[gk]) groupMap[gk] = { label: gk, total: 0, compareValues: {} };
        groupMap[gk].total++;

        // Compare dimension
        if (compareBy) {
          var cKeys;
          if (explodeDims[compareBy]) {
            cKeys = _rptSplitVals(groupKey(row, compareBy));
            if (!cKeys.length) cKeys = ['—'];
          } else {
            var ck = groupKey(row, compareBy);
            cKeys = ck ? ck.split(',').map(function(s){return s.trim();}).filter(Boolean) : ['—'];
          }
          cKeys.forEach(function(ck) {
            if (!ck) ck = '—';
            groupMap[gk].compareValues[ck] = (groupMap[gk].compareValues[ck] || 0) + 1;
          });
        }
      });
    });

    // Sort groups by total desc
    var groups = Object.values(groupMap).sort(function(a,b){ return b.total - a.total; });

    // Collect all compare keys (for chart axis labels)
    var compareKeys = [];
    if (compareBy) {
      var ckSet = {};
      groups.forEach(function(g) {
        Object.keys(g.compareValues).forEach(function(k){ ckSet[k] = true; });
      });
      compareKeys = Object.keys(ckSet).sort();
      // For month, sort chronologically
      if (compareBy === 'month') {
        compareKeys.sort(function(a,b) {
          var parseM = function(s) {
            var parts = s.split(' ');
            if (parts.length < 2) return 0;
            var mi = MONTH_NAMES.findIndex(function(m){ return m.charAt(0) + m.slice(1).toLowerCase() === parts[0]; });
            return new Date(Number(parts[1]), mi >= 0 ? mi : 0, 1).getTime();
          };
          return parseM(a) - parseM(b);
        });
      }
    }

    // Build timeline (month-over-month trend for each top group)
    var timeline = [];
    if (compareBy !== 'month' || true) { // always build timeline
      var monthSet = {};
      rows.forEach(function(row) {
        if (row.monthKey) monthSet[row.monthKey] = true;
      });
      var months = Object.keys(monthSet).sort(function(a,b) {
        var parseM = function(s) {
          var parts = s.split(' ');
          if (parts.length < 2) return 0;
          var mi = MONTH_NAMES.findIndex(function(m){ return m.charAt(0) + m.slice(1).toLowerCase() === parts[0]; });
          return new Date(Number(parts[1]), mi >= 0 ? mi : 0, 1).getTime();
        };
        return parseM(a) - parseM(b);
      });

      // Top N groups for timeline (max 8)
      var topGroups = groups.slice(0, 8);
      var timelineMap = {}; // month → { groupKey: count }
      months.forEach(function(m) { timelineMap[m] = {}; });

      rows.forEach(function(row) {
        if (!row.monthKey) return;
        var gKeys;
        if (explodeDims[groupBy]) {
          gKeys = _rptSplitVals(groupKey(row, groupBy));
          if (!gKeys.length) gKeys = ['—'];
        } else {
          var gk = groupKey(row, groupBy);
          gKeys = gk ? gk.split(',').map(function(s){return s.trim();}).filter(Boolean) : ['—'];
        }
        gKeys.forEach(function(gk) {
          if (!gk) gk = '—';
          // Only include if it's a top group
          if (topGroups.some(function(tg){ return tg.label === gk; })) {
            if (!timelineMap[row.monthKey]) timelineMap[row.monthKey] = {};
            timelineMap[row.monthKey][gk] = (timelineMap[row.monthKey][gk] || 0) + 1;
          }
        });
      });

      timeline = {
        months: months,
        series: topGroups.map(function(g) {
          return {
            label: g.label,
            values: months.map(function(m) { return timelineMap[m][g.label] || 0; })
          };
        })
      };
    }

    return {
      success: true,
      summary: summary,
      filterMeta: filterMeta,
      groups: groups,
      compareKeys: compareKeys,
      timeline: timeline,
      groupBy: groupBy,
      compareBy: compareBy
    };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ── Helper: get distinct values for custom analytics dropdowns (lightweight) ──
function getAnalyticsFilterMeta() {
  try {
    var meta = { sections:{}, labels:{}, gears:{}, reasons:{}, categories:{},
                 oemCompanies:{}, stations:{}, upDn:{}, locoMakes:{}, locoTypes:{}, locoSheds:{} };
    _getAllObsMonthSheets().forEach(function(sh) {
      var data = sh.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var r = data[i];
        if (!r[COL.SR] && !r[COL.DATE]) continue;
        var gear = String(r[COL.GEAR]||'').trim();
        var reason = String(r[COL.REASON]||'').trim();
        var label = String(r[COL.LABEL]||'').trim();
        var station = String(r[COL.STATION]||'').trim();
        if (r[COL.SECTION]) meta.sections[String(r[COL.SECTION]).trim()] = true;
        if (gear !== 'NIL') _rptSplitVals(gear).forEach(function(g){ meta.gears[g] = true; });
        if (reason) _rptSplitVals(reason).forEach(function(r){ meta.reasons[r] = true; });
        if (r[COL.OEM_COMPANY]) meta.oemCompanies[String(r[COL.OEM_COMPANY]).trim()] = true;
        if (station) _rptSplitVals(station).forEach(function(s){ meta.stations[s] = true; });
        if (r[COL.UPDN]) meta.upDn[String(r[COL.UPDN]).trim()] = true;
        if (r[COL.LOCO_MAKE]) meta.locoMakes[String(r[COL.LOCO_MAKE]).trim()] = true;
        if (r[COL.LOCO_TYPE]) meta.locoTypes[String(r[COL.LOCO_TYPE]).trim()] = true;
        if (r[COL.LOCO_SHED]) meta.locoSheds[String(r[COL.LOCO_SHED]).trim()] = true;
        if (label) label.split(',').map(function(l){return l.trim();}).filter(Boolean).forEach(function(l){ meta.labels[l] = true; });
      }
    });
    Object.keys(meta).forEach(function(k){ meta[k] = Object.keys(meta[k]).sort(); });
    return { success: true, meta: meta };
  } catch(e) { return { success: false, meta: {} }; }
}

// ══════════════════════════════════════════════════════
// KAVACH DOCUMENT EXPLORER — Drive folder browser
// ══════════════════════════════════════════════════════
var DOC_FOLDERS = {
  'BRC-URN': '1PMBs-erKC0rt5oaUQua0eOkpQPi3vuSc',
  'BRC-GDA': '1gexu1E9q11YexP27ziY-iyznMCH0MPQY'
};

function getDriveFolderContents(folderId) {
  try {
    var folder = DriveApp.getFolderById(folderId);
    var items = [];
    var fIter = folder.getFolders();
    while (fIter.hasNext()) {
      var f = fIter.next();
      var childCount = 0;
      try { childCount = f.getFiles().getContinuationIterator().hasNext() ? 1 : 0; } catch(e) {}
      try { if (!childCount) { var cf = f.getFolders(); if (cf.hasNext()) childCount = 1; } } catch(e) {}
      items.push({id:f.getId(), name:f.getName(), type:'folder', mimeType:'', size:0,
        lastUpdated:Utilities.formatDate(f.getLastUpdated(),Session.getScriptTimeZone(),'dd/MM/yyyy HH:mm'),
        url:f.getUrl(), childCount:childCount});
    }
    var fiIter = folder.getFiles();
    while (fiIter.hasNext()) {
      var fi = fiIter.next();
      var ext = fi.getName().split('.').pop().toLowerCase();
      var icon = _docIcon(ext, fi.getMimeType());
      items.push({id:fi.getId(), name:fi.getName(), type:'file', mimeType:fi.getMimeType(),
        size:fi.getSize(), ext:ext, icon:icon,
        lastUpdated:Utilities.formatDate(fi.getLastUpdated(),Session.getScriptTimeZone(),'dd/MM/yyyy HH:mm'),
        url:fi.getUrl(), thumbnailUrl:'https://drive.google.com/thumbnail?id='+fi.getId()+'&sz=w200',
        viewUrl:'https://drive.google.com/file/d/'+fi.getId()+'/preview'});
    }
    items.sort(function(a,b){
      if (a.type!==b.type) return a.type==='folder'?-1:1;
      return a.name.localeCompare(b.name);
    });
    var breadcrumb = _driveBreadcrumb(folderId);
    return {success:true, items:items, folderName:folder.getName(), breadcrumb:breadcrumb};
  } catch(e) { return {success:false, error:e.message}; }
}

function _driveBreadcrumb(folderId) {
  var crumbs = [];
  try {
    var folder = DriveApp.getFolderById(folderId);
    var allRoots = _getAllDocRoots();
    var rootKey = '';
    for (var k in allRoots) { if (allRoots[k] === folderId) { rootKey = k; break; } }
    if (rootKey) {
      crumbs.push({id:folderId, name:rootKey});
    } else {
      var current = folder;
      var path = [{id:current.getId(), name:current.getName()}];
      var safety = 0;
      while (safety < 15) {
        var parents = current.getParents();
        if (!parents.hasNext()) break;
        var p = parents.next();
        var pId = p.getId();
        var isRoot = false;
        for (var k2 in allRoots) { if (allRoots[k2] === pId) { path.unshift({id:pId, name:k2}); isRoot = true; break; } }
        if (isRoot) break;
        path.unshift({id:pId, name:p.getName()});
        current = p;
        safety++;
      }
      crumbs = path;
    }
  } catch(e) { crumbs = [{id:folderId, name:'Documents'}]; }
  return crumbs;
}

// Fetches a Drive file's raw bytes (base64-encoded) so it can be rendered
// entirely inside the app (e.g. via PDF.js) instead of embedding Drive's own
// preview — which can't be zoomed/panned properly for scanned drawings and
// refuses to load at all outside an iframe for its full-featured view.
// Native Google Docs/Sheets/Slides are NOT handled here — they have no raw
// blob and DriveApp's export conversions are too unreliable (e.g. Sheet ->
// xlsx isn't supported at all); those still redirect to their Drive link.
function getFileBlobBase64(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    return {
      success: true,
      base64: Utilities.base64Encode(blob.getBytes()),
      mimeType: blob.getContentType(),
      name: file.getName()
    };
  } catch(e) { return { success: false, error: e.message }; }
}

function _docIcon(ext, mime) {
  if (ext==='pdf') return 'pdf';
  if (['xls','xlsx','csv'].indexOf(ext)>=0) return 'excel';
  if (['ppt','pptx'].indexOf(ext)>=0) return 'ppt';
  if (['doc','docx','odt','rtf'].indexOf(ext)>=0) return 'word';
  if (['jpg','jpeg','png','gif','bmp','webp','svg'].indexOf(ext)>=0) return 'image';
  if (['mp4','avi','mov','mkv','webm'].indexOf(ext)>=0) return 'video';
  if (['zip','rar','7z','tar','gz'].indexOf(ext)>=0) return 'archive';
  if (mime && mime.indexOf('pdf')>=0) return 'pdf';
  if (mime && (mime.indexOf('spreadsheet')>=0 || mime.indexOf('excel')>=0)) return 'excel';
  if (mime && (mime.indexOf('presentation')>=0 || mime.indexOf('powerpoint')>=0)) return 'ppt';
  if (mime && (mime.indexOf('word')>=0 || mime.indexOf('document')>=0)) return 'word';
  if (mime && mime.indexOf('image')>=0) return 'image';
  return 'other';
}

// Recursively searches folder/file names under a section root (or all roots)
// for the query substring. Depth-limited to avoid runaway traversal on huge trees.
function searchDriveFiles(query, rootKey) {
  try {
    query = String(query||'').trim().toLowerCase();
    if (!query) return {success:true, results:[]};
    var allRoots = _getAllDocRoots();
    var roots = rootKey && allRoots[rootKey] ? [{key:rootKey, id:allRoots[rootKey]}]
      : Object.keys(allRoots).map(function(k){ return {key:k, id:allRoots[k]}; });
    var results = [];
    var MAX_RESULTS = 60, MAX_DEPTH = 6;

    function walk(folder, pathNames, depth) {
      if (results.length >= MAX_RESULTS || depth > MAX_DEPTH) return;
      var fIter = folder.getFolders();
      while (fIter.hasNext() && results.length < MAX_RESULTS) {
        var f = fIter.next();
        var namePath = pathNames.concat([f.getName()]);
        if (f.getName().toLowerCase().indexOf(query) >= 0) {
          results.push({id:f.getId(), name:f.getName(), type:'folder', path:pathNames.join(' / ')||'—',
            lastUpdated:Utilities.formatDate(f.getLastUpdated(),Session.getScriptTimeZone(),'dd/MM/yyyy HH:mm')});
        }
        walk(f, namePath, depth+1);
      }
      var fiIter = folder.getFiles();
      while (fiIter.hasNext() && results.length < MAX_RESULTS) {
        var fi = fiIter.next();
        if (fi.getName().toLowerCase().indexOf(query) >= 0) {
          var ext = fi.getName().split('.').pop().toLowerCase();
          results.push({id:fi.getId(), name:fi.getName(), type:'file', ext:ext, icon:_docIcon(ext, fi.getMimeType()),
            size:fi.getSize(), path:pathNames.join(' / ')||'—', mimeType:fi.getMimeType(),
            lastUpdated:Utilities.formatDate(fi.getLastUpdated(),Session.getScriptTimeZone(),'dd/MM/yyyy HH:mm'),
            viewUrl:'https://drive.google.com/file/d/'+fi.getId()+'/preview'});
        }
      }
    }

    roots.forEach(function(r) {
      if (results.length >= MAX_RESULTS) return;
      try { walk(DriveApp.getFolderById(r.id), [r.key], 0); } catch(e) {}
    });
    return {success:true, results:results, truncated: results.length >= MAX_RESULTS};
  } catch(e) { return {success:false, error:e.message, results:[]}; }
}

// ── Admin-only file operations ─────────────────────────
function _docVerifyAdmin(pin) {
  if (!pin || pin.length < 1) return false;
  return validateAdmin(pin);
}

function driveCreateFolder(parentId, name, pin) {
  if (!_docVerifyAdmin(pin)) return {success:false, error:'Invalid admin PIN'};
  try {
    var parent = DriveApp.getFolderById(parentId);
    var f = parent.createFolder(name);
    return {success:true, id:f.getId(), name:f.getName()};
  } catch(e) { return {success:false, error:e.message}; }
}

function driveCopyFile(fileId, destFolderId, pin) {
  if (!_docVerifyAdmin(pin)) return {success:false, error:'Invalid admin PIN'};
  try {
    var file = DriveApp.getFileById(fileId);
    var dest = DriveApp.getFolderById(destFolderId);
    var copy = file.makeCopy(dest);
    return {success:true, id:copy.getId(), name:copy.getName()};
  } catch(e) { return {success:false, error:e.message}; }
}

function driveMoveFile(fileId, destFolderId, pin) {
  if (!_docVerifyAdmin(pin)) return {success:false, error:'Invalid admin PIN'};
  try {
    var file = DriveApp.getFileById(fileId);
    var dest = DriveApp.getFolderById(destFolderId);
    dest.addFile(file);
    var parents = file.getParents();
    while (parents.hasNext()) {
      var p = parents.next();
      if (p.getId() !== destFolderId) { try { p.removeFile(file); } catch(e) {} }
    }
    return {success:true, name:file.getName()};
  } catch(e) { return {success:false, error:e.message}; }
}

function driveDeleteFile(fileId, pin) {
  if (!_docVerifyAdmin(pin)) return {success:false, error:'Invalid admin PIN'};
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
    return {success:true};
  } catch(e) { return {success:false, error:e.message}; }
}

function driveRenameFile(fileId, newName, pin) {
  if (!_docVerifyAdmin(pin)) return {success:false, error:'Invalid admin PIN'};
  try {
    DriveApp.getFileById(fileId).setName(newName);
    return {success:true};
  } catch(e) { return {success:false, error:e.message}; }
}

function driveGetFolderTree(rootId, pin) {
  if (!_docVerifyAdmin(pin)) return {success:false, error:'Invalid admin PIN'};
  try {
    var tree = [];
    function walk(folder, depth) {
      if (depth > 6) return;
      var node = {id:folder.getId(), name:folder.getName(), children:[]};
      var fIter = folder.getFolders();
      while (fIter.hasNext()) { node.children.push(walk(fIter.next(), depth+1)); }
      return node;
    }
    var root = DriveApp.getFolderById(rootId);
    tree.push(walk(root, 0));
    return {success:true, tree:tree};
  } catch(e) { return {success:false, error:e.message}; }
}

// ══════════════════════════════════════════════════════
// SITE MAINTENANCE — external Apps Script dashboard, embedded
// ══════════════════════════════════════════════════════
// Loaded client-side via a full-height iframe (reliable — the browser
// carries the viewer's own Google session, which a server-side fetch cannot).
var _SM_URL = 'https://script.google.com/macros/s/AKfycbxgK7BEOTB8e0_naNsJDYW6v9VEtpMSAa52hikkemp9V6y-yZ_HIyTP6eiQlu5ab1AP/exec';

function getSiteMaintenanceUrl() {
  return { success: true, url: _SM_URL };
}

// ══════════════════════════════════════════════════════
// HBL REMARK — external log of observations shared with HBL and their
// replies/actions, one tab per corridor. Read directly (not embedded) so it
// can be shown newest-first inside the app.
// ══════════════════════════════════════════════════════
const HBL_SHEET_ID_DEFAULT = '1Hh-8DlQ7hRVBvHLKG7vATbgJjkdDKCNGBWXjLqe_JHw';

function _getHblSheetId() {
  const sh   = getOrCreateSheet('CONFIG');
  const rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === 'hbl_sheet_id') {
      const v = String(rows[i][1] || '').trim();
      if (v) return v;
      break;
    }
  }
  return HBL_SHEET_ID_DEFAULT;
}

function getHblSheetIdConfig() {
  return { success: true, sheetId: _getHblSheetId() };
}

function saveHblSheetId(input) {
  try {
    input = String(input || '').trim();
    if (!input) return { success: false, error: 'Enter a Sheet ID or URL' };
    const m = input.match(/\/d\/([a-zA-Z0-9_-]+)/);
    const sheetId = m ? m[1] : input;
    let ss;
    try { ss = SpreadsheetApp.openById(sheetId); }
    catch(e) { return { success: false, error: 'Cannot open that sheet — verify the ID/URL and that it\'s shared with this script\'s account. (' + e.message + ')' }; }
    const sh   = getOrCreateSheet('CONFIG');
    const rows = sh.getDataRange().getValues();
    let found = false;
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === 'hbl_sheet_id') { sh.getRange(i + 1, 2).setValue(sheetId); found = true; break; }
    }
    if (!found) sh.appendRow(['hbl_sheet_id', sheetId]);
    return { success: true, sheetName: ss.getName() };
  } catch(e) { return { success: false, error: e.message }; }
}

function _getHblSS() {
  return SpreadsheetApp.openById(_getHblSheetId());
}

// Dates in this workbook are messy: "30.12.2025", "04-05-2026",
// "07 - 04- 2026" (stray spaces), "18-06-2026\" (stray backslash). Returns
// { key:'yyyy-MM-dd', display:'dd/MM/yyyy' } or null when not date-like.
function _hblParseDate(v) {
  if (v instanceof Date && !isNaN(v)) {
    return {
      key:     Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      display: Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy')
    };
  }
  const s = String(v || '').replace(/\\/g, '').trim();
  const m = s.match(/^(\d{1,2})\s*[.\/-]\s*(\d{1,2})\s*[.\/-]\s*(\d{2,4})$/);
  if (!m) return null;
  const d = m[1].padStart(2, '0'), mo = m[2].padStart(2, '0');
  let y = m[3];
  if (y.length === 2) y = '20' + y;
  else if (y.length === 3) y = '2' + y; // "24.03.026" — typo'd 3-digit year
  return { key: y + '-' + mo + '-' + d, display: d + '/' + mo + '/' + y };
}

function getHblSheetNames() {
  try {
    const names = _getHblSS().getSheets().map(function(s){ return s.getName(); });
    // Default tab: the BJW-ADI corridor — matched on normalized name so
    // spacing/case variants of "BJW-ADI (Post Commissioning)" still hit.
    let def = null;
    const norm = function(s){ return String(s).toUpperCase().replace(/[^A-Z0-9]/g, ''); };
    for (var i = 0; i < names.length; i++) {
      if (norm(names[i]).indexOf('BJWADI') >= 0) { def = names[i]; break; }
    }
    return { success: true, sheets: names, defaultSheet: def || names[0] || null };
  } catch(e) { return { success: false, error: e.message, sheets: [] }; }
}

function getHblSheetData(sheetName) {
  try {
    const sh = _getHblSS().getSheetByName(sheetName);
    if (!sh) return { success: false, error: 'Sheet "' + sheetName + '" not found' };
    const values = sh.getDataRange().getValues();
    if (values.length < 2) return { success: true, rows: [] };

    // Header row: first row whose cells mention both "date" and "section".
    let hdrIdx = -1, map = {};
    for (var i = 0; i < Math.min(values.length, 5); i++) {
      const cand = {};
      values[i].forEach(function(h, ci) {
        const t = String(h || '').trim().toLowerCase();
        if (!t || t.length > 40) return;
        if (cand.date     === undefined && /^date/.test(t))                cand.date = ci;
        if (cand.section  === undefined && /section/.test(t))              cand.section = ci;
        if (cand.train    === undefined && /train/.test(t))                cand.train = ci;
        if (cand.fieldObs === undefined && /field observation/.test(t))    cand.fieldObs = ci;
        if (cand.remark   === undefined && /remark by hbl/.test(t))        cand.remark = ci;
        if (cand.nms      === undefined && /nms/.test(t))                  cand.nms = ci;
        if (cand.action   === undefined && /action taken/.test(t))         cand.action = ci;
        if (cand.staff    === undefined && /onboard/.test(t))              cand.staff = ci;
        if (cand.joint    === undefined && /joint/.test(t))                cand.joint = ci;
      });
      if (cand.date !== undefined && cand.section !== undefined) { hdrIdx = i; map = cand; break; }
    }
    if (hdrIdx < 0) return { success: false, error: 'Could not find the header row (looking for "Date" and "Section" columns) in "' + sheetName + '"' };

    const out = [];
    let carry = null; // last parsed date — rows with a blank/unparsable date inherit it
    for (var r = hdrIdx + 1; r < values.length; r++) {
      const row = values[r];
      // Junk rows: nothing but a serial number in the first column.
      const hasContent = row.some(function(c, ci){ return ci !== 0 && String(c || '').trim(); });
      if (!hasContent) continue;

      // Date/Section columns swap places partway through the sheet (both
      // orders coexist row to row) — trust whichever cell actually parses
      // as a date rather than the header positions.
      let dParsed = _hblParseDate(row[map.date]);
      let sectionVal = String(row[map.section] || '').trim();
      if (!dParsed) {
        const alt = _hblParseDate(row[map.section]);
        if (alt) { dParsed = alt; sectionVal = String(row[map.date] || '').trim(); }
      }
      if (dParsed) carry = dParsed;
      const eff = dParsed || carry;

      out.push({
        date:       eff ? eff.display : '',
        dateKey:    eff ? eff.key : '',
        section:    sectionVal,
        trainLoco:  map.train    !== undefined ? String(row[map.train]    || '').trim() : '',
        fieldObs:   map.fieldObs !== undefined ? String(row[map.fieldObs] || '').trim() : '',
        remark:     map.remark   !== undefined ? String(row[map.remark]   || '').trim() : '',
        nms:        map.nms      !== undefined ? String(row[map.nms]      || '').trim() : '',
        action:     map.action   !== undefined ? String(row[map.action]   || '').trim() : '',
        staff:      map.staff    !== undefined ? String(row[map.staff]    || '').trim() : '',
        joint:      map.joint    !== undefined ? String(row[map.joint]    || '').trim() : '',
        _ord:       r
      });
    }
    // Newest date first; original sheet order preserved within the same day.
    // Dates typed beyond ~2 days into the future are typos ("24.03.027") —
    // sink those to the bottom instead of letting them squat above today's
    // rows at the top of the list forever.
    const farFuture = Utilities.formatDate(new Date(Date.now() + 2*86400000), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    out.sort(function(a, b) {
      const aS = a.dateKey > farFuture, bS = b.dateKey > farFuture;
      if (aS !== bS) return aS ? 1 : -1;
      if (a.dateKey !== b.dateKey) return a.dateKey < b.dateKey ? 1 : -1;
      return a._ord - b._ord;
    });
    out.forEach(function(o){ delete o._ord; });
    return { success: true, rows: out };
  } catch(e) { return { success: false, error: e.message }; }
}

// ── HBL/Medha remark cross-reference for All Entries ──────────────────
// The remark sheet's "Train No/Loco No" cell is free text ("59549\n30459 /
// WAP-7", "22963 / 37164 WAP-7\nBHAVNAGAR…") — extract every 4-6 digit
// number and match against the observation's train & loco for the same date.
// Rows listing both numbers must match both; rows listing only one number
// match on either, so entries like "Loco No. 22563" still resolve.
function _buildHblRemarkIndex() {
  const idx = {};
  _getHblSS().getSheets().forEach(function(sh) {
    const res = getHblSheetData(sh.getName());
    if (!res.success) return;
    res.rows.forEach(function(r) {
      if (!r.dateKey || !r.remark) return;
      const nums = String(r.trainLoco || '').match(/\d{4,6}/g) || [];
      if (!nums.length) return;
      (idx[r.dateKey] = idx[r.dateKey] || []).push({ nums: nums, remark: r.remark, sheet: sh.getName() });
    });
  });
  return idx;
}

// Display-only enrichment — never written to the observation sheet unless
// the user explicitly saves it via insertHblRemarkToObs below. Best-effort:
// failures leave entries untouched.
function _attachHblRemarks(results) {
  if (!results.length) return;
  try {
    const idx = _buildHblRemarkIndex();
    results.forEach(function(r) {
      const dk = _normDateKey(r.date);
      const cands = dk && idx[dk];
      if (!cands) return;
      const train = String(r.trainNo || '').trim(), loco = String(r.locoNo || '').trim();
      if (!train && !loco) return;
      for (var i = 0; i < cands.length; i++) {
        const nums = cands[i].nums;
        const hasT = train && nums.indexOf(train) >= 0;
        const hasL = loco && nums.indexOf(loco) >= 0;
        const ok = nums.length >= 2 ? (hasT && hasL) : (hasT || hasL);
        if (ok) { r.hblRemark = cands[i].remark; r.hblRemarkSheet = cands[i].sheet; break; }
      }
    });
  } catch(e) { /* All Entries still loads without HBL remarks */ }
}

// Explicit user action: append a matched HBL/Medha remark into the entry's
// R&D Remark column in the backend sheet, with attribution. PIN-gated like
// every other sheet write.
function insertHblRemarkToObs(data) {
  try {
    const auth = validateRailwayStaffPin(data.staffName, data.pin);
    if (!auth.valid) return { success: false, error: 'Invalid staff name or PIN' };
    const remark = String(data.remark || '').trim();
    if (!remark) return { success: false, error: 'No remark to insert' };
    const found = _findObsRowBySr(data.srNo, data.section);
    if (!found) return { success: false, error: 'Entry not found' };
    const cell = found.sh.getRange(found.rowIndex, COL.RD_REMARK + 1);
    const cur = String(cell.getValue() || '').trim();
    if (cur.indexOf(remark) >= 0) return { success: false, error: 'This remark is already on the entry' };
    const stamp = '[From HBL Remark sheet] ' + remark + ' — ' + auth.name +
      ' (' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yy HH:mm') + ')';
    cell.setValue(cur ? cur + '\n' + stamp : stamp);
    SpreadsheetApp.flush();
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}
