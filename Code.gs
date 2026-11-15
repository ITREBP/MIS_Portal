
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Admin Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
///test
// Authenticate admin user
function authenticateUser(username, password) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var usersSheet = ss.getSheetByName('Users');
    if (!usersSheet) throw new Error('Users sheet not found');
    
    var data = usersSheet.getDataRange().getValues();
    
    // 🔑 Extract REC_ID from password prefix
    var recIdFromPassword = '';
    var passwordParts = password.split('-');
    if (passwordParts.length > 1 && /^\d+$/.test(passwordParts[0])) {
      recIdFromPassword = passwordParts[0];
    }
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === username && 
          data[i][1] === password && 
          (data[i][2] === 'admin' || data[i][2] === 'superadmin')) {
        
        var userRecId = data[i][15] ? data[i][15].toString() : '';
        
        // ✅ Verify REC_ID matches password prefix
        if (recIdFromPassword && userRecId !== recIdFromPassword) {
          continue; // Skip wrong school
        }
        
        const canCorrectCol = data[0].indexOf('canCorrectAttendance');
        const canDeleteUsersCol = data[0].indexOf('canDeleteUsers');   // 👈 ADD

var recInfo = getRecInfo_(userRecId);


        return {
          success: true,
          user: {
            canDeleteUsers: canDeleteUsersCol !== -1 ? (data[i][canDeleteUsersCol] || 'No') : 'No', 
            Username: data[i][0],
            UserType: data[i][2],
            Teacher_Name: data[i][3],
            student_delete: data[i][8] || 'No',
            canQueryAttendance: data[i][9] || 'No',
            canDownload: data[i][10] || 'No',
            canAmendSIRN: data[i][16] || 'No',
            canCorrectAttendance: canCorrectCol !== -1 ? (data[i][canCorrectCol] || 'No') : 'No',
            REC_ID: userRecId,  // ← Correct REC_ID
            REC_Name: recInfo.REC_Name,
LB_Name: recInfo.LB,

            allowedTabs: data[i][7] || 'users,user_permissions,class_sections,previous_date_permissions,teacher_assignments,student_data,student_attendance,qr_management,reserve_codes'
          }
        };
      }
    }
    return { success: false, message: 'Invalid credentials or insufficient permissions' };
  } catch (e) {
    Logger.log('Authentication error: ' + e.message);
    return { success: false, message: 'Authentication error: ' + e.message };
  }
}

function getRecInfo_(recId) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('RECNames');
    if (!sheet || !recId) return { REC_Name: '', LB: '' };
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][3]).trim() === String(recId).trim()) {
        return { REC_Name: data[i][4] || '', LB: data[i][2] || '' };
      }
    }
  } catch (e) {}
  return { REC_Name: '', LB: '' };
}



// Get data from a specific sheet
function getSheetData(sheetName, recId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(sheetName + ' sheet not found');
if (isPrevPermSheet_(sheetName)) ensurePrevPermColumns_(sheet);    
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const result = [];
    
    // Find REC_ID column index
    const recIdCol = headers.indexOf('REC_ID');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        // If REC_ID column exists and we have a REC_ID filter
        if (recIdCol !== -1 && recId) {
          const rowRecId = data[i][recIdCol] ? data[i][recIdCol].toString() : '';
          // ONLY show exact matches - hide blank and different schools
          if (rowRecId !== recId) {
            continue; // Skip this row - different school or blank
          }
        }
        
        const row = {};
        for (let j = 0; j < headers.length; j++) {
          const v = data[i][j];
if (v instanceof Date) {
  row[headers[j]] = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
} else {
  row[headers[j]] = v !== null && v !== undefined ? v.toString() : '';
}

        }
        row._rowNumber = i + 1;
        result.push(row);
      }
    }
    return result;
  } catch (e) {
    Logger.log('Failed to fetch data from ' + sheetName + ': ' + e.message);
    throw new Error('Failed to fetch data from ' + sheetName + ': ' + e.message);
  }
}

// ============ CLASS_MAPPING MIGRATION (one-time, REC-aware, atomic) ============
function ensureClassMappingSheet_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log('ensureClassMappingSheet_: could not acquire lock within 30s — skipping migration this run.');
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Already migrated → nothing to do
    if (ss.getSheetByName('Class_Mapping')) return;

    const src = ss.getSheetByName('Class_Sections');
    if (!src) return;

    const data = src.getDataRange().getValues();
    if (data.length <= 1) {
      const dst = ss.insertSheet('Class_Mapping');
      dst.appendRow(['Class', 'Sections', 'REC_ID']);
      dst.setFrozenRows(1);
      SpreadsheetApp.flush();
      Logger.log('Class_Mapping sheet created (source was empty).');
      return;
    }

    const headers = data[0];
    const secIdCol = headers.indexOf('Class_Section');
    const recCol   = headers.indexOf('REC_ID');
    if (secIdCol === -1 || recCol === -1) {
      Logger.log('Migration skipped — required columns missing in Class_Sections');
      return;
    }

    // 1. Group by (Class, REC_ID) — two schools may both have class "I"
    const byKey = {};
    for (let i = 1; i < data.length; i++) {
      const combo = String(data[i][secIdCol] || '').trim();
      if (!combo) continue;

      const dashPos = combo.lastIndexOf('-');
      if (dashPos <= 0) {
        Logger.log(`Migration: skipping malformed row ${i + 1} — "${combo}" has no hyphen`);
        continue;
      }

      const cls = combo.slice(0, dashPos).trim();
      const sec = combo.slice(dashPos + 1).trim().toUpperCase();
      if (!cls || !sec) continue;

      const recId = String(data[i][recCol] || '').trim();
      if (!recId) {
        Logger.log(`Migration: skipping row ${i + 1} — no REC_ID in Col B`);
        continue;
      }

      const key = `${cls}|${recId}`;
      if (!byKey[key]) byKey[key] = { class: cls, recId: recId, sections: new Set() };
      byKey[key].sections.add(sec);
    }

    // 2. Create Class_Mapping and write one row per (Class, REC_ID)
    const dst = ss.insertSheet('Class_Mapping');
    dst.appendRow(['Class', 'Sections', 'REC_ID']);

    const expectedCount = Object.keys(byKey).length;
    const rows = Object.keys(byKey)
      .sort()
      .map(key => {
        const e = byKey[key];
        return [e.class, Array.from(e.sections).sort().join(','), e.recId];
      });

    if (rows.length > 0) {
      dst.getRange(2, 1, rows.length, 3).setValues(rows);
    }
    dst.setFrozenRows(1);
    SpreadsheetApp.flush();

    // 3. VERIFY before clearing
    const written = dst.getLastRow() - 1;
    if (written !== expectedCount) {
      Logger.log(`Migration ABORTED before clearing: wrote ${written} rows, expected ${expectedCount}. Col C/D/E preserved.`);
      ss.deleteSheet(dst);
      return;
    }

    // 4. Clear Col C/D/E in Class_Sections (headers + data, all RECs)
    const lastCol = src.getLastColumn();
    if (lastCol >= 3) {
      const totalRows = src.getLastRow();
      if (totalRows >= 1) {
        src.getRange(1, 3, totalRows, lastCol - 2).clearContent();
      }
    }

    SpreadsheetApp.flush();
    Logger.log(`Migration complete: ${written} class mappings created. Class_Sections Col C+ cleared.`);
  } catch (e) {
    Logger.log('Migration FAILED: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}


function getClassSectionData(recId) {
  try {
    ensureClassMappingSheet_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Class_Sections');
    if (!sheet) throw new Error('Class_Sections sheet not found');

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const secCol = headers.indexOf('Class_Section');
    const recCol = headers.indexOf('REC_ID');
    if (secCol === -1) throw new Error('Class_Section column not found');

    const result = [];
    for (let i = 1; i < data.length; i++) {
      const sec = String(data[i][secCol] || '').trim();
      if (!sec) continue;
      if (recId && recCol !== -1) {
        const rowRec = String(data[i][recCol] || '').trim();
        if (rowRec !== recId) continue;
      }
      result.push({ Class_Section: sec, _rowNumber: i + 1 });
    }
    return result;
  } catch (e) {
    Logger.log('Failed to fetch class section data: ' + e.message);
    throw new Error('Failed to fetch class section data: ' + e.message);
  }
}

function getClassMappingData(recId) {
  try {
    ensureClassMappingSheet_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Class_Mapping');
    if (!sheet) throw new Error('Class_Mapping sheet not found');

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const clsCol  = headers.indexOf('Class');
    const secCol  = headers.indexOf('Sections');
    const recCol  = headers.indexOf('REC_ID');
    if (clsCol === -1 || secCol === -1) throw new Error('Required columns missing in Class_Mapping');

    const result = [];
    for (let i = 1; i < data.length; i++) {
      const cls = String(data[i][clsCol] || '').trim();
      if (!cls) continue;
      if (recId && recCol !== -1) {
        const rowRec = String(data[i][recCol] || '').trim();
        if (rowRec !== recId) continue;
      }
      result.push({
        Class: cls,
        Sections: String(data[i][secCol] || '').trim(),
        REC_ID: recCol !== -1 ? String(data[i][recCol] || '').trim() : '',
        _rowNumber: i + 1
      });
    }
    return result;
  } catch (e) {
    Logger.log('Failed to fetch class mapping data: ' + e.message);
    throw new Error('Failed to fetch class mapping data: ' + e.message);
  }
}


// ============ PREVIOUS_DATE_PERMISSIONS AUTO-SETUP ============
const PREV_PERM_SHEET = 'Previous_Date_Permissions';
const PREV_PERM_NEW_COLUMNS = ['Attendance_Type', 'Granted_At', 'Expires_At'];

function isPrevPermSheet_(sheetName) {
  return String(sheetName || '').trim().toLowerCase() === PREV_PERM_SHEET.toLowerCase();
}

function ensurePrevPermColumns_(sheet) {
  if (!sheet) return [];
  let lastCol = sheet.getLastColumn();
  if (lastCol < 1) return [];
  let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const missing = PREV_PERM_NEW_COLUMNS.filter(c => headers.indexOf(c) === -1);
  if (missing.length === 0) return headers;
  sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
  const totalRows = sheet.getMaxRows();
  missing.forEach(function (name, i) {
    const colIndex = lastCol + 1 + i;
    if (name === 'Granted_At' || name === 'Expires_At') {
      if (totalRows > 1) {
        sheet.getRange(2, colIndex, totalRows - 1, 1).setNumberFormat('dd-MMM-yyyy hh:mm AM/PM');
      }
      sheet.setColumnWidth(colIndex, 180);
    }
  });
  SpreadsheetApp.flush();
  lastCol = sheet.getLastColumn();
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
}

function setupPrevPermColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PREV_PERM_SHEET);
  if (!sheet) throw new Error(PREV_PERM_SHEET + ' sheet not found');
  ensurePrevPermColumns_(sheet);
}

function stampPrevPermTimestamps_(rowData) {
  const grantedAt = new Date();
  rowData.Granted_At = grantedAt;
  rowData.Expires_At = new Date(grantedAt.getTime() + 24 * 60 * 60 * 1000);
  const t = String(rowData.Attendance_Type || 'Both').trim().toLowerCase();
  rowData.Attendance_Type = t === 'episodic' ? 'Episodic' : t === 'normal' ? 'Normal' : 'Both';
}


function sanitizeRowForClient_(row) {
  const tz = Session.getScriptTimeZone() || 'Asia/Karachi';
  Object.keys(row).forEach(function (k) {
    if (row[k] instanceof Date) {
      row[k] = Utilities.formatDate(row[k], tz, 'dd-MMM-yyyy hh:mm a');
    }
  });
  return row;
}

// ============ TERMS / CLASS_TERMS (v4) ============
// Terms sheet: Term_ID, Term_Name, Is_Active. Names ONLY — no dates.
// Class_Terms sheet: Class_Section, Term_ID, Is_Active, Start_Date, End_Date, Notes, REC_ID.
// Dates live ONLY in Class_Terms. One term name (e.g. Term-2026-2027) is shared
// across classes, but each class has its own dates. Reporting filters by date
// range from Class_Terms — never by a Term_ID stamp, and never from Terms.
// Marking is untouched. Summary dashboard is excluded (it has no dates).
const TERMS_SHEET = 'Terms';
const CLASS_TERMS_SHEET = 'Class_Terms';

function isTermsSheet_(sheetName) {
  return String(sheetName || '').trim().toLowerCase() === 'terms';
}

function isClassTermsSheet_(sheetName) {
  return String(sheetName || '').trim().toLowerCase() === 'class_terms';
}

// Is_Active is stored as the strings 'TRUE' / 'FALSE' in the sheets.
function isActiveFlag_(v) {
  return String(v == null ? '' : v).trim().toLowerCase() === 'true';
}

// Parse a sheet value to a local-midnight Date, or null.
// Accepts Date objects, 'yyyy-MM-dd', 'dd-MMM-yyyy' and other Date-parseable strings.
// 'yyyy-MM-dd' is parsed as a LOCAL date (no UTC shift), so the day is correct in PKT.
function parseTermDate_(v) {
  if (v == null || v === '') return null;
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return null;
    return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }
  const str = String(v).trim();
  if (!str) return null;
  let m = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) {
    const d = new Date(+m[1], +m[2] - 1, +m[3]);
    return isNaN(d.getTime()) ? null : d;
  }
  m = str.match(/^(\d{2})-([A-Za-z]{3})-(\d{4})/);
  if (m) {
    const months = { Jan: 0, Feb: 1, Mar: 2, Apr: 3, May: 4, Jun: 5, Jul: 6, Aug: 7, Sep: 8, Oct: 9, Nov: 10, Dec: 11 };
    const mi = months[m[2]];
    if (mi === undefined) return null;
    const d = new Date(+m[3], mi, +m[1]);
    return isNaN(d.getTime()) ? null : d;
  }
  const d = new Date(str);
  if (isNaN(d.getTime())) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

// 'yyyy-MM-dd' for a Date (local time).
function termDateKey_(d) {
  if (!d) return '';
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}

// Expand a Class_Section cell to individual class tokens.
// '{I,II,III-B}' -> ['I', 'II', 'III-B'];  'VII-A' -> ['VII-A'];  'I' -> ['I'].
function expandClassCell_(cell) {
  let str = String(cell || '').trim();
  if (!str) return [];
  if (str.charAt(0) === '{' && str.charAt(str.length - 1) === '}') str = str.slice(1, -1);
  return str.split(',').map(function (t) { return t.trim(); }).filter(Boolean);
}

// Base class of a class-section: 'VII-A' -> 'VII', 'KGI' -> 'KGI'.
function baseClassOf_(cs) {
  const str = String(cs || '').trim();
  const i = str.indexOf('-');
  return (i === -1 ? str : str.slice(0, i)).trim();
}

// Does token `tok` (from a Class_Terms row) cover class-section `cs`?
// A bare token ('I') covers every section of that class ('I-A', 'I-B', 'I' itself).
// A section token ('III-B') covers only that exact class-section. Case-insensitive.
function tokenMatches_(tok, cs) {
  const t = String(tok || '').trim();
  const c = String(cs || '').trim();
  if (!t || !c) return false;
  if (t.indexOf('-') === -1) return baseClassOf_(c).toLowerCase() === t.toLowerCase();
  return t.toLowerCase() === c.toLowerCase();
}

// Does a Class_Section cell cover class-section `cs`?
function classCellMatches_(cell, cs) {
  return expandClassCell_(cell).some(function (tok) { return tokenMatches_(tok, cs); });
}

// Do two token lists cover any common class-section? Returns the overlapping
// token from the first list, or null. Bare tokens overlap any section of their class.
function tokensOverlap_(aTokens, bTokens) {
  for (let i = 0; i < aTokens.length; i++) {
    for (let j = 0; j < bTokens.length; j++) {
      if (tokenMatches_(aTokens[i], bTokens[j]) || tokenMatches_(bTokens[j], aTokens[i])) return aTokens[i];
    }
  }
  return null;
}

// Class_Terms rows for one REC (REC-filtered by getSheetData). [] if the sheet is missing.
function getClassTermsForRec_(recId) {
  try {
    return getSheetData(CLASS_TERMS_SHEET, recId || '');
  } catch (e) {
    return [];
  }
}

// All Terms, active first. [] if the Terms sheet is missing.
function getActiveTerms() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TERMS_SHEET);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    const h = data[0];
    const idI = h.indexOf('Term_ID');
    const nmI = h.indexOf('Term_Name');
    const acI = h.indexOf('Is_Active');
    if (idI === -1) return [];
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][idI]) continue;
      rows.push({
        Term_ID: String(data[i][idI]).trim(),
        Term_Name: nmI !== -1 ? String(data[i][nmI] || '').trim() : String(data[i][idI]).trim(),
        Is_Active: acI !== -1 ? String(data[i][acI]) : ''
      });
    }
    rows.sort(function (a, b) {
      return (isActiveFlag_(b.Is_Active) ? 1 : 0) - (isActiveFlag_(a.Is_Active) ? 1 : 0);
    });
    return rows;
  } catch (e) {
    return [];
  }
}

// First matching ACTIVE Class_Terms row for this REC + class-section (+ optional termId).
// First match in sheet order wins. Returns { start, end, termId, classSection }
// with 'yyyy-MM-dd' dates, or null when nothing matches.
function getTermDateRangeForClassSection(recId, classSection, termId) {
  const rows = getClassTermsForRec_(recId);
  const wantTerm = String(termId || '').trim().toLowerCase();
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (!isActiveFlag_(r.Is_Active)) continue;
    if (wantTerm && String(r.Term_ID || '').trim().toLowerCase() !== wantTerm) continue;
    if (!classCellMatches_(String(r.Class_Section || ''), classSection)) continue;
    const s = parseTermDate_(r.Start_Date);
    const e = parseTermDate_(r.End_Date);
    if (!s || !e) continue;
    return {
      start: termDateKey_(s),
      end: termDateKey_(e),
      termId: String(r.Term_ID || '').trim(),
      classSection: String(classSection || '').trim()
    };
  }
  return null;
}

// Short alias for the report lookup.
function getTermDateRange(recId, classSection, termId) {
  return getTermDateRangeForClassSection(recId, classSection, termId);
}

// All ACTIVE term ranges covering one class-section (one entry per term,
// first matching row wins per term). Used to list the terms a class belongs to.
function getTermsForClassSection(recId, classSection) {
  const rows = getClassTermsForRec_(recId);
  const seen = {};
  const out = [];
  rows.forEach(function (r) {
    if (!isActiveFlag_(r.Is_Active)) return;
    if (!classCellMatches_(String(r.Class_Section || ''), classSection)) return;
    const tid = String(r.Term_ID || '').trim();
    if (!tid || seen[tid.toLowerCase()]) return;
    const s = parseTermDate_(r.Start_Date);
    const e = parseTermDate_(r.End_Date);
    if (!s || !e) return;
    seen[tid.toLowerCase()] = true;
    out.push({ termId: tid, start: termDateKey_(s), end: termDateKey_(e) });
  });
  return out;
}

// Validate a Terms row. excludeRow = sheet row number to skip (Edit), or 0 (Add).
// Returns an error message string, or null when OK.
function validateTermsRow_(sheet, headers, rowData, excludeRow) {
  const termId = String(rowData.Term_ID || '').trim();
  if (!termId) return 'Term_ID is required.';
  const idIdx = headers.indexOf('Term_ID');
  if (idIdx === -1) return null;
  const all = sheet.getDataRange().getValues();
  for (let i = 1; i < all.length; i++) {
    if (excludeRow && (i + 1) === excludeRow) continue;
    if (String(all[i][idIdx] || '').trim().toLowerCase() === termId.toLowerCase()) {
      return 'Term "' + termId + '" already exists. Edit the existing term instead.';
    }
  }
  return null;
}

// Validate a Class_Terms row. excludeRow = sheet row number to skip (Edit), or 0 (Add).
// Rules: Term_ID must exist in Terms; Start_Date <= End_Date; no exact duplicate
// (same Term_ID + same Class_Section + same REC_ID); and for the same REC and the
// same class, the date ranges of two DIFFERENT terms must not overlap.
// Returns an error message string, or null when OK.
function validateClassTermsRow_(sheet, headers, rowData, recId, excludeRow) {
  const termId = String(rowData.Term_ID || '').trim();
  if (!termId) return 'Term is required. Pick a term from the list.';
  // Term must exist in Terms (matches Term_ID or Term_Name, case-insensitive,
  // so both 'Term-2026-2027' and '2026-2027' spellings are accepted).
  let termExists = false;
  try {
    const ts = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TERMS_SHEET);
    if (ts) {
      const td = ts.getDataRange().getValues();
      const th = td[0] || [];
      const idI = th.indexOf('Term_ID');
      const nmI = th.indexOf('Term_Name');
      for (let i = 1; i < td.length; i++) {
        const idV = idI !== -1 ? String(td[i][idI] || '').trim().toLowerCase() : '';
        const nmV = nmI !== -1 ? String(td[i][nmI] || '').trim().toLowerCase() : '';
        if (termId.toLowerCase() === idV || termId.toLowerCase() === nmV) { termExists = true; break; }
      }
    }
  } catch (e) { termExists = false; }
  if (!termExists) return 'Term "' + termId + '" does not exist in Terms. Add it in Terms first.';
  const start = parseTermDate_(rowData.Start_Date);
  const end = parseTermDate_(rowData.End_Date);
  if (!start || !end) return 'Start Date and End Date are required and must be valid dates.';
  if (start.getTime() > end.getTime()) return 'Start Date must be on or before End Date.';
  const rec = String(recId || rowData.REC_ID || '').trim();
  const csIdx = headers.indexOf('Class_Section');
  const termIdx = headers.indexOf('Term_ID');
  const recIdx = headers.indexOf('REC_ID');
  const sIdx = headers.indexOf('Start_Date');
  const eIdx = headers.indexOf('End_Date');
  const actIdx = headers.indexOf('Is_Active');
  if (csIdx === -1 || termIdx === -1) return null;
  const newCell = String(rowData.Class_Section || '').trim();
  const newTokens = expandClassCell_(newCell);
  const all = sheet.getDataRange().getValues();
  for (let i = 1; i < all.length; i++) {
    const rowNum = i + 1;
    if (excludeRow && rowNum === excludeRow) continue;
    const exRec = recIdx !== -1 ? String(all[i][recIdx] || '').trim() : '';
    if (rec && exRec && exRec !== rec) continue;
    if (!rec && exRec) continue;
    const exTerm = String(all[i][termIdx] || '').trim();
    const exCell = String(all[i][csIdx] || '').trim();
    // Exact duplicate: same term + same class cell + same REC.
    if (exTerm.toLowerCase() === termId.toLowerCase() && exCell.toLowerCase() === newCell.toLowerCase()) {
      return 'This Class_Terms row already exists for Term "' + termId + '" and "' + newCell + '". Edit the existing row instead.';
    }
    // Overlap: same class, overlapping dates — different term, or same term
// while both rows are active. A FALSE row is history, so it never blocks,
// and a new FALSE row never gets blocked. To correct dates, set the old
// row Is_Active = FALSE first, then add the new row.
if (newTokens.length) {
  const overlapTok = tokensOverlap_(newTokens, expandClassCell_(exCell));
  if (overlapTok) {
    const exActive = actIdx === -1 ? true : isActiveFlag_(all[i][actIdx]);
    const newActive = isActiveFlag_(rowData.Is_Active);
    const sameTerm = exTerm.toLowerCase() === termId.toLowerCase();
    const blocked = sameTerm ? (exActive && newActive) : true;
    if (blocked) {
      const exS = sIdx !== -1 ? parseTermDate_(all[i][sIdx]) : null;
      const exE = eIdx !== -1 ? parseTermDate_(all[i][eIdx]) : null;
      if (exS && exE && start.getTime() <= exE.getTime() && end.getTime() >= exS.getTime()) {
        if (sameTerm) {
          return 'Class "' + overlapTok + '" already has active dates ' + termDateKey_(exS) + ' to ' + termDateKey_(exE) + ' in Term "' + exTerm + '" — set the old row Is_Active = FALSE first, then add the new dates.';
        }
        return 'Class "' + overlapTok + '" already has dates ' + termDateKey_(exS) + ' to ' + termDateKey_(exE) + ' in Term "' + exTerm + '" — dates cannot overlap for the same class over two separate terms.';
      }
    }
  }
}

  }
  return null;
}

// Add a new row to a sheet
function addSheetRow(sheetName, rowData, currentUsername) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(sheetName + ' sheet not found');
    
    // Handle special cases first
    if (sheetName === 'Student_Data') {
      // Pass REC_ID from rowData if available
      const recId = rowData.REC_ID || '';
      return addStudent(rowData, currentUsername, recId);
    }
    
    if (sheetName === 'Teacher_Assignments') {
      const existingData = sheet.getDataRange().getValues();
      const exists = existingData.some((row, i) => i > 0 && row[0] === rowData.Username);
      if (exists) {
        return { 
          success: false, 
          message: `Username ${rowData.Username} already has assigned classes. Please edit the existing record.` 
        };
      }
    }
    
    if (isPrevPermSheet_(sheetName)) ensurePrevPermColumns_(sheet);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];



if (isPrevPermSheet_(sheetName)) {
  const today = new Date();
today.setHours(0, 0, 0, 0);
const parts = String(rowData.Allowed_Date).split('-');
const allowed = new Date(parts[0], parts[1] - 1, parts[2]);
if (allowed > today) return { success: false, message: 'Allowed Date cannot be in the future.' };

  const uIdx = headers.indexOf('Username');
  const dIdx = headers.indexOf('Allowed_Date');
  if (uIdx !== -1 && dIdx !== -1 && rowData.Username && rowData.Allowed_Date) {
    const newDate = new Date(rowData.Allowed_Date);
    const newDateStr = !isNaN(newDate) ? newDate.toDateString() : String(rowData.Allowed_Date).trim();
    const all = sheet.getDataRange().getValues();
    const dup = all.some((r, i) => {
      if (i === 0) return false;
      if (String(r[uIdx] || '').trim() !== String(rowData.Username).trim()) return false;
      const d = r[dIdx] instanceof Date ? r[dIdx] : new Date(r[dIdx]);
      const dStr = !isNaN(d) ? d.toDateString() : String(r[dIdx] || '').trim();
      return dStr === newDateStr;
    });
    if (dup) return { success: false, message: 'This user already has permission for this date.' };
  }
}



    if (isPrevPermSheet_(sheetName)) stampPrevPermTimestamps_(rowData);

    // Get REC_ID from the current user
    let recId = '';
    if (currentUsername) {
      const usersSheet = ss.getSheetByName('Users');
      if (usersSheet) {
        const userData = usersSheet.getDataRange().getValues();
        const userRow = userData.find(row => row[0] === currentUsername);
        if (userRow) {
          recId = userRow[15] ? userRow[15].toString() : ''; // Column P is index 15
        }
      }
    }

    // 🔧 FIX: For Users sheet, set REC_ID in rowData BEFORE building the row
    if (sheetName === 'Users' && recId) {
      rowData.REC_ID = recId;
    }

if (sheetName.toLowerCase() === 'users') {
  const uIdx = headers.indexOf('Username');
  const tIdx = headers.indexOf('UserType');
  const rIdx = headers.indexOf('REC_ID');
  const newU = String(rowData.Username || '').trim();
  const newT = String(rowData.UserType || '').trim();
  const newR = String(rowData.REC_ID || recId || '').trim();
  if (uIdx !== -1 && tIdx !== -1 && rIdx !== -1 && newU) {
    const all = sheet.getDataRange().getValues();
    const dup = all.some((r, i) => i > 0 && String(r[uIdx] || '').trim().toLowerCase() === newU.toLowerCase() && String(r[tIdx] || '').trim() === newT && String(r[rIdx] || '').trim() === newR);
    if (dup) return { success: false, message: 'Username ' + newU + ' already exists as ' + newT + ' for this school.' };
  }
}

if (sheetName.toLowerCase() === 'users' && rowData.Password) {
  const expectedRec = String(rowData.REC_ID || recId || '').trim();
  const pwd = String(rowData.Password).trim();
  if (expectedRec) {
    const prefix = expectedRec + '-';
    if (!pwd.startsWith(prefix) || pwd === prefix) {
      return { success: false, message: 'Password must start with ' + prefix + ' and have a password after it.' };
    }
  }
}
    // Terms / Class_Terms (v4): normalise Is_Active, then validate.
    if (isTermsSheet_(sheetName)) {
      rowData.Is_Active = isActiveFlag_(rowData.Is_Active) ? 'TRUE' : 'FALSE';
      const tErr = validateTermsRow_(sheet, headers, rowData, 0);
      if (tErr) return { success: false, message: tErr };
    }
    if (isClassTermsSheet_(sheetName)) {
      rowData.Is_Active = isActiveFlag_(rowData.Is_Active) ? 'TRUE' : 'FALSE';
      const ctErr = validateClassTermsRow_(sheet, headers, rowData, recId, 0);
      if (ctErr) return { success: false, message: ctErr };
    }
    // Build new row
    const newRow = headers.map(header => {
      const value = rowData[header];
      
      // Auto-fill REC_ID if not provided
      if (header === 'REC_ID' && !value && recId) {
        return recId;
      }
      
      // Handle special formatting cases
        if (sheetName.toLowerCase() === 'user_permissions' && !['Username', 'REC_ID'].includes(header)) {
        return value === 'true' || value === true;
      }
    if (sheetName === 'Users' && ['StartTime', 'EndTime'].includes(header)) {
  return value ? value.toString() : '';
}
// 🆕 Handle all multiselect (array) fields → comma-joined strings
if (sheetName.toLowerCase() === 'users' && ['allowedTabs', 'AllowedDays'].includes(header)) {

  if (Array.isArray(value)) return value.join(',');
  return value ? value.toString() : '';
}

      if (sheetName === 'Users' && ['canQueryAttendance', 'canDownload'].includes(header)) {
        return value === 'Yes' ? 'Yes' : 'No';
      }
      
      return value !== undefined && value !== null ? value : '';
    });
    
    sheet.appendRow(newRow);
    SpreadsheetApp.flush();
    
    // Handle time format for Users sheet
    if (sheetName === 'Users') {
      const startTimeIndex = headers.indexOf('StartTime');
      const endTimeIndex = headers.indexOf('EndTime');
      const lastRow = sheet.getLastRow();
      
      if (startTimeIndex > -1) {
        sheet.getRange(lastRow, startTimeIndex + 1).setNumberFormat('@');
      }
      if (endTimeIndex > -1) {
        sheet.getRange(lastRow, endTimeIndex + 1).setNumberFormat('@');
      }
    }
    
    return { 
      success: true, 
      message: 'Record added successfully', 
      data: sanitizeRowForClient_({ ...rowData, _rowNumber: sheet.getLastRow() })
    };
  } catch (e) {
    Logger.log('Error in addSheetRow: ' + e.message);
    return { success: false, message: 'Failed to add row to ' + sheetName + ': ' + e.message };
  }
}


function updateSheetRow(sheetName, rowIndex, rowData, currentUsername) {
  try {
    // ✅ Get username from multiple sources
    let username = currentUsername;
    
    if (!username) {
      username = rowData.Username || rowData.currentUsername || '';
    }
    
    if (!username) {
      const sessionUser = Session.getActiveUser().getEmail();
      if (sessionUser) {
        username = sessionUser.split('@')[0];
      }
    }
    
    if (!username) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const usersSheet = ss.getSheetByName('Users');
      if (usersSheet) {
        const userData = usersSheet.getDataRange().getValues();
        for (let i = 1; i < userData.length; i++) {
          if (userData[i][15] && userData[i][15].toString() === rowData.REC_ID) {
            username = userData[i][0];
            break;
          }
        }
      }
    }
    
    if (!username) {
      return { 
        success: false, 
        message: 'Unable to identify user. Please re-login.' 
      };
    }
    
    Logger.log(`updateSheetRow called for sheet: ${sheetName}, rowIndex: ${rowIndex}, username: ${username}`);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('No active spreadsheet found');
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`${sheetName} sheet not found`);
    
    // Block superadmin modifications
    if (sheetName === 'Users') {
      const rowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (rowData[0] === 'superadmin') {
        return { success: false, message: 'Superadmin account cannot be modified' };
      }
    }
    
    if (!Number.isInteger(rowIndex) || rowIndex < 2 || rowIndex > sheet.getLastRow()) {
      Logger.log(`Invalid row index: ${rowIndex}. Last row: ${sheet.getLastRow()}`);
      return { success: false, message: `Invalid row index: ${rowIndex}. Must be between 2 and ${sheet.getLastRow()}` };
    }
    
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    const rangeProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    if (protections.length > 0 || rangeProtections.length > 0) {
      const userEmail = Session.getActiveUser().getEmail();
      let canEdit = true;
      protections.forEach(protection => {
        if (!protection.getEditors().includes(userEmail)) {
          canEdit = false;
        }
      });
      rangeProtections.forEach(protection => {
        const protectedRange = protection.getRange();
        if (protectedRange.getRow() <= rowIndex && rowIndex <= protectedRange.getLastRow()) {
          if (!protection.getEditors().includes(userEmail)) {
            canEdit = false;
          }
        }
      });
      if (!canEdit) {
        Logger.log(`User ${userEmail} lacks permission to edit protected ranges in ${sheetName}`);
        return { success: false, message: 'Permission denied: Sheet or range is protected' };
      }
    }

    if (isPrevPermSheet_(sheetName)) ensurePrevPermColumns_(sheet);
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!headers || headers.length === 0) {
      throw new Error('No headers found in the sheet');
    }
    

// Edit: duplicate + password prefix check for Users
if (sheetName.toLowerCase() === 'users') {
  const uIdx = headers.indexOf('Username');
  const tIdx = headers.indexOf('UserType');
  const rIdx = headers.indexOf('REC_ID');
  const existingRow = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  const newU = String(rowData.Username || '').trim();
  const newT = String(rowData.UserType || '').trim();
  const newR = String(rowData.REC_ID || (rIdx !== -1 ? existingRow[rIdx] : '') || '').trim();
  if (uIdx !== -1 && tIdx !== -1 && rIdx !== -1 && newU) {
    const all = sheet.getDataRange().getValues();
    const dup = all.some((r, i) => (i + 1) !== rowIndex && i > 0 && String(r[uIdx] || '').trim().toLowerCase() === newU.toLowerCase() && String(r[tIdx] || '').trim() === newT && String(r[rIdx] || '').trim() === newR);
    if (dup) return { success: false, message: 'Username ' + newU + ' already exists as ' + newT + ' for this school.' };
  }
  if (rowData.Password) {
    const pwd = String(rowData.Password).trim();
    if (newR) {
      const prefix = newR + '-';
      if (!pwd.startsWith(prefix) || pwd === prefix) {
        return { success: false, message: 'Password must start with ' + prefix + ' and have a password after it.' };
      }
    }
  }
}


    // Terms / Class_Terms (v4): normalise Is_Active, then validate (own row excluded).
    if (isTermsSheet_(sheetName) || isClassTermsSheet_(sheetName)) {
      rowData.Is_Active = isActiveFlag_(rowData.Is_Active) ? 'TRUE' : 'FALSE';
    }
    if (isTermsSheet_(sheetName)) {
      const tErr = validateTermsRow_(sheet, headers, rowData, rowIndex);
      if (tErr) return { success: false, message: tErr };
    }
    if (isClassTermsSheet_(sheetName)) {
      const exRow = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
      const rIdx = headers.indexOf('REC_ID');
      const ctRec = String(rowData.REC_ID || (rIdx !== -1 ? exRow[rIdx] : '') || '').trim();
      const ctErr = validateClassTermsRow_(sheet, headers, rowData, ctRec, rowIndex);
      if (ctErr) return { success: false, message: ctErr };
    }

    // ✅ STUDENT DATA SECTION
    if (sheetName === 'Student_Data') {
      const rowRange = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn());
      const rowValues = rowRange.getValues()[0];
      const originalStdId = rowValues[0];
      
      const currentUser = username;
      const currentDate = new Date();
      const formattedDate = Utilities.formatDate(currentDate, 'GMT+0500', 'dd-MMM-yyyy');
      
      // ✅ Check if SIRN is actually changing
      const newStdId = rowData.Std_ID ? rowData.Std_ID.toString() : '';
      const oldStdId = originalStdId ? originalStdId.toString() : '';
      
      if (newStdId && newStdId !== oldStdId) {
        // 🔒 CHECK PERMISSION TO AMEND SIRN
        const usersSheet = ss.getSheetByName('Users');
        if (usersSheet) {
          const userHeaders = usersSheet.getRange(1, 1, 1, usersSheet.getLastColumn()).getValues()[0];
          const canAmendSIRNCol = userHeaders.indexOf('canAmendSIRN');
          
          if (canAmendSIRNCol === -1) {
            Logger.log('WARNING: canAmendSIRN column not found in Users sheet');
            return { 
              success: false, 
              message: 'System configuration error: canAmendSIRN permission not configured. Please contact administrator.' 
            };
          }
          
          const userData = usersSheet.getDataRange().getValues();
          const userRow = userData.find(row => row[0] === currentUser);
          
          if (!userRow) {
            Logger.log(`User ${currentUser} not found in Users sheet`);
            return { 
              success: false, 
              message: 'User not found. Please re-login.' 
            };
          }
          
          // ✅ IMPORTANT: Blank = "No" (default)
          const canAmendSIRN = userRow[canAmendSIRNCol] || 'No';
          
          if (canAmendSIRN !== 'Yes') {
            Logger.log(`User ${currentUser} does not have permission to amend SIRN. canAmendSIRN: "${canAmendSIRN}"`);
            return { 
              success: false, 
              message: 'You do not have permission to amend SIRN (Student ID)' 
            };
          }
          
          Logger.log(`User ${currentUser} has permission to amend SIRN: ${canAmendSIRN}`);
        } else {
          Logger.log('ERROR: Users sheet not found');
          return { 
            success: false, 
            message: 'System error: Users sheet not found' 
          };
        }
        
        // ✅ Validate the new SIRN
        if (!/^\d+$/.test(newStdId)) {
          Logger.log(`Std_ID validation failed: ${newStdId} is not numeric`);
          return { success: false, message: 'Std_ID must be numeric' };
        }
        
        // 🔒 Block the change if the OLD Std_ID already has attendance marked
        const attendanceSheet = ss.getSheetByName('Student_Attendance');
        if (attendanceSheet) {
          const attData = attendanceSheet.getDataRange().getValues();
          const attStdIdCol = attData[0].indexOf('Std_ID');
          if (attStdIdCol !== -1) {
            const hasAttendance = attData.some((row, idx) => idx > 0 && row[attStdIdCol] && row[attStdIdCol].toString() === oldStdId);
            if (hasAttendance) {
              Logger.log(`SIRN change blocked: Std_ID ${oldStdId} already has attendance records`);
              return { success: false, message: `Cannot change SIRN — attendance records already exist for Std_ID ${oldStdId}. SIRN can only be amended before any attendance is marked.` };
            }
          }
        }
        
        // ✅ Check duplicate ONLY within the same REC_ID
        const stdIdCol = headers.indexOf('Std_ID');
        const recIdCol = headers.indexOf('REC_ID');
        const currentRowRecId = rowValues[recIdCol] ? rowValues[recIdCol].toString() : '';
        
        if (stdIdCol !== -1 && recIdCol !== -1) {
          const existingData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
          
          for (let i = 0; i < existingData.length; i++) {
            const rowNum = i + 2;
            if (rowNum !== rowIndex) {
              const existingStdId = existingData[i][stdIdCol] ? existingData[i][stdIdCol].toString() : '';
              const existingRecId = existingData[i][recIdCol] ? existingData[i][recIdCol].toString() : '';
              
              if (existingStdId === newStdId && existingRecId === currentRowRecId) {
                Logger.log(`Duplicate Std_ID found: ${newStdId} at row ${rowNum}`);
                return { success: false, message: 'SIRN already exists in your school' };
              }
            }
          }
        }
      } else {
        Logger.log(`SIRN not changed. Original: "${oldStdId}", New: "${newStdId}"`);
      }
      
      // ✅ Check duplicate Unique_ID (global — across all schools) if it's being changed
      const uniqueIdColIdx = headers.indexOf('Unique_ID');
      Logger.log(`Unique_ID check: colIdx=${uniqueIdColIdx}, incoming rowData.Unique_ID="${rowData.Unique_ID}", headers=${JSON.stringify(headers)}`);
      if (uniqueIdColIdx !== -1 && rowData.Unique_ID) {
        const newUniqueId = rowData.Unique_ID.toString();
        const oldUniqueId = rowValues[uniqueIdColIdx] ? rowValues[uniqueIdColIdx].toString() : '';
        Logger.log(`Unique_ID compare: old="${oldUniqueId}" new="${newUniqueId}" changed=${newUniqueId !== oldUniqueId}`);
        if (newUniqueId !== oldUniqueId) {
          const allStudentData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
          for (let i = 0; i < allStudentData.length; i++) {
            const rowNum = i + 2;
            if (rowNum !== rowIndex) {
              const existingUniqueId = allStudentData[i][uniqueIdColIdx] ? allStudentData[i][uniqueIdColIdx].toString() : '';
              if (existingUniqueId === newUniqueId) {
                Logger.log(`Duplicate Unique_ID found: ${newUniqueId} at row ${rowNum}`);
                return { success: false, message: 'Unique_ID already exists' };
              }
            }
          }
        }
      }
      
      // ✅ Format Creation Date
      let creationDate = rowValues[9] || '';
      if (creationDate instanceof Date) {
        creationDate = Utilities.formatDate(creationDate, 'GMT+0500', 'dd-MMM-yyyy');
      }
      
      // ✅ Handle Inactive Date logic
      let inactiveDate = rowValues[8] || '';
      if (rowData.Status === 'Inactive' && rowValues[5] !== 'Inactive') {
        inactiveDate = formattedDate;
      } else if (rowData.Status === 'Active' && rowValues[5] === 'Inactive') {
        inactiveDate = '';
      }
      
      // ✅ BUILD THE UPDATED ROW - Dynamically match sheet columns
      const updatedRow = [];
      
      for (let i = 0; i < headers.length; i++) {
        const header = headers[i];
        
        switch(header) {
          case 'Std_ID':
            updatedRow.push(rowData.Std_ID || originalStdId);
            break;
          case 'Barcode_ID':
            updatedRow.push(rowData.Barcode_ID || rowValues[1] || '');
            break;
          case 'Student_Name':
            updatedRow.push(rowData.Student_Name || rowValues[2] || '');
            break;
          case 'Student_Class':
            updatedRow.push(rowData.Student_Class || rowValues[3] || '');
            break;
          case 'Student_Section':
            updatedRow.push(rowData.Student_Section || rowValues[4] || '');
            break;
          case 'Status':
            updatedRow.push(rowData.Status || rowValues[5] || 'Active');
            break;
          case 'Gender':
            updatedRow.push(rowData.Gender || rowValues[6] || '');
            break;
          case 'Date_of_Joining':
            updatedRow.push(rowData.Date_of_Joining || rowValues[7] || '');
            break;
          case 'Inactive_Date':
            updatedRow.push(inactiveDate);
            break;
          case 'Creation_Date':
            updatedRow.push(creationDate);
            break;
          case 'Created_By':
            updatedRow.push(rowValues[10] || '');
            break;
          case 'Last_Modified_Date':
            updatedRow.push(formattedDate);
            break;
          case 'Last_Modified_By':
            updatedRow.push(currentUser);
            break;
          case 'REC_ID':
            updatedRow.push(rowValues[13] || '');
            break;
          case 'QR_Printed':
            updatedRow.push(rowValues[14] || '');
            break;
          case 'Unique_ID':
            updatedRow.push(rowData.Unique_ID || rowValues[15] || '');
            break;
          default:
            // For any other columns, preserve the original value
            updatedRow.push(rowValues[i] !== undefined && rowValues[i] !== null ? rowValues[i] : '');
            break;
        }
      }
      
      // ✅ Log the column count for debugging
      Logger.log(`Updated row has ${updatedRow.length} columns, sheet has ${headers.length} columns`);
      
      // ✅ Update the row
      let attempts = 3;
      while (attempts > 0) {
        try {
          rowRange.setValues([updatedRow]);
          SpreadsheetApp.flush();
          Logger.log(`Row ${rowIndex} updated successfully for Std_ID: ${updatedRow[0]}`);
          break;
        } catch (e) {
          Logger.log(`Attempt ${4 - attempts} failed: ${e.message}`);
          attempts--;
          if (attempts === 0) {
            return { success: false, message: `Failed to update row after retries: ${e.message}` };
          }
          Utilities.sleep(1000);
        }
      }
      
      // ✅ Return success with all data
return {
  success: true,
  message: 'Student updated successfully: ' + (rowData.Student_Name || rowValues[2] || 'Unknown') + ' (SIRN: ' + (rowData.Std_ID || rowValues[0] || 'Unknown') + ')',
  data: {
    Std_ID: updatedRow[0],
    Barcode_ID: updatedRow[1],
    Student_Name: updatedRow[2],
    Student_Class: updatedRow[3],
    Student_Section: updatedRow[4],
    Status: updatedRow[5],
    Gender: updatedRow[6],
    Date_of_Joining: updatedRow[7],
    Inactive_Date: updatedRow[8],
    Creation_Date: updatedRow[9],
    Created_By: updatedRow[10],
    Last_Modified_Date: updatedRow[11],
    Last_Modified_By: updatedRow[12],
    REC_ID: updatedRow[13] || '',
    QR_Printed: updatedRow[14] || '',
    Unique_ID: updatedRow[15] || '',
    _rowNumber: rowIndex
  }
};
    }
    
if (isPrevPermSheet_(sheetName)) stampPrevPermTimestamps_(rowData);


   

    // ✅ OTHER SHEETS (Users, User_Permissions, etc.)

if (sheetName === 'Users') {
  const usersSheet = ss.getSheetByName('Users');
  const userHeaders = usersSheet.getRange(1, 1, 1, usersSheet.getLastColumn()).getValues()[0];
  const userData = usersSheet.getDataRange().getValues();
  const callerRow = userData.find(r => r[0] === username);
  if (!callerRow) {
    return { success: false, message: 'User not found. Please re-login.' };
  }

  // Header-index lookup for UserType (safer than hardcoded [2])
  const userTypeCol = userHeaders.indexOf('UserType');
  const isSuperadmin = userTypeCol !== -1
    ? String(callerRow[userTypeCol] || '').trim() === 'superadmin'
    : false;

 if (!isSuperadmin) {
  // Only these 4 columns are superadmin-grantable.
  // Everything else (canQueryAttendance, canDownload, canAmendSIRN,
  // canCorrectAttendance, allowedTabs) can be edited by any admin.
  const PROTECTED = [
    'student_delete',
    'canDeleteUsers',
    'UserType',
    'REC_ID',
  ];

  const existingRow = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

  for (const colName of PROTECTED) {
    const idx = userHeaders.indexOf(colName);
    if (idx === -1) continue;

    // Skip fields the client didn't send (type-filtered form hides them)
    if (rowData[colName] === undefined) continue;

    const oldVal = String(existingRow[idx] ?? '');
    const newVal = String(rowData[colName] ?? '');

    if (oldVal !== newVal) {
      return {
        success: false,
        message: 'Only superadmin can change "' + colName + '".'
      };
    }
  }
}
}  

    const newRow = headers.map(function(header) {
      const value = rowData[header];
        if (sheetName.toLowerCase() === 'user_permissions' && !['Username', 'REC_ID'].includes(header)) {
        return value === 'true' || value === true;
      }
      if (sheetName === 'Users' && ['StartTime', 'EndTime'].includes(header)) {
        return value ? value.toString() : '';
      }
      if (sheetName.toLowerCase() === 'users' && ['allowedTabs', 'AllowedDays'].includes(header)) {

  if (Array.isArray(value)) return value.join(',');
  return value ? value.toString() : '';
}

      if (sheetName === 'Users' && ['canQueryAttendance', 'canDownload'].includes(header)) {
        return value === 'Yes' ? 'Yes' : 'No';
      }
      return value !== undefined && value !== null ? value : '';
    });
    
    let attempts = 3;
    while (attempts > 0) {
      try {
        sheet.getRange(rowIndex, 1, 1, headers.length).setValues([newRow]);
        SpreadsheetApp.flush();
        Logger.log(`Row ${rowIndex} updated successfully in ${sheetName}`);
        break;
      } catch (e) {
        Logger.log(`Attempt ${4 - attempts} failed: ${e.message}`);
        attempts--;
        if (attempts === 0) {
          return { success: false, message: `Failed to update row after retries: ${e.message}` };
        }
        Utilities.sleep(1000);
      }
    }
    
    if (sheetName === 'Users') {
      const startTimeIndex = headers.indexOf('StartTime') + 1;
      const endTimeIndex = headers.indexOf('EndTime') + 1;
      if (startTimeIndex > 0) {
        sheet.getRange(rowIndex, startTimeIndex).setNumberFormat('@');
      }
      if (endTimeIndex > 0) {
        sheet.getRange(rowIndex, endTimeIndex).setNumberFormat('@');
      }
    }
    
    return {
      success: true,
      message: 'Record updated successfully',
      data: sanitizeRowForClient_({ ...rowData, _rowNumber: rowIndex })
    };
    
  } catch (e) {
    Logger.log(`Error in updateSheetRow for ${sheetName}: ${e.message}`);
    return {
      success: false,
      message: `Failed to update record: ${e.message}`
    };
  }
}

// Delete a row from a sheet
function deleteSheetRow(sheetName, rowNumber, username) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(sheetName + ' sheet not found');

    // Block deleting the superadmin row itself
    if (sheetName === 'Users') {
      const rowData = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (rowData[0] === 'superadmin') {
        return { success: false, message: 'Superadmin account cannot be deleted' };
      }
    }

    // Fetch caller
    const usersSheet = ss.getSheetByName('Users');
    let userRow = null, userHeaders = null;
    if (usersSheet && username) {
      userHeaders = usersSheet.getRange(1, 1, 1, usersSheet.getLastColumn()).getValues()[0];
      const userData = usersSheet.getDataRange().getValues();
      userRow = userData.find(r => r[0] === username);
    }
    if (!userRow) {
      return { success: false, message: 'User not found. Please re-login.' };
    }

    const isSuperadmin = userRow[2] === 'superadmin';

    // Student delete — student_delete flag OR superadmin
    if (sheetName === 'Student_Data') {
      const col = userHeaders.indexOf('student_delete');
      const flag = col !== -1 ? (userRow[col] || 'No') : 'No';
      if (!isSuperadmin && flag !== 'Yes') {
        return { success: false, message: 'You do not have permission to delete student records' };
      }
    }

    // User delete — canDeleteUsers flag OR superadmin
    if (sheetName === 'Users') {
      const col = userHeaders.indexOf('canDeleteUsers');
      const flag = col !== -1 ? (userRow[col] || 'No') : 'No';
      if (!isSuperadmin && flag !== 'Yes') {
        return { success: false, message: 'You do not have permission to delete user accounts' };
      }
    }

    sheet.deleteRow(rowNumber);
    SpreadsheetApp.flush();
     return { success: true, message: 'Record deleted successfully' };
  } catch (e) {
    Logger.log('Error in deleteSheetRow: ' + e.message);
    return { success: false, message: 'Failed to delete record: ' + e.message };
  }
}

// Get all class sections for teacher assignments
function getClassSections() {
  try {
    var data = getClassSectionData();
    return data.map(row => row.Class_Section);
  } catch (e) {
    Logger.log('Failed to fetch class sections: ' + e.message);
    throw new Error('Failed to fetch class sections: ' + e.message);
  }
}

// Add a student with barcode generation and audit fields
// Add a student with barcode generation and audit fields
function addStudent(studentData, currentUsername, recId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Student_Data');
    if (!sheet) throw new Error('Student_Data sheet not found');
    
    // ✅ VALIDATE Std_ID
    if (!studentData.Std_ID || !/^\d+$/.test(studentData.Std_ID)) {
      return { success: false, message: 'Std_ID must be numeric and required' };
    }
    
    // ✅ VALIDATE Unique_ID
    if (!studentData.Unique_ID || !/^\d{6}$/.test(studentData.Unique_ID)) {
      return { success: false, message: 'Unique_ID must be exactly 6 digits' };
    }
    
    const lastRow = sheet.getLastRow();
    
    // Get headers to find column indices
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const stdIdCol = headers.indexOf('Std_ID') + 1; // Column A
    const recIdCol = headers.indexOf('REC_ID') + 1; // Column N
    
    // ✅ FIX: Check duplicate Std_ID ONLY within the same REC_ID (school)
    if (lastRow > 1) {
      const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
      
      // Check if any row has the same Std_ID AND same REC_ID
      let duplicateFound = false;
      let existingRow = 0;
      
      for (let i = 0; i < dataRange.length; i++) {
        const rowStdId = dataRange[i][stdIdCol - 1] ? dataRange[i][stdIdCol - 1].toString() : '';
        const rowRecId = dataRange[i][recIdCol - 1] ? dataRange[i][recIdCol - 1].toString() : '';
        
        // Same Std_ID AND same REC_ID => duplicate
        if (rowStdId === studentData.Std_ID.toString() && rowRecId === recId) {
          duplicateFound = true;
          existingRow = i + 2; // +2 because dataRange starts at row 2
          break;
        }
      }
      
      if (duplicateFound) {
        return { 
          success: false, 
          message: `SIRN ${studentData.Std_ID} already exists in your school (Row ${existingRow}). Please use a different SIRN.` 
        };
      }
    }
    
    // Check duplicate Unique_ID (global - across all schools)
    const uniqueIdCol = headers.indexOf('Unique_ID') + 1;
    if (uniqueIdCol > 0 && lastRow > 1) {
      const existingUniqueIds = sheet.getRange(2, uniqueIdCol, lastRow - 1, 1).getValues();
      for (let i = 0; i < existingUniqueIds.length; i++) {
        if (existingUniqueIds[i][0] && existingUniqueIds[i][0].toString() === studentData.Unique_ID.toString()) {
          return { success: false, message: 'Unique_ID already exists' };
        }
      }
    }
    
    // Get current user and timestamp
    const currentUser = currentUsername;
    const currentDate = new Date();
    const formattedDate = Utilities.formatDate(currentDate, 'GMT+0500', 'dd-MMM-yyyy');
    
    // ✅ USE PASSED REC_ID - Backend verification
    let finalRecId = recId || '';
    
    // 🔒 BACKEND VERIFICATION: Verify recId matches the user's REC_ID
    let isValidRecId = false;
    if (finalRecId && currentUsername) {
      const usersSheet = ss.getSheetByName('Users');
      if (usersSheet) {
        const userData = usersSheet.getDataRange().getValues();
        for (var i = 0; i < userData.length; i++) {
          if (userData[i][0] === currentUsername) {
            var userRecId = userData[i][15] ? userData[i][15].toString() : '';
            if (userRecId === finalRecId) {
              isValidRecId = true;
              break;
            }
          }
        }
      }
    }
    
    // If no valid REC_ID found, try to look it up (fallback)
    if (!isValidRecId && currentUsername) {
      const usersSheet = ss.getSheetByName('Users');
      if (usersSheet) {
        const userData = usersSheet.getDataRange().getValues();
        for (var i = 0; i < userData.length; i++) {
          if (userData[i][0] === currentUsername) {
            finalRecId = userData[i][15] ? userData[i][15].toString() : '';
            if (finalRecId) {
              isValidRecId = true;
            }
            break;
          }
        }
      }
    }
    
    // If still no valid REC_ID, return error
    if (!isValidRecId || !finalRecId) {
      return { success: false, message: 'Unable to determine school (REC_ID). Please contact administrator.' };
    }
    
    // If we have a REC_ID and it's not in studentData, add it
    if (finalRecId && !studentData.REC_ID) {
      studentData.REC_ID = finalRecId;
    }
    
    // 🎫 Assign Barcode: prefer an already-generated reserve code for this REC, else auto-generate
    let assignedReserveRow = -1;
    const reserveSheetForAssign = ss.getSheetByName('Reserve_Codes');
    if (reserveSheetForAssign && reserveSheetForAssign.getLastRow() > 1) {
      const reserveData = reserveSheetForAssign.getDataRange().getValues();
      const reserveHeaders = reserveData[0];
      const rBarcodeCol = reserveHeaders.indexOf('Barcode_ID');
      const rRecCol = reserveHeaders.indexOf('REC_ID');
      const rStatusCol = reserveHeaders.indexOf('Status');
      for (let i = 1; i < reserveData.length; i++) {
        const rowRecId = reserveData[i][rRecCol] ? reserveData[i][rRecCol].toString() : '';
        const rowStatus = reserveData[i][rStatusCol] || '';
        if (rowRecId === finalRecId && rowStatus === 'Unused') {
          studentData.Barcode_ID = reserveData[i][rBarcodeCol];
          assignedReserveRow = i + 1; // actual sheet row number
          break;
        }
      }
    }
    if (assignedReserveRow === -1) {
      studentData.Barcode_ID = 'BC' + getNextBarcodeNumber(ss);
    }
    
    // Ensure Unique_ID column exists
    if (uniqueIdCol === 0) {
      sheet.getRange(1, headers.length + 1).setValue('Unique_ID');
    }
    
    // NEW ROW with REC_ID
    const newRow = [
      studentData.Std_ID,                // A - 0
      studentData.Barcode_ID,            // B - 1
      studentData.Student_Name || '',    // C - 2
      studentData.Student_Class || '',   // D - 3
      studentData.Student_Section || '', // E - 4
      studentData.Status || 'Active',    // F - 5
      studentData.Gender || '',          // G - 6
      studentData.Date_of_Joining || '', // H - 7
      '',                                 // I - 8 Inactive Date
      formattedDate,                     // J - 9 Creation Date
      currentUser,                       // K - 10 Created By
      formattedDate,                     // L - 11 Last Modified Date
      currentUser,                       // M - 12 Last Modified By
      finalRecId,                        // N - 13 REC_ID (passed value)
      '',                                 // O - 14 QR_Printed
      studentData.Unique_ID              // P - 15 Unique_ID
    ];
    
    sheet.getRange(lastRow + 1, 1, 1, 16).setValues([newRow]);
    SpreadsheetApp.flush();
    
    // 🎫 Mark the reserve code as used, now that the student row is saved
    if (assignedReserveRow !== -1) {
      const reserveHeaders2 = reserveSheetForAssign.getRange(1, 1, 1, reserveSheetForAssign.getLastColumn()).getValues()[0];
      const statusCol2 = reserveHeaders2.indexOf('Status') + 1;
      const usedByCol2 = reserveHeaders2.indexOf('Used_By_Std_ID') + 1;
      const usedDateCol2 = reserveHeaders2.indexOf('Used_Date') + 1;
      if (statusCol2 > 0) reserveSheetForAssign.getRange(assignedReserveRow, statusCol2).setValue('Used');
      if (usedByCol2 > 0) reserveSheetForAssign.getRange(assignedReserveRow, usedByCol2).setValue(studentData.Std_ID);
      if (usedDateCol2 > 0) reserveSheetForAssign.getRange(assignedReserveRow, usedDateCol2).setValue(formattedDate);
      SpreadsheetApp.flush();
    }
    
    return {
      success: true,
      message: 'Student added successfully: ' + studentData.Student_Name + ' (SIRN: ' + studentData.Std_ID + ', Barcode: ' + studentData.Barcode_ID + ')',
      data: { 
        ...studentData, 
        Creation_Date: formattedDate,
        Created_By: currentUser,
        Last_Modified_Date: formattedDate,
        Last_Modified_By: currentUser,
        _rowNumber: lastRow + 1 
      }
    };
  } catch (e) {
    Logger.log('Error in addStudent: ' + e.message);
    return { success: false, message: 'Error adding student: ' + e.message };
  }
}

// Get student data
function getStudentData(recId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Student_Data');
    if (!sheet) throw new Error('Student_Data sheet not found');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const result = [];
    
    // Find REC_ID column index
    const recIdCol = headers.indexOf('REC_ID');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        // Filter by REC_ID
        if (recIdCol !== -1 && recId) {
          const rowRecId = data[i][recIdCol] ? data[i][recIdCol].toString() : '';
          if (rowRecId !== recId) {
            continue; // Skip - different school
          }
        }
        
        let row = {};
        for (let j = 0; j < headers.length; j++) {
          if (headers[j] === 'Creation_Date' && data[i][j]) {
            try {
              const dateObj = new Date(data[i][j]);
              row[headers[j]] = Utilities.formatDate(dateObj, 'GMT+0500', 'dd-MMM-yyyy');
            } catch (e) {
              row[headers[j]] = data[i][j] !== null && data[i][j] !== undefined ? data[i][j].toString() : '';
            }
          } else if (['Date_of_Joining', 'Last_Modified_Date', 'Inactive_Date'].includes(headers[j]) && data[i][j] instanceof Date) {
            row[headers[j]] = Utilities.formatDate(data[i][j], 'GMT+0500', 'dd-MMM-yyyy');
          } else {
            row[headers[j]] = data[i][j] !== null && data[i][j] !== undefined ? data[i][j].toString() : '';
          }
        }
        row._rowNumber = i + 1;
        result.push(row);
      }
    }
    return result;
  } catch (e) {
    Logger.log('Error in getStudentData: ' + e.message);
    throw new Error('Failed to fetch student data: ' + e.message);
  }
}

// Get gender options for dropdown
function getGenderOptions() {
  return ['Male', 'Female'];
}

// Get student attendance based on search criteria
function getStudentAttendance(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Permission check
    const usersSheet = ss.getSheetByName('Users');
    if (!usersSheet) throw new Error('Users sheet not found');
    
    const userData = usersSheet.getDataRange().getValues();
    const userRow = userData.find(row => row[0] === formData.username);
    if (!userRow || userRow[9] !== 'Yes') {
      throw new Error('You do not have permission to query attendance');
    }
    
    const sheet = ss.getSheetByName('Student_Attendance');
    if (!sheet) throw new Error('Student_Attendance sheet not found');
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return []; // No data
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    
    // Pre-process search criteria
    const searchValue = formData.searchValue ? formData.searchValue.toString().toLowerCase() : '';
    const searchType = formData.searchType;
    const statusFilter = formData.statusFilter;
    
    // Date range setup
    let startDate, endDate;
    if (formData.startDate && formData.endDate) {
      startDate = new Date(formData.startDate);
      endDate = new Date(formData.endDate);
      startDate.setHours(0, 0, 0, 0);
      endDate.setHours(23, 59, 59, 999);
    }
    
    // Cache column indices
    const columnIndices = {
      date: headers.indexOf('Date'),
      stdId: headers.indexOf('Std_ID'),
      barcodeId: headers.indexOf('Barcode_ID'),
      studentName: headers.indexOf('Student_Name'),
      status: headers.indexOf('Status')
    };
    
    return data
      .filter(row => {
        if (!row[columnIndices.date]) return false; // Skip empty rows
        
        // Date filtering
        if (startDate && endDate) {
          const attendanceDate = new Date(row[columnIndices.date]);
          if (isNaN(attendanceDate)) return false;
          attendanceDate.setHours(0, 0, 0, 0);
          if (attendanceDate < startDate || attendanceDate > endDate) return false;
        }
        
        // Search value filtering
        if (searchValue) {
          let matches = false;
          switch (searchType) {
            case 'Std_ID':
              matches = row[columnIndices.stdId].toString().toLowerCase() === searchValue;
              break;
            case 'Barcode_ID':
              matches = row[columnIndices.barcodeId].toString().toLowerCase() === searchValue;
              break;
            case 'Student_Name':
              matches = row[columnIndices.studentName].toString().toLowerCase().includes(searchValue);
              break;
          }
          if (!matches) return false;
        }
        
        // Status filtering
        if (statusFilter !== 'All' && row[columnIndices.status].toString() !== statusFilter) {
          return false;
        }
        
        return true;
      })
      .map((row, index) => {
        const resultRow = {};
        headers.forEach((header, j) => {
          resultRow[header] = header === 'Date' && row[j] instanceof Date 
            ? Utilities.formatDate(row[j], 'GMT+0500', 'yyyy-MMM-dd')
            : (row[j] !== null && row[j] !== undefined ? row[j].toString() : '');
        });
        resultRow._rowNumber = index + 2;
        return resultRow;
      });
  } catch (e) {
    Logger.log('Error in getStudentAttendance: ' + e.message);
    throw new Error('Failed to fetch attendance data: ' + e.message);
  }
}

// Download attendance data in CSV or PDF format
function downloadAttendance(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const usersSheet = ss.getSheetByName('Users');
    if (!usersSheet) throw new Error('Users sheet not found');
    
    const userData = usersSheet.getDataRange().getValues();
    const userRow = userData.find(row => row[0] === formData.username);
    if (!userRow || userRow[10] !== 'Yes') {
      throw new Error('You do not have permission to download attendance reports');
    }
    
    const sheet = ss.getSheetByName('Student_Attendance');
    if (!sheet) throw new Error('Student_Attendance sheet not found');
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const requiredColumns = ['Attendance_ID', 'Date', 'Std_ID', 'Barcode_ID', 'Student_Name', 'Student_Class', 'Student_Section', 'Status', 'Timestamp', 'Teacher_ID', 'Class_Section'];
    const columnIndices = requiredColumns.map(col => headers.indexOf(col) + 1).filter(idx => idx > 0);
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    let filteredData = [headers];
    
    // Pre-parse search criteria
    const searchValue = formData.searchValue ? formData.searchValue.toString().toLowerCase() : '';
    const searchType = formData.searchType;
    const statusFilter = formData.statusFilter;
    const dateIndex = headers.indexOf('Date');

    // Normalize startDate and endDate once
    let startDate, endDate;
    if (formData.startDate && formData.endDate) {
      startDate = new Date(formData.startDate);
      endDate = new Date(formData.endDate);
      startDate.setHours(0, 0, 0, 0); // Start of day
      endDate.setHours(23, 59, 59, 999); // End of day
    }

    for (let i = 0; i < data.length; i++) {
      if (!data[i][0]) continue;
      
      let matches = true;

      // Date range filtering first
      if (startDate && endDate) {
        const attendanceDate = new Date(data[i][dateIndex]);
        if (isNaN(attendanceDate)) {
          Logger.log(`Invalid date at row ${i + 2}: ${data[i][dateIndex]}`);
          continue;
        }
        attendanceDate.setHours(0, 0, 0, 0); // Normalize to start of day
        if (attendanceDate < startDate || attendanceDate > endDate) {
          matches = false;
          continue;
        }
      }

      // Search value filtering
      if (searchValue) {
        if (searchType === 'Std_ID' && data[i][2].toString().toLowerCase() !== searchValue) {
          matches = false;
        } else if (searchType === 'Barcode_ID' && data[i][3].toString().toLowerCase() !== searchValue) {
          matches = false;
        } else if (searchType === 'Student_Name' && !data[i][4].toString().toLowerCase().includes(searchValue)) {
          matches = false;
        }
      }

      // Status filtering
      if (statusFilter !== 'All' && data[i][7].toString() !== statusFilter) {
        matches = false;
      }

      if (matches) {
        let row = data[i].slice();
        if (dateIndex !== -1 && row[dateIndex] instanceof Date) {
          row[dateIndex] = Utilities.formatDate(row[dateIndex], 'GMT+0500', 'yyyy-MMM-dd');
        }
        filteredData.push(row);
      }
    }
    
    if (formData.format === 'CSV') {
      let csv = headers.join(',') + '\n';
      for (let i = 1; i < filteredData.length; i++) {
        csv += filteredData[i].map(val => `"${val !== null && val !== undefined ? val.toString().replace(/"/g, '""') : ''}"`).join(',') + '\n';
      }
      return { success: true, data: csv };
    } else if (formData.format === 'PDF') {
      const doc = DocumentApp.create('Attendance_Report_' + new Date().toISOString());
      const body = doc.getBody();
      body.appendParagraph('Attendance Report').setHeading(DocumentApp.ParagraphHeading.HEADING1);
      const table = body.appendTable(filteredData);
      table.setBorderWidth(1);
      table.getRow(0).setAttributes({ BOLD: true });
      doc.saveAndClose();
      const pdf = doc.getAs('application/pdf');
      const pdfBase64 = Utilities.base64Encode(pdf.getBytes());
      DriveApp.getFileById(doc.getId()).setTrashed(true);
      return { success: true, data: pdfBase64 };
    }
    return { success: false, message: 'Invalid format specified' };
  } catch (e) {
    Logger.log('Error in downloadAttendance: ' + e.message);
    return { success: false, message: 'Failed to download attendance: ' + e.message };
  }
}

// Generate QR code data URL for a student
function generateStudentQRCode(stdId, barcodeId, studentName) {
  try {
    // Create QR code data with student information
    const qrData = {
      type: 'student',
      stdId: stdId,
      barcodeId: barcodeId,
      name: studentName,
      timestamp: new Date().toISOString()
    };
    
    // Convert to JSON string for QR code
    const qrText = JSON.stringify(qrData);
    
    // Use Google Charts API to generate QR code
    const qrCodeUrl = `https://chart.googleapis.com/chart?chs=150x150&cht=qr&chl=${encodeURIComponent(qrText)}&choe=UTF-8`;
    
    return qrCodeUrl;
  } catch (e) {
    Logger.log('Error generating QR code: ' + e.message);
    throw e;
  }
}

// Get unprinted students for QR code generation
function getUnprintedStudents(recId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Student_Data');
    if (!sheet) throw new Error('Student_Data sheet not found');
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Check if QR_Printed column exists, if not create it
    if (!headers.includes('QR_Printed')) {
      const lastCol = sheet.getLastColumn();
      sheet.getRange(1, lastCol + 1).setValue('QR_Printed');
      headers.push('QR_Printed');
    }
    
    const stdIdCol = headers.indexOf('Std_ID');
    const barcodeCol = headers.indexOf('Barcode_ID');
    const nameCol = headers.indexOf('Student_Name');
    const classCol = headers.indexOf('Student_Class');
    const sectionCol = headers.indexOf('Student_Section');
    const qrPrintedCol = headers.indexOf('QR_Printed');
    const recIdCol = headers.indexOf('REC_ID');  // ← ADD THIS
    
    const unprintedStudents = [];
    
    for (let i = 1; i < data.length; i++) {
      const qrPrinted = data[i][qrPrintedCol] || '';
      
      // 🔒 FILTER BY REC_ID
      if (recIdCol !== -1 && recId) {
        const rowRecId = data[i][recIdCol] ? data[i][recIdCol].toString() : '';
        if (rowRecId !== recId) {
          continue; // Skip - different school
        }
      }
      
      // Only include students whose QR hasn't been printed
      if (qrPrinted.toString().toLowerCase() !== 'yes') {
        unprintedStudents.push({
          Std_ID: data[i][stdIdCol],
          Barcode_ID: data[i][barcodeCol],
          Student_Name: data[i][nameCol],
          Student_Class: data[i][classCol],
          Student_Section: data[i][sectionCol],
          _rowNumber: i + 1
        });
      }
    }
    
    return unprintedStudents;
    
  } catch (e) {
    Logger.log('Error getting unprinted students: ' + e.message);
    throw e;
  }
}

// Mark students as printed
function markStudentsAsPrinted(studentIds) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Student_Data');
    if (!sheet) throw new Error('Student_Data sheet not found');
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let qrPrintedCol = headers.indexOf('QR_Printed');
    
    // If column doesn't exist, create it
    if (qrPrintedCol === -1) {
      qrPrintedCol = headers.length;
      sheet.getRange(1, qrPrintedCol + 1).setValue('QR_Printed');
    }
    
    const data = sheet.getDataRange().getValues();
    const stdIdCol = headers.indexOf('Std_ID');
    
    let markedCount = 0;
    
    studentIds.forEach(stdId => {
      for (let i = 1; i < data.length; i++) {
        if (data[i][stdIdCol].toString() === stdId.toString()) {
          sheet.getRange(i + 1, qrPrintedCol + 1).setValue('Yes');
          markedCount++;
          break;
        }
      }
    });
    
    SpreadsheetApp.flush();
    return { success: true, count: markedCount };
    
  } catch (e) {
    Logger.log('Error marking students as printed: ' + e.message);
    return { success: false, message: e.message };
  }
}

// Get or create the Reserve_Codes sheet with headers
function getOrCreateReserveCodesSheet(ss) {
  let sheet = ss.getSheetByName('Reserve_Codes');
  if (!sheet) {
    sheet = ss.insertSheet('Reserve_Codes');
    sheet.appendRow(['Barcode_ID', 'REC_ID', 'Status', 'Generated_By', 'Generated_Date', 'Used_By_Std_ID', 'Used_Date', 'Printed']);
  }
  return sheet;
}

// Compute the next available BC number, checking BOTH Student_Data and Reserve_Codes
// so reserved-but-unused codes never get re-issued to a different purpose
function getNextBarcodeNumber(ss) {
  let maxNumber = 1000000; // Start from 1000001

  const studentSheet = ss.getSheetByName('Student_Data');
  if (studentSheet) {
    const lastRow = studentSheet.getLastRow();
    if (lastRow > 1) {
      const barcodes = studentSheet.getRange('B2:B' + lastRow).getValues();
      barcodes.forEach(row => {
        if (row[0] && typeof row[0] === 'string' && row[0].startsWith('BC')) {
          const num = parseInt(row[0].replace('BC', ''));
          if (!isNaN(num) && num > maxNumber) maxNumber = num;
        }
      });
    }
  }

  const reserveSheet = getOrCreateReserveCodesSheet(ss);
  const reserveLastRow = reserveSheet.getLastRow();
  if (reserveLastRow > 1) {
    const barcodes = reserveSheet.getRange('A2:A' + reserveLastRow).getValues();
    barcodes.forEach(row => {
      if (row[0] && typeof row[0] === 'string' && row[0].startsWith('BC')) {
        const num = parseInt(row[0].replace('BC', ''));
        if (!isNaN(num) && num > maxNumber) maxNumber = num;
      }
    });
  }

  return maxNumber + 1;
}

// Generate reserve codes (bulk barcode generation) and SAVE them to Reserve_Codes
function generateReserveCodes(count, recId, username) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const reserveSheet = getOrCreateReserveCodesSheet(ss);

    const nextNumber = getNextBarcodeNumber(ss);
    const formattedDate = Utilities.formatDate(new Date(), 'GMT+0500', 'dd-MMM-yyyy');
    const reserveCodes = [];
    const rowsToAppend = [];

    for (let i = 0; i < count; i++) {
      const barcode = 'BC' + (nextNumber + i);
      reserveCodes.push(barcode);
      rowsToAppend.push([barcode, recId || '', 'Unused', username || '', formattedDate, '', '', 'No']);
    }

    reserveSheet.getRange(reserveSheet.getLastRow() + 1, 1, rowsToAppend.length, 8).setValues(rowsToAppend);
    SpreadsheetApp.flush();

    return reserveCodes;

  } catch (e) {
    Logger.log('Error generating reserve codes: ' + e.message);
    throw e;
  }
}

// Generate PDF with multiple QR codes
function generateQRCodePDF(selectedStudents) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Create HTML template for QR code printing
    const html = HtmlService.createHtmlOutputFromFile('QRPrintTemplate');
    html.setTitle('Student QR Codes');
    
    // Generate QR codes for selected students
    const studentsWithQR = selectedStudents.map(student => ({
      ...student,
      qrCodeUrl: generateStudentQRCode(student.Std_ID, student.Barcode_ID, student.Student_Name)
    }));
    
    // Pass data to template
    const output = HtmlService.createTemplate(html.getContent());
    output.students = studentsWithQR;
    output.pageTitle = 'Student QR Codes';
    
    return output.evaluate().getContent();
    
  } catch (e) {
    Logger.log('Error generating QR PDF: ' + e.message);
    throw e;
  }
}


// Generate PDF with blank QR cards for reserve barcodes (no student info assigned yet)
function generateBlankQRCodePDF(selectedCodes) {
  try {
    const html = HtmlService.createHtmlOutputFromFile('QRPrintTemplate');
    html.setTitle('Reserve Code QR Cards');

    // Only Barcode_ID is known — everything else stays blank on the card
    const blankCards = selectedCodes.map(code => ({
      Std_ID: '',
      Barcode_ID: code,
      Student_Name: '',
      Student_Class: '',
      Student_Section: ''
    }));

    const output = HtmlService.createTemplate(html.getContent());
    output.students = blankCards;
    output.pageTitle = 'Reserve Code QR Cards';

    return output.evaluate().getContent();

  } catch (e) {
    Logger.log('Error generating blank QR PDF: ' + e.message);
    throw e;
  }
}

// Mark reserve codes as printed. Always sets 'Yes' regardless of current value,
// so calling this again (reprint) is a safe no-op on the sheet.
function markReserveCodesAsPrinted(barcodeIds) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateReserveCodesSheet(ss);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const barcodeCol = headers.indexOf('Barcode_ID');
    const printedCol = headers.indexOf('Printed');

    const data = sheet.getDataRange().getValues();
    let markedCount = 0;

    barcodeIds.forEach(code => {
      for (let i = 1; i < data.length; i++) {
        if (data[i][barcodeCol] && data[i][barcodeCol].toString() === code.toString()) {
          sheet.getRange(i + 1, printedCol + 1).setValue('Yes');
          markedCount++;
          break;
        }
      }
    });

    SpreadsheetApp.flush();
    return { success: true, count: markedCount };

  } catch (e) {
    Logger.log('Error marking reserve codes as printed: ' + e.message);
    return { success: false, message: e.message };
  }
}


// Get or create the Attendance_Correction_Log sheet with headers
function getOrCreateCorrectionLogSheet(ss) {
  let sheet = ss.getSheetByName('Attendance_Correction_Log');
  const requiredHeaders = [
    'Attendance_ID', 'Std_ID', 'Student_Name', 'Date', 'Class_Section',
    'Old_Status', 'New_Status', 'Corrected_By', 'REC_ID', 'Timestamp',
    'Reason', 'Comment', 'Old_Hours', 'New_Hours',
    'Old_Category', 'New_Category'  // 🆕 Added
  ];
  
  if (!sheet) {
    // Create new sheet with all headers
    sheet = ss.insertSheet('Attendance_Correction_Log');
    sheet.appendRow(requiredHeaders);
  } else {
    // Sheet exists — check if new columns are missing and append them
    const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const missingHeaders = requiredHeaders.filter(h => !existingHeaders.includes(h));
    
    if (missingHeaders.length > 0) {
      const startCol = existingHeaders.length + 1;
      sheet.getRange(1, startCol, 1, missingHeaders.length).setValues([missingHeaders]);
      // Optional: make the new headers bold to match the rest
      sheet.getRange(1, startCol, 1, missingHeaders.length).setFontWeight('bold');
    }
  }
  return sheet;
}

// Correct a single attendance record's Status, with a full audit trail
function correctAttendanceRecord(attendanceId, newStatus, username, reason, comment) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 🔒 Permission check
    const usersSheet = ss.getSheetByName('Users');
    if (!usersSheet) throw new Error('Users sheet not found');
    const userData = usersSheet.getDataRange().getValues();
    const userHeaders = userData[0];
    const canCorrectCol = userHeaders.indexOf('canCorrectAttendance');
    const userRow = userData.find(row => row[0] === username);
    if (canCorrectCol === -1 || !userRow || userRow[canCorrectCol] !== 'Yes') {
      return { success: false, message: 'You do not have permission to correct attendance records' };
    }

    if (!['Present', 'Absent', 'Leave'].includes(newStatus)) {
      return { success: false, message: 'Invalid status value' };
    }

    // 🆕 Validate reason
    if (!reason || String(reason).trim() === '') {
      return { success: false, message: 'A reason is required for correction' };
    }

    // Find the exact row by Attendance_ID
    const sheet = ss.getSheetByName('Student_Attendance');
    if (!sheet) throw new Error('Student_Attendance sheet not found');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const idCol = headers.indexOf('Attendance_ID');
    const statusCol = headers.indexOf('Status');
    if (idCol === -1 || statusCol === -1) throw new Error('Required columns not found in Student_Attendance');

    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] && data[i][idCol].toString() === attendanceId.toString()) {
        rowIndex = i + 1;
        break;
      }
    }
    if (rowIndex === -1) {
      return { success: false, message: 'Attendance record not found' };
    }

    const rowValues = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    const oldStatus = rowValues[statusCol] || '';

    if (oldStatus === newStatus) {
      return { success: false, message: 'New status is the same as the current status' };
    }

    // Update the status
    sheet.getRange(rowIndex, statusCol + 1).setValue(newStatus);
    SpreadsheetApp.flush();

    // 📝 Audit log
    const dateCol = headers.indexOf('Date');
    const stdIdCol = headers.indexOf('Std_ID');
    const nameCol = headers.indexOf('Student_Name');
    const classSectionCol = headers.indexOf('Class_Section');
    const recIdCol = headers.indexOf('REC_ID');

    const logSheet = getOrCreateCorrectionLogSheet(ss);
    const timestamp = Utilities.formatDate(new Date(), 'GMT+0500', 'dd-MMM-yyyy HH:mm');
    const logDate = (dateCol !== -1 && rowValues[dateCol] instanceof Date)
      ? Utilities.formatDate(rowValues[dateCol], 'GMT+0500', 'yyyy-MMM-dd')
      : (dateCol !== -1 ? rowValues[dateCol] : '');

    // 🆕 Append with Reason and Comment
    logSheet.appendRow([
      attendanceId,
      stdIdCol !== -1 ? rowValues[stdIdCol] : '',
      nameCol !== -1 ? rowValues[nameCol] : '',
      logDate,
      classSectionCol !== -1 ? rowValues[classSectionCol] : '',
      oldStatus,
      newStatus,
      username,
      recIdCol !== -1 ? rowValues[recIdCol] : '',
      timestamp,
      reason,                                 // 🆕
      comment ? String(comment).trim() : ''   // 🆕
    ]);
    SpreadsheetApp.flush();

    return {
      success: true,
      message: `Status corrected: ${oldStatus} → ${newStatus} (Reason: ${reason})`,
      data: { Attendance_ID: attendanceId, Status: newStatus, _rowNumber: rowIndex }
    };

  } catch (e) {
    Logger.log('Error in correctAttendanceRecord: ' + e.message);
    return { success: false, message: 'Failed to correct attendance record: ' + e.message };
  }
}



/**
 * Correct ONLY the Hours of a single attendance row.
 * - Category is NEVER touched (it belongs to the class-section for the day).
 * - New hours must be > 0, a 0.25 multiple, and <= the day's total class hours.
 * - Row must be Present.
 * - Logs Old_Hours / New_Hours (plus reason/comment) to Attendance_Correction_Log.
 */
function correctAttendanceHours(attendanceId, newHours, newCategory, username, reason, comment) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return { success: false, message: 'System busy. Please try again.' };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Permission
    const usersSheet = ss.getSheetByName('Users');
    if (!usersSheet) throw new Error('Users sheet not found');
    const userData = usersSheet.getDataRange().getValues();
    const userHeaders = userData[0];
    const canCorrectCol = userHeaders.indexOf('canCorrectAttendance');
    const userRow = userData.find(row => row[0] === username);
    if (canCorrectCol === -1 || !userRow || userRow[canCorrectCol] !== 'Yes') {
      return { success: false, message: 'You do not have permission to correct attendance records' };
    }
    if (!reason || String(reason).trim() === '') {
      return { success: false, message: 'A reason is required for correction' };
    }

    // 2. Locate the attendance row
    const sheet = ss.getSheetByName('Student_Attendance');
    if (!sheet) throw new Error('Student_Attendance sheet not found');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const idCol     = headers.indexOf('Attendance_ID');
    const statusCol = headers.indexOf('Status');
    const hoursCol  = headers.indexOf('Hours');
    const catCol    = headers.indexOf('Category');
    const dateCol   = headers.indexOf('Date');
    const classCol  = headers.indexOf('Class_Section');
    if (idCol === -1 || hoursCol === -1 || statusCol === -1) {
      throw new Error('Required columns not found in Student_Attendance');
    }

    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] && data[i][idCol].toString() === attendanceId.toString()) {
        rowIndex = i + 1;
        break;
      }
    }
    if (rowIndex === -1) return { success: false, message: 'Attendance record not found' };

    const rowValues = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    const currentStatus = rowValues[statusCol];
    const oldHours      = rowValues[hoursCol];
    const oldCategory   = catCol !== -1 ? rowValues[catCol] : '';
    const rowDate       = dateCol   !== -1 ? rowValues[dateCol]   : '';
    const rowClassSec   = classCol  !== -1 ? rowValues[classCol]  : '';

    // 3. Only Present rows can carry editable hours
    if (String(currentStatus).toLowerCase().trim() !== 'present') {
      return {
        success: false,
        message: `Hours can only be corrected for Present rows. This row is "${currentStatus}".`
      };
    }

    // 4. Fetch the day's total hours from Daily_Hours_Setup
    const classHoursCheck = checkDailyHours(rowDate, rowClassSec, 'normal');
    if (!classHoursCheck.exists) {
      return {
        success: false,
        message: 'No class hours configured for this date and section. Cannot validate the correction.'
      };
    }
    const dayTotal = Number(classHoursCheck.hours) || 0;
    if (dayTotal <= 0) {
      return { success: false, message: 'Invalid configured hours for this day.' };
    }

    // 5. Validate the new hours
    const numHours = Number(newHours);
    if (isNaN(numHours) || numHours <= 0) {
      return { success: false, message: 'Hours must be a positive number.' };
    }
    if (Math.abs(Math.round(numHours / 0.25) * 0.25 - numHours) > 0.001) {
      return { success: false, message: 'Hours must be in multiples of 0.25.' };
    }
    if (numHours > dayTotal + 0.0001) {
      return {
        success: false,
        message: `Hours cannot exceed the class's configured total for the day (${dayTotal}).`
      };
    }

    // 6. Validate Category Names and Split
    if (catCol !== -1) {
      const daySetup = checkDailyHours(rowDate, rowClassSec, 'normal');
      const dayItems = daySetup.exists ? parseCategoryBreakdown_(daySetup.category) : null;
      const newItems = parseCategoryBreakdown_(newCategory);

      if (dayItems) {
        // Multi-category day validation
        if (!newItems) {
          return { success: false, message: `Category must contain all day categories: ${dayItems.map(d => d.name).join(', ')}.` };
        }
        // 6a. New items must use existing day category names
        for (const n of newItems) {
          if (!dayItems.some(d => d.name === n.name)) {
            return { success: false, message: `Category "${n.name}" is not part of this day's setup. Allowed: ${dayItems.map(d => d.name).join(', ')}.` };
          }
        }
        // 6b. Every day category must be present
        for (const d of dayItems) {
          if (!newItems.some(n => n.name === d.name)) {
            return { success: false, message: `Missing category "${d.name}". All day categories must be present.` };
          }
        }
        // 6c. Hours must sum up to total
        if (Math.abs(sumBreakdown_(newItems) - numHours) > 0.001) {
          return { success: false, message: `Category hours must add up to the total hours (${numHours}).` };
        }
      } else {
        // Single-category day validation
        const dayCat = daySetup.exists ? String(daySetup.category || '').trim() : '';
        if (dayCat && String(newCategory || '').trim() !== dayCat) {
          return { success: false, message: `Category cannot be changed from "${dayCat}".` };
        }
      }
    }

    if (String(oldHours) === String(numHours) && String(oldCategory) === String(newCategory)) {
      return { success: false, message: 'No change — values equal the current values.' };
    }

    // 7. Apply — Update both Hours and Category
    sheet.getRange(rowIndex, hoursCol + 1).setValue(numHours);
    if (catCol !== -1) {
      sheet.getRange(rowIndex, catCol + 1).setValue(newCategory);
    }
    SpreadsheetApp.flush();

    // 8. Audit log
    const nameCol  = headers.indexOf('Student_Name');
    const stdIdCol = headers.indexOf('Std_ID');
    const recIdCol = headers.indexOf('REC_ID');
    const logSheet = getOrCreateCorrectionLogSheet(ss);
    const timestamp = Utilities.formatDate(new Date(), 'GMT+0500', 'dd-MMM-yyyy HH:mm');
    const logDate = (rowDate instanceof Date)
      ? Utilities.formatDate(rowDate, 'GMT+0500', 'yyyy-MMM-dd')
      : (rowDate || '');

    logSheet.appendRow([
      attendanceId,
      stdIdCol !== -1 ? rowValues[stdIdCol] : '',
      nameCol  !== -1 ? rowValues[nameCol]  : '',
      logDate,
      rowClassSec,
      currentStatus,           // Old_Status
      currentStatus,           // New_Status (unchanged)
      username,
      recIdCol !== -1 ? rowValues[recIdCol] : '',
      timestamp,
      reason,
      comment ? String(comment).trim() : '',
      oldHours,                // Old_Hours
      numHours,                // New_Hours
      oldCategory,             // Old_Category
      newCategory              // New_Category
    ]);
    SpreadsheetApp.flush();

    return {
      success: true,
      message: `Hours corrected: ${oldHours || '—'} → ${numHours}`,
      data: { Attendance_ID: attendanceId, Hours: numHours, Category: newCategory }
    };

  } catch (e) {
    Logger.log('Error in correctAttendanceHours: ' + e.message);
    return { success: false, message: 'Failed to correct hours: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}


// ---------------------------------------------------------------------------
// Daily_Hours_Setup helpers (were missing from Admin project)
// ---------------------------------------------------------------------------

/**
 * Look up the hours/category setup for a given date + class-section.
 * type: 'normal' | 'episodic'  (episodic returns Events if present)
 * Returns { exists, hours, category, events }
 */
function checkDailyHours(date, classSection, type) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Daily_Hours_Setup');
    if (!sheet) return { exists: false, hours: 0, category: '', events: '' };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { exists: false, hours: 0, category: '', events: '' };

    const headers = data[0];
    const dateCol      = headers.indexOf('Date');
    const classSecCol  = headers.indexOf('Class_Section');
    const hoursCol     = headers.indexOf('Hours');
    const categoryCol  = headers.indexOf('Category');
    const eventsCol    = headers.indexOf('Events');
    const typeCol      = headers.indexOf('Type');

    if (dateCol === -1 || classSecCol === -1) {
      return { exists: false, hours: 0, category: '', events: '' };
    }

    // Normalise incoming date to yyyy-MM-dd for comparison
    const target = Utilities.formatDate(new Date(date), 'GMT+0500', 'yyyy-MM-dd');
    const wantedType = String(type || 'normal').toLowerCase();

    for (let i = 1; i < data.length; i++) {
      const rowDate = data[i][dateCol];
      if (!rowDate) continue;

      const rowDateStr = (rowDate instanceof Date)
        ? Utilities.formatDate(rowDate, 'GMT+0500', 'yyyy-MM-dd')
        : String(rowDate).trim().slice(0, 10);

      if (rowDateStr !== target) continue;
      if (String(data[i][classSecCol] || '').trim() !== String(classSection || '').trim()) continue;

      // If the sheet has a Type column, respect it (episodic vs normal)
      if (typeCol !== -1) {
        const rowType = String(data[i][typeCol] || 'normal').toLowerCase();
        if (rowType !== wantedType) continue;
      }

      return {
        exists: true,
        hours: hoursCol !== -1 ? Number(data[i][hoursCol] || 0) : 0,
        category: categoryCol !== -1 ? String(data[i][categoryCol] || '') : '',
        events: eventsCol !== -1 ? String(data[i][eventsCol] || '') : ''
      };
    }

    return { exists: false, hours: 0, category: '', events: '' };
  } catch (e) {
    Logger.log('checkDailyHours error: ' + e.message);
    return { exists: false, hours: 0, category: '', events: '' };
  }
}

/**
 * Parse "Name (hours) + Name (hours) + ..." into [{name, hours}, ...].
 * Returns null for plain single-category strings (no " + " and no "(h)").
 */
function parseCategoryBreakdown_(categoryStr) {
  const s = String(categoryStr || '').trim();
  if (!s) return null;
  if (s.indexOf(' + ') === -1) return null;

  const items = [];
  for (const part of s.split(' + ')) {
    const m = part.trim().match(/^(.+) \((\d+(?:\.\d+)?)\)$/);
    if (m) items.push({ name: m[1].trim(), hours: Number(m[2]) });
  }
  return items.length ? items : null;
}

/**
 * Sum hours from an array returned by parseCategoryBreakdown_.
 */
function sumBreakdown_(items) {
  if (!Array.isArray(items)) return 0;
  return items.reduce((sum, i) => sum + (Number(i.hours) || 0), 0);
}



// ============ CLASS MAPPINGS — BULK ADD ============
const MASTER_CLASSES = [
  'IBT', 'NUR', 'KGI', 'KGII',
  'I', 'II', 'III', 'IV', 'V', 'VI',
  'VII', 'VIII', 'IX', 'X', 'XI', 'HRE'
];

/**
 * Add a Class + all its Sections in ONE call.
 * Writes to Class_Sections sheet:
 *   - Column C: Class
 *   - Column D: Sections (comma-separated)
 *   - Column A: one row per Class-Section combo (I-A, I-B, ...)
 */
function addClassMappingWithSections(className, sectionsStr, recId) {
  try {
    ensureClassMappingSheet_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mapSheet = ss.getSheetByName('Class_Mapping');
    const secSheet = ss.getSheetByName('Class_Sections');
    if (!mapSheet) throw new Error('Class_Mapping sheet not found');
    if (!secSheet) throw new Error('Class_Sections sheet not found');

    const cls = String(className || '').trim();
    if (!MASTER_CLASSES.includes(cls)) {
      throw new Error(`"${cls}" is not a valid class. Choose from: ${MASTER_CLASSES.join(', ')}`);
    }

    const sections = String(sectionsStr || '')
      .split(',').map(s => s.trim().toUpperCase()).filter(Boolean);

    if (!sections.length) throw new Error('At least one section is required');
    if (!sections.every(s => /^[A-Z]$/.test(s))) {
      throw new Error('Sections must be single letters (e.g., A,B,C)');
    }

    const finalRecId = String(recId || '').trim();
    if (!finalRecId) throw new Error('REC_ID is required to add a class mapping');

    // 1. Duplicate check scoped to (Class, REC_ID)
    const mapData = mapSheet.getDataRange().getValues();
    for (let i = 1; i < mapData.length; i++) {
      if (String(mapData[i][0] || '').trim() === cls &&
          String(mapData[i][2] || '').trim() === finalRecId) {
        throw new Error(`Class "${cls}" is already mapped for this school. Edit the existing mapping instead.`);
      }
    }

    // 2. Write ONE row to Class_Mapping
    mapSheet.appendRow([cls, sections.join(','), finalRecId]);

    // 3. Write N rows to Class_Sections (skip existing)
    const secData = secSheet.getDataRange().getValues();
    const existing = new Set();
    for (let i = 1; i < secData.length; i++) {
      const combo = String(secData[i][0] || '').trim();
      const rec   = String(secData[i][1] || '').trim();
      if (combo) existing.add(`${combo}|${rec}`);
    }

    const newRows = [];
    for (const sec of sections) {
      const combo = `${cls}-${sec}`;
      if (existing.has(`${combo}|${finalRecId}`)) continue;
      newRows.push([combo, finalRecId]);
    }

    if (newRows.length > 0) {
      secSheet.getRange(secSheet.getLastRow() + 1, 1, newRows.length, 2).setValues(newRows);
    }

    SpreadsheetApp.flush();
    return {
      success: true,
      message: `Added Class "${cls}" with sections ${sections.join(', ')} (${newRows.length} new section rows)`,
      data: { Class: cls, Sections: sections.join(','), REC_ID: finalRecId }
    };
  } catch (e) {
    Logger.log('Error in addClassMappingWithSections: ' + e.message);
    return { success: false, message: e.message };
  }
}

function updateClassMapping(className, newSectionsStr, recId, oldSectionsStr) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    return { success: false, message: 'System busy. Please try again.' };
  }
  try {
    ensureClassMappingSheet_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mapSheet = ss.getSheetByName('Class_Mapping');
    const secSheet = ss.getSheetByName('Class_Sections');
    if (!mapSheet) throw new Error('Class_Mapping sheet not found');
    if (!secSheet) throw new Error('Class_Sections sheet not found');

    const cls = String(className || '').trim();
    const finalRecId = String(recId || '').trim();
    if (!cls || !finalRecId) throw new Error('className and recId are required');

    const newSections = String(newSectionsStr || '')
      .split(',').map(s => s.trim().toUpperCase()).filter(Boolean);
    const oldSections = String(oldSectionsStr || '')
      .split(',').map(s => s.trim().toUpperCase()).filter(Boolean);

    if (!newSections.length) throw new Error('At least one section is required');
    if (!newSections.every(s => /^[A-Z]$/.test(s))) {
      throw new Error('Sections must be single letters (e.g., A,B,C)');
    }

    // 1. Find row in Class_Mapping
    const mapData = mapSheet.getDataRange().getValues();
    let mapRowIndex = -1;
    for (let i = 1; i < mapData.length; i++) {
      if (String(mapData[i][0] || '').trim() === cls &&
          String(mapData[i][2] || '').trim() === finalRecId) {
        mapRowIndex = i + 1;
        break;
      }
    }
    if (mapRowIndex === -1) {
      return { success: false, message: `Class "${cls}" mapping not found for this school` };
    }

    // 2. Update Class_Mapping row
    mapSheet.getRange(mapRowIndex, 2).setValue(newSections.join(','));

    // 3. APPEND-ONLY: add newly-added sections to Class_Sections.
    //    Removed sections stay in Class_Sections (orphaned but preserved).
    const added = newSections.filter(s => !oldSections.includes(s));

    if (added.length > 0) {
      const secData = secSheet.getDataRange().getValues();
      const existing = new Set();
      for (let i = 1; i < secData.length; i++) {
        const combo = String(secData[i][0] || '').trim();
        const rec   = String(secData[i][1] || '').trim();
        if (combo) existing.add(`${combo}|${rec}`);
      }
      const newRows = added
        .map(sec => [`${cls}-${sec}`, finalRecId])
        .filter(([combo]) => !existing.has(`${combo}|${finalRecId}`));
      if (newRows.length > 0) {
        secSheet.getRange(secSheet.getLastRow() + 1, 1, newRows.length, 2).setValues(newRows);
      }
    }

    SpreadsheetApp.flush();
    return {
      success: true,
      message: `Updated Class "${cls}": +${added.length} new section(s) appended to Class_Sections (removals preserved)`,
      data: { Class: cls, Sections: newSections.join(','), REC_ID: finalRecId }
    };
  } catch (e) {
    Logger.log('Error in updateClassMapping: ' + e.message);
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}


function deleteClassMappingWithSections(className, recId) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    return { success: false, message: 'System busy. Please try again.' };
  }
  try {
    ensureClassMappingSheet_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mapSheet = ss.getSheetByName('Class_Mapping');
    if (!mapSheet) throw new Error('Class_Mapping sheet not found');

    const cls = String(className || '').trim();
    const finalRecId = String(recId || '').trim();
    if (!cls || !finalRecId) throw new Error('className and recId are required');

    // 1. Delete only from Class_Mapping
    const mapData = mapSheet.getDataRange().getValues();
    const mapRowsToDelete = [];
    for (let i = mapData.length - 1; i >= 1; i--) {
      if (String(mapData[i][0] || '').trim() === cls &&
          String(mapData[i][2] || '').trim() === finalRecId) {
        mapRowsToDelete.push(i + 1);
      }
    }
    if (mapRowsToDelete.length === 0) {
      return { success: false, message: `Class "${cls}" mapping not found for this school` };
    }
    mapRowsToDelete.forEach(r => mapSheet.deleteRow(r));

    // 2. Class_Sections is NOT touched — sections remain (append-only registry)

    SpreadsheetApp.flush();
    return {
      success: true,
      message: `Removed class "${cls}" from Class Mappings (sections preserved in Class Sections)`
    };
  } catch (e) {
    Logger.log('Error in deleteClassMappingWithSections: ' + e.message);
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

