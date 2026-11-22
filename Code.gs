
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

        // v7.20 (F1): an Inactive Users row cannot log in, for admin and
        // superadmin alike — same refusal as the marking-app login
        // ('User account is inactive. Please contact administrator.').
        // Status is found by header name; a missing Status column or a
        // blank Status means Active, so existing rows are grandfathered.
        // Returned (not thrown) so the message reaches the login screen
        // verbatim via setError(result.message); a throw would be caught
        // below and prefixed 'Authentication error: '.
        const statusCol = data[0].indexOf('Status');
        const userStatus = statusCol !== -1 ? String(data[i][statusCol] || '').trim() : '';
        if (userStatus.toLowerCase() === 'inactive') {
          return { success: false, message: 'User account is inactive. Please contact administrator.' };
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

// ============ v7.8: CALLER GATE (Fix 1) ============
// The portal login is a custom username/password check against the Users
// sheet (authenticateUser above) — NOT Google sign-in — so
// Session.getActiveUser() cannot prove who is calling a write function.
// The strongest check available without redesigning login is: the caller
// name the client sends must belong to a Users row whose UserType is
// admin or superadmin. That stops a teacher (or an unknown name) calling
// a write function directly.
// Residual gap, stated plainly: a direct caller who knows a valid admin
// username can still act as that admin, because the server never sees a
// password proof on write calls. Closing that needs a login-issued token
// (minted at login, sent on every write, validated server-side).
const ADMIN_USER_TYPES = ['admin', 'superadmin'];
const ALLOWED_USER_TYPES = ['admin', 'teacher', 'superadmin'];
// The 24 tab keys the portal UI offers in the allowedTabs picker
// (HTML v7.7, UserManagement 'Admin Access' section, flat options list).
// v7.23: 'outage_permissions' added for the Outage Permissions tab (see grantOutagePermissions).
const ALLOWED_TABS = ['users', 'user_permissions', 'class_sections', 'previous_date_permissions', 'outage_permissions', 'teacher_assignments', 'terms', 'class_terms', 'student_data', 'student_attendance', 'student_reports', 'audit_reports', 'attendance_reports', 'monthly_bar_report', 'access_hours_report', 'daily_hours_report', 'itreb_report', 'batch_hours_log', 'attendance_correction', 'attendance_correction_log', 'qr_management', 'reserve_codes', 'hours_setup', 'hour_categories', 'transfer_certificates', 'leaving_reasons', 'countries', 'rec_names'];
const ADMIN_ONLY_MSG = 'Only an admin can make this change.';

// Look up a caller by username in the Users sheet. Returns
// { username, userType, recId } or null when there is no such user.
function getCallerUser_(ss, username) {
  const name = String(username || '').trim();
  if (!name) return null;
  const usersSheet = ss.getSheetByName('Users');
  if (!usersSheet) return null;
  const userData = usersSheet.getDataRange().getValues();
  if (!userData.length) return null;
  const headers = userData[0] || [];
  const typeCol = headers.indexOf('UserType');
  const recCol = headers.indexOf('REC_ID');
  for (let i = 1; i < userData.length; i++) {
    if (String(userData[i][0] || '').trim() === name) {
      return {
        username: name,
        userType: typeCol !== -1 ? String(userData[i][typeCol] || '').trim() : '',
        recId: recCol !== -1 ? String(userData[i][recCol] || '').trim() : ''
      };
    }
  }
  return null;
}

// Refuse unless the caller is a Users row with UserType admin or superadmin.
// Returns { ok: true, caller } or { ok: false, error: { success: false, message } }.
function requireAdminCaller_(ss, username) {
  const caller = getCallerUser_(ss, username);
  if (!caller) return { ok: false, error: { success: false, message: 'User not found. Please re-login.' } };
  if (ADMIN_USER_TYPES.indexOf(caller.userType) === -1) {
    return { ok: false, error: { success: false, message: ADMIN_ONLY_MSG } };
  }
  return { ok: true, caller: caller };
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



// ============ TRANSFER CERTIFICATES (v7.10, v7.11, v7.14, v7.15) ============
// A student is deactivated only with a reason: a Transfer (leaving for
// another REC / jurisdiction / country) or an Error (wrong enrolment /
// duplicate). The reason is captured in the portal form, validated here
// server-side, and stored as one row in Transfer_Certificates plus
// Inactive_Type / Inactive_Reason on the student row. The certificate
// itself is a snapshot: re-printing a TC never changes, even if the
// student row is edited later.
// New Student_Data columns (appended at the END, never in the middle):
//   Global_ID, Date_of_Birth, Father_Name, Address, Phone,
//   Inactive_Type, Inactive_Reason
// v7.14: Father_Name is no longer collected or stored (the sample
// certificate has no Father Name field) — the column stays on the sheet.
// Last_Class_Passed / Class_Studying are derived from the student row,
// never taken from the popup. A student with any certificate record
// (Transfer or Error) is frozen: no Edit, no Delete, no Reactivate
// (server-enforced here; the client hides the buttons). Date_of_Birth
// is now on the Add/Edit form too.
//   Previous_Std_ID, Previous_REC_ID (v7.11: re-joining merge link)
// Global_ID is STU-<year>-<NNNN>, one sequence per year across ALL
// schools, generated server-side only, stable across Std_ID changes.
// TC_Number is <REC Code>-<NNNN>, one running sequence per REC.
// Extra snapshot columns sit at the END of Transfer_Certificates
// (New_Address, Involvement, then the v7.15 Country / City /
// Destination_REC_Name / contact columns): the certificate prints them,
// and the spec'd 29 columns hold none of them, so they are snapshotted
// there. Involvement is no longer collected (v7.15) — the column stays
// for old rows and always stores '' for new ones.
const TC_SHEET = 'Transfer_Certificates';
const STUDENT_NEW_COLUMNS = ['Global_ID', 'Date_of_Birth', 'Father_Name', 'Address', 'Phone', 'Inactive_Type', 'Inactive_Reason', 'Previous_Std_ID', 'Previous_REC_ID', 'Secondary_Phone', 'Primary_Contact_Relation', 'Secondary_Contact_Relation'];
const TC_HEADERS = ['TC_Number', 'Type', 'Global_ID', 'Std_ID', 'Student_Name', 'REC_ID', 'REC_Name', 'LB', 'Term_ID', 'Class', 'Section', 'Date_of_Birth', 'Father_Name', 'Address', 'Phone', 'Last_Class_Passed', 'Class_Studying', 'Leaving_Reason', 'Destination_Type', 'Destination', 'Expected_Joining_Date', 'Issue_Date', 'Error_Reason', 'Duplicate_Of_Std_ID', 'Justification', 'Documents_Issued', 'Remarks', 'Issued_By', 'Created_At', 'New_Address', 'Involvement', 'Country', 'City', 'Destination_REC_Name', 'Primary_Contact_Relation', 'Secondary_Contact_No', 'Secondary_Contact_Relation'];
// v7.15: leaving reasons no longer live in a hard-coded list — they are
// rows of the Leaving_Reasons sheet (managed in the portal, active only).
// Category drives the popup/server rules: Moving + Temporary need a REC
// destination (Country is always Pakistan); Moving-Abroad needs a country
// (never Pakistan) + city; Medical + Left REC need a justification and
// have no destination. Requires_Destination / Requires_Justification are
// derived from Category, never typed in by hand.
const LEAVING_REASONS_SHEET = 'Leaving_Reasons';
const LEAVING_REASONS_HEADERS = ['Reason', 'Category', 'Requires_Destination', 'Requires_Justification', 'Is_Active'];
const LEAVING_REASON_CATEGORIES = ['Moving', 'Moving-Abroad', 'Temporary', 'Medical', 'Left REC'];
const LEAVING_REASONS_SEED = [
  ['Shifting', 'Moving', 'Yes', 'No', 'TRUE'],
  ['Migration', 'Moving-Abroad', 'Yes', 'No', 'TRUE'],
  ['Temporary Shifting to Another REC', 'Temporary', 'Yes', 'No', 'TRUE'],
  ['Illness', 'Medical', 'No', 'Yes', 'TRUE'],
  ['REC far from Home', 'Left REC', 'No', 'Yes', 'TRUE'],
  ['Secular School Work Over Load', 'Left REC', 'No', 'Yes', 'TRUE'],
  ['Other (Please Specify)', 'Left REC', 'No', 'Yes', 'TRUE']
];
// v7.15: Countries sheet (managed in the portal, active only).
const COUNTRIES_SHEET = 'Countries';
const COUNTRIES_HEADERS = ['Country', 'Is_Active'];
const COUNTRIES_SEED = ['Pakistan', 'Canada', 'USA', 'United Kingdom', 'UAE', 'Qatar', 'Australia', 'Portugal', 'Kenya', 'Tanzania', 'India'];
const REC_NAMES_SHEET = 'RECNames';
// v7.15: every contact number (primary + secondary, and the student-form
// Phone when it exists) is exactly 11 digits starting with 03.
const TC_PHONE_RE = /^03\d{9}$/;
const TC_CONTACT_RELATIONS = ['Father', 'Mother', 'Guardian'];
const TC_ERROR_REASONS = ['Wrong enrolment', 'Duplicate'];
// v7.15: 'Progress File' is renamed 'Progress Card' (Annexure 5B wording);
// 'Activity Book' joins the issued set. Old rows storing 'Progress File'
// are read as 'Progress Card' when the certificate prints.
const TC_DOCUMENTS = ['Evaluation Card', 'Progress Card', 'Curriculum Set', 'Activity Book'];
const TC_TZ = 'Asia/Karachi';

function isStudentDataSheet_(sheetName) {
  return String(sheetName || '').trim().toLowerCase() === 'student_data';
}

function isTransferCertificatesSheet_(sheetName) {
  return String(sheetName || '').trim().toLowerCase() === 'transfer_certificates';
}

// Append the v7.10/v7.11 Student_Data columns at the end when any is missing.
// Lazy only — called from write paths (updateSheetRow / addStudent), never
// from setupSpreadsheet() and never on read. Mirrors ensureHoursConfigColumns_.
function ensureStudentColumns_(sheet) {
  if (!sheet) return [];
  let lastCol = sheet.getLastColumn();
  if (lastCol < 1) return [];
  let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  const missing = STUDENT_NEW_COLUMNS.filter(function (c) { return headers.indexOf(c) === -1; });
  if (missing.length === 0) return headers;
  sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
  SpreadsheetApp.flush();
  lastCol = sheet.getLastColumn();
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
}

// Create Transfer_Certificates with its headers when missing; append any
// missing header at the end when the sheet already exists. Never called
// from setupSpreadsheet() — only on first TC/error issue.
function getOrCreateTransferCertificatesSheet_(ss) {
  let sheet = ss.getSheetByName(TC_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(TC_SHEET);
    sheet.appendRow(TC_HEADERS);
    sheet.setFrozenRows(1);
    SpreadsheetApp.flush();
    return sheet;
  }
  const existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
  const missing = TC_HEADERS.filter(function (h) { return existing.indexOf(h) === -1; });
  if (missing.length > 0) {
    sheet.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
    SpreadsheetApp.flush();
  }
  return sheet;
}

// ============ MANAGED TABLES (v7.15): Leaving_Reasons / Countries / RECNames ============
// Leaving_Reasons and Countries are created lazily (headers + seed rows)
// on FIRST ACCESS — never by setupSpreadsheet(). A sheet that exists but
// has no data rows is returned empty on purpose: the admin emptied it.
// RECNames is a pre-existing master sheet; only its Is_Active column is
// appended lazily (a missing column, or a blank cell, means active).
function isLeavingReasonsSheet_(sheetName) {
  return String(sheetName || '').trim().toLowerCase() === 'leaving_reasons';
}

function isCountriesSheet_(sheetName) {
  return String(sheetName || '').trim().toLowerCase() === 'countries';
}

function isRecNamesSheet_(sheetName) {
  return String(sheetName || '').trim().toLowerCase() === 'rec_names';
}

function isManagedTableSheet_(sheetName) {
  return isLeavingReasonsSheet_(sheetName) || isCountriesSheet_(sheetName) || isRecNamesSheet_(sheetName);
}

// v7.15: the portal tab key 'rec_names' has no underscore twin — the
// sheet is 'RECNames', so a plain case-insensitive lookup misses it.
// Every generic write/read path fetches its sheet through here.
function resolveSheet_(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet && isRecNamesSheet_(sheetName)) sheet = ss.getSheetByName(REC_NAMES_SHEET);
  return sheet;
}

// Create the sheet with headers + seed rows when it does not exist yet.
// Returns the sheet (new or existing). Never re-seeds an existing sheet.
function getOrCreateLeavingReasonsSheet_(ss) {
  let sheet = ss.getSheetByName(LEAVING_REASONS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(LEAVING_REASONS_SHEET);
    sheet.appendRow(LEAVING_REASONS_HEADERS);
    sheet.getRange(2, 1, LEAVING_REASONS_SEED.length, LEAVING_REASONS_HEADERS.length).setValues(LEAVING_REASONS_SEED);
    sheet.setFrozenRows(1);
    SpreadsheetApp.flush();
    return sheet;
  }
  return sheet;
}

function getOrCreateCountriesSheet_(ss) {
  let sheet = ss.getSheetByName(COUNTRIES_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(COUNTRIES_SHEET);
    sheet.appendRow(COUNTRIES_HEADERS);
    sheet.getRange(2, 1, COUNTRIES_SEED.length, COUNTRIES_HEADERS.length).setValues(COUNTRIES_SEED.map(function (c) { return [c, 'TRUE']; }));
    sheet.setFrozenRows(1);
    SpreadsheetApp.flush();
    return sheet;
  }
  return sheet;
}

// All rows of a managed sheet as objects keyed by header; [] when the
// sheet does not exist or has no data rows. `seedFirst` creates + seeds
// a missing Leaving_Reasons / Countries sheet (first access).
function getManagedRows_(ss, sheetName, seedFirst) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    if (!seedFirst) return [];
    if (isLeavingReasonsSheet_(sheetName)) sheet = getOrCreateLeavingReasonsSheet_(ss);
    else if (isCountriesSheet_(sheetName)) sheet = getOrCreateCountriesSheet_(ss);
    else return [];
  }
  if (sheet.getLastRow() <= 1) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i].some(function (v) { return v !== '' && v !== null && v !== undefined; })) continue;
    const row = {};
    for (let j = 0; j < headers.length; j++) {
      const v = data[i][j];
      row[headers[j]] = v !== null && v !== undefined ? String(v).trim() : '';
    }
    row._rowNumber = i + 1;
    out.push(row);
  }
  return out;
}

// One Leaving_Reasons row by exact Reason text, or null. Seeds the sheet
// on first access so a direct deactivation call works before any popup load.
function getLeavingReasonRow_(ss, reason) {
  const want = String(reason || '').trim();
  if (!want) return null;
  const rows = getManagedRows_(ss, LEAVING_REASONS_SHEET, true);
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i].Reason || '').trim() === want) return rows[i];
  }
  return null;
}

// Active countries (seeded on first access). Pakistan is included here —
// callers that need "abroad only" filter it out themselves.
function getActiveCountries_(ss) {
  return getManagedRows_(ss, COUNTRIES_SHEET, true).filter(function (r) {
    return isActiveFlag_(r.Is_Active);
  });
}

// Client-facing, admin-gated: ACTIVE leaving reasons / countries only.
// These are what the Deactivate popup offers.
function getLeavingReasons(currentUsername) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gate = requireAdminCaller_(ss, currentUsername);
  if (!gate.ok) return gate.error;
  return getManagedRows_(ss, LEAVING_REASONS_SHEET, true).filter(function (r) {
    return isActiveFlag_(r.Is_Active);
  });
}

function getCountries(currentUsername) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gate = requireAdminCaller_(ss, currentUsername);
  if (!gate.ok) return gate.error;
  return getActiveCountries_(ss);
}

// Client-facing, admin-gated: ALL rows (active + inactive) for the
// management screens, so an admin can re-activate a switched-off row.
function getLeavingReasonsAll(currentUsername) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gate = requireAdminCaller_(ss, currentUsername);
  if (!gate.ok) return gate.error;
  return getManagedRows_(ss, LEAVING_REASONS_SHEET, true);
}

function getCountriesAll(currentUsername) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gate = requireAdminCaller_(ss, currentUsername);
  if (!gate.ok) return gate.error;
  return getManagedRows_(ss, COUNTRIES_SHEET, true);
}

// Append the Is_Active column to RECNames when it is missing. A missing
// column (or a blank cell) means the REC is active.
function ensureRecNamesActiveColumn_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
  if (headers.indexOf('Is_Active') !== -1) return;
  sheet.getRange(1, headers.length + 1, 1, 1).setValues([['Is_Active']]);
  SpreadsheetApp.flush();
}

// True when a value is a valid contact number: exactly 11 digits, 03....
function isValidTcPhone_(v) {
  return TC_PHONE_RE.test(String(v == null ? '' : v).trim());
}

// Requires flags derived from Category — the single source of truth, so
// the sheet can never hold a row whose flags disagree with its category.
function leavingReasonFlags_(category) {
  if (category === 'Moving' || category === 'Moving-Abroad' || category === 'Temporary') {
    return { Requires_Destination: 'Yes', Requires_Justification: 'No' };
  }
  return { Requires_Destination: 'No', Requires_Justification: 'Yes' };
}

// Validate a Leaving_Reasons / Countries / RECNames row on Add
// (excludeRow = 0 on Add, or the own row number on Edit). Normalises
// Is_Active in rowData and, for reasons, derives Requires_Destination /
// Requires_Justification from Category. Returns an error message string,
// or null when OK.
function validateManagedTableRow_(sheetName, sheet, headers, rowData, excludeRow) {
  if (isLeavingReasonsSheet_(sheetName)) {
    // v7.20 (F2): on Edit (excludeRow = the own row number) an absent
    // Reason/Category keeps the stored value (v7.8 Fix-4 rule), so an
    // Is_Active-only edit is not refused for a missing Reason. This is
    // why Countries passed and Leaving_Reasons did not: the edit path
    // calls this validator only for Leaving_Reasons, while Countries /
    // RECNames skip it on Edit — the fix belongs here, in the shared
    // validator, not in a new edit-path branch. On Add (excludeRow 0)
    // there is no stored row, so a blank Reason is still refused.
    // Only undefined/null are filled; an explicitly sent '' is still a
    // blank Reason and is refused below (and by the immutable check).
    if (excludeRow) {
      const storedVals = sheet.getDataRange().getValues()[excludeRow - 1] || [];
      const rIdx0 = headers.indexOf('Reason');
      const cIdx0 = headers.indexOf('Category');
      if ((rowData.Reason === undefined || rowData.Reason === null) && rIdx0 !== -1) rowData.Reason = storedVals[rIdx0];
      if ((rowData.Category === undefined || rowData.Category === null) && cIdx0 !== -1) rowData.Category = storedVals[cIdx0];
    }
    const reason = String(rowData.Reason == null ? '' : rowData.Reason).trim();
    if (!reason) return 'Reason is required.';
    const category = String(rowData.Category == null ? '' : rowData.Category).trim();
    if (LEAVING_REASON_CATEGORIES.indexOf(category) === -1) {
      return 'Category must be one of: ' + LEAVING_REASON_CATEGORIES.join(', ') + '.';
    }
    const flags = leavingReasonFlags_(category);
    rowData.Requires_Destination = flags.Requires_Destination;
    rowData.Requires_Justification = flags.Requires_Justification;
    const rIdx = headers.indexOf('Reason');
    if (rIdx !== -1) {
      const all = sheet.getDataRange().getValues();
      for (let i = 1; i < all.length; i++) {
        if (excludeRow && (i + 1) === excludeRow) continue;
        if (String(all[i][rIdx] || '').trim().toLowerCase() === reason.toLowerCase()) {
          return 'Reason "' + reason + '" already exists. Re-activate the existing row instead of adding a duplicate.';
        }
      }
    }
    return null;
  }
  if (isCountriesSheet_(sheetName)) {
    const country = String(rowData.Country == null ? '' : rowData.Country).trim();
    if (!country) return 'Country is required.';
    const cIdx = headers.indexOf('Country');
    if (cIdx !== -1) {
      const all = sheet.getDataRange().getValues();
      for (let i = 1; i < all.length; i++) {
        if (excludeRow && (i + 1) === excludeRow) continue;
        if (String(all[i][cIdx] || '').trim().toLowerCase() === country.toLowerCase()) {
          return 'Country "' + country + '" already exists. Re-activate the existing row instead of adding a duplicate.';
        }
      }
    }
    return null;
  }
  // RECNames: LB Code, LB, REC Code and REC Name are all required; the REC
  // Code must be numeric and unique; S.# is auto-filled by addSheetRow.
  const recCode = String(rowData['REC Code'] == null ? '' : rowData['REC Code']).trim();
  if (!String(rowData['LB Code'] == null ? '' : rowData['LB Code']).trim()) return 'LB Code is required.';
  if (!String(rowData.LB == null ? '' : rowData.LB).trim()) return 'LB is required.';
  if (!recCode) return 'REC Code is required.';
  if (!/^\d+$/.test(recCode)) return 'REC Code must be numeric.';
  if (!String(rowData['REC Name'] == null ? '' : rowData['REC Name']).trim()) return 'REC Name is required.';
  const codeIdx = headers.indexOf('REC Code');
  if (codeIdx !== -1) {
    const all = sheet.getDataRange().getValues();
    for (let i = 1; i < all.length; i++) {
      if (excludeRow && (i + 1) === excludeRow) continue;
      if (String(all[i][codeIdx] || '').trim() === recCode) {
        return 'REC Code "' + recCode + '" already exists. Re-activate the existing row instead of adding a duplicate.';
      }
    }
  }
  return null;
}

// Managed tables (v7.15): rows are immutable once added — only Is_Active
// may change. Returns an error message string, or null when OK. Normalises
// Is_Active in rowData (a blank value keeps the stored one). An absent
// field counts as unchanged (v7.8 Fix-4 rule).
function checkManagedTableImmutable_(sheetName, sheet, headers, rowData, rowIndex) {
  const storedRow = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  const normVal = function (v) { return String(v == null ? '' : v).trim().toLowerCase(); };
  const changed = headers.some(function (h) {
    if (h === 'Is_Active') return false;
    const si = headers.indexOf(h);
    const oldV = storedRow[si];
    const newV = rowData[h];
    if (newV === undefined) return false;
    return normVal(newV) !== normVal(oldV);
  });
  if (changed) return 'Rows cannot be amended once added — only Active/Inactive can change.';
  const iaIdx = headers.indexOf('Is_Active');
  if (rowData.Is_Active === undefined || rowData.Is_Active === null || String(rowData.Is_Active).trim() === '') {
    rowData.Is_Active = iaIdx !== -1 ? storedRow[iaIdx] : 'TRUE';
  } else {
    rowData.Is_Active = isActiveFlag_(rowData.Is_Active) ? 'TRUE' : 'FALSE';
  }
  return null;
}

// One stored cell as a trimmed string; a Date cell becomes 'yyyy-MM-dd'
// so the date checks below see a real date instead of a long date string.
function tcStoredStr_(v) {
  if (v == null) return '';
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return '';
    return v.getFullYear() + '-' + String(v.getMonth() + 1).padStart(2, '0') + '-' + String(v.getDate()).padStart(2, '0');
  }
  return String(v).trim();
}

// Next Global_ID for the current year (Asia/Karachi). CALLER MUST HOLD
// THE SCRIPT LOCK — reads the max NNNN among existing STU-<year>-<NNNN>
// values across all schools and returns max + 1.
function nextGlobalIdLocked_(ss) {
  const year = Number(Utilities.formatDate(new Date(), TC_TZ, 'yyyy'));
  const prefix = 'STU-' + year + '-';
  let maxN = 0;
  const sheet = ss.getSheetByName('Student_Data');
  if (sheet && sheet.getLastRow() > 1) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const gIdx = headers.indexOf('Global_ID');
    if (gIdx !== -1) {
      const vals = sheet.getRange(2, gIdx + 1, sheet.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < vals.length; i++) {
        const m = /^STU-(\d{4})-(\d{4})$/.exec(String(vals[i][0] || '').trim());
        if (m && Number(m[1]) === year) maxN = Math.max(maxN, Number(m[2]));
      }
    }
  }
  return prefix + String(maxN + 1).padStart(4, '0');
}

// Next TC number for one REC. CALLER MUST HOLD THE SCRIPT LOCK — reads
// the max NNNN among existing <REC Code>-<NNNN> values for that REC.
function nextTcNumberLocked_(ss, recCode) {
  const code = String(recCode || '').trim();
  const prefix = code + '-';
  let maxN = 0;
  const sheet = ss.getSheetByName(TC_SHEET);
  if (sheet && sheet.getLastRow() > 1) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const tIdx = headers.indexOf('TC_Number');
    if (tIdx !== -1) {
      const vals = sheet.getRange(2, tIdx + 1, sheet.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < vals.length; i++) {
        const v = String(vals[i][0] || '').trim();
        if (v.indexOf(prefix) !== 0) continue;
        const n = Number(v.slice(prefix.length));
        if (!isNaN(n)) maxN = Math.max(maxN, n);
      }
    }
  }
  return prefix + String(maxN + 1).padStart(4, '0');
}

// RECNames rows: [{ REC_Code, REC_Name, LB }]. [] when the sheet is missing.
// Columns: S.#, LB Code, LB, REC Code, REC Name, LB_REC.
function getRecNamesList_() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('RECNames');
    if (!sheet || sheet.getLastRow() <= 1) return [];
    // v7.15: lazily append the Is_Active column; FALSE rows are hidden.
    ensureRecNamesActiveColumn_(sheet);
    const data = sheet.getDataRange().getValues();
    const hdrRow = data[0].map(function (h) { return String(h).trim(); });
    const iaIdx = hdrRow.indexOf('Is_Active');
    const out = [];
    for (let i = 1; i < data.length; i++) {
      const code = String(data[i][3] == null ? '' : data[i][3]).trim();
      if (!code) continue;
      // A missing column (or a blank cell) means active; an explicit
      // FALSE deactivates the REC.
      if (iaIdx !== -1) {
        const iaVal = String(data[i][iaIdx] == null ? '' : data[i][iaIdx]).trim();
        if (iaVal !== '' && !isActiveFlag_(iaVal)) continue;
      }
      out.push({ REC_Code: code, REC_Name: String(data[i][4] || '').trim(), LB: String(data[i][2] || '').trim() });
    }
    // v7.14: numeric sort by REC_Code, so 999 sorts before 1102 and 1114
    // before 1201 (a string sort would put 1201 before 1114). The
    // RECNames sheet order itself is untouched.
    out.sort(function (a, b) { return (Number(a.REC_Code) || 0) - (Number(b.REC_Code) || 0); });
    return out;
  } catch (e) { return []; }
}

// Client-facing: the REC list for the Transfer destination dropdown.
function getRecNames() {
  return getRecNamesList_();
}

// Previous class in the 16-class master order; '' for IBT (nothing before it).
function prevClassOf_(cls) {
  const i = MASTER_CLASSES.indexOf(String(cls || '').trim());
  return i > 0 ? MASTER_CLASSES[i - 1] : '';
}

// The term a student was in when leaving: the ACTIVE Class_Terms row for
// the REC + class-section whose dates cover the leaving date; else the
// first ACTIVE row covering the class-section; else null.
function resolveTermForLeaving_(recId, classSection, leaveDate) {
  const rows = getClassTermsForRec_(recId);
  const cs = String(classSection || '').trim();
  if (!cs) return null;
  const leave = leaveDate instanceof Date ? leaveDate : parseTermDate_(leaveDate);
  let first = null;
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (!isActiveFlag_(r.Is_Active)) continue;
    if (!classCellMatches_(String(r.Class_Section || ''), cs)) continue;
    const s = parseTermDate_(r.Start_Date);
    const e = parseTermDate_(r.End_Date);
    if (!s || !e) continue;
    if (!first) first = { termId: String(r.Term_ID || '').trim(), start: s, end: e };
    if (leave && leave >= s && leave <= e) return { termId: String(r.Term_ID || '').trim(), start: s, end: e };
  }
  return first;
}

// Validate a deactivation payload (_tcPayload). `stored` is the student
// row as an object keyed by header (values may be '' when blank).
// Returns { error } or { error: null, clean } where clean is the
// normalised payload the write path uses. Pure validation — writes nothing.
function validateTcPayload_(ss, payload, stored, rowRecId) {
  const p = payload || {};
  const type = String(p.Type || '').trim();
  if (type !== 'Transfer' && type !== 'Error') {
    return { error: 'Deactivation type is required: Transfer or Error.' };
  }
  const clean = { Type: type };
  const eff = function (key) {
    const v = p[key];
    if (v !== undefined && v !== null && String(v).trim() !== '') return tcStoredStr_(v);
    return stored ? tcStoredStr_(stored[key]) : '';
  };
  if (type === 'Transfer') {
    // v7.15: the reason must be an ACTIVE row of the Leaving_Reasons
    // table (exact match). An old client sending a v7.14 reason value is
    // refused here — this is why the new Code.gs deploys before the new
    // HTML. The row's Category drives every rule below.
    const reason = String(p.Leaving_Reason || '').trim();
    const reasonRow = getLeavingReasonRow_(ss, reason);
    if (!reasonRow) {
      return { error: 'Leaving_Reason "' + reason + '" is not a known leaving reason. Pick one from the list.' };
    }
    if (!isActiveFlag_(reasonRow.Is_Active)) {
      return { error: 'Leaving_Reason "' + reason + '" is no longer active. Pick another reason.' };
    }
    clean.Leaving_Reason = reason;
    const category = String(reasonRow.Category || '').trim();
    clean.Category = category;
    const needsJust = String(reasonRow.Requires_Justification || '').trim().toLowerCase() === 'yes';
    // Destination, by category. The stored Destination_Type is derived
    // here — any value the client sent is ignored.
    clean.Country = '';
    clean.City = '';
    clean.Destination_REC_Name = '';
    clean.New_Address = '';
    if (category === 'Moving' || category === 'Temporary') {
      // Shifting (within Pakistan): Country is fixed to Pakistan, City
      // and a destination REC are mandatory, and the new resident
      // address is mandatory.
      const country = String(p.Country || '').trim();
      if (country && country !== 'Pakistan') {
        return { error: 'Country must be Pakistan for "' + reason + '".' };
      }
      clean.Country = 'Pakistan';
      const city = String(p.City || '').trim();
      if (!city) return { error: 'City is required for "' + reason + '".' };
      clean.City = city;
      const dest = String(p.Destination || '').trim();
      const typedRec = String(p.Destination_REC_Name || '').trim();
      if (dest) {
        const names = getRecNamesList_();
        const ok = names.some(function (r) { return String(r.REC_Code || '').trim() === dest; });
        if (!ok) return { error: 'Destination REC "' + dest + '" is not a known active REC Code.' };
        clean.Destination = dest;
        clean.Destination_REC_Name = '';
      } else if (typedRec) {
        // 'REC not in list' — the typed name is stored as-is.
        clean.Destination = typedRec;
        clean.Destination_REC_Name = typedRec;
      } else {
        return { error: 'Destination REC is required for "' + reason + '". Pick a REC, or choose "REC not in list" and type its name.' };
      }
      clean.Destination_Type = 'REC';
      const newAddr = String(p.New_Address || '').trim();
      if (!newAddr) return { error: 'New Address is required for "' + reason + '".' };
      clean.New_Address = newAddr;
    } else if (category === 'Moving-Abroad') {
      // Migration: an ACTIVE country other than Pakistan, plus a city.
      // No REC and no new address.
      const country = String(p.Country || '').trim();
      if (!country) return { error: 'Country is required for "' + reason + '".' };
      if (country === 'Pakistan') {
        return { error: 'Country must not be Pakistan for "' + reason + '". Shifting within Pakistan uses "Shifting".' };
      }
      const activeCountries = getActiveCountries_(ss).map(function (r) { return String(r.Country || '').trim(); });
      if (activeCountries.indexOf(country) === -1) {
        return { error: 'Country "' + country + '" is not an active country. Pick one from the list.' };
      }
      clean.Country = country;
      const city = String(p.City || '').trim();
      if (!city) return { error: 'City is required for "' + reason + '".' };
      clean.City = city;
      clean.Destination_Type = 'Out of Country';
      clean.Destination = country + ' — ' + city;
    } else {
      // Medical / Left REC: no destination, no new address.
      clean.Destination_Type = 'None';
      clean.Destination = '';
    }
    // Justification / specification, by category. For 'Other (Please
    // Specify)' the specification IS the justification.
    if (needsJust) {
      const just = String(p.Justification || '').trim();
      if (!just) {
        return { error: (reason === 'Other (Please Specify)' ? 'Please specify the reason for leaving.' : 'Justification is required for "' + reason + '".') };
      }
      clean.Justification = just;
    } else {
      clean.Justification = '';
    }
    // v7.15: Expected Date of Joining is KEPT (optional; prints on the
    // certificate under Issued For). It was never dropped.
    const exp = String(p.Expected_Joining_Date || '').trim();
    if (exp && !isRealDateString_(exp)) return { error: 'Expected_Joining_Date is not a real date.' };
    clean.Expected_Joining_Date = exp;
    // Documents: 'Progress File' from an old client reads as 'Progress Card'.
    const docsRaw = Array.isArray(p.Documents_Issued) ? p.Documents_Issued : (p.Documents_Issued ? String(p.Documents_Issued).split(',') : []);
    const docs = docsRaw.map(function (d) {
      const t = String(d).trim();
      return t === 'Progress File' ? 'Progress Card' : t;
    }).filter(Boolean);
    const badDoc = docs.filter(function (d) { return TC_DOCUMENTS.indexOf(d) === -1; });
    if (badDoc.length) return { error: 'Unknown document "' + badDoc[0] + '". Choose from: ' + TC_DOCUMENTS.join(', ') + '.' };
    clean.Documents_Issued = docs.join(', ');
    clean.Remarks = String(p.Remarks || '').trim();
    // v7.15: Involvement is no longer collected — any value sent by an
    // old client is ignored, never stored.
    clean.Involvement = '';
    // A Transfer needs the student's identity/contact details — taken
    // from the payload when given, else from the stored row.
    // v7.14: Father_Name is no longer collected (the sample certificate
    // has no Father Name field) — any value sent by an old client is
    // ignored, never stored. Last_Class_Passed is derived server-side
    // in appendTcRowLocked_ (prevClassOf_ the student's class) — any
    // value sent by an old client is ignored, never validated here.
    clean.Date_of_Birth = eff('Date_of_Birth');
    clean.Father_Name = '';
    clean.Address = eff('Address');
    clean.Phone = eff('Phone');
    if (!clean.Date_of_Birth) return { error: 'Date of Birth is required for a Transfer.' };
    if (!isRealDateString_(clean.Date_of_Birth)) return { error: 'Date of Birth is not a real date.' };
    if (!clean.Address) return { error: 'Address is required for a Transfer.' };
    if (!clean.Phone) return { error: 'Phone is required for a Transfer.' };
    // v7.15: contacts. The primary number IS the Phone above; its
    // relation is mandatory (payload, else the stored row from an
    // earlier TC). A secondary number needs its relation too.
    if (!isValidTcPhone_(clean.Phone)) {
      return { error: 'Phone must be exactly 11 digits starting with 03 (e.g. 03XXXXXXXXX).' };
    }
    const primRel = eff('Primary_Contact_Relation');
    if (TC_CONTACT_RELATIONS.indexOf(primRel) === -1) {
      return { error: 'Primary Contact "Belongs To" is required: Father, Mother or Guardian.' };
    }
    clean.Primary_Contact_Relation = primRel;
    const secNo = String(p.Secondary_Contact_No || '').trim() || (stored ? tcStoredStr_(stored.Secondary_Phone) : '');
    const secRel = eff('Secondary_Contact_Relation');
    if (secNo) {
      if (!isValidTcPhone_(secNo)) {
        return { error: 'Secondary Contact No. must be exactly 11 digits starting with 03 (e.g. 03XXXXXXXXX).' };
      }
      if (TC_CONTACT_RELATIONS.indexOf(secRel) === -1) {
        return { error: 'Secondary Contact "Belongs To" is required when a secondary number is given: Father, Mother or Guardian.' };
      }
      clean.Secondary_Contact_No = secNo;
      clean.Secondary_Contact_Relation = secRel;
    } else {
      clean.Secondary_Contact_No = '';
      clean.Secondary_Contact_Relation = '';
    }
    clean.Last_Class_Passed = '';
    clean.Error_Reason = '';
    clean.Duplicate_Of_Std_ID = '';
  } else {
    const errReason = String(p.Error_Reason || '').trim();
    if (TC_ERROR_REASONS.indexOf(errReason) === -1) {
      return { error: 'Error_Reason must be one of: ' + TC_ERROR_REASONS.join(', ') + '.' };
    }
    clean.Error_Reason = errReason;
    clean.Justification = String(p.Justification || '').trim();
    if (!clean.Justification) return { error: 'Justification is required for an Error deactivation.' };
    clean.Duplicate_Of_Std_ID = '';
    if (errReason === 'Duplicate') {
      const dup = String(p.Duplicate_Of_Std_ID || '').trim();
      if (!dup) return { error: 'Duplicate_Of_Std_ID is required when the error is a Duplicate.' };
      // Must be a real Std_ID in the same school.
      let found = false;
      const stu = ss.getSheetByName('Student_Data');
      if (stu && stu.getLastRow() > 1) {
        const hd = stu.getRange(1, 1, 1, stu.getLastColumn()).getValues()[0];
        const sIdx = hd.indexOf('Std_ID');
        const rIdx = hd.indexOf('REC_ID');
        if (sIdx !== -1) {
          const all = stu.getRange(2, 1, stu.getLastRow() - 1, stu.getLastColumn()).getValues();
          found = all.some(function (r) {
            if (String(r[sIdx] || '').trim() !== dup) return false;
            if (rIdx !== -1 && rowRecId) return String(r[rIdx] || '').trim() === String(rowRecId).trim();
            return true;
          });
        }
      }
      if (!found) return { error: 'Duplicate_Of_Std_ID "' + dup + '" is not a student in this school.' };
      clean.Duplicate_Of_Std_ID = dup;
    }
    clean.Leaving_Reason = '';
    clean.Destination_Type = '';
    clean.Destination = '';
    clean.Expected_Joining_Date = '';
    clean.Documents_Issued = '';
    clean.Remarks = String(p.Remarks || '').trim();
    clean.Involvement = ''; // v7.15: no longer collected.
    clean.New_Address = ''; // v7.15: an Error has no destination.
    clean.Date_of_Birth = eff('Date_of_Birth');
    clean.Father_Name = ''; // v7.14: no longer collected or stored.
    clean.Address = eff('Address');
    clean.Phone = eff('Phone');
    if (clean.Date_of_Birth && !isRealDateString_(clean.Date_of_Birth)) return { error: 'Date of Birth is not a real date.' };
    clean.Last_Class_Passed = '';
  }
  return { error: null, clean: clean };
}

// Append one Transfer_Certificates row from a validated (clean) payload.
// CALLER MUST HOLD THE SCRIPT LOCK. Returns the stored row as an object
// (including _rowNumber) so the client can open the certificate.
function appendTcRowLocked_(ss, clean, student, issuedBy) {
  const sheet = getOrCreateTransferCertificatesSheet_(ss);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const now = new Date();
  const issueDate = Utilities.formatDate(now, TC_TZ, 'dd-MMM-yyyy');
  const createdAt = Utilities.formatDate(now, TC_TZ, 'dd-MMM-yyyy HH:mm');
  const tcNumber = clean.Type === 'Transfer' ? nextTcNumberLocked_(ss, student.REC_ID) : '';
  const classSection = String(student.Student_Class || '').trim() + (String(student.Student_Section || '').trim() ? '-' + String(student.Student_Section).trim() : '');
  const term = resolveTermForLeaving_(student.REC_ID, classSection, now);
  const lastPassed = clean.Type === 'Transfer'
    ? (clean.Last_Class_Passed || prevClassOf_(student.Student_Class))
    : '';
  const values = {
    TC_Number: tcNumber,
    Type: clean.Type,
    Global_ID: student.Global_ID || '',
    Std_ID: student.Std_ID || '',
    Student_Name: student.Student_Name || '',
    REC_ID: student.REC_ID || '',
    REC_Name: student.REC_Name || '',
    LB: student.LB || '',
    Term_ID: term ? term.termId : '',
    Class: student.Student_Class || '',
    Section: student.Student_Section || '',
    Date_of_Birth: clean.Date_of_Birth || '',
    Father_Name: clean.Father_Name || '',
    Address: clean.Address || '',
    Phone: clean.Phone || '',
    Last_Class_Passed: lastPassed,
    Class_Studying: student.Student_Class || '',
    Leaving_Reason: clean.Leaving_Reason || '',
    Destination_Type: clean.Destination_Type || '',
    Destination: clean.Destination || '',
    Expected_Joining_Date: clean.Expected_Joining_Date || '',
    Issue_Date: issueDate,
    Error_Reason: clean.Error_Reason || '',
    Duplicate_Of_Std_ID: clean.Duplicate_Of_Std_ID || '',
    Justification: clean.Justification || '',
    Documents_Issued: clean.Documents_Issued || '',
    Remarks: clean.Remarks || '',
    Issued_By: issuedBy || '',
    Created_At: createdAt,
    New_Address: clean.New_Address || '',
    Involvement: clean.Involvement || '',
    Country: clean.Country || '',
    City: clean.City || '',
    Destination_REC_Name: clean.Destination_REC_Name || '',
    Primary_Contact_Relation: clean.Primary_Contact_Relation || '',
    Secondary_Contact_No: clean.Secondary_Contact_No || '',
    Secondary_Contact_Relation: clean.Secondary_Contact_Relation || ''
  };
  const row = headers.map(function (h) { return values[h] !== undefined ? values[h] : ''; });
  sheet.appendRow(row);
  SpreadsheetApp.flush();
  const out = {};
  headers.forEach(function (h, j) { out[h] = row[j]; });
  out._rowNumber = sheet.getLastRow();
  return out;
}

// v7.16: does ANY certificate record exist for this student — Transfer
// or Error (including a Duplicate error)? Match by Global_ID when the
// stored row has one, else by Std_ID + REC_ID. The server backstop behind
// the client hiding Edit/Delete/Reactivate for any student with a TC row —
// a direct call is refused by the same check.
function hasAnyTc_(ss, globalId, stdId, recId) {
  const sheet = ss.getSheetByName(TC_SHEET);
  if (!sheet || sheet.getLastRow() <= 1) return false;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const gIdx = headers.indexOf('Global_ID');
  const sIdx = headers.indexOf('Std_ID');
  const rIdx = headers.indexOf('REC_ID');
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const gid = String(globalId || '').trim();
  const sid = String(stdId || '').trim();
  const rid = String(recId || '').trim();
  for (let i = 0; i < rows.length; i++) {
    if (gid && gIdx !== -1) {
      if (String(rows[i][gIdx] || '').trim() === gid) return true;
    } else if (sIdx !== -1 && rIdx !== -1) {
      if (String(rows[i][sIdx] || '').trim() === sid && String(rows[i][rIdx] || '').trim() === rid) return true;
    }
  }
  return false;
}

// All Transfer_Certificates rows for one school ('' = all schools, superadmin).
// Reads the sheet directly: Error rows have a blank TC_Number, so the
// usual getSheetData first-column check would drop them.
function getTransferCertificates(recId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TC_SHEET);
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    const headers = data[0];
    const recCol = headers.indexOf('REC_ID');
    const out = [];
    for (let i = 1; i < data.length; i++) {
      const hasAny = data[i].some(function (v) { return v !== '' && v !== null && v !== undefined; });
      if (!hasAny) continue;
      if (recCol !== -1 && recId) {
        if (String(data[i][recCol] || '').trim() !== String(recId).trim()) continue;
      }
      const row = {};
      for (let j = 0; j < headers.length; j++) {
        const v = data[i][j];
        row[headers[j]] = v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : (v !== null && v !== undefined ? v.toString() : '');
      }
      row._rowNumber = i + 1;
      out.push(row);
    }
    return out;
  } catch (e) {
    Logger.log('Failed to fetch transfer certificates: ' + e.message);
    throw new Error('Failed to fetch transfer certificates: ' + e.message);
  }
}

// One certificate's data: the TC row (snapshot), the student row, the
// issuing school's REC/LB names, and the attendance history.
// tcNumberOrRow is a TC_Number ('1114-0001') for transfers, or
// 'row:<sheet row number>' for an Error record (which has no number).
// v7.18: the 12-month attendance history shared by the printed certificate
// (getTransferCertificateData) and the Deactivate-popup preview
// (getTcAttendancePreview) — one computation, so the preview can never
// disagree with the print. Inputs: the student, the class-section, the
// joining date (null when unknown) and the leaving date. Returns the
// per-month array (month label, monthKey, workingDays, daysPresent,
// termLabel) and the last Present day on or before the leaving date.
function buildTcAttendanceHistory_(ss, stdId, recId, classSection, joinDate, leaveDate) {
  // v7.14: attendance history = the 12 calendar months ending at the
  // leaving month, clipped to start no earlier than the student's
  // Date_of_Joining month (a student who joined 3 months ago shows 3
  // months, not 12). Months with no attendance still print blank, as the
  // sample does. The per-month Working Days / Days of Present
  // computation below is unchanged — only the month list changed.
  // Sample check: leaving Aug 2025 -> Sep 2024..Aug 2025.
  const months = [];
  let startY = leaveDate.getFullYear(), startM = leaveDate.getMonth() - 11;
  while (startM < 0) { startM += 12; startY--; }
  if (joinDate) {
    const jY = joinDate.getFullYear(), jM = joinDate.getMonth();
    if (jY > startY || (jY === startY && jM > startM)) { startY = jY; startM = jM; }
  }
  const endY = leaveDate.getFullYear(), endM = leaveDate.getMonth();
  if (startY > endY || (startY === endY && startM > endM)) {
    months.push(endY + '-' + String(endM + 1).padStart(2, '0'));
  } else {
    let y = startY, m = startM;
    while (y < endY || (y === endY && m <= endM)) {
      months.push(y + '-' + String(m + 1).padStart(2, '0'));
      m++; if (m > 11) { m = 0; y++; }
    }
  }
  // Working Days (ITREB definition): distinct Daily_Hours_Setup dates in
  // the month for the class-section; fallback to distinct attendance dates
  // when the section has no setup rows at all. Days Present = Present rows
  // that month for this student.
  const setupDays = {};
  const attDays = {};
  const setupSheet = ss.getSheetByName('Daily_Hours_Setup');
  if (setupSheet && setupSheet.getLastRow() > 1) {
    const sh = setupSheet.getRange(1, 1, 1, setupSheet.getLastColumn()).getValues()[0];
    const dIdx = sh.indexOf('Date'), cIdx = sh.indexOf('Class_Section');
    if (dIdx !== -1 && cIdx !== -1) {
      const all = setupSheet.getRange(2, 1, setupSheet.getLastRow() - 1, setupSheet.getLastColumn()).getValues();
      all.forEach(function (r) {
        if (String(r[cIdx] || '').trim() !== classSection) return;
        const d = parseTermDate_(r[dIdx]);
        if (!d) return;
        const k = termDateKey_(d);
        const mk = k.slice(0, 7);
        if (!setupDays[mk]) setupDays[mk] = {};
        setupDays[mk][k] = true;
      });
    }
  }
  const hasSetup = Object.keys(setupDays).length > 0;
  const presentByMonth = {};
  let lastPresent = null; // v7.15: latest Present day on/before leaving.
  const attSheet = ss.getSheetByName('Student_Attendance');
  if (attSheet && attSheet.getLastRow() > 1) {
    const sh = attSheet.getRange(1, 1, 1, attSheet.getLastColumn()).getValues()[0];
    const dIdx = sh.indexOf('Date'), sIdx = sh.indexOf('Std_ID'), stIdx = sh.indexOf('Status'), cIdx = sh.indexOf('Class_Section'), rIdx = sh.indexOf('REC_ID');
    if (dIdx !== -1 && sIdx !== -1) {
      const all = attSheet.getRange(2, 1, attSheet.getLastRow() - 1, attSheet.getLastColumn()).getValues();
      all.forEach(function (r) {
        if (String(r[sIdx] || '').trim() !== String(stdId || '').trim()) return;
        if (rIdx !== -1 && recId && String(r[rIdx] || '').trim() && String(r[rIdx] || '').trim() !== String(recId).trim()) return;
        const d = parseTermDate_(r[dIdx]);
        if (!d) return;
        const k = termDateKey_(d), mk = k.slice(0, 7);
        if (!hasSetup) {
          const cs = cIdx !== -1 ? String(r[cIdx] || '').trim() : classSection;
          if (cs !== classSection) return;
          if (!attDays[mk]) attDays[mk] = {};
          attDays[mk][k] = true;
        }
        if (stIdx !== -1 && String(r[stIdx] || '').trim().toLowerCase() === 'present') {
          presentByMonth[mk] = (presentByMonth[mk] || 0) + 1;
          // v7.15: latest Present day on or before the leaving date.
          if (d <= leaveDate && (!lastPresent || d > lastPresent)) lastPresent = d;
        }
      });
    }
  }
  const MONTHS_SHORT = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const attendance = months.map(function (mk) {
    const parts = mk.split('-');
    const label = MONTHS_SHORT[Number(parts[1]) - 1] + ' ' + parts[0].slice(2);
    const working = hasSetup ? Object.keys(setupDays[mk] || {}).length : Object.keys(attDays[mk] || {}).length;
    return { month: label, monthKey: mk, workingDays: working, daysPresent: presentByMonth[mk] || 0 };
  });
  // v7.15: the term label under each month — the term name from the
  // ACTIVE Class_Terms row for the REC + class-section covering that
  // month, blank when none. (Term names live in Terms; Class_Terms
  // carries only the Term_ID.)
  const termNamesById = {};
  try {
    getActiveTerms().forEach(function (t) {
      termNamesById[String(t.Term_ID || '').trim()] = String(t.Term_Name || t.Term_ID || '').trim();
    });
  } catch (e) {}
  const classTermRows = getClassTermsForRec_(recId);
  attendance.forEach(function (a) {
    const mp = a.monthKey.split('-');
    const mStart = new Date(Number(mp[0]), Number(mp[1]) - 1, 1);
    const mEnd = new Date(Number(mp[0]), Number(mp[1]), 0);
    let termLabel = '';
    for (let i = 0; i < classTermRows.length; i++) {
      const r = classTermRows[i];
      if (!isActiveFlag_(r.Is_Active)) continue;
      if (!classCellMatches_(String(r.Class_Section || ''), classSection)) continue;
      const s = parseTermDate_(r.Start_Date);
      const e = parseTermDate_(r.End_Date);
      if (!s || !e) continue;
      if (s <= mEnd && e >= mStart) {
        termLabel = termNamesById[String(r.Term_ID || '').trim()] || String(r.Term_ID || '').trim();
        break;
      }
    }
    a.termLabel = termLabel;
  });
  return { attendance: attendance, lastDayOfAttendance: lastPresent ? termDateKey_(lastPresent) : '' };
}

function getTransferCertificateData(tcNumberOrRow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TC_SHEET);
  if (!sheet) throw new Error('Transfer_Certificates sheet not found');
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error('Transfer certificate not found: ' + tcNumberOrRow);
  const headers = data[0];
  const key = String(tcNumberOrRow || '').trim();
  let tcRow = null;
  if (key.toLowerCase().indexOf('row:') === 0) {
    const rn = Number(key.slice(4));
    if (Number.isInteger(rn) && rn >= 2 && rn <= data.length) tcRow = data[rn - 1];
  } else {
    const tIdx = headers.indexOf('TC_Number');
    for (let i = 1; i < data.length; i++) {
      if (tIdx !== -1 && String(data[i][tIdx] || '').trim() === key) { tcRow = data[i]; break; }
    }
  }
  if (!tcRow) throw new Error('Transfer certificate not found: ' + key);
  const tc = {};
  headers.forEach(function (h, j) {
    const v = tcRow[j];
    tc[h] = v instanceof Date ? Utilities.formatDate(v, TC_TZ, 'yyyy-MM-dd') : (v !== null && v !== undefined ? v.toString() : '');
  });
  // Student row (current values — the certificate prints the TC snapshot;
  // the student row is only used for the leaving date and fallbacks).
  let student = null;
  const stu = ss.getSheetByName('Student_Data');
  if (stu && stu.getLastRow() > 1) {
    const sh = stu.getRange(1, 1, 1, stu.getLastColumn()).getValues()[0];
    const sIdx = sh.indexOf('Std_ID');
    const rIdx = sh.indexOf('REC_ID');
    const all = stu.getRange(2, 1, stu.getLastRow() - 1, stu.getLastColumn()).getValues();
    for (let i = 0; i < all.length; i++) {
      if (sIdx !== -1 && String(all[i][sIdx] || '').trim() === String(tc.Std_ID || '').trim() &&
          (rIdx === -1 || !tc.REC_ID || String(all[i][rIdx] || '').trim() === String(tc.REC_ID).trim())) {
        student = {};
        sh.forEach(function (h, j) {
          const v = all[i][j];
          student[h] = v instanceof Date ? Utilities.formatDate(v, TC_TZ, 'yyyy-MM-dd') : (v !== null && v !== undefined ? v.toString() : '');
        });
        break;
      }
    }
  }
  const rec = getRecInfo_(tc.REC_ID);
  // Attendance history: the term resolution below (stored Term_ID when it
  // still covers the class-section, else the term covering the leaving
  // date) is unchanged — it feeds the certificate heading only.
  const classSection = String(tc.Class || '').trim() + (String(tc.Section || '').trim() ? '-' + String(tc.Section).trim() : '');
  const leaveDate = parseTermDate_(student && student.Inactive_Date ? student.Inactive_Date : tc.Issue_Date) || new Date();
  let term = null;
  if (tc.Term_ID) {
    const rows = getClassTermsForRec_(tc.REC_ID);
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (!isActiveFlag_(r.Is_Active)) continue;
      if (String(r.Term_ID || '').trim() !== String(tc.Term_ID).trim()) continue;
      if (!classCellMatches_(String(r.Class_Section || ''), classSection)) continue;
      const s = parseTermDate_(r.Start_Date);
      const e = parseTermDate_(r.End_Date);
      if (!s || !e) continue;
      term = { termId: String(r.Term_ID || '').trim(), start: s, end: e };
      break;
    }
  }
  if (!term) term = resolveTermForLeaving_(tc.REC_ID, classSection, leaveDate);
  let termStart = term ? term.start : null;
  let termEnd = term ? term.end : null;
  if (!termStart) {
    // Fallback: September of the academic year containing the leaving date.
    const y = leaveDate.getMonth() + 1 >= 9 ? leaveDate.getFullYear() : leaveDate.getFullYear() - 1;
    termStart = new Date(y, 8, 1);
    termEnd = leaveDate;
  }
  if (termEnd > leaveDate) termEnd = leaveDate;
  const joinDate = parseTermDate_(student && student.Date_of_Joining ? student.Date_of_Joining : null);
  // v7.18: history via the shared helper (also used by the popup preview).
  const history = buildTcAttendanceHistory_(ss, tc.Std_ID, tc.REC_ID, classSection, joinDate, leaveDate);
  const attendance = history.attendance;
  const lastDayOfAttendance = history.lastDayOfAttendance;
  // v7.15: the reason's Category (for the shifting block). An old row
  // whose reason is not in the table gets '' — the client falls back to
  // Destination_Type + New_Address for those.
  let reasonCategory = '';
  try {
    const rr = getLeavingReasonRow_(ss, tc.Leaving_Reason);
    if (rr) reasonCategory = String(rr.Category || '').trim();
  } catch (e) {}
  let destinationName = String(tc.Destination || '');
  if (String(tc.Destination_Type || '') === 'REC' && destinationName) {
    const found = getRecNamesList_().filter(function (r) { return String(r.REC_Code || '').trim() === destinationName; })[0];
    if (found) destinationName = found.REC_Name;
    // v7.15: a typed 'REC not in list' name is stored in
    // Destination_REC_Name (Destination holds the same typed name).
    else if (tc.Destination_REC_Name) destinationName = String(tc.Destination_REC_Name);
  }
  return {
    tc: tc,
    student: student,
    recName: rec.REC_Name || tc.REC_Name || '',
    lb: rec.LB || tc.LB || '',
    termId: term ? term.termId : String(tc.Term_ID || ''),
    termStart: termStart ? termDateKey_(termStart) : '',
    termEnd: termEnd ? termDateKey_(termEnd) : '',
    attendance: attendance,
    // v7.14: the student's Unique_ID (B-Form last 6 digits), from the
    // student row — the TC snapshot has no Unique_ID column.
    uniqueId: student ? String(student.Unique_ID || '') : '',
    registrationNumber: String(tc.REC_ID || '') + '-' + String(tc.Std_ID || ''),
    destinationName: destinationName,
    // v7.15: last Present day on/before leaving ('' when none), and the
    // reason's Category for the shifting block.
    lastDayOfAttendance: lastDayOfAttendance,
    reasonCategory: reasonCategory
  };
}

// v7.18: attendance preview for the Deactivate popup — the history that
// WOULD print on the certificate if this student were deactivated today.
// READ-ONLY: it writes nothing — no TC row, no student change, no sheet
// creation. Admin-gated like the other TC functions; a teacher calling it
// directly gets {success:false}. The leaving month is the current month;
// the history is built by the same helper as the printed certificate.
function getTcAttendancePreview(stdId, recId, username) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gate = requireAdminCaller_(ss, username);
  if (!gate.ok) return gate.error;
  const sid = String(stdId || '').trim();
  const rid = String(recId || '').trim();
  if (!sid) return { success: false, message: 'Student not found.' };
  let student = null;
  const stu = ss.getSheetByName('Student_Data');
  if (stu && stu.getLastRow() > 1) {
    const sh = stu.getRange(1, 1, 1, stu.getLastColumn()).getValues()[0];
    const sIdx = sh.indexOf('Std_ID');
    const rIdx = sh.indexOf('REC_ID');
    const all = stu.getRange(2, 1, stu.getLastRow() - 1, stu.getLastColumn()).getValues();
    for (let i = 0; i < all.length; i++) {
      if (sIdx !== -1 && String(all[i][sIdx] || '').trim() === sid &&
          (rIdx === -1 || !rid || String(all[i][rIdx] || '').trim() === rid)) {
        student = {};
        sh.forEach(function (h, j) {
          const v = all[i][j];
          student[h] = v instanceof Date ? Utilities.formatDate(v, TC_TZ, 'yyyy-MM-dd') : (v !== null && v !== undefined ? v.toString() : '');
        });
        break;
      }
    }
  }
  if (!student) return { success: false, message: 'Student not found.' };
  const classSection = String(student.Student_Class || '').trim() + (String(student.Student_Section || '').trim() ? '-' + String(student.Student_Section).trim() : '');
  const joinDate = parseTermDate_(student.Date_of_Joining ? student.Date_of_Joining : null);
  const leaveDate = new Date();
  const history = buildTcAttendanceHistory_(ss, sid, rid, classSection, joinDate, leaveDate);
  let totalWorking = 0, totalPresent = 0;
  history.attendance.forEach(function (a) { totalWorking += Number(a.workingDays) || 0; totalPresent += Number(a.daysPresent) || 0; });
  return {
    success: true,
    stdId: sid,
    recId: rid,
    leaveMonth: leaveDate.getFullYear() + '-' + String(leaveDate.getMonth() + 1).padStart(2, '0'),
    attendance: history.attendance,
    lastDayOfAttendance: history.lastDayOfAttendance,
    totalWorking: totalWorking,
    totalPresent: totalPresent
  };
}

// Get data from a specific sheet
function getSheetData(sheetName, recId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = resolveSheet_(ss, sheetName);
    // v7.15: a managed sheet that does not exist yet reads as empty —
    // it is created + seeded on first access (popup load or Add), never here.
    if (!sheet) {
      if (isLeavingReasonsSheet_(sheetName) || isCountriesSheet_(sheetName)) return [];
      throw new Error(sheetName + ' sheet not found');
    }
if (isPrevPermSheet_(sheetName)) ensurePrevPermColumns_(sheet);    
    if (isHoursConfigSheet_(sheetName)) ensureHoursConfigColumns_(sheet, sheetName);
    if (isRecNamesSheet_(sheetName)) ensureRecNamesActiveColumn_(sheet);
    // v7.10: Error rows have a blank TC_Number, so the generic first-column
    // check below would drop them — read through the dedicated function.
    if (isTransferCertificatesSheet_(sheetName)) return getTransferCertificates(recId);
    
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

// ============ v7.9 VALIDATION HELPERS ============
// Class_Mapping rows for one school, read directly (no ensureClassMappingSheet_
// call — that helper migrates Class_Sections when Class_Mapping is missing,
// which a validation must never trigger as a side effect). Returns null when
// the sheet cannot be read (callers fail open), else [{Class, Sections}].
function getMappingRowsForRec_(ss, recId) {
  try {
    const sh = ss.getSheetByName('Class_Mapping');
    if (!sh) return null;
    const data = sh.getDataRange().getValues();
    if (!data.length) return [];
    const hd = data[0] || [];
    const clsI = hd.indexOf('Class');
    const secI = hd.indexOf('Sections');
    const recI = hd.indexOf('REC_ID');
    if (clsI === -1 || secI === -1) return null;
    const out = [];
    for (let i = 1; i < data.length; i++) {
      const rowRec = recI !== -1 ? String(data[i][recI] || '').trim() : '';
      if (recId && rowRec !== recId) continue;
      const cls = String(data[i][clsI] || '').trim();
      if (!cls) continue;
      out.push({ Class: cls, Sections: String(data[i][secI] || '').trim() });
    }
    return out;
  } catch (e) { return null; }
}

// Class_Sections combos for one school, read directly. Returns null when the
// sheet cannot be read (callers fail open), else a list of 'VII-A' strings.
// A row with a blank REC_ID counts for any school (legacy tolerance).
function getClassSectionsForRec_(ss, recId) {
  try {
    const sh = ss.getSheetByName('Class_Sections');
    if (!sh) return null;
    const data = sh.getDataRange().getValues();
    if (!data.length) return [];
    const hd = data[0] || [];
    const rIdx = hd.indexOf('REC_ID');
    const scopeRec = String(recId || '').trim();
    const out = [];
    for (let i = 1; i < data.length; i++) {
      const combo = String(data[i][0] || '').trim();
      if (!combo) continue;
      const rowRec = rIdx !== -1 ? String(data[i][rIdx] || '').trim() : '';
      if (scopeRec && rowRec && rowRec !== scopeRec) continue;
      out.push(combo);
    }
    return out;
  } catch (e) { return null; }
}

// Class-sections assigned to one teacher (Teacher_Assignments rows store
// Class_Section comma-joined). [] when the sheet cannot be read or the
// teacher has no assignments — callers treat an empty list as no constraint,
// matching the portal form's fallback to all sections.
function getAssignedClassesForTeacher_(ss, username, recId) {
  try {
    const sh = ss.getSheetByName('Teacher_Assignments');
    if (!sh) return [];
    const data = sh.getDataRange().getValues();
    if (!data.length) return [];
    const hd = data[0] || [];
    const rIdx = hd.indexOf('REC_ID');
    const want = String(username || '').trim();
    const scopeRec = String(recId || '').trim();
    const out = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() !== want) continue;
      const rowRec = rIdx !== -1 ? String(data[i][rIdx] || '').trim() : '';
      if (scopeRec && rowRec && rowRec !== scopeRec) continue;
      String(data[i][1] || '').split(',').forEach(function (s) {
        const v = s.trim();
        if (v && out.indexOf(v) === -1) out.push(v);
      });
    }
    return out;
  } catch (e) { return []; }
}

// 'VII-A' -> {cls:'VII', sec:'A'}; anything else -> null.
function parseClassSectionCombo_(combo) {
  const m = /^([A-Za-z]+)-([A-Za-z])$/.exec(String(combo == null ? '' : combo).trim());
  if (!m) return null;
  return { cls: m[1].toUpperCase(), sec: m[2].toUpperCase() };
}

// Validate one Class_Sections value: a real class + a single section letter,
// and the class must be set up in Class_Mapping for the school when the
// mapping can be read. Returns an error string, or null when OK.
function checkClassSectionCombo_(combo, ss, recId) {
  const val = String(combo == null ? '' : combo).trim();
  if (!val) return 'Class_Section is required (e.g. VII-A).';
  const parsed = parseClassSectionCombo_(val);
  if (!parsed) return 'Class_Section must be a class and a single section letter (e.g. VII-A).';
  if (MASTER_CLASSES.indexOf(parsed.cls) === -1) return 'Unknown class "' + parsed.cls + '". Use a class from the master list.';
  const mapping = getMappingRowsForRec_(ss, recId);
  if (mapping && mapping.length > 0) {
    const names = mapping.map(function (x) { return String(x.Class || '').trim(); });
    if (names.indexOf(parsed.cls) === -1) return 'Class "' + parsed.cls + '" is not set up in Class Mappings for this school.';
  }
  return null;
}

// True when another row already has this Class_Section for the same school.
// A blank REC_ID on either side counts as the same school (legacy rows).
function classSectionDuplicateExists_(sheet, headers, combo, recId, excludeRow) {
  const csIdx = headers.indexOf('Class_Section');
  const rIdx = headers.indexOf('REC_ID');
  if (csIdx === -1) return false;
  const want = String(combo || '').trim().toUpperCase();
  if (!want) return false;
  const scopeRec = String(recId || '').trim();
  const all = sheet.getDataRange().getValues();
  for (let i = 1; i < all.length; i++) {
    if (excludeRow && (i + 1) === excludeRow) continue;
    if (String(all[i][csIdx] || '').trim().toUpperCase() !== want) continue;
    const exRec = rIdx !== -1 ? String(all[i][rIdx] || '').trim() : '';
    if (scopeRec && exRec && exRec !== scopeRec) continue;
    return true;
  }
  return false;
}

// True when another Teacher_Assignments row already has this Username for
// the same school. A blank REC_ID on either side counts as the same school.
function teacherAssignmentDuplicateExists_(sheet, username, recId, excludeRow) {
  const want = String(username || '').trim();
  if (!want) return false;
  const data = sheet.getDataRange().getValues();
  if (!data.length) return false;
  const hd = data[0] || [];
  const rIdx = hd.indexOf('REC_ID');
  const scopeRec = String(recId || '').trim();
  for (let i = 1; i < data.length; i++) {
    if (excludeRow && (i + 1) === excludeRow) continue;
    if (String(data[i][0] || '').trim() !== want) continue;
    const exRec = rIdx !== -1 ? String(data[i][rIdx] || '').trim() : '';
    if (scopeRec && exRec && exRec !== scopeRec) continue;
    return true;
  }
  return false;
}

// True when another User_Permissions row already has this Username.
// Permissions are keyed by username alone (no school scoping).
function userPermissionsDuplicateExists_(sheet, headers, username, excludeRow) {
  const uIdx = headers.indexOf('Username');
  if (uIdx === -1) return false;
  const want = String(username || '').trim();
  if (!want) return false;
  const all = sheet.getDataRange().getValues();
  for (let i = 1; i < all.length; i++) {
    if (excludeRow && (i + 1) === excludeRow) continue;
    if (String(all[i][uIdx] || '').trim() === want) return true;
  }
  return false;
}

// Previous Date Permissions Edit checks (v7.9 fix 7): real date, not future,
// non-empty password, a changed username must be a teacher of the caller's
// school, a changed class must be one of that teacher's assigned classes,
// plus the duplicate check on the effective values. An unchanged username
// or class is grandfathered. Returns an error string, or null when OK.
function validatePrevPermEdit_(ss, sheet, headers, rowData, rowIndex, caller) {
  const stored = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  const uIdx = headers.indexOf('Username');
  const dIdx = headers.indexOf('Allowed_Date');
  const pIdx = headers.indexOf('Password');
  const cIdx = headers.indexOf('Class_Section');
  const rawVal = function (field, idx) {
    if (rowData[field] !== undefined && rowData[field] !== null) return rowData[field];
    return idx !== -1 ? stored[idx] : '';
  };
  const effU = String(rawVal('Username', uIdx) || '').trim();
  const effD = rawVal('Allowed_Date', dIdx);
  const effP = String(rawVal('Password', pIdx) || '').trim();
  const effC = String(rawVal('Class_Section', cIdx) || '').trim();
  const storedU = uIdx !== -1 ? String(stored[uIdx] || '').trim() : '';
  const storedC = cIdx !== -1 ? String(stored[cIdx] || '').trim() : '';
  if (!effU) return 'Username is required.';
  if (!isRealDateString_(effD)) return 'Allowed_Date is not a real date.';
  const allowedDate = parseTermDate_(effD);
  const today = new Date(); today.setHours(0, 0, 0, 0);
  if (allowedDate && allowedDate.getTime() > today.getTime()) return 'Allowed Date cannot be in the future.';
  if (!effP) return 'Password is required.';
  if (effU !== storedU) {
    let teacherOk = false; let teacherReadable = false;
    try {
      const us = ss.getSheetByName('Users');
      if (us) {
        teacherReadable = true;
        const ud = us.getDataRange().getValues();
        const uh = ud[0] || [];
        const tI = uh.indexOf('UserType');
        const rI = uh.indexOf('REC_ID');
        const callerRec = caller && caller.recId ? String(caller.recId).trim() : '';
        const isSuper = !!(caller && caller.userType === 'superadmin');
        for (let i = 1; i < ud.length; i++) {
          if (String(ud[i][0] || '').trim() !== effU) continue;
          const ut = tI !== -1 ? String(ud[i][tI] || '').trim() : '';
          const ur = rI !== -1 ? String(ud[i][rI] || '').trim() : '';
          if (ut === 'teacher' && (isSuper || !callerRec || !ur || ur === callerRec)) { teacherOk = true; break; }
        }
      }
    } catch (e) { teacherReadable = false; }
    if (teacherReadable && !teacherOk) return 'Username "' + effU + '" is not a teacher of this school.';
  }
  if (effC && effC !== storedC) {
    const callerRec = caller && caller.recId ? String(caller.recId).trim() : '';
    const assigned = getAssignedClassesForTeacher_(ss, effU, callerRec);
    if (assigned.length > 0 && assigned.indexOf(effC) === -1) {
      return 'Class "' + effC + '" is not assigned to ' + effU + '. Assigned classes: ' + assigned.join(', ') + '.';
    }
  }
  if (prevPermDuplicateExists_(sheet, headers, { Username: effU, Allowed_Date: effD, Class_Section: effC }, rowIndex)) {
    return 'This user already has permission for this date for this class.';
  }
  return null;
}

// Teacher_Assignments Edit checks (v7.9 fix 8): a changed username must
// exist in Users of the caller's school; changed class-sections must exist
// in Class_Sections for the school; plus the duplicate check (own row
// excluded). Unchanged values are grandfathered. Returns an error string,
// or null when OK.
function validateTeacherAssignmentEdit_(ss, sheet, headers, rowData, rowIndex, caller) {
  const stored = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  const uIdx = headers.indexOf('Username');
  const cIdx = headers.indexOf('Class_Section');
  const rIdx = headers.indexOf('REC_ID');
  const storedU = uIdx !== -1 ? String(stored[uIdx] || '').trim() : '';
  const storedC = cIdx !== -1 ? String(stored[cIdx] || '').trim() : '';
  const effU = rowData.Username !== undefined && rowData.Username !== null && String(rowData.Username).trim() !== '' ? String(rowData.Username).trim() : storedU;
  const effC = rowData.Class_Section !== undefined && rowData.Class_Section !== null ? String(rowData.Class_Section).trim() : storedC;
  const effRec = rowData.REC_ID !== undefined && rowData.REC_ID !== null && String(rowData.REC_ID).trim() !== '' ? String(rowData.REC_ID).trim() : (rIdx !== -1 ? String(stored[rIdx] || '').trim() : '');
  if (!effU) return 'Username is required.';
  if (!effC) return 'Class_Section is required.';
  if (effU !== storedU) {
    let userOk = false; let userReadable = false;
    try {
      const us = ss.getSheetByName('Users');
      if (us) {
        userReadable = true;
        const ud = us.getDataRange().getValues();
        const uh = ud[0] || [];
        const rI = uh.indexOf('REC_ID');
        const callerRec = caller && caller.recId ? String(caller.recId).trim() : '';
        const isSuper = !!(caller && caller.userType === 'superadmin');
        for (let i = 1; i < ud.length; i++) {
          if (String(ud[i][0] || '').trim() !== effU) continue;
          const ur = rI !== -1 ? String(ud[i][rI] || '').trim() : '';
          if (isSuper || !callerRec || !ur || ur === callerRec) { userOk = true; break; }
        }
      }
    } catch (e) { userReadable = false; }
    if (userReadable && !userOk) return 'Username "' + effU + '" does not exist for this school.';
  }
  const effTokens = effC.split(',').map(function (s) { return s.trim(); }).filter(Boolean);
  const storedTokens = storedC.split(',').map(function (s) { return s.trim(); }).filter(Boolean);
  const newTokens = effTokens.filter(function (t) { return storedTokens.indexOf(t) === -1; });
  if (newTokens.length > 0) {
    const sections = getClassSectionsForRec_(ss, effRec);
    if (sections) {
      for (let ti = 0; ti < newTokens.length; ti++) {
        if (sections.indexOf(newTokens[ti]) === -1) {
          return 'Class_Section "' + newTokens[ti] + '" does not exist for this school.';
        }
      }
    }
  }
  if (teacherAssignmentDuplicateExists_(sheet, effU, effRec, rowIndex)) {
    return 'Username ' + effU + ' already has assigned classes. Please edit the existing record.';
  }
  return null;
}


// ============ PREVIOUS_DATE_PERMISSIONS AUTO-SETUP ============
const PREV_PERM_SHEET = 'Previous_Date_Permissions';
const PREV_PERM_NEW_COLUMNS = ['Attendance_Type', 'Granted_At', 'Expires_At', 'Class_Section'];

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

// Previous Date Permissions: a saved row only blocks a new/edited permission while it is LIVE.
// Live = effective expiry (Expires_At, or Granted_At + 24h when Expires_At is blank/unparseable)
// is still in the future. A row with neither a usable Expires_At nor a usable Granted_At is
// NOT live. Expired rows are dead and never block.
/** Effective expiry: Expires_At ho to wohi; warna Granted_At + 24h; dono na hon to null. */
function prevPermEffectiveExpiry_(expiresVal, grantedVal) {
  if (expiresVal !== null && expiresVal !== undefined && String(expiresVal).trim() !== '') {
    const exp = expiresVal instanceof Date ? expiresVal : new Date(expiresVal);
    if (!isNaN(exp.getTime())) return exp;
  }
  if (grantedVal !== null && grantedVal !== undefined && String(grantedVal).trim() !== '') {
    const g = grantedVal instanceof Date ? grantedVal : new Date(grantedVal);
    if (!isNaN(g.getTime())) return new Date(g.getTime() + 24 * 60 * 60 * 1000);
  }
  return null;
}

function isPrevPermRowLive_(expiresVal, grantedVal) {
  const exp = prevPermEffectiveExpiry_(expiresVal, grantedVal);
  if (!exp) return false;
  return exp.getTime() > new Date().getTime();
}

// True when a LIVE permission already exists for the same Username + Allowed_Date + Class_Section.
// excludeRow = 1-based sheet row to skip (the row being edited); 0 = exclude nothing.
// Blank Class_Section equals blank (old unscoped rows match only other blank rows).
function prevPermDuplicateExists_(sheet, headers, rowData, excludeRow) {
  const uIdx = headers.indexOf('Username');
  const dIdx = headers.indexOf('Allowed_Date');
  const cIdx = headers.indexOf('Class_Section');
  const eIdx = headers.indexOf('Expires_At');
  const gIdx = headers.indexOf('Granted_At');
  if (uIdx === -1 || dIdx === -1) return false;
  if (!rowData.Username || !rowData.Allowed_Date) return false;
  const newDate = new Date(rowData.Allowed_Date);
  const newDateStr = !isNaN(newDate) ? newDate.toDateString() : String(rowData.Allowed_Date).trim();
  const newCls = String(rowData.Class_Section || '').trim();
  const all = sheet.getDataRange().getValues();
  return all.some(function (r, i) {
    if (i === 0) return false;
    if (excludeRow && (i + 1) === excludeRow) return false;
    if (String(r[uIdx] || '').trim() !== String(rowData.Username).trim()) return false;
    const d = r[dIdx] instanceof Date ? r[dIdx] : new Date(r[dIdx]);
    const dStr = !isNaN(d) ? d.toDateString() : String(r[dIdx] || '').trim();
    if (dStr !== newDateStr) return false;
    const cls = cIdx !== -1 ? String(r[cIdx] || '').trim() : '';
    if (cls !== newCls) return false;
    return isPrevPermRowLive_(eIdx !== -1 ? r[eIdx] : '', gIdx !== -1 ? r[gIdx] : '');
  });
}


// ============ OUTAGE PERMISSIONS (v7.23) ============
// After an outage day, a designated admin grants previous-date access in
// bulk from the portal's Outage Permissions tab: ONE date, several
// teachers, several class sections, one shared password, and a
// 24/48/72-hour window. This function writes ORDINARY
// Previous_Date_Permissions rows -- one per teacher x class -- in exactly
// the shape the single-row Add writes, so the marking app honours them
// unchanged. A (teacher, class) pair is written only when the class is
// among that teacher's Teacher_Assignments at the caller's school and no
// LIVE permission already covers it; every skip is reported back in its
// own bucket so the portal can show the admin exactly what happened.
function grantOutagePermissions(payload, currentUsername) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('No active spreadsheet found');

    // Gate: admin/superadmin only -- the same gate as every other write.
    const gate = requireAdminCaller_(ss, currentUsername);
    if (!gate.ok) return gate.error;
    const caller = gate.caller;

    // The caller's school comes from their own Users row. A caller with no
    // school cannot grant for anyone (standing rule, same as addStudent).
    const callerRec = String(caller.recId || '').trim();
    if (!callerRec) {
      return { success: false, message: 'Your user account has no school (REC_ID) assigned, so you cannot grant outage permissions.' };
    }

    // Designated-person check: a superadmin may always grant; anyone else
    // must carry the outage_permissions tab in their own allowedTabs.
    if (caller.userType !== 'superadmin') {
      const usersSheet0 = ss.getSheetByName('Users');
      const uData0 = usersSheet0 ? usersSheet0.getDataRange().getValues() : [];
      const uHead0 = uData0[0] || [];
      const tabsCol = uHead0.indexOf('allowedTabs');
      let designated = false;
      for (let i = 1; i < uData0.length; i++) {
        if (String(uData0[i][0] || '').trim() === caller.username) {
          const rawTabs = tabsCol !== -1 ? uData0[i][tabsCol] : '';
          const tabList = Array.isArray(rawTabs)
            ? rawTabs.map(function (x) { return String(x || '').trim(); })
            : String(rawTabs || '').split(',').map(function (x) { return x.trim(); });
          designated = tabList.indexOf('outage_permissions') !== -1;
          break;
        }
      }
      if (!designated) {
        return { success: false, message: 'You are not designated to grant outage permissions. Please ask an admin to add the Outage Permissions tab to your allowed tabs.' };
      }
    }

    payload = payload || {};

    // Date: required, a real date, not in the future, at most 7 days back.
    // Today IS allowed -- an admin may prepare in the evening for the
    // next morning.
    const dateRaw = String(payload.date || '').trim();
    if (!dateRaw) return { success: false, message: 'Please select the outage date.' };
    const parts = dateRaw.split('-');
    if (parts.length !== 3) return { success: false, message: 'The outage date is not a valid date.' };
    const dy = parseInt(parts[0], 10);
    const dm = parseInt(parts[1], 10);
    const dd = parseInt(parts[2], 10);
    const allowed = new Date(dy, dm - 1, dd);
    if (isNaN(allowed.getTime()) || allowed.getFullYear() !== dy || allowed.getMonth() !== dm - 1 || allowed.getDate() !== dd) {
      return { success: false, message: 'The outage date is not a valid date.' };
    }
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    if (allowed.getTime() > today.getTime()) {
      return { success: false, message: 'The outage date cannot be in the future.' };
    }
    const earliest = new Date(today);
    earliest.setDate(earliest.getDate() - 7);
    if (allowed.getTime() < earliest.getTime()) {
      return { success: false, message: 'The outage date cannot be more than 7 days back.' };
    }
    const pad2 = function (n) { return ('0' + n).slice(-2); };
    const allowedStr = dy + '-' + pad2(dm) + '-' + pad2(dd);
    const dayName = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'][allowed.getDay()];

    // Selections: at least one teacher and one class section.
    const dedupe = function (arr) {
      const out = [];
      (Array.isArray(arr) ? arr : []).forEach(function (x) {
        const v = String(x || '').trim();
        if (v && out.indexOf(v) === -1) out.push(v);
      });
      return out;
    };
    const usernames = dedupe(payload.usernames);
    if (!usernames.length) return { success: false, message: 'Please select at least one teacher.' };
    const classSections = dedupe(payload.classSections);
    if (!classSections.length) return { success: false, message: 'Please select at least one class section.' };

    // One shared password, stored on every row exactly as the single-row
    // Add stores it.
    const password = String(payload.password || '').trim();
    if (!password) return { success: false, message: 'Please enter the password the teachers will type for this date.' };

    // Window + type. The type normalisation mirrors
    // stampPrevPermTimestamps_ (Both / Normal / Episodic).
    let hours = parseInt(payload.hours, 10);
    if ([24, 48, 72].indexOf(hours) === -1) hours = 48;
    const typeRaw = String(payload.attendanceType || 'Both').trim().toLowerCase();
    const attendanceType = typeRaw === 'episodic' ? 'Episodic' : typeRaw === 'normal' ? 'Normal' : 'Both';

    const sheet = ss.getSheetByName(PREV_PERM_SHEET);
    if (!sheet) throw new Error(PREV_PERM_SHEET + ' sheet not found');
    ensurePrevPermColumns_(sheet);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Teachers of the caller's school: Users rows with UserType 'teacher'
    // whose REC_ID is the caller's school.
    const usersSheet = ss.getSheetByName('Users');
    const uData = usersSheet ? usersSheet.getDataRange().getValues() : [];
    const uHead = uData[0] || [];
    const uTypeCol = uHead.indexOf('UserType');
    const uRecCol = uHead.indexOf('REC_ID');
    const teacherSet = {};
    for (let i = 1; i < uData.length; i++) {
      const uname = String(uData[i][0] || '').trim();
      if (!uname) continue;
      const utype = uTypeCol !== -1 ? String(uData[i][uTypeCol] || '').trim() : '';
      const urec = uRecCol !== -1 ? String(uData[i][uRecCol] || '').trim() : '';
      if (utype === 'teacher' && urec === callerRec) teacherSet[uname] = true;
    }

    const now = new Date();
    const expiresAt = new Date(now.getTime() + hours * 60 * 60 * 1000);

    const createdPairs = [];
    const skippedDuplicate = [];
    const skippedNotAssigned = [];
    const skippedUnknownUser = [];

    usernames.forEach(function (uname) {
      if (!teacherSet[uname]) { skippedUnknownUser.push(uname); return; }
      const assigned = getAssignedClassesForTeacher_(ss, uname, callerRec);
      classSections.forEach(function (cs) {
        const pair = uname + ' \u00B7 ' + cs;
        if (assigned.indexOf(cs) === -1) { skippedNotAssigned.push(pair); return; }
        if (prevPermDuplicateExists_(sheet, headers, { Username: uname, Allowed_Date: allowedStr, Class_Section: cs }, 0)) {
          skippedDuplicate.push(pair);
          return;
        }
        const rowData = {
          Username: uname,
          Allowed_Date: allowedStr,
          Day: dayName,
          Password: password,
          Attendance_Type: attendanceType,
          Class_Section: cs,
          Granted_At: now,
          Expires_At: expiresAt,
          REC_ID: callerRec
        };
        const newRow = headers.map(function (header) {
          const v = rowData[String(header).trim()];
          return v !== undefined && v !== null ? v : '';
        });
        sheet.appendRow(newRow);
        createdPairs.push(pair);
      });
    });
    SpreadsheetApp.flush();

    const summary = [];
    summary.push('Created ' + createdPairs.length + ' permission' + (createdPairs.length === 1 ? '' : 's') + '.');
    if (skippedDuplicate.length) summary.push('Skipped ' + skippedDuplicate.length + ' \u2014 already had access for this date.');
    if (skippedNotAssigned.length) summary.push('Skipped ' + skippedNotAssigned.length + ' \u2014 class not assigned to that teacher.');
    if (skippedUnknownUser.length) summary.push('Skipped ' + skippedUnknownUser.length + ' \u2014 teacher not found at your school.');

    return {
      success: true,
      created: createdPairs.length,
      createdPairs: createdPairs,
      skippedDuplicate: skippedDuplicate,
      skippedNotAssigned: skippedNotAssigned,
      skippedUnknownUser: skippedUnknownUser,
      message: summary.join(' ')
    };
  } catch (e) {
    Logger.log('Error in grantOutagePermissions: ' + e.message);
    return { success: false, message: 'Failed to grant outage permissions: ' + e.message };
  }
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

// v7.9: day key for comparing two date values that may be stored in
// different shapes (a Date object, 'yyyy-MM-dd' or 'dd-MMM-yyyy').
// '' when the value is not a real date.
function dateKeyForCompare_(v) {
  const d = parseTermDate_(v);
  return d ? termDateKey_(d) : '';
}

// v7.9: strict real-date check for a yyyy-MM-dd or dd-MMM-yyyy string
// (or a Date object). A rolled-over date such as 2026-02-31 is refused —
// parseTermDate_ alone lets it roll into March, so the parsed parts are
// compared back against the input.
function isRealDateString_(v) {
  if (v == null) return false;
  if (v instanceof Date) return !isNaN(v.getTime());
  const str = String(v).trim();
  if (!str) return false;
  const iso = /^(\d{4})-(\d{2})-(\d{2})$/.exec(str);
  if (iso) {
    const d = new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
    return d.getFullYear() === Number(iso[1]) && d.getMonth() === Number(iso[2]) - 1 && d.getDate() === Number(iso[3]);
  }
  const dmy = /^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/.exec(str);
  if (!dmy) return false;
  const d = parseTermDate_(str);
  if (!d) return false;
  const monMap = { Jan: 0, Feb: 1, Mar: 2, Apr: 3, May: 4, Jun: 5, Jul: 6, Aug: 7, Sep: 8, Oct: 9, Nov: 10, Dec: 11 };
  const monKey = dmy[2].charAt(0).toUpperCase() + dmy[2].slice(1).toLowerCase();
  const mi = monMap[monKey];
  return mi !== undefined && d.getDate() === Number(dmy[1]) && d.getMonth() === mi && d.getFullYear() === Number(dmy[3]);
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
  const termName = String(rowData.Term_Name == null ? '' : rowData.Term_Name).trim();
  if (!termName) return 'Term_Name is required.';
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
function validateClassTermsRow_(sheet, headers, rowData, recId, excludeRow, caller) {
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
  // v7.9 (fix 10): the row's own REC_ID must be the caller's school
  // (superadmin skipped; fail-open when the caller's school is unknown).
  const ctCallerRec = caller && caller.recId ? String(caller.recId).trim() : '';
  const ctIsSuper = !!(caller && caller.userType === 'superadmin');
  const ctRowRec = String(rowData.REC_ID || '').trim();
  if (!ctIsSuper && ctCallerRec && ctRowRec && ctRowRec !== ctCallerRec) {
    return 'This record belongs to another school. You can only manage Class_Terms for your own school.';
  }
  const csIdx = headers.indexOf('Class_Section');
  const termIdx = headers.indexOf('Term_ID');
  const recIdx = headers.indexOf('REC_ID');
  const sIdx = headers.indexOf('Start_Date');
  const eIdx = headers.indexOf('End_Date');
  const actIdx = headers.indexOf('Is_Active');
  if (csIdx === -1 || termIdx === -1) return null;
  const newCell = String(rowData.Class_Section || '').trim();
  const newTokens = expandClassCell_(newCell);
  // v7.9 (fix 10): every class token in Class_Section must be a real class.
  if (newCell) {
    for (let ti = 0; ti < newTokens.length; ti++) {
      const tokBase = baseClassOf_(newTokens[ti]);
      if (MASTER_CLASSES.indexOf(tokBase) === -1) {
        return 'Unknown class "' + newTokens[ti] + '" in Class_Section. Use a class from the master list.';
      }
    }
  }
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
// ============ HOURS CONFIG TABLES (v6.9) ============
// Hours_Setup sheet: Hours, Attendance_Type, Max_Custom, REC_ID (+ Is_Active auto-appended).
// Hour_Categories sheet: Category_ID, Category_Name, Is_Active, Attendance_Type, REC_ID, Applicable_Classes.
// Rows are immutable once added — only Is_Active and Applicable_Classes may change (v7.0;
// enforced in updateSheetRow; Delete is hidden in the UI and blocked in deleteSheetRow).
// Is_Active defaults TRUE on Add. Applicable_Classes blank = all classes.
const HOURS_SETUP_SHEET = 'Hours_Setup';
const HOUR_CATEGORIES_SHEET = 'Hour_Categories';
const HOURS_SETUP_COLUMNS = ['Hours', 'Attendance_Type', 'Max_Custom', 'REC_ID', 'Is_Active'];
const HOUR_CATEGORIES_COLUMNS = ['Category_ID', 'Category_Name', 'Is_Active', 'Attendance_Type', 'REC_ID', 'Applicable_Classes'];
const HOURS_CONFIG_IMMUTABLE_MSG = 'Rows cannot be amended once added — only Active/Inactive and Applicable Classes can change.';

function isHoursSetupSheet_(sheetName) {
  return String(sheetName || '').trim().toLowerCase() === 'hours_setup';
}

function isHourCategoriesSheet_(sheetName) {
  return String(sheetName || '').trim().toLowerCase() === 'hour_categories';
}

function isHoursConfigSheet_(sheetName) {
  return isHoursSetupSheet_(sheetName) || isHourCategoriesSheet_(sheetName);
}

// Append any missing columns at the end of the sheet (Is_Active is new for Hours_Setup).
// Never creates the sheet itself — callers throw 'sheet not found' when it is missing.
function ensureHoursConfigColumns_(sheet, sheetName) {
  if (!sheet) return [];
  const wanted = isHoursSetupSheet_(sheetName) ? HOURS_SETUP_COLUMNS : HOUR_CATEGORIES_COLUMNS;
  let lastCol = sheet.getLastColumn();
  if (lastCol < 1) return [];
  let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const missing = wanted.filter(c => headers.indexOf(c) === -1);
  if (missing.length === 0) return headers;
  sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
  SpreadsheetApp.flush();
  lastCol = sheet.getLastColumn();
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
}

// Normalise an Attendance_Type value to lowercase 'normal' / 'episodic' / 'both'; '' when
// invalid. Lowercase (not Title case like prevPerm) because this sheet's existing rows
// and the form options are lowercase — so the value round-trips through the Edit form.
function normHoursConfigType_(v) {
  const t = String(v == null ? '' : v).trim().toLowerCase();
  if (t === 'normal') return 'normal';
  if (t === 'episodic') return 'episodic';
  if (t === 'both') return 'both';
  return '';
}

// Normalise an Hours value for comparison: numbers via String(Number(v)) so 2, '2' and
// '2.0' all match; 'custom' lowercased; anything else lowercased + trimmed.
function normHoursValue_(v) {
  const s = String(v == null ? '' : v).trim();
  if (s.toLowerCase() === 'custom') return 'custom';
  const n = Number(s);
  if (s !== '' && !isNaN(n)) return String(n);
  return s.toLowerCase();
}

// Validate a Hours_Setup / Hour_Categories row on Add (excludeRow = 0 on Add).
// Also normalises Attendance_Type in rowData and clears Max_Custom when Hours is not custom.
// Duplicate check is global (not REC-scoped): the marking app reads these sheets without a
// REC filter, so a duplicate in another school would collide there too.
// Returns an error message string, or null when OK.
// ---------- Applicable_Classes (v7.0) ----------
// Parse an Applicable_Classes cell into an array of class tokens. Accepts '{VIII,IX,X,XI}',
// a bare 'VII', '' (blank = all classes) or an array. Same brace rules as Class_Terms. Never throws.
function parseApplicableClasses_(cell) {
  if (Array.isArray(cell)) return cell.map(function (t) { return String(t == null ? '' : t).trim(); }).filter(Boolean);
  const str = String(cell == null ? '' : cell).trim();
  if (!str) return [];
  let s = str;
  if (s.charAt(0) === '{' && s.charAt(s.length - 1) === '}') s = s.slice(1, -1);
  return s.split(',').map(function (t) { return t.trim(); }).filter(Boolean);
}

// Format a class list for storage: 0 → '' (all classes), 1 → bare value, 2+ → '{a,b,c}' (same as Class_Terms).
function formatApplicableClasses_(list) {
  const arr = Array.isArray(list) ? list : parseApplicableClasses_(list);
  if (arr.length === 0) return '';
  if (arr.length === 1) return arr[0];
  return '{' + arr.join(',') + '}';
}

// Validate + normalise an Applicable_Classes value: every listed class must be in MASTER_CLASSES
// (matched case-insensitively, stored in canonical spelling); duplicates collapse.
// Returns { value, error } — error is null when OK.
function normaliseApplicableClasses_(v) {
  const parsed = parseApplicableClasses_(v);
  const seen = {};
  const out = [];
  for (let i = 0; i < parsed.length; i++) {
    const tok = String(parsed[i]).trim();
    if (!tok) continue;
    const canonical = MASTER_CLASSES.filter(function (c) { return c.toLowerCase() === tok.toLowerCase(); })[0];
    if (!canonical) return { value: '', error: 'Unknown class "' + tok + '". Choose from: ' + MASTER_CLASSES.join(', ') + '.' };
    if (!seen[canonical]) { seen[canonical] = true; out.push(canonical); }
  }
  return { value: formatApplicableClasses_(out), error: null };
}

function validateHoursConfigRow_(sheetName, sheet, headers, rowData, excludeRow, scopeRec) {
  if (isHoursSetupSheet_(sheetName)) {
    const hRaw = String(rowData.Hours == null ? '' : rowData.Hours).trim();
    if (!hRaw) return 'Hours is required.';
    const isCustom = hRaw.toLowerCase() === 'custom';
    if (!isCustom) {
      const n = Number(hRaw);
      if (isNaN(n) || n <= 0) return 'Hours must be a number greater than 0, or "custom".';
    }
    const type = normHoursConfigType_(rowData.Attendance_Type);
    if (!type) return 'Attendance_Type must be normal, episodic or both.';
    rowData.Attendance_Type = type;
    if (isCustom) {
      const mc = rowData.Max_Custom;
      if (mc === undefined || mc === null || String(mc).trim() === '' || isNaN(Number(mc)) || Number(mc) <= 0) {
        return 'Max_Custom is required and must be a positive number when Hours is "custom".';
      }
    } else {
      rowData.Max_Custom = '';
    }
    const hIdx = headers.indexOf('Hours');
    const tIdx = headers.indexOf('Attendance_Type');
    const hrIdx = headers.indexOf('REC_ID');
    if (hIdx !== -1 && tIdx !== -1) {
      const all = sheet.getDataRange().getValues();
      for (let i = 1; i < all.length; i++) {
        if (excludeRow && (i + 1) === excludeRow) continue;
        // v7.9 (fix 14): a duplicate only blocks within the same school.
        const exRecH = hrIdx !== -1 ? String(all[i][hrIdx] || '').trim() : '';
        if (scopeRec && exRecH && exRecH !== scopeRec) continue;
        if (normHoursValue_(all[i][hIdx]) === normHoursValue_(hRaw) &&
            normHoursConfigType_(all[i][tIdx]) === type) {
          return 'Hours "' + hRaw + '" (' + type + ') already exists. Re-activate the existing row instead of adding a duplicate.';
        }
      }
    }
    return null;
  }
  // Hour_Categories
  const catId = String(rowData.Category_ID == null ? '' : rowData.Category_ID).trim();
  if (!catId) return 'Category_ID is required.';
  if (!String(rowData.Category_Name == null ? '' : rowData.Category_Name).trim()) return 'Category_Name is required.';
  const type = normHoursConfigType_(rowData.Attendance_Type);
  if (!type) return 'Attendance_Type must be normal, episodic or both.';
  rowData.Attendance_Type = type;
  // v7.0: Applicable_Classes — blank = all classes; every listed class must be in MASTER_CLASSES.
  const normAC = normaliseApplicableClasses_(rowData.Applicable_Classes);
  if (normAC.error) return normAC.error;
  rowData.Applicable_Classes = normAC.value;
  const cIdx = headers.indexOf('Category_ID');
  const crIdx = headers.indexOf('REC_ID');
  if (cIdx !== -1) {
    const all = sheet.getDataRange().getValues();
    for (let i = 1; i < all.length; i++) {
      if (excludeRow && (i + 1) === excludeRow) continue;
      // v7.9 (fix 14): a duplicate only blocks within the same school.
      const exRecC = crIdx !== -1 ? String(all[i][crIdx] || '').trim() : '';
      if (scopeRec && exRecC && exRecC !== scopeRec) continue;
      if (String(all[i][cIdx] || '').trim().toLowerCase() === catId.toLowerCase()) {
        return 'Category_ID "' + catId + '" already exists. Re-activate the existing row instead of adding a duplicate.';
      }
    }
  }
  return null;
}

function addSheetRow(sheetName, rowData, currentUsername) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = resolveSheet_(ss, sheetName);
    // v7.15: a managed sheet that does not exist yet is created + seeded
    // on first Add (same lazy rule as Transfer_Certificates).
    if (!sheet && isLeavingReasonsSheet_(sheetName)) sheet = getOrCreateLeavingReasonsSheet_(ss);
    if (!sheet && isCountriesSheet_(sheetName)) sheet = getOrCreateCountriesSheet_(ss);
    if (!sheet) throw new Error(sheetName + ' sheet not found');
    
    // v7.8 (Fix 1): only an admin/superadmin Users row may write.
    const addGate = requireAdminCaller_(ss, currentUsername);
    if (!addGate.ok) return addGate.error;
    const addCaller = addGate.caller;

    // Handle special cases first
    if (String(sheetName || '').trim().toLowerCase() === 'student_data') {
      // Pass REC_ID from rowData if available
      const recId = rowData.REC_ID || '';
      return addStudent(rowData, currentUsername, recId);
    }

    // v7.10: Transfer_Certificates rows are created only by a student
    // deactivation (Student_Data branch of updateSheetRow / addStudent).
    if (isTransferCertificatesSheet_(sheetName)) {
      return { success: false, message: 'Transfer certificates are created only when a student is deactivated from the Student Data form.' };
    }
    
    // v7.9 (fix 10): the duplicate check now actually runs — the old
    // exact-case 'Teacher_Assignments' never matched the lowercase name
    // the client sends — and it is scoped by school, so one school's row
    // can no longer block another school's.
    if (String(sheetName || '').trim().toLowerCase() === 'teacher_assignments') {
      const taScopeRec = addCaller.userType === 'superadmin' ? String(rowData.REC_ID || '').trim() : (addCaller.recId || '');
      if (teacherAssignmentDuplicateExists_(sheet, rowData.Username, taScopeRec, 0)) {
        return {
          success: false,
          message: `Username ${rowData.Username} already has assigned classes. Please edit the existing record.`
        };
      }
    }
    
    if (isPrevPermSheet_(sheetName)) ensurePrevPermColumns_(sheet);
    if (isHoursConfigSheet_(sheetName)) ensureHoursConfigColumns_(sheet, sheetName);
    if (isRecNamesSheet_(sheetName)) ensureRecNamesActiveColumn_(sheet);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];



if (isPrevPermSheet_(sheetName)) {
  const today = new Date();
today.setHours(0, 0, 0, 0);
const parts = String(rowData.Allowed_Date).split('-');
const allowed = new Date(parts[0], parts[1] - 1, parts[2]);
if (allowed > today) return { success: false, message: 'Allowed Date cannot be in the future.' };

  if (prevPermDuplicateExists_(sheet, headers, rowData, 0)) return { success: false, message: 'This user already has permission for this date for this class.' };
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

    // v7.8 (Fix 2): Users Add hardening — UserType allow-list, only a
    // superadmin may create a superadmin, allowedTabs must be known keys.
    if (sheetName.toLowerCase() === 'users') {
      const newUserType = String(rowData.UserType || '').trim();
      if (ALLOWED_USER_TYPES.indexOf(newUserType) === -1) {
        return { success: false, message: 'UserType must be admin, teacher or superadmin.' };
      }
      if (newUserType === 'superadmin' && addCaller.userType !== 'superadmin') {
        return { success: false, message: 'Only a superadmin can create a superadmin user.' };
      }
      const tabsRaw = rowData.allowedTabs;
      const tabsList = Array.isArray(tabsRaw) ? tabsRaw : (tabsRaw == null || tabsRaw === '' ? [] : String(tabsRaw).split(','));
      for (let tabIdx = 0; tabIdx < tabsList.length; tabIdx++) {
        const tabKey = String(tabsList[tabIdx] || '').trim();
        if (tabKey && ALLOWED_TABS.indexOf(tabKey) === -1) {
          return { success: false, message: 'Unknown tab "' + tabKey + '" in allowedTabs.' };
        }
      }
    }
    // v7.8 (Fix 2): a User_Permissions row needs a real user — of the
    // caller's school, or of any school when the caller is superadmin.
    if (sheetName.toLowerCase() === 'user_permissions') {
      const permUser = String(rowData.Username || '').trim();
      const permUsersSheet = ss.getSheetByName('Users');
      const permUsersData = permUsersSheet ? permUsersSheet.getDataRange().getValues() : [];
      const permHeaders = permUsersData[0] || [];
      const permRecCol = permHeaders.indexOf('REC_ID');
      let permUserFound = false;
      for (let pu = 1; pu < permUsersData.length; pu++) {
        if (String(permUsersData[pu][0] || '').trim() === permUser) {
          const permUserRec = permRecCol !== -1 ? String(permUsersData[pu][permRecCol] || '').trim() : '';
          if (addCaller.userType === 'superadmin' || !recId || !permUserRec || permUserRec === recId) { permUserFound = true; break; }
        }
      }
      if (!permUserFound) return { success: false, message: 'Username "' + permUser + '" does not exist for this school.' };
      // v7.9 (fix 6): one permissions row per username.
      if (userPermissionsDuplicateExists_(sheet, headers, permUser, 0)) {
        return { success: false, message: 'Username "' + permUser + '" already has a permissions row. Edit the existing row instead.' };
      }
    }
    // v7.8 (Fix 3): the row's school must be the caller's school
    // (superadmin skipped; fail-open when the caller's school is unknown).
    // Sheets with no REC_ID column (Terms is global by design) skip the check.
    if (headers.indexOf('REC_ID') !== -1 && addCaller.userType !== 'superadmin' && recId) {
      const incomingRec = String(rowData.REC_ID || '').trim();
      if (incomingRec && incomingRec !== recId) {
        return { success: false, message: 'This record belongs to another school. You can only add records for your own school.' };
      }
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
    // v7.9 (fix 9): Class_Sections server checks — the value must be a real
    // class + single section letter, and no duplicate for the same school.
    if (String(sheetName || '').trim().toLowerCase() === 'class_sections') {
      const csEffRec = String(rowData.REC_ID || '').trim() || recId;
      const csAddErr = checkClassSectionCombo_(rowData.Class_Section, ss, csEffRec);
      if (csAddErr) return { success: false, message: csAddErr };
      if (classSectionDuplicateExists_(sheet, headers, rowData.Class_Section, csEffRec, 0)) {
        return { success: false, message: 'Section "' + String(rowData.Class_Section || '').trim() + '" already exists for this school.' };
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
      const ctErr = validateClassTermsRow_(sheet, headers, rowData, recId, 0, addCaller);
      if (ctErr) return { success: false, message: ctErr };
    }
    // Hours config tables (v6.9): default Is_Active TRUE on Add, then validate + duplicate check.
    if (isHoursConfigSheet_(sheetName)) {
      if (rowData.Is_Active === undefined || rowData.Is_Active === null || String(rowData.Is_Active).trim() === '') rowData.Is_Active = 'TRUE';
      else rowData.Is_Active = isActiveFlag_(rowData.Is_Active) ? 'TRUE' : 'FALSE';
      const hcErr = validateHoursConfigRow_(sheetName, sheet, headers, rowData, 0, recId);
      if (hcErr) return { success: false, message: hcErr };
    }
    // Managed tables (v7.15): default Is_Active TRUE on Add, then validate
    // + duplicate check. RECNames also gets a blank S.# filled with the
    // next serial number.
    if (isManagedTableSheet_(sheetName)) {
      if (rowData.Is_Active === undefined || rowData.Is_Active === null || String(rowData.Is_Active).trim() === '') rowData.Is_Active = 'TRUE';
      else rowData.Is_Active = isActiveFlag_(rowData.Is_Active) ? 'TRUE' : 'FALSE';
      const mgErr = validateManagedTableRow_(sheetName, sheet, headers, rowData, 0);
      if (mgErr) return { success: false, message: mgErr };
      if (isRecNamesSheet_(sheetName) && (rowData['S.#'] === undefined || rowData['S.#'] === null || String(rowData['S.#']).trim() === '')) {
        rowData['S.#'] = String(sheet.getLastRow());
      }
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

      // v7.21: all six admin rights are stored as 'Yes'/'No', never
      // blank. (The old check used exact-case 'Users' while the client
      // sends 'users', so it never ran, and it covered only two rights.)
      if (sheetName.toLowerCase() === 'users' && ['student_delete', 'canDeleteUsers', 'canQueryAttendance', 'canDownload', 'canAmendSIRN', 'canCorrectAttendance'].includes(header)) {
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
    // v7.8 (Fix 1, corrected): gate on the username the caller actually
    // sent, BEFORE the fallbacks below. Gating after the fallbacks let a
    // direct call with no username pass as the first Users row of the
    // school named in rowData.REC_ID.
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('No active spreadsheet found');
    const updGate = requireAdminCaller_(ss, currentUsername);
    if (!updGate.ok) return updGate.error;
    const updCaller = updGate.caller;

    // Username for stamping (Last_Modified_By etc.). The gate above has
    // already proved currentUsername is an admin; the fallbacks below are
    // kept only as a backstop and no longer decide who may write.
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
    
    const sheet = resolveSheet_(ss, sheetName);
    if (!sheet) throw new Error(`${sheetName} sheet not found`);

    // v7.10: Transfer_Certificates is read-only through the generic path —
    // rows are created only by a student deactivation (see the Student_Data
    // branch below). Editing one would rewrite the certificate snapshot.
    if (isTransferCertificatesSheet_(sheetName)) {
      return { success: false, message: 'Transfer certificates cannot be edited. They are created when a student is deactivated.' };
    }
    
    // Block superadmin modifications. v7.9 (fix 13): the old exact-case
    // 'Users' never matched the lowercase name the client sends, so this
    // block never ran from the portal — it does now.
    if (String(sheetName || '').trim().toLowerCase() === 'users') {
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
    if (isHoursConfigSheet_(sheetName)) ensureHoursConfigColumns_(sheet, sheetName);
    if (isRecNamesSheet_(sheetName)) ensureRecNamesActiveColumn_(sheet);
    
    // v7.10: lazily append the new Student_Data columns (Global_ID etc.)
    // before the headers are read, so the Student_Data branch below sees them.
    if (isStudentDataSheet_(sheetName)) ensureStudentColumns_(sheet);
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
  // v7.9 (fix 13): Users Edit hardening — UserType allow-list, allowedTabs
  // must be known keys, and only a superadmin may set UserType superadmin.
  // A UserType the caller did not send keeps the stored value (Fix-4 rule),
  // so only a sent value is checked.
  if (newT && ALLOWED_USER_TYPES.indexOf(newT) === -1) {
    return { success: false, message: 'UserType must be admin, teacher or superadmin.' };
  }
  if (newT === 'superadmin' && updCaller.userType !== 'superadmin') {
    return { success: false, message: 'Only a superadmin can set UserType to superadmin.' };
  }
  if (rowData.allowedTabs !== undefined && rowData.allowedTabs !== null) {
    const tabsRaw = rowData.allowedTabs;
    const tabsList = Array.isArray(tabsRaw) ? tabsRaw : (tabsRaw === '' ? [] : String(tabsRaw).split(','));
    for (let tabIdx = 0; tabIdx < tabsList.length; tabIdx++) {
      const tabKey = String(tabsList[tabIdx] || '').trim();
      if (tabKey && ALLOWED_TABS.indexOf(tabKey) === -1) {
        return { success: false, message: 'Unknown tab "' + tabKey + '" in allowedTabs.' };
      }
    }
  }
}


// Edit: Previous Date Permissions — full Add checks restored (v7.9 fix 7):
// real date, not future, non-empty password, teacher of the caller's
// school, class one of their assigned classes, plus the duplicate check.
if (isPrevPermSheet_(sheetName)) {
  const ppErr = validatePrevPermEdit_(ss, sheet, headers, rowData, rowIndex, updCaller);
  if (ppErr) return { success: false, message: ppErr };
}

// v7.9 (fix 8): Teacher_Assignments Edit — duplicate, class existence,
// username existence (see validateTeacherAssignmentEdit_).
if (String(sheetName || '').trim().toLowerCase() === 'teacher_assignments') {
  const taErr = validateTeacherAssignmentEdit_(ss, sheet, headers, rowData, rowIndex, updCaller);
  if (taErr) return { success: false, message: taErr };
}

// v7.9 (fix 9): Class_Sections Edit — same checks as Add, own row excluded.
// An unchanged combo is grandfathered (the format check only runs when
// the value changed); the duplicate check always runs.
if (String(sheetName || '').trim().toLowerCase() === 'class_sections') {
  const csStored = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  const csComboIdx = headers.indexOf('Class_Section');
  const csRecIdx = headers.indexOf('REC_ID');
  const csStoredCombo = csComboIdx !== -1 ? String(csStored[csComboIdx] || '').trim() : '';
  const csEffCombo = rowData.Class_Section !== undefined && rowData.Class_Section !== null ? String(rowData.Class_Section).trim() : csStoredCombo;
  const csEffRec = rowData.REC_ID !== undefined && rowData.REC_ID !== null && String(rowData.REC_ID).trim() !== '' ? String(rowData.REC_ID).trim() : (csRecIdx !== -1 ? String(csStored[csRecIdx] || '').trim() : '');
  if (csEffCombo !== csStoredCombo) {
    const csEditErr = checkClassSectionCombo_(csEffCombo, ss, csEffRec);
    if (csEditErr) return { success: false, message: csEditErr };
  }
  if (classSectionDuplicateExists_(sheet, headers, csEffCombo, csEffRec, rowIndex)) {
    return { success: false, message: 'Section "' + csEffCombo + '" already exists for this school.' };
  }
}

// v7.9 (fix 6): User_Permissions Edit — one permissions row per username.
if (sheetName.toLowerCase() === 'user_permissions') {
  const upStored = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  const upUIdx = headers.indexOf('Username');
  const upEffU = rowData.Username !== undefined && rowData.Username !== null && String(rowData.Username).trim() !== '' ? String(rowData.Username).trim() : (upUIdx !== -1 ? String(upStored[upUIdx] || '').trim() : '');
  if (upEffU && userPermissionsDuplicateExists_(sheet, headers, upEffU, rowIndex)) {
    return { success: false, message: 'Username "' + upEffU + '" already has a permissions row. Edit the existing row instead.' };
  }
}


    // Terms / Class_Terms (v4): normalise Is_Active, then validate (own row excluded).
    // v7.8 (Fix 4): only normalise Is_Active when the caller sent it — an
    // absent value keeps the stored one in the generic write below.
    if ((isTermsSheet_(sheetName) || isClassTermsSheet_(sheetName)) && rowData.Is_Active !== undefined && rowData.Is_Active !== null) {
      rowData.Is_Active = isActiveFlag_(rowData.Is_Active) ? 'TRUE' : 'FALSE';
    }
    if (isTermsSheet_(sheetName)) {
      // v7.9: an omitted Term_Name on Edit keeps the stored one (Fix-4 rule),
      // so the required-name check sees the effective value.
      const exRowT = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
      const nmIdx = headers.indexOf('Term_Name');
      const effTermsRow = Object.assign({}, rowData);
      if ((rowData.Term_Name === undefined || rowData.Term_Name === null) && nmIdx !== -1) effTermsRow.Term_Name = exRowT[nmIdx];
      const idIdxT = headers.indexOf('Term_ID');
      if ((rowData.Term_ID === undefined || rowData.Term_ID === null) && idIdxT !== -1) effTermsRow.Term_ID = exRowT[idIdxT];
      const tErr = validateTermsRow_(sheet, headers, effTermsRow, rowIndex);
      if (tErr) return { success: false, message: tErr };
    }
    if (isClassTermsSheet_(sheetName)) {
      const exRow = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
      const rIdx = headers.indexOf('REC_ID');
      const ctRec = String(rowData.REC_ID || (rIdx !== -1 ? exRow[rIdx] : '') || '').trim();
      const ctErr = validateClassTermsRow_(sheet, headers, rowData, ctRec, rowIndex, updCaller);
      if (ctErr) return { success: false, message: ctErr };
    }

    // Hours config tables (v6.9, v7.0): rows are immutable once added — only Is_Active and Applicable_Classes may change.
    if (isHoursConfigSheet_(sheetName)) {
      const storedRow = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
      const normVal = function (v) { return String(v == null ? '' : v).trim().toLowerCase(); };
      const changed = headers.some(function (h) {
        if (h === 'Is_Active' || h === 'REC_ID' || h === 'Applicable_Classes') return false;
        const si = headers.indexOf(h);
        const oldV = storedRow[si];
        const newV = rowData[h];
        if (h === 'Hours') return normHoursValue_(newV) !== normHoursValue_(oldV);
        if (h === 'Max_Custom') {
          const on = (newV === '' || newV == null) ? '' : String(Number(newV));
          const oo = (oldV === '' || oldV == null) ? '' : String(Number(oldV));
          return on !== oo;
        }
        return normVal(newV) !== normVal(oldV);
      });
      if (changed) return { success: false, message: HOURS_CONFIG_IMMUTABLE_MSG };
      // Normalise Is_Active; a blank value keeps the stored one.
      const iaIdx = headers.indexOf('Is_Active');
      if (rowData.Is_Active === undefined || rowData.Is_Active === null || String(rowData.Is_Active).trim() === '') {
        rowData.Is_Active = iaIdx !== -1 ? storedRow[iaIdx] : 'TRUE';
      } else {
        rowData.Is_Active = isActiveFlag_(rowData.Is_Active) ? 'TRUE' : 'FALSE';
      }
      // v7.0: validate the editable fields on Edit too (own row excluded) — unknown classes are refused.
      const hcRecIdx = headers.indexOf('REC_ID');
      const hcScopeRec = hcRecIdx !== -1 ? String(storedRow[hcRecIdx] || '').trim() : (updCaller.recId || '');
      const hcErr = validateHoursConfigRow_(sheetName, sheet, headers, rowData, rowIndex, hcScopeRec);
      if (hcErr) return { success: false, message: hcErr };
    }

    // Managed tables (v7.15): rows are immutable once added — only
    // Is_Active may change. A reason's flags are re-derived on Edit too,
    // so a direct call cannot smuggle in flags that disagree with Category.
    if (isManagedTableSheet_(sheetName)) {
      const mgImErr = checkManagedTableImmutable_(sheetName, sheet, headers, rowData, rowIndex);
      if (mgImErr) return { success: false, message: mgImErr };
      if (isLeavingReasonsSheet_(sheetName)) {
        const mgValErr = validateManagedTableRow_(sheetName, sheet, headers, rowData, rowIndex);
        if (mgValErr) return { success: false, message: mgValErr };
      }
    }

    // ✅ STUDENT DATA SECTION
    if (String(sheetName || '').trim().toLowerCase() === 'student_data') {
      const rowRange = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn());
      const rowValues = rowRange.getValues()[0];
      const originalStdId = rowValues[0];

      // v7.7: Student-edit validations — school, stale row, class, section,
      // gender, date, Unique_ID. Each failure rejects before anything is written.
      const sdRecIdx = headers.indexOf('REC_ID');
      const sdSecIdx = headers.indexOf('Student_Section');
      const sdGenIdx = headers.indexOf('Gender');
      const sdClsIdx = headers.indexOf('Student_Class');
      const rowRecId = sdRecIdx !== -1 && rowValues[sdRecIdx] != null ? String(rowValues[sdRecIdx]).trim() : '';
      const storedSection = sdSecIdx !== -1 && rowValues[sdSecIdx] != null ? String(rowValues[sdSecIdx]).trim() : '';
      const storedGender = sdGenIdx !== -1 && rowValues[sdGenIdx] != null ? String(rowValues[sdGenIdx]).trim() : '';
      const storedClass = sdClsIdx !== -1 && rowValues[sdClsIdx] != null ? String(rowValues[sdClsIdx]).trim() : '';
      const newClass = rowData.Student_Class != null ? String(rowData.Student_Class).trim() : '';
      const newSection = rowData.Student_Section != null ? String(rowData.Student_Section).trim() : '';
      const newGender = rowData.Gender != null ? String(rowData.Gender).trim() : '';
      const newDoj = rowData.Date_of_Joining != null ? String(rowData.Date_of_Joining).trim() : '';
      const newUid = rowData.Unique_ID != null ? String(rowData.Unique_ID).trim() : '';
      // v7.14: Date_of_Birth is optional on Add/Edit — validated as a
      // real date only when a value is given; blank is allowed.
      const newDob = rowData.Date_of_Birth != null ? String(rowData.Date_of_Birth).trim() : '';

      // (a) School check — the row must belong to the caller's school.
      // Fail-open when the caller's school cannot be determined.
      if (username !== 'superadmin') {
        let userRecId = '';
        try {
          const usersSheet = ss.getSheetByName('Users');
          if (usersSheet) {
            const usersData = usersSheet.getDataRange().getValues();
            for (let ui = 1; ui < usersData.length; ui++) {
              if (String(usersData[ui][0] || '').trim() === username) {
                userRecId = usersData[ui][15] != null ? String(usersData[ui][15]).trim() : '';
                break;
              }
            }
          }
        } catch (e) {
          userRecId = '';
        }
        if (userRecId && rowRecId && userRecId !== rowRecId) {
          return { success: false, message: 'This student belongs to another school. You can only edit students of your own school.' };
        }
      }

      // (b) Stale-row check — the record must still be the one the editor opened.
      // Skipped when the client did not send _originalStdId (old clients).
      const sentOrigStdId = rowData._originalStdId != null ? String(rowData._originalStdId).trim() : '';
      if (sentOrigStdId && sentOrigStdId !== String(originalStdId != null ? originalStdId : '').trim()) {
        return { success: false, message: 'This record changed since you opened it. Please refresh and try again.' };
      }
      // (b2) Concurrent-edit check (v7.9) — the row must still be exactly
      // as the editor opened it. Last_Modified_Date is date-only, so field
      // values are compared, not the date. Skipped when the client sent no
      // _openedRow (old clients).
      if (rowData._openedRow && typeof rowData._openedRow === 'object') {
        const opened = rowData._openedRow;
        const sdNameIdx = headers.indexOf('Student_Name');
        const sdStatIdx = headers.indexOf('Status');
        const sdBcIdx = headers.indexOf('Barcode_ID');
        const sdDojIdx = headers.indexOf('Date_of_Joining');
        const sdUidIdx = headers.indexOf('Unique_ID');
        const lmByIdx = headers.indexOf('Last_Modified_By');
        const lmDateIdx = headers.indexOf('Last_Modified_Date');
        const openedVals = {
          Std_ID: String(originalStdId != null ? originalStdId : '').trim(),
          Student_Name: sdNameIdx !== -1 ? String(rowValues[sdNameIdx] || '').trim() : '',
          Student_Class: storedClass,
          Student_Section: storedSection,
          Status: sdStatIdx !== -1 ? String(rowValues[sdStatIdx] || '').trim() : '',
          Barcode_ID: sdBcIdx !== -1 ? String(rowValues[sdBcIdx] || '').trim() : '',
          Gender: storedGender,
          Date_of_Joining: sdDojIdx !== -1 ? rowValues[sdDojIdx] : '',
          Unique_ID: sdUidIdx !== -1 ? String(rowValues[sdUidIdx] || '').trim() : ''
        };
        const sentVals = {
          Std_ID: String(opened.Std_ID != null ? opened.Std_ID : '').trim(),
          Student_Name: String(opened.Student_Name != null ? opened.Student_Name : '').trim(),
          Student_Class: String(opened.Student_Class != null ? opened.Student_Class : '').trim(),
          Student_Section: String(opened.Student_Section != null ? opened.Student_Section : '').trim(),
          Status: String(opened.Status != null ? opened.Status : '').trim(),
          Barcode_ID: String(opened.Barcode_ID != null ? opened.Barcode_ID : '').trim(),
          Gender: String(opened.Gender != null ? opened.Gender : '').trim(),
          Date_of_Joining: opened.Date_of_Joining != null ? opened.Date_of_Joining : '',
          Unique_ID: String(opened.Unique_ID != null ? opened.Unique_ID : '').trim()
        };
        // Dates may arrive as dd-MMM-yyyy (stored) vs yyyy-MM-dd (sent) —
        // compare the day, not the spelling.
        const dojStoredKey = dateKeyForCompare_(openedVals.Date_of_Joining);
        const dojSentKey = dateKeyForCompare_(sentVals.Date_of_Joining);
        const dojChanged = (dojStoredKey || String(openedVals.Date_of_Joining || '').trim()) !== (dojSentKey || String(sentVals.Date_of_Joining || '').trim());
        const anyChanged = sentVals.Std_ID !== openedVals.Std_ID ||
          sentVals.Student_Name !== openedVals.Student_Name ||
          sentVals.Student_Class !== openedVals.Student_Class ||
          sentVals.Student_Section !== openedVals.Student_Section ||
          sentVals.Status !== openedVals.Status ||
          sentVals.Barcode_ID !== openedVals.Barcode_ID ||
          sentVals.Gender !== openedVals.Gender ||
          dojChanged ||
          sentVals.Unique_ID !== openedVals.Unique_ID;
        if (anyChanged) {
          const lmBy = lmByIdx !== -1 ? String(rowValues[lmByIdx] || '').trim() : '';
          const lmDate = lmDateIdx !== -1 ? String(rowValues[lmDateIdx] || '').trim() : '';
          const who = (lmBy && lmDate) ? (' by ' + lmBy + ' on ' + lmDate) : (lmBy ? (' by ' + lmBy) : (lmDate ? (' on ' + lmDate) : ''));
          return { success: false, message: 'This student was changed' + who + '. Please refresh and try again.' };
        }
      }

      // (c) + (d) Class and section checks against Class_Mapping for the row's school.
      // Fail-open if the mapping sheet cannot be read.
      let mappedClasses = null;
      try {
        mappedClasses = getClassMappingData(rowRecId || '');
      } catch (e) {
        mappedClasses = null;
      }
      if (mappedClasses) {
        const classNames = mappedClasses.map(function (m) { return String(m.Class || '').trim(); });
        const effClass = newClass || storedClass;
        if (effClass && effClass !== storedClass && classNames.length > 0 && classNames.indexOf(effClass) === -1) {
          return { success: false, message: 'Unknown class "' + effClass + '". Choose from: ' + classNames.join(', ') };
        }
        const mapRow = effClass ? mappedClasses.filter(function (m) { return String(m.Class || '').trim() === effClass; })[0] : null;
        const allowedSecs = mapRow ? String(mapRow.Sections || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean) : [];
        const effSection = newSection || storedSection;
        if (effSection && effSection !== storedSection && allowedSecs.indexOf(effSection) === -1) {
          if (mapRow && allowedSecs.length > 0) {
            return { success: false, message: 'Section must be one of ' + allowedSecs.join(', ') + ' for class ' + effClass + '.' };
          }
          return { success: false, message: 'Section "' + effSection + '" is not set up for class ' + effClass + '. The stored section is ' + (storedSection || '(none)') + '.' };
        }
      }

      // (e) Gender check — a changed Gender must be Male or Female; the
      // stored value is always allowed (grandfathering).
      if (newGender && newGender !== storedGender && newGender !== 'Male' && newGender !== 'Female') {
        return { success: false, message: 'Gender must be Male or Female.' };
      }

      // (f) Date_of_Joining must be a real date (yyyy-MM-dd or dd-MMM-yyyy).
      if (newDoj) {
        let dojOk = false;
        const isoM = /^(\d{4})-(\d{2})-(\d{2})$/.exec(newDoj);
        const dmyM = /^(\d{2})-([A-Za-z]{3})-(\d{4})$/.exec(newDoj);
        if (isoM) {
          const dd = new Date(Number(isoM[1]), Number(isoM[2]) - 1, Number(isoM[3]));
          dojOk = dd.getFullYear() === Number(isoM[1]) && dd.getMonth() === Number(isoM[2]) - 1 && dd.getDate() === Number(isoM[3]);
        } else if (dmyM) {
          const monMap = { Jan: 0, Feb: 1, Mar: 2, Apr: 3, May: 4, Jun: 5, Jul: 6, Aug: 7, Sep: 8, Oct: 9, Nov: 10, Dec: 11 };
          const monKey = dmyM[2].charAt(0).toUpperCase() + dmyM[2].slice(1).toLowerCase();
          const mi = monMap[monKey];
          if (mi !== undefined) {
            const dd = new Date(Number(dmyM[3]), mi, Number(dmyM[1]));
            dojOk = dd.getFullYear() === Number(dmyM[3]) && dd.getMonth() === mi && dd.getDate() === Number(dmyM[1]);
          }
        }
        if (!dojOk) {
          return { success: false, message: 'Date of Joining is not a real date.' };
        }
      }

      // (g) Unique_ID must be exactly 6 digits when provided.
      if (newUid && !/^\d{6}$/.test(newUid)) {
        return { success: false, message: 'Unique_ID must be exactly 6 digits' };
      }

      // (h) v7.14: Date_of_Birth, when given, must be a real date.
      if (newDob && !isRealDateString_(newDob)) {
        return { success: false, message: 'Date of Birth is not a real date.' };
      }
      
      // v7.9 (fix 12): Barcode_ID must be unique — another row with the same
      // barcode is refused. A blank barcode is never a duplicate.
      const sdBcIdx2 = headers.indexOf('Barcode_ID');
      const effBc = rowData.Barcode_ID != null ? String(rowData.Barcode_ID).trim() : '';
      if (sdBcIdx2 !== -1 && effBc && sheet.getLastRow() > 1) {
        const allBc = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
        for (let bi = 0; bi < allBc.length; bi++) {
          if ((bi + 2) === rowIndex) continue;
          if (String(allBc[bi][sdBcIdx2] || '').trim() === effBc) {
            return { success: false, message: 'Barcode_ID "' + effBc + '" already exists for another student.' };
          }
        }
      }

      // v7.10: TRANSFER CERTIFICATE — a NEW deactivation (Status becomes
      // Inactive from anything else) needs a validated reason payload
      // (rowData._tcPayload). An already-Inactive student edited without
      // one is grandfathered: no new TC is created.
      const sdStatusIdx = headers.indexOf('Status');
      const storedStatus = sdStatusIdx !== -1 ? String(rowValues[sdStatusIdx] || '').trim() : '';
      const effStatus = rowData.Status != null && String(rowData.Status).trim() !== '' ? String(rowData.Status).trim() : storedStatus;
      const isNewDeactivation = effStatus === 'Inactive' && storedStatus !== 'Inactive';
      const sdStoredByHeader = {};
      headers.forEach(function (h, hi) { sdStoredByHeader[h] = tcStoredStr_(rowValues[hi]); });
      let tcClean = null;
      if (isNewDeactivation) {
        const tcCheck = validateTcPayload_(ss, rowData._tcPayload, sdStoredByHeader, rowRecId);
        if (tcCheck.error) return { success: false, message: tcCheck.error };
        tcClean = tcCheck.clean;
      }
      // v7.10: Global_ID — never accept a client-supplied value. Keep the
      // stored one, or allocate a new one (under the script lock, together
      // with the TC row) when the row has none.
      const sdGlobalIdx = headers.indexOf('Global_ID');
      const storedGlobalId = sdGlobalIdx !== -1 ? tcStoredStr_(rowValues[sdGlobalIdx]) : '';
      const needsGlobalId = !storedGlobalId;
      // v7.16: REACTIVATION BLOCK — a student with ANY certificate
      // record (Transfer or Error) can never be reactivated. Only a
      // student deactivated before v7.10 (no TC-sheet row at all), or one
      // whose TC append failed (student Inactive, no TC row), may be.
      // The client hides the Reactivate button for students with a
      // certificate record; this is the server backstop, so a direct
      // call is refused too.
      const isReactivation = effStatus === 'Active' && storedStatus === 'Inactive';
      if (isReactivation && hasAnyTc_(ss, storedGlobalId, originalStdId, rowRecId)) {
        return { success: false, message: 'A certificate record was issued for this student. Reactivation is not allowed.' };
      }
      // v7.16: FROZEN RECORD — a student with ANY certificate record
      // (Transfer or Error) cannot be edited at all (the certificate is a
      // snapshot of this row). The deactivation that creates the TC is
      // unaffected: the student is Active then and no TC exists yet.
      // The lazy Global_ID allocation for a pre-TC inactive student is
      // unaffected too (no TC record exists for them). A correction to
      // a frozen student's details after issue is possible only
      // directly in the sheet. The client hides the Edit button for
      // these rows; this is the server backstop for a direct call.
      if (hasAnyTc_(ss, storedGlobalId, originalStdId, rowRecId)) {
        return { success: false, message: 'A certificate record was issued for this student. The record is locked.' };
      }

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
      
      // v7.10: allocate the Global_ID (when the row has none) and hold the
      // script lock until the TC row is appended below, so the Global_ID
      // and TC number sequences cannot collide with a concurrent save.
      let tcLock = null;
      let newGlobalId = storedGlobalId;
      if (needsGlobalId || isNewDeactivation) {
        tcLock = LockService.getScriptLock();
        if (!tcLock.tryLock(30000)) {
          return { success: false, message: 'The system is busy — please try again.' };
        }
        if (needsGlobalId) newGlobalId = nextGlobalIdLocked_(ss);
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
          case 'Global_ID':
            // v7.10: never accept a client-supplied Global_ID.
            updatedRow.push(newGlobalId || '');
            break;
          case 'Date_of_Birth':
            updatedRow.push(tcClean && tcClean.Date_of_Birth ? tcClean.Date_of_Birth : (rowData.Date_of_Birth || rowValues[i] || ''));
            break;
          case 'Father_Name':
            // v7.14: Father_Name is no longer collected or stored — keep
            // the stored value (the column stays on the sheet).
            updatedRow.push(rowValues[i] || '');
            break;
          case 'Address':
            updatedRow.push(tcClean && tcClean.Address ? tcClean.Address : (rowData.Address || rowValues[i] || ''));
            break;
          case 'Phone':
            updatedRow.push(tcClean && tcClean.Phone ? tcClean.Phone : (rowData.Phone || rowValues[i] || ''));
            break;
          case 'Secondary_Phone':
            // v7.15: the secondary contact number, saved from the TC
            // payload on deactivation; otherwise keep the stored value.
            updatedRow.push(tcClean && tcClean.Secondary_Contact_No ? tcClean.Secondary_Contact_No : (rowValues[i] || ''));
            break;
          case 'Primary_Contact_Relation':
            updatedRow.push(tcClean && tcClean.Primary_Contact_Relation ? tcClean.Primary_Contact_Relation : (rowValues[i] || ''));
            break;
          case 'Secondary_Contact_Relation':
            updatedRow.push(tcClean && tcClean.Secondary_Contact_Relation ? tcClean.Secondary_Contact_Relation : (rowValues[i] || ''));
            break;
          case 'Inactive_Type':
            // v7.10: set on deactivation; cleared again on reactivation
            // (like Inactive_Date). Otherwise keep the stored value.
            updatedRow.push(isNewDeactivation ? tcClean.Type : (effStatus === 'Active' && storedStatus === 'Inactive' ? '' : (rowData.Inactive_Type || rowValues[i] || '')));
            break;
          case 'Inactive_Reason':
            updatedRow.push(isNewDeactivation ? (tcClean.Type === 'Transfer' ? tcClean.Leaving_Reason : tcClean.Error_Reason) : (effStatus === 'Active' && storedStatus === 'Inactive' ? '' : (rowData.Inactive_Reason || rowValues[i] || '')));
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
            if (tcLock) { try { tcLock.releaseLock(); } catch (e2) {} }
            return { success: false, message: `Failed to update row after retries: ${e.message}` };
          }
          Utilities.sleep(1000);
        }
      }
      
      // v7.7: Read-back verification — re-read the row and confirm the write
      // actually landed before reporting success.
      const verifyRow = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
      const verifySecIdx = headers.indexOf('Student_Section');
      const verifyStdIdx = headers.indexOf('Std_ID');
      const verifyGlobalIdx = headers.indexOf('Global_ID');
      const verifyNameIdx = headers.indexOf('Student_Name');
      const verifyClassIdx = headers.indexOf('Student_Class');
      const intendedSection = verifySecIdx !== -1 ? updatedRow[verifySecIdx] : undefined;
      const intendedStdId = verifyStdIdx !== -1 ? updatedRow[verifyStdIdx] : undefined;
      const intendedGlobal = verifyGlobalIdx !== -1 ? updatedRow[verifyGlobalIdx] : undefined;
      const savedSection = verifySecIdx !== -1 ? verifyRow[verifySecIdx] : undefined;
      const savedStdId = verifyStdIdx !== -1 ? verifyRow[verifyStdIdx] : undefined;
      const savedGlobal = verifyGlobalIdx !== -1 ? verifyRow[verifyGlobalIdx] : undefined;
      if (String(savedSection != null ? savedSection : '') !== String(intendedSection != null ? intendedSection : '') ||
          String(savedStdId != null ? savedStdId : '') !== String(intendedStdId != null ? intendedStdId : '') ||
          (verifyGlobalIdx !== -1 && String(savedGlobal != null ? savedGlobal : '') !== String(intendedGlobal != null ? intendedGlobal : ''))) {
        Logger.log('Student update read-back mismatch at row ' + rowIndex + ': section saved="' + savedSection + '" intended="' + intendedSection + '", Std_ID saved="' + savedStdId + '" intended="' + intendedStdId + '", Global_ID saved="' + savedGlobal + '" intended="' + intendedGlobal + '"');
        if (tcLock) { try { tcLock.releaseLock(); } catch (e2) {} }
        return { success: false, message: 'Save could not be verified — please refresh and check the sheet.' };
      }

      // v7.10: create the TC row (still inside the lock that allocated the
      // Global_ID / TC number), then return it so the client can open the
      // certificate. The student row is already written and verified above.
      let tcRecord = null;
      if (isNewDeactivation) {
        try {
          const tcRecInfo = getRecInfo_(rowRecId);
          tcRecord = appendTcRowLocked_(ss, tcClean, {
            Global_ID: newGlobalId,
            Std_ID: verifyStdIdx !== -1 ? String(verifyRow[verifyStdIdx] || '') : '',
            Student_Name: verifyNameIdx !== -1 ? String(verifyRow[verifyNameIdx] || '') : '',
            REC_ID: rowRecId,
            REC_Name: tcRecInfo.REC_Name || '',
            LB: tcRecInfo.LB || '',
            Term_ID: '',
            Student_Class: verifyClassIdx !== -1 ? String(verifyRow[verifyClassIdx] || '') : '',
            Student_Section: verifySecIdx !== -1 ? String(verifyRow[verifySecIdx] || '') : ''
          }, username);
        } catch (tcErr) {
          if (tcLock) { try { tcLock.releaseLock(); } catch (e2) {} }
          return { success: false, message: 'Student deactivated, but the transfer certificate could not be created: ' + tcErr.message + '. Please re-activate and deactivate again, or contact the administrator.' };
        }
      }
      if (tcLock) { try { tcLock.releaseLock(); } catch (e2) {} }

      // ✅ Return success with all data (from the re-read row, same shape as before)
      const fmtCell = function (v) { return v == null ? '' : String(v); };
      const fmtDateCell = function (v) {
        if (v instanceof Date) return Utilities.formatDate(v, 'GMT+0500', 'dd-MMM-yyyy');
        return fmtCell(v);
      };
      const fmtCreationCell = function (v) {
        if (v instanceof Date) return Utilities.formatDate(v, 'GMT+0500', 'dd-MMM-yyyy');
        if (v) {
          try {
            const d = new Date(v);
            if (!isNaN(d.getTime())) return Utilities.formatDate(d, 'GMT+0500', 'dd-MMM-yyyy');
          } catch (e) {}
        }
        return fmtCell(v);
      };
      const vr = verifyRow;
return {
  success: true,
  message: 'Student updated successfully: ' + (rowData.Student_Name || rowValues[2] || 'Unknown') + ' (SIRN: ' + (rowData.Std_ID || rowValues[0] || 'Unknown') + ')',
  data: {
    Std_ID: fmtCell(vr[0]),
    Barcode_ID: fmtCell(vr[1]),
    Student_Name: fmtCell(vr[2]),
    Student_Class: fmtCell(vr[3]),
    Student_Section: fmtCell(vr[4]),
    Status: fmtCell(vr[5]),
    Gender: fmtCell(vr[6]),
    Date_of_Joining: fmtDateCell(vr[7]),
    Inactive_Date: fmtDateCell(vr[8]),
    Creation_Date: fmtCreationCell(vr[9]),
    Created_By: fmtCell(vr[10]),
    Last_Modified_Date: fmtDateCell(vr[11]),
    Last_Modified_By: fmtCell(vr[12]),
    REC_ID: fmtCell(vr[13]) || '',
    QR_Printed: fmtCell(vr[14]) || '',
    Unique_ID: fmtCell(vr[15]) || '',
    Global_ID: fmtCell(verifyGlobalIdx !== -1 ? vr[verifyGlobalIdx] : ''),
    Date_of_Birth: fmtDateCell(headers.indexOf('Date_of_Birth') !== -1 ? vr[headers.indexOf('Date_of_Birth')] : ''),
    Father_Name: fmtCell(headers.indexOf('Father_Name') !== -1 ? vr[headers.indexOf('Father_Name')] : ''),
    Address: fmtCell(headers.indexOf('Address') !== -1 ? vr[headers.indexOf('Address')] : ''),
    Phone: fmtCell(headers.indexOf('Phone') !== -1 ? vr[headers.indexOf('Phone')] : ''),
    Inactive_Type: fmtCell(headers.indexOf('Inactive_Type') !== -1 ? vr[headers.indexOf('Inactive_Type')] : ''),
    Inactive_Reason: fmtCell(headers.indexOf('Inactive_Reason') !== -1 ? vr[headers.indexOf('Inactive_Reason')] : ''),
    Previous_Std_ID: fmtCell(headers.indexOf('Previous_Std_ID') !== -1 ? vr[headers.indexOf('Previous_Std_ID')] : ''),
    Previous_REC_ID: fmtCell(headers.indexOf('Previous_REC_ID') !== -1 ? vr[headers.indexOf('Previous_REC_ID')] : ''),
    Secondary_Phone: fmtCell(headers.indexOf('Secondary_Phone') !== -1 ? vr[headers.indexOf('Secondary_Phone')] : ''),
    Primary_Contact_Relation: fmtCell(headers.indexOf('Primary_Contact_Relation') !== -1 ? vr[headers.indexOf('Primary_Contact_Relation')] : ''),
    Secondary_Contact_Relation: fmtCell(headers.indexOf('Secondary_Contact_Relation') !== -1 ? vr[headers.indexOf('Secondary_Contact_Relation')] : ''),
    _rowNumber: rowIndex
  },
  tc: tcRecord
};
    }
    
if (isPrevPermSheet_(sheetName)) stampPrevPermTimestamps_(rowData);


   

    // ✅ OTHER SHEETS (Users, User_Permissions, etc.)

// v7.9 (fix 13): the old exact-case 'Users' never matched the lowercase
// name the client sends, so this guard never ran from the portal — it does now.
if (String(sheetName || '').trim().toLowerCase() === 'users') {
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

    // v7.22: for the two rights, compare EFFECTIVE states, not raw
    // strings -- blank and 'No' both mean the right is off, so a
    // blank -> 'No' save (what Edit sends since HTML v7.23) is not a
    // change; switching a right OFF is allowed for any admin; only
    // switching one ON (off -> 'Yes') stays superadmin-only.
    // UserType / REC_ID keep the exact comparison below.
    if (colName === 'student_delete' || colName === 'canDeleteUsers') {
      const oldEff = oldVal.trim() === 'Yes' ? 'Yes' : 'No';
      const newEff = newVal.trim() === 'Yes' ? 'Yes' : 'No';
      if (oldEff === newEff || oldEff === 'Yes') continue;
      return {
        success: false,
        message: 'Only superadmin can change "' + colName + '".'
      };
    }

    if (oldVal !== newVal) {
      return {
        success: false,
        message: 'Only superadmin can change "' + colName + '".'
      };
    }
  }
}
}  

    // v7.8 (Fix 3): the row being edited must belong to the caller's school
    // (superadmin skipped; fail-open when the caller's school is unknown or
    // the row has no school). The Student_Data branch above has its own
    // v7.7 school check and returns before this point.
    // v7.8 (Fix 4): an absent field keeps the stored value instead of being
    // wiped to ''. An explicitly sent '' still clears the field.
    const existingRowForEdit = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    if (headers.indexOf('REC_ID') !== -1 && updCaller.userType !== 'superadmin' && updCaller.recId) {
      const storedRec = String(existingRowForEdit[headers.indexOf('REC_ID')] || '').trim();
      const incomingRec = String(rowData.REC_ID || '').trim();
      if ((storedRec && storedRec !== updCaller.recId) || (incomingRec && incomingRec !== updCaller.recId)) {
        return { success: false, message: 'This record belongs to another school. You can only edit records of your own school.' };
      }
    }
    const newRow = headers.map(function(header, hIdx) {
      const value = rowData[header];
      // v7.8 (Fix 4): an absent field keeps the stored value.
      if (value === undefined) {
        const kept = existingRowForEdit[hIdx];
        return kept !== undefined && kept !== null ? kept : '';
      }
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

      // v7.21: all six admin rights are stored as 'Yes'/'No', never
      // blank. (The old check used exact-case 'Users' while the client
      // sends 'users', so it never ran, and it covered only two rights.)
      if (sheetName.toLowerCase() === 'users' && ['student_delete', 'canDeleteUsers', 'canQueryAttendance', 'canDownload', 'canAmendSIRN', 'canCorrectAttendance'].includes(header)) {
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
    var sheet = resolveSheet_(ss, sheetName);
    if (!sheet) throw new Error(sheetName + ' sheet not found');

    // v7.8 (Fix 5): row 1 is the header row and can never be deleted.
    const delRowNum = Number(rowNumber);
    if (!Number.isInteger(delRowNum) || delRowNum < 2) {
      return { success: false, message: 'Row 1 is the header row and cannot be deleted.' };
    }

    // Hours config tables (v6.9): rows cannot be deleted — deactivate via Is_Active instead.
    if (isHoursConfigSheet_(sheetName)) return { success: false, message: 'Rows cannot be deleted — set Is_Active to FALSE to deactivate instead.' };

    // Managed tables (v7.15): same rule — deactivate via Is_Active instead.
    if (isManagedTableSheet_(sheetName)) return { success: false, message: 'Rows cannot be deleted — set Is_Active to FALSE to deactivate instead.' };

    // v7.10: Transfer_Certificates rows are history — they are never deleted.
    if (isTransferCertificatesSheet_(sheetName)) return { success: false, message: 'Transfer certificates cannot be deleted. They are kept as history.' };

    // Block deleting the superadmin row itself
    if (String(sheetName || '').trim().toLowerCase() === 'users') {
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

    // v7.8 (Fix 1): only an admin/superadmin Users row may delete.
    const delTypeCol = userHeaders.indexOf('UserType');
    const delCallerType = delTypeCol !== -1 ? String(userRow[delTypeCol] || '').trim() : '';
    const delRecCol = userHeaders.indexOf('REC_ID');
    const delCallerRec = delRecCol !== -1 ? String(userRow[delRecCol] || '').trim() : '';
    if (delCallerType !== 'admin' && delCallerType !== 'superadmin') {
      return { success: false, message: 'Only an admin can make this change.' };
    }

    const isSuperadmin = userRow[2] === 'superadmin';

    // Student delete — student_delete flag OR superadmin
    if (String(sheetName || '').trim().toLowerCase() === 'student_data') {
      const col = userHeaders.indexOf('student_delete');
      const flag = col !== -1 ? (userRow[col] || 'No') : 'No';
      if (!isSuperadmin && flag !== 'Yes') {
        return { success: false, message: 'You do not have permission to delete student records' };
      }
    }

    // User delete — canDeleteUsers flag OR superadmin
    if (String(sheetName || '').trim().toLowerCase() === 'users') {
      const col = userHeaders.indexOf('canDeleteUsers');
      const flag = col !== -1 ? (userRow[col] || 'No') : 'No';
      if (!isSuperadmin && flag !== 'Yes') {
        return { success: false, message: 'You do not have permission to delete user accounts' };
      }
    }

    // v7.8 (Fix 3 + Fix 5): the row must belong to the caller's school
    // (superadmin skipped; fail-open when the caller's school is unknown
    // or the row has no school).
    if (delCallerType !== 'superadmin' && delCallerRec) {
      const delTargetHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const delTargetRecCol = delTargetHeaders.indexOf('REC_ID');
      if (delTargetRecCol !== -1) {
        const delTargetRow = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
        const delTargetRec = String(delTargetRow[delTargetRecCol] || '').trim();
        if (delTargetRec && delTargetRec !== delCallerRec) {
          return { success: false, message: 'This record belongs to another school. You can only delete records of your own school.' };
        }
      }
    }

    // v7.16: FROZEN RECORD — a student with ANY certificate record
    // (Transfer or Error) cannot be deleted. The client hides the Delete
    // button for these rows; this is the server backstop for a direct call.
    if (String(sheetName || '').trim().toLowerCase() === 'student_data') {
      const delHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const delRow = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
      const delG = delHeaders.indexOf('Global_ID') !== -1 ? tcStoredStr_(delRow[delHeaders.indexOf('Global_ID')]) : '';
      const delS = delHeaders.indexOf('Std_ID') !== -1 ? String(delRow[delHeaders.indexOf('Std_ID')] || '').trim() : '';
      const delR = delHeaders.indexOf('REC_ID') !== -1 ? String(delRow[delHeaders.indexOf('REC_ID')] || '').trim() : '';
      if (hasAnyTc_(ss, delG, delS, delR)) {
        return { success: false, message: 'A certificate record was issued for this student. The record is locked.' };
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

// v7.11: RE-JOINING LOOKUP — find inactive students who might be re-joining,
// across ALL schools on purpose (the Global_ID is a cross-school identity
// and a child may re-join from another REC). Admin-gated. query is
// { Std_ID } for an exact Std_ID match, or { Student_Name, Date_of_Birth? }
// for a case-insensitive name contains, narrowed by DOB when given.
// Returns at most 10 rows; [] when nothing matches. A non-admin caller gets
// { success: false, message } instead of an array.
function findPreviousStudent(query, currentUsername) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const gate = requireAdminCaller_(ss, currentUsername);
    if (!gate.ok) return gate.error;
    const q = query || {};
    const qStdId = String(q.Std_ID || '').trim();
    const qName = String(q.Student_Name || '').trim().toLowerCase();
    const qDob = String(q.Date_of_Birth || '').trim();
    if (!qStdId && !qName) return [];
    const sheet = ss.getSheetByName('Student_Data');
    if (!sheet || sheet.getLastRow() <= 1) return [];
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const cIdx = function (h) { return headers.indexOf(h); };
    const out = [];
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (String(r[cIdx('Status')] || '').trim() !== 'Inactive') continue;
      const stdId = String(r[cIdx('Std_ID')] || '').trim();
      const name = String(r[cIdx('Student_Name')] || '').trim();
      if (qStdId) {
        if (stdId !== qStdId) continue;
      } else {
        if (name.toLowerCase().indexOf(qName) === -1) continue;
        if (qDob) {
          const storedDob = tcStoredStr_(r[cIdx('Date_of_Birth')]);
          const qKey = dateKeyForCompare_(qDob);
          const sKey = dateKeyForCompare_(storedDob);
          if ((qKey || qDob) !== (sKey || storedDob)) continue;
        }
      }
      const recId = String(r[cIdx('REC_ID')] || '').trim();
      const recInfo = getRecInfo_(recId);
      out.push({
        Std_ID: stdId,
        Student_Name: name,
        Global_ID: String(r[cIdx('Global_ID')] || '').trim(),
        REC_ID: recId,
        REC_Name: recInfo.REC_Name || '',
        Date_of_Birth: tcStoredStr_(r[cIdx('Date_of_Birth')]),
        Inactive_Date: tcStoredStr_(r[cIdx('Inactive_Date')]),
        Inactive_Type: String(r[cIdx('Inactive_Type')] || '').trim()
      });
      if (out.length >= 10) break;
    }
    return out;
  } catch (e) {
    Logger.log('Error in findPreviousStudent: ' + e.message);
    return { success: false, message: 'Previous-student lookup failed: ' + e.message };
  }
}

// Add a student with barcode generation and audit fields
// Add a student with barcode generation and audit fields
function addStudent(studentData, currentUsername, recId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Student_Data');
    if (!sheet) throw new Error('Student_Data sheet not found');

    // v7.8 (Fix 1): only an admin/superadmin Users row may add students.
    const stuGate = requireAdminCaller_(ss, currentUsername);
    if (!stuGate.ok) return stuGate.error;

    // v7.10: lazily append the new Student_Data columns (Global_ID etc.).
    ensureStudentColumns_(sheet);
    // v7.20 (O1): no cross-school adds — for anyone, superadmin included.
    // A passed recId that differs from the caller's school is refused
    // instead of silently filing the student under the caller's school.
    // A missing recId defaults to the caller's school (and the duplicate
    // Std_ID check below already uses that effective school). The
    // caller's school is the Users-row REC_ID behind the v7.8 admin gate
    // above (first matching row — the same first-row-wins limit as the
    // other v7.8 school checks; only a login token could do better).
    // A caller with no school on their Users row (blank REC_ID) has no
    // school to add to: any explicit recId is a mismatch, and a missing
    // recId cannot be resolved — such a caller needs a Users row at the
    // school they work at.
    let finalRecId = String(recId || '').trim();
    const callerRecId = stuGate.caller ? String(stuGate.caller.recId || '').trim() : '';
    if (finalRecId && finalRecId !== callerRecId) {
      return { success: false, message: 'You can only add students to your own school' + (callerRecId ? ' (REC_ID ' + callerRecId + ')' : '') + '.' };
    }
    if (!finalRecId) finalRecId = callerRecId;
    if (!finalRecId) {
      return { success: false, message: 'Unable to determine school (REC_ID). Please contact administrator.' };
    }

    
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
        if (rowStdId === studentData.Std_ID.toString() && rowRecId === finalRecId) {
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
    
    
    // v7.9 (fix 12): mirror the v7.7 Edit checks on Add — name, class and
    // section against Class_Mapping (fail-open if the mapping cannot be
    // read), Gender Male/Female, Date_of_Joining a real date.
    const addName = String(studentData.Student_Name == null ? '' : studentData.Student_Name).trim();
    if (!addName) return { success: false, message: 'Student_Name is required.' };
    const addGender = String(studentData.Gender == null ? '' : studentData.Gender).trim();
    if (addGender !== 'Male' && addGender !== 'Female') return { success: false, message: 'Gender must be Male or Female.' };
    const addDoj = String(studentData.Date_of_Joining == null ? '' : studentData.Date_of_Joining).trim();
    if (!isRealDateString_(addDoj)) return { success: false, message: 'Date of Joining is not a real date.' };
    // v7.14: Date_of_Birth is optional on Add — a real date only when given.
    const addDob = String(studentData.Date_of_Birth == null ? '' : studentData.Date_of_Birth).trim();
    if (addDob && !isRealDateString_(addDob)) return { success: false, message: 'Date of Birth is not a real date.' };
    const addClass = String(studentData.Student_Class == null ? '' : studentData.Student_Class).trim();
    const addSec = String(studentData.Student_Section == null ? '' : studentData.Student_Section).trim();
    if (!addClass) return { success: false, message: 'Student_Class is required.' };
    if (!addSec) return { success: false, message: 'Student_Section is required.' };
    let addMapping = null;
    try { addMapping = getMappingRowsForRec_(ss, finalRecId); } catch (e) { addMapping = null; }
    if (addMapping && addMapping.length > 0) {
      const addClassNames = addMapping.map(function (m) { return String(m.Class || '').trim(); });
      if (addClassNames.indexOf(addClass) === -1) {
        return { success: false, message: 'Unknown class "' + addClass + '". Choose from: ' + addClassNames.join(', ') };
      }
      const addMapRow = addMapping.filter(function (m) { return String(m.Class || '').trim() === addClass; })[0];
      const addAllowedSecs = addMapRow ? String(addMapRow.Sections || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean) : [];
      if (addAllowedSecs.indexOf(addSec) === -1) {
        return { success: false, message: 'Section must be one of ' + addAllowedSecs.join(', ') + ' for class ' + addClass + '.' };
      }
    }

    // v7.10: a student added as Inactive needs the same reason payload as
    // a deactivation (studentData._tcPayload). A client-supplied Global_ID
    // is ignored — it is allocated server-side below.
    const addStatus = String(studentData.Status || 'Active').trim();
    const addIsInactive = addStatus === 'Inactive';
    let addTcClean = null;
    if (addIsInactive) {
      const addTcCheck = validateTcPayload_(ss, studentData._tcPayload, {}, finalRecId);
      if (addTcCheck.error) return { success: false, message: addTcCheck.error };
      addTcClean = addTcCheck.clean;
    }

    // v7.11: RE-JOINING MERGE — when the Add carries a previous-student
    // link (_previousStdId, _previousRecId if known), the old row must
    // exist and be Inactive (any school). The new row then REUSES the old
    // row's Global_ID and records the link (see the lock block below).
    // The new enrolment gets a NEW Std_ID in its school: re-using the old
    // Std_ID in the same school is still refused by the per-school
    // duplicate rule above, which is why Previous_Std_ID exists.
    const prevStdId = String(studentData._previousStdId || '').trim();
    const prevRecId = String(studentData._previousRecId || '').trim();
    let prevLink = null;
    if (prevStdId) {
      const lStdIdx = headers.indexOf('Std_ID');
      const lRecIdx = headers.indexOf('REC_ID');
      const lStatusIdx = headers.indexOf('Status');
      const lGlobalIdx = headers.indexOf('Global_ID');
      const matches = [];
      if (lastRow > 1) {
        const allRows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
        for (let li = 0; li < allRows.length; li++) {
          const rStd = lStdIdx !== -1 ? String(allRows[li][lStdIdx] || '').trim() : '';
          if (rStd !== prevStdId) continue;
          const rRec = lRecIdx !== -1 ? String(allRows[li][lRecIdx] || '').trim() : '';
          if (prevRecId && rRec !== prevRecId) continue;
          matches.push({
            rowNumber: li + 2,
            recId: rRec,
            status: lStatusIdx !== -1 ? String(allRows[li][lStatusIdx] || '').trim() : '',
            globalId: lGlobalIdx !== -1 ? String(allRows[li][lGlobalIdx] || '').trim() : ''
          });
        }
      }
      if (matches.length === 0) {
        return { success: false, message: 'The previous student (Std_ID ' + prevStdId + ') was not found. Remove the link and try again, or add as a new student.' };
      }
      if (matches.length > 1) {
        return { success: false, message: 'More than one student has Std_ID ' + prevStdId + '. Pick the right school from the lookup and try again.' };
      }
      if (matches[0].status !== 'Inactive') {
        return { success: false, message: 'The previous student (Std_ID ' + prevStdId + ') is still Active, so it cannot be linked as a re-joining student.' };
      }
      prevLink = matches[0];
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
    
    // v7.10/v7.11: build the row by header name (the sheet now has 26 columns).
    // Global_ID is allocated here, under the script lock, together with
    // the TC row when the student is added as Inactive.
    let addGlobalId = '';
    let addTcRecord = null;
    let addRowNum = lastRow + 1;
    const addLock = LockService.getScriptLock();
    if (!addLock.tryLock(30000)) {
      return { success: false, message: 'The system is busy — please try again.' };
    }
    try {
      // v7.11: a re-joining student reuses the old row's Global_ID. If the
      // old row has none yet (deactivated before v7.10), allocate one here
      // and stamp it back on the old row too, so the shared identity stays
      // one Global_ID.
      if (prevLink && prevLink.globalId) {
        addGlobalId = prevLink.globalId;
      } else {
        addGlobalId = nextGlobalIdLocked_(ss);
        if (prevLink) {
          const gCol = headers.indexOf('Global_ID') + 1;
          if (gCol > 0) sheet.getRange(prevLink.rowNumber, gCol).setValue(addGlobalId);
          prevLink.globalId = addGlobalId;
        }
      }
      addRowNum = sheet.getLastRow() + 1;
      const addVals = {
        Std_ID: studentData.Std_ID,
        Barcode_ID: studentData.Barcode_ID,
        Student_Name: studentData.Student_Name || '',
        Student_Class: studentData.Student_Class || '',
        Student_Section: studentData.Student_Section || '',
        Status: studentData.Status || 'Active',
        Gender: studentData.Gender || '',
        Date_of_Joining: studentData.Date_of_Joining || '',
        Inactive_Date: addIsInactive ? formattedDate : '',
        Creation_Date: formattedDate,
        Created_By: currentUser,
        Last_Modified_Date: formattedDate,
        Last_Modified_By: currentUser,
        REC_ID: finalRecId,
        QR_Printed: '',
        Unique_ID: studentData.Unique_ID,
        Global_ID: addGlobalId,
        Date_of_Birth: addTcClean ? addTcClean.Date_of_Birth : (studentData.Date_of_Birth || ''),
        Father_Name: '', // v7.14: no longer collected or stored.
        Address: addTcClean ? addTcClean.Address : (studentData.Address || ''),
        Phone: addTcClean ? addTcClean.Phone : (studentData.Phone || ''),
        Inactive_Type: addTcClean ? addTcClean.Type : '',
        Inactive_Reason: addTcClean ? (addTcClean.Type === 'Transfer' ? addTcClean.Leaving_Reason : addTcClean.Error_Reason) : '',
        Previous_Std_ID: prevLink ? prevStdId : '',
        Previous_REC_ID: prevLink ? prevLink.recId : '',
        Secondary_Phone: addTcClean ? addTcClean.Secondary_Contact_No : '',
        Primary_Contact_Relation: addTcClean ? addTcClean.Primary_Contact_Relation : '',
        Secondary_Contact_Relation: addTcClean ? addTcClean.Secondary_Contact_Relation : ''
      };
      const addRow = headers.map(function (h) { return addVals[h] !== undefined ? addVals[h] : ''; });
      sheet.getRange(addRowNum, 1, 1, headers.length).setValues([addRow]);
      SpreadsheetApp.flush();
      if (addIsInactive) {
        const addRecInfo = getRecInfo_(finalRecId);
        addTcRecord = appendTcRowLocked_(ss, addTcClean, {
          Global_ID: addGlobalId,
          Std_ID: String(studentData.Std_ID),
          Student_Name: String(studentData.Student_Name || ''),
          REC_ID: finalRecId,
          REC_Name: addRecInfo.REC_Name || '',
          LB: addRecInfo.LB || '',
          Term_ID: '',
          Student_Class: String(studentData.Student_Class || ''),
          Student_Section: String(studentData.Student_Section || '')
        }, currentUser);
      }
    } finally {
      try { addLock.releaseLock(); } catch (e) {}
    }
    
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
    
    const addReturnData = {
      ...studentData,
      Global_ID: addGlobalId,
      Date_of_Birth: addTcClean ? addTcClean.Date_of_Birth : (studentData.Date_of_Birth || ''),
      Father_Name: '', // v7.14: no longer collected or stored.
      Address: addTcClean ? addTcClean.Address : (studentData.Address || ''),
      Phone: addTcClean ? addTcClean.Phone : (studentData.Phone || ''),
      Inactive_Type: addTcClean ? addTcClean.Type : '',
      Inactive_Reason: addTcClean ? (addTcClean.Type === 'Transfer' ? addTcClean.Leaving_Reason : addTcClean.Error_Reason) : '',
      Previous_Std_ID: prevLink ? prevStdId : '',
      Previous_REC_ID: prevLink ? prevLink.recId : '',
      Inactive_Date: addIsInactive ? formattedDate : '',
      Creation_Date: formattedDate,
      Created_By: currentUser,
      Last_Modified_Date: formattedDate,
      Last_Modified_By: currentUser,
      _rowNumber: addRowNum
    };
    delete addReturnData._tcPayload;
    delete addReturnData._previousStdId;
    delete addReturnData._previousRecId;
    return {
      success: true,
      message: 'Student added successfully: ' + studentData.Student_Name + ' (SIRN: ' + studentData.Std_ID + ', Barcode: ' + studentData.Barcode_ID + ')',
      data: addReturnData,
      tc: addTcRecord
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
          } else if (['Date_of_Joining', 'Last_Modified_Date', 'Inactive_Date', 'Date_of_Birth'].includes(headers[j]) && data[i][j] instanceof Date) {
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
    // v7.20 (O3): school-filter the rows by the caller's school. The
    // Daily / Monthly / Monthly Day Wise reports already read REC-filtered
    // data, but this function was unfiltered and ACCESS, ITREB and the
    // attendance query all read through it, scoping rows by Std_ID only —
    // so another school's same-Std_ID rows leaked into a school's totals.
    // A row whose REC_ID is blank matches any school (grandfathered old
    // rows). When the caller's school cannot be determined the filter is
    // not applied — the same fail-open rule the file uses elsewhere.
    const callerHeaders = userData[0] || [];
    const callerRecCol = callerHeaders.indexOf('REC_ID');
    const callerRecId = callerRecCol !== -1 && userRow[callerRecCol] != null ? String(userRow[callerRecCol]).trim() : '';
    
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
      status: headers.indexOf('Status'),
      recId: headers.indexOf('REC_ID')
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
        
        // v7.20 (O3): school filter — blank REC_ID matches any school.
        if (callerRecId && columnIndices.recId !== -1) {
          const rowRecId = row[columnIndices.recId] != null ? String(row[columnIndices.recId]).trim() : '';
          if (rowRecId && rowRecId !== callerRecId) return false;
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
function markStudentsAsPrinted(studentIds, username) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // v7.9 (fix 13): only an admin/superadmin may mark, and a student is
    // matched by Std_ID + the caller's school (not Std_ID globally).
    const mpGate = requireAdminCaller_(ss, username);
    if (!mpGate.ok) return mpGate.error;
    const mpCaller = mpGate.caller;
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
    const mpRecCol = headers.indexOf('REC_ID');
    
    let markedCount = 0;
    
    studentIds.forEach(stdId => {
      for (let i = 1; i < data.length; i++) {
        const mpRowRec = mpRecCol !== -1 ? String(data[i][mpRecCol] || '').trim() : '';
        const mpRecOk = mpCaller.userType === 'superadmin' || !mpCaller.recId || (mpRowRec !== '' && mpRowRec === mpCaller.recId);
        if (data[i][stdIdCol].toString() === stdId.toString() && mpRecOk) {
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
    // v7.8 (Fix 1): only an admin/superadmin may generate reserve codes.
    // This function returns a code array or throws, so a refused call throws
    // instead of returning {success:false} — the client shows it as a failure.
    const resGate = requireAdminCaller_(ss, username);
    if (!resGate.ok) throw new Error(resGate.error.message);
    // v7.9 (fix 13): count is capped at 1–100, and the codes must be for the
    // caller's school (superadmin skipped; fail-open when the caller's
    // school is unknown). Array-or-throw contract kept.
    const resCount = Number(count);
    if (!Number.isInteger(resCount) || resCount < 1 || resCount > 100) {
      throw new Error('Count must be between 1 and 100.');
    }
    const resRec = String(recId || '').trim();
    if (resGate.caller.userType !== 'superadmin' && resGate.caller.recId && resRec && resRec !== resGate.caller.recId) {
      throw new Error('This record belongs to another school. You can only generate codes for your own school.');
    }
    count = resCount;

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
function markReserveCodesAsPrinted(barcodeIds, username) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // v7.9 (fix 13): only an admin/superadmin may mark, and a code is
    // matched by Barcode_ID + the caller's school.
    const mrGate = requireAdminCaller_(ss, username);
    if (!mrGate.ok) return mrGate.error;
    const mrCaller = mrGate.caller;
    const sheet = getOrCreateReserveCodesSheet(ss);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const barcodeCol = headers.indexOf('Barcode_ID');
    const printedCol = headers.indexOf('Printed');
    const mrRecCol = headers.indexOf('REC_ID');

    const data = sheet.getDataRange().getValues();
    let markedCount = 0;

    barcodeIds.forEach(code => {
      for (let i = 1; i < data.length; i++) {
        const mrRowRec = mrRecCol !== -1 ? String(data[i][mrRecCol] || '').trim() : '';
        const mrRecOk = mrCaller.userType === 'superadmin' || !mrCaller.recId || (mrRowRec !== '' && mrRowRec === mrCaller.recId);
        if (data[i][barcodeCol] && data[i][barcodeCol].toString() === code.toString() && mrRecOk) {
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

    // v7.8 (Fix 1): the caller must also be an admin/superadmin, not just
    // hold the canCorrectAttendance flag.
    const corrTypeCol = userHeaders.indexOf('UserType');
    const corrCallerType = corrTypeCol !== -1 ? String(userRow[corrTypeCol] || '').trim() : '';
    if (corrCallerType !== 'admin' && corrCallerType !== 'superadmin') {
      return { success: false, message: 'Only an admin can make this change.' };
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
    // v7.9 (fix 11): the record must belong to the caller's school
    // (superadmin skipped; fail-open when the caller's school is unknown).
    const corrRecCol = headers.indexOf('REC_ID');
    const corrRecIdx = userHeaders.indexOf('REC_ID');
    const corrCallerRec = corrRecIdx !== -1 ? String(userRow[corrRecIdx] || '').trim() : '';
    if (corrCallerType !== 'superadmin' && corrCallerRec && corrRecCol !== -1) {
      const corrRowRec = String(rowValues[corrRecCol] || '').trim();
      if (corrRowRec && corrRowRec !== corrCallerRec) {
        return { success: false, message: 'This record belongs to another school. You can only correct records of your own school.' };
      }
    }

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

    // v7.8 (Fix 1): the caller must also be an admin/superadmin, not just
    // hold the canCorrectAttendance flag.
    const corrTypeCol = userHeaders.indexOf('UserType');
    const corrCallerType = corrTypeCol !== -1 ? String(userRow[corrTypeCol] || '').trim() : '';
    if (corrCallerType !== 'admin' && corrCallerType !== 'superadmin') {
      return { success: false, message: 'Only an admin can make this change.' };
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
    // v7.9 (fix 11): the record must belong to the caller's school
    // (superadmin skipped; fail-open when the caller's school is unknown).
    const chRecCol = headers.indexOf('REC_ID');
    const chRecIdx = userHeaders.indexOf('REC_ID');
    const chCallerRec = chRecIdx !== -1 ? String(userRow[chRecIdx] || '').trim() : '';
    if (corrCallerType !== 'superadmin' && chCallerRec && chRecCol !== -1) {
      const chRowRec = String(rowValues[chRecCol] || '').trim();
      if (chRowRec && chRowRec !== chCallerRec) {
        return { success: false, message: 'This record belongs to another school. You can only correct records of your own school.' };
      }
    }

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
function addClassMappingWithSections(className, sectionsStr, recId, username) {
  try {
    // v7.9 (fix 12): only an admin/superadmin may manage class mappings.
    // The gate runs before ensureClassMappingSheet_() so a refused call
    // has no side effects.
    const cmAddSs = SpreadsheetApp.getActiveSpreadsheet();
    const cmAddGate = requireAdminCaller_(cmAddSs, username);
    if (!cmAddGate.ok) return cmAddGate.error;
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
    // v7.9 (fix 12): the mapping must be for the caller's school.
    if (cmAddGate.caller.userType !== 'superadmin' && cmAddGate.caller.recId && finalRecId !== cmAddGate.caller.recId) {
      return { success: false, message: 'This record belongs to another school. You can only manage class mappings of your own school.' };
    }

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

function updateClassMapping(className, newSectionsStr, recId, oldSectionsStr, username) {
  // v7.9 (fix 12): only an admin/superadmin may manage class mappings.
  // The gate runs before the lock so a refused call never holds it.
  const cmUpdSs = SpreadsheetApp.getActiveSpreadsheet();
  const cmUpdGate = requireAdminCaller_(cmUpdSs, username);
  if (!cmUpdGate.ok) return cmUpdGate.error;
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

    // v7.9 (fix 12): the mapping must be for the caller's school.
    if (cmUpdGate.caller.userType !== 'superadmin' && cmUpdGate.caller.recId && finalRecId !== cmUpdGate.caller.recId) {
      return { success: false, message: 'This record belongs to another school. You can only manage class mappings of your own school.' };
    }

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


function deleteClassMappingWithSections(className, recId, username) {
  // v7.9 (fix 12): only an admin/superadmin may manage class mappings.
  // The gate runs before the lock so a refused call never holds it.
  const cmDelSs = SpreadsheetApp.getActiveSpreadsheet();
  const cmDelGate = requireAdminCaller_(cmDelSs, username);
  if (!cmDelGate.ok) return cmDelGate.error;
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

    // v7.9 (fix 12): the mapping must be for the caller's school.
    if (cmDelGate.caller.userType !== 'superadmin' && cmDelGate.caller.recId && finalRecId !== cmDelGate.caller.recId) {
      return { success: false, message: 'This record belongs to another school. You can only manage class mappings of your own school.' };
    }

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
// ============ DAILY HOURS REPORT (v6.6) ============
// Term-wise Daily Hours Report for admins (no teacher scoping — all classes
// visible, like the portal's other reports). Reads Daily_Hours_Setup,
// Episodic_Events and Student_Attendance, all REC-filtered via getSheetData.
// Read-only: creates no sheets or columns and changes no existing function.
// One setup row = one detail row; summaries count each distinct setup row
// once (identical rows are flagged 'Duplicate setup row' in the detail).
// A setup row belongs to the picked term only when its class-section has an
// ACTIVE Class_Terms row for that term covering its date (first match wins,
// via getTermDateRangeForClassSection); classes with no range for the term
// are collected into skippedClasses instead of failing the report.

function safeDailyHoursSheetData_(sheetName, recId) {
  try {
    return getSheetData(sheetName, recId || '');
  } catch (e) {
    return [];
  }
}

// generateDailyHoursReport(termId, recId, classFilter, sectionFilter, periodType, periodValue) -> {
//   termId, termStart, termEnd, detail, classSummary, groupSummary,
//   categorySummary, orphanEvents, skippedClasses }
// Optional filters (v6.7): classFilter = base class ('VII') or ''; sectionFilter = full
// class-section ('VII-A') or ''; periodType = 'full' | 'month' | 'qtd'; periodValue =
// 'yyyy-MM' for month, 'Q1'..'Q4' for qtd. Filters apply after the term-range
// filter; summaries, orphans and skippedClasses are computed over filtered rows.
// termStart/termEnd = overall span of the term (earliest Start_Date to latest
// End_Date among ACTIVE Class_Terms rows for this term + REC); '' when the
// term has no Class_Terms rows at all (then every setup class is skipped).
function generateDailyHoursReport(termId, recId, classFilter, sectionFilter, periodType, periodValue) {
  const term = String(termId || '').trim();
  if (!term) throw new Error('Please select a Term.');
  const rec = String(recId || '').trim();

  const setupRows = safeDailyHoursSheetData_('Daily_Hours_Setup', rec);
  const eventRows = safeDailyHoursSheetData_('Episodic_Events', rec);
  const attRows = safeDailyHoursSheetData_('Student_Attendance', rec);

  const normDate = function (v) {
    const m = String(v == null ? '' : v).trim().match(/^(\d{4})-(\d{2})-(\d{2})/);
    return m ? m[1] + '-' + m[2] + '-' + m[3] : '';
  };
  const normText = function (v) { return String(v == null ? '' : v).trim(); };
  const round2 = function (n) { return Math.round((Number(n) || 0) * 100) / 100; };

  // Overall term span (for the orphan-events filter).
  let termStart = '';
  let termEnd = '';
  getClassTermsForRec_(rec).forEach(function (r) {
    if (!isActiveFlag_(r.Is_Active)) return;
    if (String(r.Term_ID || '').trim().toLowerCase() !== term.toLowerCase()) return;
    const s = parseTermDate_(r.Start_Date);
    const e = parseTermDate_(r.End_Date);
    if (!s || !e) return;
    const sk = termDateKey_(s), ek = termDateKey_(e);
    if (!termStart || sk < termStart) termStart = sk;
    if (!termEnd || ek > termEnd) termEnd = ek;
  });

  // Optional filters (added v6.7; all '' = no filter). They apply AFTER the
  // term-range filter, over the same deduped rows.
  const clsFilter = String(classFilter || '').trim();
  const secFilter = String(sectionFilter || '').trim();
  const perType = String(periodType || '').trim().toLowerCase();
  const perValue = String(periodValue || '').trim();
  // Term-start year (same rule as the ACCESS report): first 4-digit year in
  // the Term_ID, else the earliest Class_Terms Start_Date year.
  const termStartYear = (function () {
    const m = term.match(/(\d{4})/);
    if (m) return parseInt(m[1], 10);
    const s = termStart ? parseInt(termStart.slice(0, 4), 10) : NaN;
    return s || new Date().getFullYear();
  })();
  // Month keys ('yyyy-MM') covered by the picked period, or null = whole term.
  // Quarter ends reuse the ACCESS report's definition exactly:
  // Q1 ends Nov (start year), Q2 ends Feb, Q3 ends May, Q4 ends Aug (start year + 1).
  const DHR_QUARTERS = { Q1: { month: 11, yearOff: 0 }, Q2: { month: 2, yearOff: 1 }, Q3: { month: 5, yearOff: 1 }, Q4: { month: 8, yearOff: 1 } };
  const periodMonths = (function () {
    if (perType === 'month') {
      return /^\d{4}-\d{2}$/.test(perValue) ? [perValue] : null;
    }
    if (perType === 'qtd') {
      const q = DHR_QUARTERS[perValue] || null;
      if (!q) return null;
      const endY = termStartYear + q.yearOff, endM = q.month;
      const keys = [];
      for (let i = 2; i >= 0; i--) {
        let y = endY, m = endM - i;
        while (m < 1) { m += 12; y -= 1; }
        keys.push(y + '-' + ('0' + m).slice(-2));
      }
      return keys;
    }
    return null;
  })();
  const inPeriod = function (dateStr) {
    if (!periodMonths) return true;
    return periodMonths.indexOf(String(dateStr || '').slice(0, 7)) !== -1;
  };

  // Per-class-section term range, cached (first matching ACTIVE row wins).
  const rangeCache = {};
  const rangeFor = function (cs) {
    const key = String(cs || '').trim().toLowerCase();
    if (!(key in rangeCache)) rangeCache[key] = getTermDateRangeForClassSection(rec, cs, term);
    return rangeCache[key];
  };

  // Episodic events, normalised once. Class_Sections is a comma list
  // ('VII-A,VII-B'); the raw string is kept for display.
  const events = eventRows.map(function (ev) {
    return {
      date: normDate(ev.Event_Date),
      name: normText(ev.Event_Name),
      category: normText(ev.Category),
      hours: Number(ev.Hours || 0),
      classes: String(ev.Class_Sections || '').split(',').map(function (c) { return c.trim().toLowerCase(); }).filter(Boolean),
      rawClasses: normText(ev.Class_Sections)
    };
  }).filter(function (ev) { return ev.date && ev.name; });

  // Event <-> setup match: same date, event's class list contains the setup
  // class-section, same category (case-insensitive).
  const eventMatchesSetup = function (ev, date, cs, category) {
    if (ev.date !== date) return false;
    if (ev.classes.indexOf(String(cs || '').trim().toLowerCase()) === -1) return false;
    return ev.category.toLowerCase() === String(category || '').trim().toLowerCase();
  };

  // Attendance index: 'yyyy-MM-dd|class-section(lower)' -> counts.
  // Episodic attendance rows carry a non-blank Event_Name; normal rows do not.
  const attIndex = {};
  attRows.forEach(function (a) {
    const d = normDate(a.Date);
    if (!d) return;
    let cs = normText(a.Class_Section);
    if (!cs) {
      const c = normText(a.Student_Class), s = normText(a.Student_Section);
      if (!c) return;
      cs = s ? c + '-' + s : c;
    }
    const key = d + '|' + cs.toLowerCase();
    const evName = normText(a.Event_Name);
    if (!attIndex[key]) attIndex[key] = { normal: 0, episodicAny: 0, byEvent: {} };
    if (evName) {
      attIndex[key].episodicAny++;
      const ek = evName.toLowerCase();
      attIndex[key].byEvent[ek] = (attIndex[key].byEvent[ek] || 0) + 1;
    } else {
      attIndex[key].normal++;
    }
  });

  const detail = [];
  const deduped = [];
  const skippedSet = {};
  const seenKeys = {};
  setupRows.forEach(function (r) {
    const date = normDate(r.Date);
    const cs = normText(r.Class_Section);
    if (!date || !cs) return;
    // v6.7 filters: class / section first (a class with no term range is only
    // reported as skipped when it passes the class filter), then the period.
    if (secFilter) {
      if (cs.toLowerCase() !== secFilter.toLowerCase()) return;
    } else if (clsFilter) {
      if (baseClassOf_(cs).toLowerCase() !== clsFilter.toLowerCase()) return;
    }
    const range = rangeFor(cs);
    if (!range) { skippedSet[cs] = true; return; }
    if (date < range.start || date > range.end) return;
    if (!inPeriod(date)) return;

    // The sheet's type column is Attendance_Type; fall back to the older
    // 'Type' column that checkDailyHours looks for.
    const typeRaw = normText(r.Attendance_Type) || normText(r.Type) || 'normal';
    const type = typeRaw.toLowerCase() === 'episodic' ? 'Episodic' : 'Normal';
    const category = normText(r.Category);
    const hours = Number(r.Hours || 0);
    const teacher = normText(r.Teacher_ID);

    const dupKey = [date, cs.toLowerCase(), category.toLowerCase(), type, hours, teacher.toLowerCase()].join('|');
    const isDup = !!seenKeys[dupKey];
    seenKeys[dupKey] = true;

    const issues = [];
    if (isDup) issues.push('Duplicate setup row');

    // Episodic link: first matching event wins; normal rows stay blank.
    let linkedEvent = null;
    let eventName = '';
    let eventHours = '';
    if (type === 'Episodic') {
      const matches = events.filter(function (ev) { return eventMatchesSetup(ev, date, cs, category); });
      if (matches.length > 1) issues.push('Multiple events match');
      if (matches.length) {
        linkedEvent = matches[0];
        eventName = linkedEvent.name;
        eventHours = linkedEvent.hours;
        if (Number(linkedEvent.hours) !== hours) {
          issues.push('Event hours differ (setup ' + hours + ' vs event ' + linkedEvent.hours + ')');
        }
      } else {
        issues.push('No matching event');
      }
    }

    // Marked check: normal rows count blank-Event_Name rows; episodic rows
    // with a linked event count rows for that event name; episodic rows
    // with no linked event count any non-blank-Event_Name rows.
    const att = attIndex[date + '|' + cs.toLowerCase()];
    let markedCount = 0;
    if (att) {
      if (type === 'Normal') markedCount = att.normal;
      else if (linkedEvent) markedCount = att.byEvent[linkedEvent.name.toLowerCase()] || 0;
      else markedCount = att.episodicAny;
    }
    const marked = markedCount > 0;

    detail.push({
      Date: date,
      Class_Section: cs,
      Category: category,
      Hours: hours,
      Type: type,
      Teacher_ID: teacher,
      Event: eventName,
      Event_Hours: eventHours,
      Marked: marked ? 'Yes (' + markedCount + ')' : 'No',
      Data_Issues: issues.join('; ')
    });
    if (!isDup) deduped.push({ date: date, cs: cs, hours: hours, type: type, category: category, marked: marked });
  });

  detail.sort(function (a, b) {
    if (a.Date !== b.Date) return a.Date < b.Date ? -1 : 1;
    if (a.Class_Section !== b.Class_Section) return a.Class_Section < b.Class_Section ? -1 : 1;
    return String(a.Category).localeCompare(String(b.Category));
  });

  const classOrder = function (a, b) {
    const ai = MASTER_CLASSES.indexOf(baseClassOf_(a));
    const bi = MASTER_CLASSES.indexOf(baseClassOf_(b));
    if (ai !== bi) return (ai === -1 ? 999 : ai) - (bi === -1 ? 999 : bi);
    return a < b ? -1 : a > b ? 1 : 0;
  };

  // 1. Per class-section (from DEDUPED rows only).
  // A date counts as marked only when EVERY setup row on that date is
  // marked, so a partly-marked day (e.g. normal marked, episodic not)
  // still shows up under Days_Not_Marked instead of hiding the gap.
  const byClass = {};
  deduped.forEach(function (d) {
    if (!byClass[d.cs]) byClass[d.cs] = { dates: {}, unmarkedDates: {}, normal: 0, episodic: 0 };
    const b = byClass[d.cs];
    b.dates[d.date] = true;
    if (!d.marked) b.unmarkedDates[d.date] = true;
    if (d.type === 'Episodic') b.episodic += d.hours; else b.normal += d.hours;
  });
  const classSummary = Object.keys(byClass).sort(classOrder).map(function (cs) {
    const b = byClass[cs];
    const days = Object.keys(b.dates).length;
    const markedDays = days - Object.keys(b.unmarkedDates).length;
    return {
      Class_Section: cs,
      Days_Setup: days,
      Normal_Hours: round2(b.normal),
      Episodic_Hours: round2(b.episodic),
      Total_Hours: round2(b.normal + b.episodic),
      Days_Marked: markedDays,
      Days_Not_Marked: days - markedDays
    };
  });

  // 2. Per group (same group definitions as the ACCESS Hours Report) + Grand Total.
  const GROUPS = [
    { name: 'Pre-Primary', classes: ['IBT', 'NUR', 'KGI', 'KGII'] },
    { name: 'Lower Primary', classes: ['I', 'II', 'III'] },
    { name: 'Upper Primary', classes: ['IV', 'V', 'VI'] },
    { name: 'Secondary', classes: ['VII', 'VIII', 'IX', 'X', 'XI', 'HRE'] }
  ];
  const groupSummary = [];
  const grand = { Group: 'Grand Total', Days_Setup: 0, Normal_Hours: 0, Episodic_Hours: 0, Total_Hours: 0, Days_Marked: 0, Days_Not_Marked: 0 };
  GROUPS.forEach(function (g) {
    const rows = classSummary.filter(function (r) { return g.classes.indexOf(baseClassOf_(r.Class_Section)) !== -1; });
    if (!rows.length) return;
    const tot = { Group: g.name, Days_Setup: 0, Normal_Hours: 0, Episodic_Hours: 0, Total_Hours: 0, Days_Marked: 0, Days_Not_Marked: 0 };
    rows.forEach(function (r) {
      tot.Days_Setup += r.Days_Setup;
      tot.Normal_Hours += r.Normal_Hours;
      tot.Episodic_Hours += r.Episodic_Hours;
      tot.Total_Hours += r.Total_Hours;
      tot.Days_Marked += r.Days_Marked;
      tot.Days_Not_Marked += r.Days_Not_Marked;
    });
    tot.Normal_Hours = round2(tot.Normal_Hours);
    tot.Episodic_Hours = round2(tot.Episodic_Hours);
    tot.Total_Hours = round2(tot.Total_Hours);
    groupSummary.push(tot);
    grand.Days_Setup += tot.Days_Setup;
    grand.Normal_Hours += tot.Normal_Hours;
    grand.Episodic_Hours += tot.Episodic_Hours;
    grand.Total_Hours += tot.Total_Hours;
    grand.Days_Marked += tot.Days_Marked;
    grand.Days_Not_Marked += tot.Days_Not_Marked;
  });
  grand.Normal_Hours = round2(grand.Normal_Hours);
  grand.Episodic_Hours = round2(grand.Episodic_Hours);
  grand.Total_Hours = round2(grand.Total_Hours);
  groupSummary.push(grand);

  // 3. Hours by category (from DEDUPED rows only).
  const byCat = {};
  deduped.forEach(function (d) {
    if (!byCat[d.category]) byCat[d.category] = { normal: 0, episodic: 0 };
    if (d.type === 'Episodic') byCat[d.category].episodic += d.hours;
    else byCat[d.category].normal += d.hours;
  });
  const categorySummary = Object.keys(byCat).sort().map(function (c) {
    const b = byCat[c];
    return {
      Category: c,
      Normal_Hours: round2(b.normal),
      Episodic_Hours: round2(b.episodic),
      Total_Hours: round2(b.normal + b.episodic)
    };
  });

  // 4. Episodic events without a setup row: events inside the term's overall
  // span with no in-range episodic setup row matching them (same match rule
  // reversed). Each event row is listed separately, even if re-created.
  const orphanEvents = [];
  if (termStart && termEnd) {
    events.forEach(function (ev) {
      if (ev.date < termStart || ev.date > termEnd) return;
      if (!inPeriod(ev.date)) return;
      if (secFilter) {
        if (ev.classes.indexOf(secFilter.toLowerCase()) === -1) return;
      } else if (clsFilter) {
        const coversClass = ev.classes.some(function (c) { return baseClassOf_(c).toLowerCase() === clsFilter.toLowerCase(); });
        if (!coversClass) return;
      }
      const hasSetup = deduped.some(function (d) {
        return d.type === 'Episodic' && eventMatchesSetup(ev, d.date, d.cs, d.category);
      });
      if (!hasSetup) {
        orphanEvents.push({
          Event_Date: ev.date,
          Event_Name: ev.name,
          Category: ev.category,
          Hours: ev.hours,
          Class_Sections: ev.rawClasses
        });
      }
    });
  }
  orphanEvents.sort(function (a, b) {
    if (a.Event_Date !== b.Event_Date) return a.Event_Date < b.Event_Date ? -1 : 1;
    return String(a.Event_Name).localeCompare(String(b.Event_Name));
  });

  return {
    termId: term,
    termStart: termStart,
    termEnd: termEnd,
    detail: detail,
    classSummary: classSummary,
    groupSummary: groupSummary,
    categorySummary: categorySummary,
    orphanEvents: orphanEvents,
    skippedClasses: Object.keys(skippedSet).sort(classOrder)
  };
}
