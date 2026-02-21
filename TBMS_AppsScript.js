// ============================================================
//  TBMS - The Bap Management System v2.4
//  Google Apps Script Backend (Code.gs)
//  Deployed: 2026-02-17
//  URL: https://script.google.com/macros/s/AKfycbwsTnzSbU67SH2xC6Hlh7eTgv8eYdYvCFPG-W7wwp5qOg7wipvfY1x9slBCLwR2WtE/exec
// ============================================================
//  SETUP:
//  1. Google Drive > New > Google Sheets > Name "TBMS Database"
//  2. Extensions > Apps Script
//  3. Delete Code.gs contents, paste this code, Save
//  4. Run > initSheets (grant permissions on first run)
//  5. Deploy > New Deployment > Web app
//     - Execute as: Me / Who has access: Anyone
//  6. Copy Web App URL > paste into TBMS.html setup
// ============================================================

const SHEETS = {
  Users:         ['id','username','password','name','role','email','storeId','permissions'],
  Stores:        ['id','code','name','company','companyNo','vatNo','vatQuarter','address','phone','email','manager','memo','active'],
  Staff:         ['id','storeId','name','nickName','clothSize','kioskPwd','dob','address','niNo','eVisa','mobile','startDate','leftDate','rate','sortCode','accountNo','email','memo','active'],
  Attendance:    ['id','staffId','storeId','date','clockIn','clockOut','photoIn','photoOut','source'],
  Suppliers:     ['id','name','email','phone','address','website','memo','active'],
  StockTemplate: ['id','category','name','unit','min','sortOrder','supplier1','supplier2','supplier3','memo'],
  StoreStock:    ['storeId','itemId','category','name','unit','min','qty'],
  StockCount:    ['id','storeId','week','countDate','itemId','category','name','unit','qty','submittedBy','submittedAt'],
  WeeklySales:   ['id','storeId','week','weekStart','totalSales','notes','submittedBy','submittedAt'],
  StaffMessages: ['id','storeId','staffId','type','message','active','createdBy','createdAt'],
  EditLog:        ['id','attendanceId','staffId','storeId','date','fieldChanged','oldValue','newValue','editTimestamp','kioskVersion'],
  TimeChangeReq:  ['id','attendanceId','staffId','storeId','date','field','currentValue','requestedValue','reason','status','kioskVersion','createdAt','reviewedBy','reviewedAt','acknowledgedAt'],
  DiaryConfig:    ['id','storeId','configType','name','sortOrder','active'],
  DiaryEntry:     ['id','storeId','date','openChecks','fridgeTemps','deliveries','cookingTemps','coolingRecords','leftovers','closingChecks','holdTemps','extraChecks','issues','signedBy','submittedBy','submittedAt']
};

// Fields that should remain numeric
const NUMERIC_FIELDS = ['rate','min','qty','totalSales','sortOrder'];
// Fields that are boolean
const BOOL_FIELDS = ['active'];
// Fields that are time (HH:mm format)
const TIME_FIELDS = ['clockIn','clockOut'];
// Fields that are date (yyyy-MM-dd format)
const DATE_FIELDS = ['dob','startDate','leftDate','date','countDate','weekStart'];

// Normalize header for fuzzy matching: "Nick Name" → "nickname", "nickName" → "nickname"
function normalizeHeader(h) {
  return String(h).replace(/[\s_-]+/g, '').toLowerCase();
}

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'ping';
    var sheet = (e && e.parameter) ? e.parameter.sheet : null;
    var result;
    switch(action) {
      case 'getAll':   result = getAllData(); break;
      case 'getSheet': result = getSheetData(sheet); break;
      case 'getStoreData': result = getStoreData(e.parameter.store, e.parameter.sheets); break;
      case 'getSetting': result = getSetting(e.parameter.key); break;
      case 'init':     result = initSheets(); break;
      case 'ping':     result = {status:'ok', time: new Date().toISOString(), version:'TBMS 2.1'}; break;
      default:         result = {error:'Unknown action: '+action};
    }
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({error: err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var data = JSON.parse(e.postData.contents);
    var result;
    switch(data.action) {
      case 'saveSheet':  result = saveSheet(data.sheet, data.rows); break;
      case 'upsert':     result = upsertRow(data.sheet, data.row); break;
      case 'deleteRow':  result = deleteRow(data.sheet, data.id); break;
      case 'appendRow':  result = appendNewRow(data.sheet, data.row); break;
      case 'clockInPhoto':  result = clockInWithPhoto(data); break;
      case 'clockOutPhoto': result = clockOutWithPhoto(data); break;
      case 'saveStockCount': result = saveStockCount(data); break;
      case 'editLog':    result = appendEditLog(data); break;
      case 'timeChangeReq': result = createTimeChangeReq(data); break;
      case 'reviewTimeReq': result = reviewTimeChangeReq(data); break;
      case 'ackTimeReq':   result = ackTimeChangeReq(data); break;
      case 'saveSetting': result = saveSetting(data.key, data.value); break;
      case 'initData':   result = initWithData(data); break;
      default:           result = {error:'Unknown action: '+data.action};
    }
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({error: err.toString()})).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
//  Settings (PropertiesService) — persistent key-value store
// ============================================================
function saveSetting(key, value) {
  if (!key) return {error: 'key is required'};
  PropertiesService.getScriptProperties().setProperty('setting_' + key, JSON.stringify(value));
  return {status: 'ok', key: key};
}
function getSetting(key) {
  if (!key) return {error: 'key is required'};
  var raw = PropertiesService.getScriptProperties().getProperty('setting_' + key);
  return {status: 'ok', key: key, value: raw ? JSON.parse(raw) : null};
}

// ============================================================
//  initSheets — creates sheets, fixes headers, remaps data
//  Uses normalizeHeader() to match "Nick Name" → "nickName" etc.
// ============================================================
function initSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var created = [];
  var updated = [];
  for (var name in SHEETS) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) { sheet = ss.insertSheet(name); created.push(name); }
    var headers = SHEETS[name];
    var lastCol = sheet.getLastColumn();
    var firstRow = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    if (!firstRow[0]) {
      // Brand new sheet — write headers
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    } else {
      // Existing sheet — read current headers (stop at first empty)
      var currentHeaders = [];
      for (var ci = 0; ci < firstRow.length; ci++) {
        var ch = String(firstRow[ci]).trim();
        if (ch === '') break;
        currentHeaders.push(ch);
      }
      // Check if headers match EXACTLY (name and count)
      var needsUpdate = false;
      if (currentHeaders.length !== headers.length) {
        needsUpdate = true;
      } else {
        for (var hi = 0; hi < headers.length; hi++) {
          if (currentHeaders[hi] !== headers[hi]) {
            needsUpdate = true;
            break;
          }
        }
      }
      if (needsUpdate) {
        // Build normalized mapping: normalizedOldHeader → column index
        var normMap = {};
        for (var mi = 0; mi < currentHeaders.length; mi++) {
          normMap[normalizeHeader(currentHeaders[mi])] = mi;
        }
        // Remap existing data to new column order
        var dataRows = sheet.getLastRow() - 1;
        if (dataRows > 0) {
          var oldData = sheet.getRange(2, 1, dataRows, currentHeaders.length).getValues();
          var newData = oldData.map(function(row) {
            return headers.map(function(h) {
              var normH = normalizeHeader(h);
              var oldIdx = normMap[normH];
              return (oldIdx !== undefined && oldIdx < row.length) ? row[oldIdx] : '';
            });
          });
          // Clear extra columns if old sheet had more
          if (currentHeaders.length > headers.length) {
            sheet.getRange(1, headers.length + 1, sheet.getLastRow(), currentHeaders.length - headers.length).clearContent();
          }
          // Write corrected headers and remapped data
          sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
          sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
          sheet.getRange(2, 1, dataRows, headers.length).setValues(newData);
        } else {
          // No data rows — just fix headers
          if (currentHeaders.length > headers.length) {
            sheet.getRange(1, headers.length + 1, 1, currentHeaders.length - headers.length).clearContent();
          }
          sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
          sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
        }
        sheet.setFrozenRows(1);
        updated.push(name);
      }
    }
  }
  // Set clockIn/clockOut columns to plain text format to prevent auto-conversion
  var attSheet = ss.getSheetByName('Attendance');
  if (attSheet) {
    var attHeaders = SHEETS.Attendance;
    var ciCol = attHeaders.indexOf('clockIn') + 1;  // 1-based
    var coCol = attHeaders.indexOf('clockOut') + 1;
    if (ciCol > 0) attSheet.getRange(2, ciCol, attSheet.getMaxRows() - 1, 1).setNumberFormat('@');
    if (coCol > 0) attSheet.getRange(2, coCol, attSheet.getMaxRows() - 1, 1).setNumberFormat('@');
  }
  var def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) { try { ss.deleteSheet(def); } catch(e) {} }
  return {status:'ok', created: created, updated: updated};
}

function getAllData() {
  var result = {};
  var skip = {DiaryConfig:1, DiaryEntry:1};
  for (var name in SHEETS) { if (!skip[name]) result[name] = readSheet(name); }
  return {status:'ok', data: result};
}

// ============================================================
//  readSheet — Uses SHEETS[name] for property keys + header
//  name matching for correct column mapping.
//  "Nick Name" in sheet → matched to "nickName" in SHEETS → obj.nickName
// ============================================================
function readSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var canonicalHeaders = SHEETS[name];
  // Read actual sheet headers (stop at first empty)
  var actualHeaders = [];
  for (var k = 0; k < data[0].length; k++) {
    var hh = String(data[0][k]).trim();
    if (hh === '') break;
    actualHeaders.push(hh);
  }
  // Build normalized map: normalizedName → actual column index
  var normToCol = {};
  for (var ai = 0; ai < actualHeaders.length; ai++) {
    normToCol[normalizeHeader(actualHeaders[ai])] = ai;
  }
  // Build column map: canonical field name → actual column index
  var colMap = [];
  for (var ci = 0; ci < canonicalHeaders.length; ci++) {
    var normKey = normalizeHeader(canonicalHeaders[ci]);
    colMap.push(normToCol[normKey] !== undefined ? normToCol[normKey] : -1);
  }
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < canonicalHeaders.length; j++) {
      var h = canonicalHeaders[j];
      var colIdx = colMap[j];
      var val = (colIdx >= 0 && colIdx < data[i].length) ? data[i][colIdx] : '';
      if (val instanceof Date && TIME_FIELDS.indexOf(h) >= 0) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'HH:mm');
      } else if (val instanceof Date && DATE_FIELDS.indexOf(h) >= 0) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else if (BOOL_FIELDS.indexOf(h) >= 0) {
        val = (val === true || val === 'true' || val === 'TRUE');
      } else if (NUMERIC_FIELDS.indexOf(h) >= 0) {
        val = Number(val) || 0;
      } else {
        val = (val === '' || val === null || val === undefined) ? '' : String(val);
      }
      obj[h] = val;
    }
    // Check if row has any meaningful data (skip truly empty rows)
    var hasData = false;
    for (var hk = 0; hk < canonicalHeaders.length; hk++) {
      var fld = canonicalHeaders[hk];
      if (BOOL_FIELDS.indexOf(fld) >= 0 || NUMERIC_FIELDS.indexOf(fld) >= 0) continue;
      if (obj[fld] !== '' && obj[fld] !== null && obj[fld] !== undefined) { hasData = true; break; }
    }
    // Auto-generate id if missing
    if (hasData && canonicalHeaders[0] === 'id' && !obj.id) {
      obj.id = name.toLowerCase().substring(0,3) + '_' + Utilities.getUuid().substring(0,8);
    }
    if (hasData) {
      rows.push(obj);
    }
  }
  return rows;
}

function getSheetData(name) {
  if (!name || !SHEETS[name]) return {error: 'Invalid sheet name'};
  return {status:'ok', data: readSheet(name)};
}

// Lightweight endpoint: fetch only specific sheets, filtered by storeId
// Usage: ?action=getStoreData&store=PAB&sheets=Staff,Attendance,TimeChangeReq
function getStoreData(storeId, sheetNames) {
  if (!storeId) return {error: 'Missing store parameter'};
  var names = sheetNames ? sheetNames.split(',') : ['Staff','Attendance'];
  var result = {};
  for (var i = 0; i < names.length; i++) {
    var name = names[i].trim();
    if (!SHEETS[name]) continue;
    var rows = readSheet(name);
    // Filter by storeId for sheets that have it
    if (SHEETS[name].indexOf('storeId') >= 0) {
      rows = rows.filter(function(r) { return String(r.storeId) === String(storeId); });
    }
    result[name] = rows;
  }
  return {status:'ok', data: result};
}

function saveSheet(name, rows) {
  if (!name || !SHEETS[name]) return {error: 'Invalid sheet name'};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) { sheet = ss.insertSheet(name); }
  var headers = SHEETS[name];
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows && rows.length > 0) {
    var values = rows.map(function(row) {
      return headers.map(function(h) { var v = row[h]; return (v === null || v === undefined) ? '' : v; });
    });
    sheet.getRange(2, 1, values.length, headers.length).setValues(values);
  }
  return {status:'ok', count: rows ? rows.length : 0};
}

// ============================================================
//  upsertRow — Fixes headers first, then writes data using
//  SHEETS[name] for column mapping. This ensures "Nick Name"
//  gets corrected to "nickName" on every save.
// ============================================================
function upsertRow(name, row) {
  if (!name || !SHEETS[name] || !row) return {error: 'Invalid parameters'};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) return {error: 'Sheet not found: ' + name};
  var headers = SHEETS[name];
  // Always ensure header row matches canonical names
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  // Now read data (headers are guaranteed correct)
  var data = sheet.getDataRange().getValues();
  var idValue = row[headers[0]];
  var found = false;
  var targetRow = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(idValue)) {
      // Merge: preserve existing values when new value is empty
      var vals = headers.map(function(h, idx) {
        var newVal = row[h];
        var existingVal = data[i][idx];
        // If new value is empty/null/undefined but existing has data, keep existing
        if ((newVal === null || newVal === undefined || newVal === '') && existingVal !== '' && existingVal !== null && existingVal !== undefined) {
          return existingVal;
        }
        return (newVal === null || newVal === undefined) ? '' : newVal;
      });
      sheet.getRange(i + 1, 1, 1, headers.length).setValues([vals]);
      found = true; targetRow = i + 1; break;
    }
  }
  if (!found) {
    var vals = headers.map(function(h) { var v = row[h]; return (v === null || v === undefined) ? '' : v; });
    sheet.appendRow(vals);
    targetRow = sheet.getLastRow();
  }
  // Force plain text for time fields
  if (name === 'Attendance' && targetRow > 0) {
    var ciCol = headers.indexOf('clockIn') + 1;
    var coCol = headers.indexOf('clockOut') + 1;
    if (ciCol > 0) { sheet.getRange(targetRow, ciCol).setNumberFormat('@'); if (row.clockIn) sheet.getRange(targetRow, ciCol).setValue(String(row.clockIn)); }
    if (coCol > 0) { sheet.getRange(targetRow, coCol).setNumberFormat('@'); if (row.clockOut) sheet.getRange(targetRow, coCol).setValue(String(row.clockOut)); }
  }
  return {status:'ok', action: found ? 'updated' : 'inserted'};
}

function deleteRow(name, id) {
  if (!name || !SHEETS[name]) return {error: 'Invalid parameters'};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) return {error: 'Sheet not found'};
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) { sheet.deleteRow(i + 1); return {status:'ok', deleted: true}; }
  }
  return {status:'ok', deleted: false};
}

// ============================================================
//  appendNewRow — Fixes headers, then appends using SHEETS[name]
// ============================================================
function appendNewRow(name, row) {
  if (!name || !SHEETS[name] || !row) return {error: 'Invalid parameters'};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) return {error: 'Sheet not found'};
  var headers = SHEETS[name];
  // Always ensure header row matches canonical names
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  var vals = headers.map(function(h) { var v = row[h]; return (v === null || v === undefined) ? '' : v; });
  sheet.appendRow(vals);
  // Force plain text for time fields to prevent Google Sheets auto-conversion
  if (name === 'Attendance') {
    var lastRow = sheet.getLastRow();
    var ciCol = headers.indexOf('clockIn') + 1;
    var coCol = headers.indexOf('clockOut') + 1;
    if (ciCol > 0) { sheet.getRange(lastRow, ciCol).setNumberFormat('@'); if (row.clockIn) sheet.getRange(lastRow, ciCol).setValue(String(row.clockIn)); }
    if (coCol > 0) { sheet.getRange(lastRow, coCol).setNumberFormat('@'); if (row.clockOut) sheet.getRange(lastRow, coCol).setValue(String(row.clockOut)); }
  }
  return {status:'ok'};
}

function initWithData(data) {
  initSheets();
  var results = {};
  if (data.sheets) {
    for (var name in data.sheets) {
      if (SHEETS[name]) {
        var existing = readSheet(name);
        if (existing.length === 0) {
          results[name] = saveSheet(name, data.sheets[name]);
        } else {
          results[name] = {status:'skipped', reason:'data exists', count: existing.length};
        }
      }
    }
  }
  return {status:'ok', results: results};
}

// ============================================================
//  Fix existing time data in Attendance sheet
//  Run this ONCE after deploying updated code to convert
//  Date objects (1899-12-30 23:02:00) back to plain text "23:02"
// ============================================================
function fixTimeColumns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Attendance');
  if (!sheet) return {status:'error', message:'Attendance sheet not found'};
  var headers = SHEETS.Attendance;
  var ciCol = headers.indexOf('clockIn') + 1;
  var coCol = headers.indexOf('clockOut') + 1;
  if (ciCol <= 0) return {status:'error', message:'clockIn column not found'};
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return {status:'ok', message:'No data rows', fixed: 0};
  var fixed = 0;
  // Set entire columns to plain text first
  sheet.getRange(2, ciCol, lastRow - 1, 1).setNumberFormat('@');
  if (coCol > 0) sheet.getRange(2, coCol, lastRow - 1, 1).setNumberFormat('@');
  // Now read and fix each cell
  for (var row = 2; row <= lastRow; row++) {
    // Fix clockIn
    var ciCell = sheet.getRange(row, ciCol);
    var ciVal = ciCell.getValue();
    if (ciVal instanceof Date) {
      var timeStr = Utilities.formatDate(ciVal, Session.getScriptTimeZone(), 'HH:mm');
      ciCell.setValue(timeStr);
      fixed++;
    }
    // Fix clockOut
    if (coCol > 0) {
      var coCell = sheet.getRange(row, coCol);
      var coVal = coCell.getValue();
      if (coVal instanceof Date) {
        var timeStr2 = Utilities.formatDate(coVal, Session.getScriptTimeZone(), 'HH:mm');
        coCell.setValue(timeStr2);
        fixed++;
      }
    }
  }
  return {status:'ok', message:'Fixed ' + fixed + ' cells', fixed: fixed};
}

// ============================================================
//  Selfie Photo Functions — Save photos to Google Drive
// ============================================================
function getPhotoFolder() {
  var folders = DriveApp.getFoldersByName('TBMS_Photos');
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder('TBMS_Photos');
}

function savePhotoToDrive(base64Data, fileName) {
  try {
    var folder = getPhotoFolder();
    var decoded = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(decoded, 'image/jpeg', fileName + '.jpg');
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getId();
  } catch(e) {
    Logger.log('Photo save error: ' + e.toString());
    return '';
  }
}

function clockInWithPhoto(data) {
  if (!data.row) return {error: 'Missing row data'};
  if (data.photo) {
    var fileName = data.row.staffId + '_' + data.row.date + '_in_' + Date.now();
    var fileId = savePhotoToDrive(data.photo, fileName);
    data.row.photoIn = fileId;
  }
  var result = appendNewRow('Attendance', data.row);
  return {status: 'ok', row: {photoIn: data.row.photoIn || ''}};
}

function clockOutWithPhoto(data) {
  if (!data.row) return {error: 'Missing row data'};
  var photoOutId = '';
  if (data.photo) {
    var fileName = data.row.staffId + '_' + data.row.date + '_out_' + Date.now();
    photoOutId = savePhotoToDrive(data.photo, fileName);
  }
  // PARTIAL UPDATE: only update clockOut + photoOut columns, never touch other fields
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Attendance');
  if (!sheet) return {error: 'Attendance sheet not found'};
  var headers = SHEETS['Attendance'];
  var allData = sheet.getDataRange().getValues();
  var rowId = data.row[headers[0]];
  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][0]) === String(rowId)) {
      var coCol = headers.indexOf('clockOut') + 1;
      var poCol = headers.indexOf('photoOut') + 1;
      if (coCol > 0 && data.row.clockOut) {
        sheet.getRange(i + 1, coCol).setNumberFormat('@');
        sheet.getRange(i + 1, coCol).setValue(String(data.row.clockOut));
      }
      if (poCol > 0 && photoOutId) {
        sheet.getRange(i + 1, poCol).setValue(photoOutId);
      }
      return {status: 'ok', action: 'updated'};
    }
  }
  return {error: 'Attendance record not found: ' + rowId};
}

// ============================================================
//  Stock Count — Batch save weekly stock count + sync StoreStock
// ============================================================
function getISOWeek(dateStr) {
  var d = new Date(dateStr);
  d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  var dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  var weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return d.getUTCFullYear() * 100 + weekNo;
}

function saveStockCount(data) {
  if (!data.storeId || !data.countDate || !data.items || !data.items.length) {
    return {error: 'Missing required fields: storeId, countDate, items'};
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scSheet = ss.getSheetByName('StockCount');
  if (!scSheet) return {error: 'StockCount sheet not found. Run initSheets() first.'};
  var headers = SHEETS.StockCount;
  // Ensure headers
  scSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var week = getISOWeek(data.countDate);
  var timestamp = new Date().toISOString();
  var submittedBy = data.submittedBy || '';

  // Build rows for batch append
  var rows = [];
  for (var i = 0; i < data.items.length; i++) {
    var item = data.items[i];
    var id = 'sc' + Date.now().toString(36) + Math.random().toString(36).substr(2, 4) + i;
    rows.push([
      id,
      data.storeId,
      week,
      data.countDate,
      item.itemId,
      item.category || '',
      item.name || '',
      item.unit || '',
      Number(item.qty) || 0,
      submittedBy,
      timestamp
    ]);
  }

  // Batch append all rows at once
  if (rows.length > 0) {
    scSheet.getRange(scSheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
    // Force text format for date column
    var lastRow = scSheet.getLastRow();
    var dateCol = headers.indexOf('countDate') + 1;
    if (dateCol > 0) {
      scSheet.getRange(lastRow - rows.length + 1, dateCol, rows.length, 1).setNumberFormat('@');
    }
  }

  // Sync StoreStock qty for each item
  var ssSheet = ss.getSheetByName('StoreStock');
  if (ssSheet) {
    var ssData = ssSheet.getDataRange().getValues();
    var ssHeaders = SHEETS.StoreStock;
    var storeCol = ssHeaders.indexOf('storeId');
    var itemCol = ssHeaders.indexOf('itemId');
    var qtyCol = ssHeaders.indexOf('qty');
    for (var j = 0; j < data.items.length; j++) {
      for (var r = 1; r < ssData.length; r++) {
        if (String(ssData[r][storeCol]) === String(data.storeId) && String(ssData[r][itemCol]) === String(data.items[j].itemId)) {
          ssSheet.getRange(r + 1, qtyCol + 1).setValue(Number(data.items[j].qty) || 0);
          break;
        }
      }
    }
  }

  return {status: 'ok', week: week, count: rows.length};
}

// ============================================================
//  EditLog — Append edit audit log entry
// ============================================================
function appendEditLog(data) {
  if (!data.attendanceId || !data.staffId || !data.fieldChanged) {
    return {error: 'Missing required fields: attendanceId, staffId, fieldChanged'};
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('EditLog');
  if (!sheet) { sheet = ss.insertSheet('EditLog'); }
  var headers = SHEETS.EditLog;
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  var row = {
    id: 'el' + Date.now().toString(36) + Math.random().toString(36).substr(2, 4),
    attendanceId: data.attendanceId || '',
    staffId: data.staffId || '',
    storeId: data.storeId || '',
    date: data.date || '',
    fieldChanged: data.fieldChanged || '',
    oldValue: data.oldValue || '',
    newValue: data.newValue || '',
    editTimestamp: new Date().toISOString(),
    kioskVersion: data.kioskVersion || ''
  };
  var vals = headers.map(function(h) { return row[h] || ''; });
  var newRow = sheet.getLastRow() + 1;
  var range = sheet.getRange(newRow, 1, 1, vals.length);
  range.setNumberFormat('@');
  range.setValues([vals]);
  return {status: 'ok', id: row.id};
}

// ============================================================
//  TimeChangeReq — Create a time change request
// ============================================================
function createTimeChangeReq(data) {
  if (!data.attendanceId || !data.staffId || !data.field) {
    return {error: 'Missing required fields'};
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('TimeChangeReq');
  if (!sheet) { sheet = ss.insertSheet('TimeChangeReq'); }
  var headers = SHEETS.TimeChangeReq;
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  var row = {
    id: 'tcr' + Date.now().toString(36) + Math.random().toString(36).substr(2, 4),
    attendanceId: data.attendanceId || '',
    staffId: data.staffId || '',
    storeId: data.storeId || '',
    date: data.date || '',
    field: data.field || '',
    currentValue: data.currentValue || '',
    requestedValue: data.requestedValue || '',
    reason: data.reason || '',
    status: 'pending',
    kioskVersion: data.kioskVersion || '',
    createdAt: new Date().toISOString(),
    reviewedBy: '',
    reviewedAt: ''
  };
  var vals = headers.map(function(h) { return row[h] || ''; });
  var newRow = sheet.getLastRow() + 1;
  var range = sheet.getRange(newRow, 1, 1, vals.length);
  range.setNumberFormat('@');
  range.setValues([vals]);
  return {status: 'ok', id: row.id};
}

// ============================================================
//  reviewTimeChangeReq — Approve or reject, update Attendance if approved
// ============================================================
function reviewTimeChangeReq(data) {
  if (!data.id || !data.status) return {error: 'Missing id or status'};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('TimeChangeReq');
  if (!sheet) return {error: 'TimeChangeReq sheet not found'};
  var headers = SHEETS.TimeChangeReq;
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  var allData = sheet.getDataRange().getValues();
  var found = false;
  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][0]) === String(data.id)) {
      // Update status, reviewedBy, reviewedAt
      var statusCol = headers.indexOf('status') + 1;
      var revByCol = headers.indexOf('reviewedBy') + 1;
      var revAtCol = headers.indexOf('reviewedAt') + 1;
      sheet.getRange(i + 1, statusCol).setValue(data.status);
      sheet.getRange(i + 1, revByCol).setValue(data.reviewedBy || '');
      sheet.getRange(i + 1, revAtCol).setValue(new Date().toISOString());
      // Attendance update is handled by TBMS frontend via upsert after approval
      found = true;
      break;
    }
  }
  if (!found) return {error: 'Request not found'};
  return {status: 'ok', action: data.status};
}

// ============================================================
//  ackTimeChangeReq — Staff acknowledges approved/rejected CR
// ============================================================
function ackTimeChangeReq(data) {
  if (!data.id) return {error: 'Missing id'};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('TimeChangeReq');
  if (!sheet) return {error: 'TimeChangeReq sheet not found'};
  var headers = SHEETS.TimeChangeReq;
  var allData = sheet.getDataRange().getValues();
  var ackCol = headers.indexOf('acknowledgedAt') + 1;
  if (ackCol < 1) return {error: 'acknowledgedAt column not found'};
  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][0]) === String(data.id)) {
      sheet.getRange(i + 1, ackCol).setValue(new Date().toISOString());
      return {status: 'ok'};
    }
  }
  return {error: 'Request not found'};
}
