// ============================================================
//  TBMS - The Bap Management System v2.0
//  Google Apps Script Backend (Code.gs)
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
  Users:         ['id','username','password','name','role','email','storeId'],
  Stores:        ['id','code','name','company','companyNo','address','phone','email','manager','memo','active'],
  Staff:         ['id','storeId','name','clothSize','kioskPwd','dob','address','niNo','eVisa','mobile','startDate','rate','sortCode','accountNo','email','memo','active'],
  Attendance:    ['id','staffId','storeId','date','clockIn','clockOut'],
  StockTemplate: ['id','category','name','unit','min'],
  StoreStock:    ['storeId','itemId','category','name','unit','min','qty']
};

// Fields that should remain numeric
const NUMERIC_FIELDS = ['rate','min','qty'];
// Fields that are boolean
const BOOL_FIELDS = ['active'];
// Fields that are time (HH:mm format)
const TIME_FIELDS = ['clockIn','clockOut'];
// Fields that are date (yyyy-MM-dd format)
const DATE_FIELDS = ['dob','startDate','date'];

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'ping';
    var sheet = (e && e.parameter) ? e.parameter.sheet : null;
    var result;
    switch(action) {
      case 'getAll':   result = getAllData(); break;
      case 'getSheet': result = getSheetData(sheet); break;
      case 'init':     result = initSheets(); break;
      case 'ping':     result = {status:'ok', time: new Date().toISOString(), version:'TBMS 2.0'}; break;
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

function initSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var created = [];
  for (var name in SHEETS) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) { sheet = ss.insertSheet(name); created.push(name); }
    var headers = SHEETS[name];
    var firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    if (!firstRow[0]) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
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
  return {status:'ok', created: created};
}

function getAllData() {
  var result = {};
  for (var name in SHEETS) { result[name] = readSheet(name); }
  return {status:'ok', data: result};
}

function readSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var h = headers[j];
      var val = data[i][j];
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
    if (obj[headers[0]] !== '' && obj[headers[0]] !== null && obj[headers[0]] !== undefined) {
      rows.push(obj);
    }
  }
  return rows;
}

function getSheetData(name) {
  if (!name || !SHEETS[name]) return {error: 'Invalid sheet name'};
  return {status:'ok', data: readSheet(name)};
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

function upsertRow(name, row) {
  if (!name || !SHEETS[name] || !row) return {error: 'Invalid parameters'};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) return {error: 'Sheet not found: ' + name};
  var headers = SHEETS[name];
  var idValue = row[headers[0]];
  var data = sheet.getDataRange().getValues();
  var found = false;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(idValue)) {
      var vals = headers.map(function(h) { var v = row[h]; return (v === null || v === undefined) ? '' : v; });
      sheet.getRange(i + 1, 1, 1, headers.length).setValues([vals]);
      found = true; break;
    }
  }
  if (!found) {
    var vals = headers.map(function(h) { var v = row[h]; return (v === null || v === undefined) ? '' : v; });
    sheet.appendRow(vals);
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

function appendNewRow(name, row) {
  if (!name || !SHEETS[name] || !row) return {error: 'Invalid parameters'};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) return {error: 'Sheet not found'};
  var headers = SHEETS[name];
  var vals = headers.map(function(h) { var v = row[h]; return (v === null || v === undefined) ? '' : v; });
  sheet.appendRow(vals);
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
