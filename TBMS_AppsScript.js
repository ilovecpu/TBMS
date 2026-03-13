// ============================================================
//  TBMS - The Bap Management System v3.4
//  Google Apps Script Backend (Code.gs)
//  Deployed: 2026-03-13
//  URL: https://script.google.com/macros/s/AKfycbycwxP1jF_ZyY8F2ICm1j_d80Nh9Vyjf8boRwhFJLFSoPdCVG1c7txPp8awGnTuhFfa/exec
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
  Staff:         ['id','storeId','name','nickName','clothSize','kioskPwd','dob','address','niNo','eVisa','mobile','startDate','leftDate','rate','taxCode','sortCode','accountNo','email','memo','active','kioskLogin'],
  Attendance:    ['id','staffId','storeId','date','clockIn','clockOut','photoIn','photoOut','source'],
  Suppliers:     ['id','name','email','phone','address','website','memo','active'],
  StockTemplate: ['id','category','name','unit','min','minCondition','orderQty','orderUnit','alwaysOrder','sortOrder','supplier1','supplier2','supplier3','memo','photo'],
  StoreStock:    ['storeId','itemId','category','name','unit','min','minCondition','orderQty','orderUnit','alwaysOrder','qty','supplier1','supplier2','supplier3'],
  StockCount:    ['id','storeId','week','countDate','itemId','category','name','unit','qty','submittedBy','submittedAt'],
  WeeklySales:   ['id','storeId','week','weekStart','totalSales','notes','submittedBy','submittedAt'],
  StaffMessages: ['id','storeId','staffId','type','message','active','createdBy','createdAt'],
  EditLog:        ['id','attendanceId','staffId','storeId','date','fieldChanged','oldValue','newValue','editTimestamp','kioskVersion'],
  TimeChangeReq:  ['id','attendanceId','staffId','storeId','date','field','currentValue','requestedValue','reason','status','kioskVersion','createdAt','reviewedBy','reviewedAt','acknowledgedAt'],
  DiaryConfig:    ['id','storeId','configType','name','sortOrder','active'],
  DiaryEntry:     ['id','storeId','date','openChecks','fridgeTemps','deliveries','cookingTemps','coolingRecords','leftovers','closingChecks','holdTemps','extraChecks','issues','signedBy','submittedBy','submittedAt'],
  StoreInfo:      ['id','storeId','leaseStart','leaseEnd','monthlyRent','serviceCharge','rentReviewYears','landlordName','landlordPhone','landlordEmail','estateAgent','estateAgentPhone','estateAgentEmail','councilName','councilEmail','councilPhone','businessRateAnnual','hygieneInspDate','hygieneRating','electricCompany','electricContractStart','electricContractEnd','electricKwhRate','electricDailyCharge','phoneCompany','phoneContractStart','phoneContractEnd','waterCompany','cardMachineCompany','cardContractStart','cardContractEnd','cardRate','fridgeCleanDate','fridgeCleanNext','airconCleanDate','airconCleanNext','memo','custom1Label','custom1Value','custom2Label','custom2Value','custom3Label','custom3Value','custom4Label','custom4Value','custom5Label','custom5Value','custom6Label','custom6Value','custom7Label','custom7Value','custom8Label','custom8Value','custom9Label','custom9Value','custom10Label','custom10Value'],
  KnowledgeBase:  ['id','category','title','content','tags','source','createdBy','createdAt','updatedAt','active','version','accessLevel'],
  // ★ POS Sales Data — pushed from branch servers
  DailySales:     ['date','branch','branchName','totalOrders','main_cashTotal','main_cardTotal','main_grandTotal','main_vatTotal','main_vatBreakdown','sub_cashPct','sub_cashTotal','sub_cardTotal','sub_grandTotal','sub_vatTotal','sub_vatBreakdown','cashCount','cardCount','itemBreakdown','pushedAt'],
  LiveSales:      ['date','branch','branchName','main_grandTotal','main_vatTotal','sub_grandTotal','sub_vatTotal','totalOrders','cashCount','cardCount','lastUpdated'],
  EndSales:       ['id','branch','branchName','periodFrom','periodTo','totalOrders','main_cashTotal','main_cardTotal','main_grandTotal','main_vatTotal','sub_cashPct','sub_grandTotal','sub_vatTotal','itemBreakdown','staff','pushedAt']
};

// Fields that should remain numeric
const NUMERIC_FIELDS = ['rate','min','orderQty','qty','totalSales','sortOrder'];
// Fields that are boolean
const BOOL_FIELDS = ['active','canViewRate','alwaysOrder'];
// Fields that are time (HH:mm format)
const TIME_FIELDS = ['clockIn','clockOut'];
// Fields that are date (yyyy-MM-dd format)
const DATE_FIELDS = ['dob','startDate','leftDate','date','countDate','weekStart'];

// ===== SECRET KEY — change this to your own random string =====
const API_SECRET = 'tBaP2026xKr!mGt9Qz';

function checkKey(key) {
  return key === API_SECRET;
}

// Normalize header for fuzzy matching: "Nick Name" → "nickname", "nickName" → "nickname"
function normalizeHeader(h) {
  return String(h).replace(/[\s_-]+/g, '').toLowerCase();
}

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'ping';
    if (action === 'ping') {
      return ContentService.createTextOutput(JSON.stringify({status:'ok'})).setMimeType(ContentService.MimeType.JSON);
    }
    if (!checkKey(e && e.parameter ? e.parameter.apikey : '')) {
      return ContentService.createTextOutput(JSON.stringify({error:'Unauthorized'})).setMimeType(ContentService.MimeType.JSON);
    }
    var sheet = (e && e.parameter) ? e.parameter.sheet : null;
    var result;
    switch(action) {
      case 'getAll':   result = getAllData(); break;
      case 'getSheet': result = getSheetData(sheet); break;
      case 'getStoreData': result = getStoreData(e.parameter.store, e.parameter.sheets); break;
      case 'getSetting': result = getSetting(e.parameter.key); break;
      case 'init':     result = initSheets(); break;
      case 'ping':     result = {status:'ok', time: new Date().toISOString(), version:'TBMS 2.5'}; break;
      case 'diagSheets': result = {status:'ok', sheets: Object.keys(SHEETS)}; break;
      case 'getArchive': result = getArchiveData(e.parameter.sheet, e.parameter.store, e.parameter.from, e.parameter.to); break;
      case 'getKB':      result = getKB(e.parameter.category); break;
      case 'searchKB':   result = searchKB(e.parameter.q); break;
      // ★ POS Sales Data — read endpoints for TBMS dashboard
      case 'getSalesReport': result = getSalesReport(e.parameter.branch, e.parameter.from, e.parameter.to); break;
      case 'getLiveSales':   result = getLiveSales(e.parameter.date); break;
      case 'getEndSalesLog': result = getEndSalesLog(e.parameter.branch, e.parameter.from, e.parameter.to); break;
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
    if (!checkKey(data.apikey)) {
      return ContentService.createTextOutput(JSON.stringify({error:'Unauthorized'})).setMimeType(ContentService.MimeType.JSON);
    }
    var result;
    switch(data.action) {
      case 'saveSheet':  result = saveSheet(data.sheet, data.rows); break;
      case 'upsert':     result = upsertRow(data.sheet, data.row); break;
      case 'deleteRow':  result = deleteRow(data.sheet, data.id); break;
      case 'appendRow':  result = appendNewRow(data.sheet, data.row); break;
      case 'clockInPhoto':  result = clockInWithPhoto(data); break;
      case 'clockOutPhoto': result = clockOutWithPhoto(data); break;
      case 'saveStockCount': result = saveStockCount(data); break;
      case 'saveItemPhoto': result = saveItemPhoto(data); break;
      case 'deleteItemPhoto': result = deleteItemPhoto(data); break;
      case 'editLog':    result = appendEditLog(data); break;
      case 'timeChangeReq': result = createTimeChangeReq(data); break;
      case 'reviewTimeReq': result = reviewTimeChangeReq(data); break;
      case 'ackTimeReq':   result = ackTimeChangeReq(data); break;
      case 'saveSetting': result = saveSetting(data.key, data.value); break;
      case 'saveKB':     result = saveKB(data.entry); break;
      case 'deleteKB':   result = deleteKB(data.id); break;
      case 'runArchive': result = archiveOldData(); break;
      case 'initData':   result = initWithData(data); break;
      // ★ POS Sales Data — push endpoints from branch servers
      case 'pushDailySales': result = pushDailySales(data); break;
      case 'pushLiveSales':  result = pushLiveSales(data); break;
      case 'pushEndSales':   result = pushEndSales(data); break;
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
  var skip = {DiaryConfig:1, DiaryEntry:1, KnowledgeBase:1};
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
  if (!name) return {error: 'Missing sheet name'};
  if (!SHEETS[name]) return {error: 'Unknown sheet: ' + name};
  if (!row) return {error: 'Missing row data'};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    // Auto-create sheet with headers
    sheet = ss.insertSheet(name);
    var headers = SHEETS[name];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
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
  if (!name || !SHEETS[name]) return {error: 'Unknown sheet: ' + (name||'(empty)')};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) return {error: 'Sheet not found: ' + name};
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
  if (!name) return {error: 'Missing sheet name'};
  if (!SHEETS[name]) return {error: 'Unknown sheet: ' + name};
  if (!row) return {error: 'Missing row data'};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    var h = SHEETS[name];
    sheet.getRange(1, 1, 1, h.length).setValues([h]);
    sheet.getRange(1, 1, 1, h.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
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

  // Delete existing rows for same storeId + countDate (batch method — much faster than deleteRow loop)
  var storeCol = headers.indexOf('storeId');
  var dateCol = headers.indexOf('countDate');
  if (storeCol >= 0 && dateCol >= 0) {
    var allData = scSheet.getDataRange().getValues();
    var keepRows = [allData[0]]; // keep header
    for (var d = 1; d < allData.length; d++) {
      if (String(allData[d][storeCol]) === String(data.storeId) && String(allData[d][dateCol]) === String(data.countDate)) {
        continue; // skip rows to delete
      }
      keepRows.push(allData[d]);
    }
    // Rewrite sheet with kept rows only
    if (keepRows.length < allData.length) {
      scSheet.clearContents();
      if (keepRows.length > 0) {
        scSheet.getRange(1, 1, keepRows.length, keepRows[0].length).setValues(keepRows);
      }
    }
  }

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
    if (dateCol > 0) {
      scSheet.getRange(lastRow - rows.length + 1, dateCol, rows.length, 1).setNumberFormat('@');
    }
  }

  // Sync StoreStock qty only if countDate is today (batch method for speed)
  var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (data.countDate === todayStr) {
    var ssSheet = ss.getSheetByName('StoreStock');
    if (ssSheet && ssSheet.getLastRow() > 1) {
      var ssData = ssSheet.getDataRange().getValues();
      var ssHeaders = SHEETS.StoreStock;
      var ssStoreCol = ssHeaders.indexOf('storeId');
      var ssItemCol = ssHeaders.indexOf('itemId');
      var ssQtyCol = ssHeaders.indexOf('qty');
      var itemMap = {};
      for (var j = 0; j < data.items.length; j++) {
        itemMap[String(data.items[j].itemId)] = Number(data.items[j].qty) || 0;
      }
      var changed = false;
      for (var r = 1; r < ssData.length; r++) {
        if (String(ssData[r][ssStoreCol]) === String(data.storeId) && itemMap[String(ssData[r][ssItemCol])] !== undefined) {
          ssData[r][ssQtyCol] = itemMap[String(ssData[r][ssItemCol])];
          changed = true;
        }
      }
      if (changed) {
        ssSheet.getRange(2, 1, ssData.length - 1, ssData[0].length).setValues(ssData.slice(1));
      }
    }
  }

  return {status: 'ok', week: week, count: rows.length};
}

// ============================================================
//  saveItemPhoto — Save photo for a StockTemplate item
// ============================================================
function saveItemPhoto(data) {
  if (!data.itemId || !data.photo) return {error: 'Missing itemId or photo'};
  var fileName = 'stock_' + data.itemId + '_' + Date.now();
  var fileId = savePhotoToDrive(data.photo, fileName);
  if (!fileId) return {error: 'Failed to save photo'};
  // Update StockTemplate photo field
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('StockTemplate');
  if (!sheet) return {error: 'StockTemplate sheet not found'};
  var headers = SHEETS.StockTemplate;
  var photoCol = headers.indexOf('photo') + 1;
  if (photoCol < 1) return {error: 'photo column not found'};
  var allData = sheet.getDataRange().getValues();
  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][0]) === String(data.itemId)) {
      sheet.getRange(i + 1, photoCol).setValue(fileId);
      return {status: 'ok', fileId: fileId};
    }
  }
  return {error: 'Item not found: ' + data.itemId};
}

// ============================================================
//  deleteItemPhoto — Remove photo from a StockTemplate item
// ============================================================
function deleteItemPhoto(data) {
  if (!data.itemId) return {error: 'Missing itemId'};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('StockTemplate');
  if (!sheet) return {error: 'StockTemplate sheet not found'};
  var headers = SHEETS.StockTemplate;
  var photoCol = headers.indexOf('photo') + 1;
  if (photoCol < 1) return {error: 'photo column not found'};
  var allData = sheet.getDataRange().getValues();
  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][0]) === String(data.itemId)) {
      var oldFileId = allData[i][photoCol - 1];
      sheet.getRange(i + 1, photoCol).setValue('');
      // Try to delete the file from Drive
      if (oldFileId) {
        try { DriveApp.getFileById(String(oldFileId)).setTrashed(true); } catch(e) {}
      }
      return {status: 'ok'};
    }
  }
  return {error: 'Item not found: ' + data.itemId};
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

// ============================================================
//  Knowledge Base — AI Brain persistent storage
// ============================================================
function getKB(category) {
  var rows = readSheet('KnowledgeBase');
  var active = rows.filter(function(r) { return r.active !== 'false' && r.active !== false; });
  if (category) {
    active = active.filter(function(r) { return r.category === category; });
  }
  return {status: 'ok', data: active, total: active.length};
}

function searchKB(query) {
  if (!query) return {error: 'Missing search query'};
  var q = query.toLowerCase();
  var rows = readSheet('KnowledgeBase');
  var results = rows.filter(function(r) {
    if (r.active === 'false' || r.active === false) return false;
    return (r.title && r.title.toLowerCase().indexOf(q) >= 0) ||
           (r.content && r.content.toLowerCase().indexOf(q) >= 0) ||
           (r.tags && r.tags.toLowerCase().indexOf(q) >= 0) ||
           (r.category && r.category.toLowerCase().indexOf(q) >= 0);
  });
  return {status: 'ok', data: results, total: results.length};
}

function saveKB(entry) {
  if (!entry) return {error: 'Missing entry data'};
  var now = new Date().toISOString();
  if (!entry.id) {
    // New entry
    entry.id = 'kb_' + Date.now().toString(36) + Math.random().toString(36).substr(2, 4);
    entry.createdAt = now;
    entry.updatedAt = now;
    entry.active = 'true';
    entry.version = '1';
  } else {
    // Update existing
    entry.updatedAt = now;
    // Increment version
    var ver = parseInt(entry.version) || 0;
    entry.version = String(ver + 1);
  }
  var result = upsertRow('KnowledgeBase', entry);
  return {status: 'ok', entry: entry, action: result.action};
}

function deleteKB(id) {
  if (!id) return {error: 'Missing id'};
  // Soft delete — set active to false
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('KnowledgeBase');
  if (!sheet) return {error: 'KnowledgeBase sheet not found'};
  var headers = SHEETS.KnowledgeBase;
  var activeCol = headers.indexOf('active') + 1;
  var allData = sheet.getDataRange().getValues();
  for (var i = 1; i < allData.length; i++) {
    if (String(allData[i][0]) === String(id)) {
      sheet.getRange(i + 1, activeCol).setValue('false');
      return {status: 'ok', deleted: true};
    }
  }
  return {status: 'ok', deleted: false};
}

// ============================================================
//  DATA ARCHIVE — moves records older than 3 months to archive sheets
//  Sheets archived: StockCount, Attendance, EditLog, DiaryEntry
//  Archive sheet names: StockCount_Archive, Attendance_Archive, etc.
//  Run manually or via time-based trigger (monthly recommended)
// ============================================================
function archiveOldData() {
  var cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - 3);
  var cutoffStr = Utilities.formatDate(cutoff, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetsToArchive = [
    {name: 'StockCount', dateField: 'countDate'},
    {name: 'Attendance', dateField: 'date'},
    {name: 'EditLog', dateField: 'date'},
    {name: 'DiaryEntry', dateField: 'date'}
  ];
  var summary = [];
  sheetsToArchive.forEach(function(cfg) {
    var sheet = ss.getSheetByName(cfg.name);
    if (!sheet || sheet.getLastRow() < 2) return;
    var headers = SHEETS[cfg.name];
    if (!headers) return;
    var dateCol = headers.indexOf(cfg.dateField);
    if (dateCol < 0) return;
    var allData = sheet.getDataRange().getValues();
    var headerRow = allData[0];
    // Split into keep vs archive
    var keep = [];
    var archive = [];
    for (var i = 1; i < allData.length; i++) {
      var dateVal = String(allData[i][dateCol] || '');
      if (dateVal.length >= 10) dateVal = dateVal.substring(0, 10);
      if (dateVal && dateVal < cutoffStr) {
        archive.push(allData[i]);
      } else {
        keep.push(allData[i]);
      }
    }
    if (archive.length === 0) { summary.push(cfg.name + ': 0 archived'); return; }
    // Get or create archive sheet
    var archName = cfg.name + '_Archive';
    var archSheet = ss.getSheetByName(archName);
    if (!archSheet) {
      archSheet = ss.insertSheet(archName);
      archSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    // Append archived rows
    var lastRow = archSheet.getLastRow();
    archSheet.getRange(lastRow + 1, 1, archive.length, headers.length).setValues(archive);
    // Rewrite main sheet with kept rows only
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }
    if (keep.length > 0) {
      sheet.getRange(2, 1, keep.length, headers.length).setValues(keep);
    }
    summary.push(cfg.name + ': ' + archive.length + ' archived, ' + keep.length + ' kept');
  });
  return {status: 'ok', cutoffDate: cutoffStr, results: summary};
}

// Retrieve archived data for a specific sheet, optionally filtered by store and date range
function getArchiveData(sheetName, storeId, dateFrom, dateTo) {
  if (!sheetName) return {error: 'Missing sheetName'};
  var archName = sheetName + '_Archive';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var archSheet = ss.getSheetByName(archName);
  if (!archSheet || archSheet.getLastRow() < 2) return {status: 'ok', data: []};
  var headers = SHEETS[sheetName];
  if (!headers) return {error: 'Unknown sheet: ' + sheetName};
  var storeCol = headers.indexOf('storeId');
  var dateField = sheetName === 'StockCount' ? 'countDate' : 'date';
  var dateCol = headers.indexOf(dateField);
  var allData = archSheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < allData.length; i++) {
    // Filter by store
    if (storeId && storeCol >= 0 && String(allData[i][storeCol]) !== String(storeId)) continue;
    // Filter by date range
    if (dateCol >= 0) {
      var d = String(allData[i][dateCol] || '');
      if (d.length >= 10) d = d.substring(0, 10);
      if (dateFrom && d < dateFrom) continue;
      if (dateTo && d > dateTo) continue;
    }
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      row[headers[j]] = allData[i][j] !== undefined ? allData[i][j] : '';
    }
    rows.push(row);
  }
  return {status: 'ok', data: rows, total: rows.length};
}

// ============================================================
//  ★ POS Sales Data — Push & Read Functions
//  Branch servers push daily/live/end-sales data here
//  TBMS reads from here for Sales Report dashboard
// ============================================================

// ─── Helper: upsert row in sheet by key columns ───
function _upsertSalesRow(sheetName, keyFields, rowData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    // Auto-create sheet with headers
    sheet = ss.insertSheet(sheetName);
    var headers = SHEETS[sheetName];
    if (headers) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }
  var headers = SHEETS[sheetName];
  if (!headers) return {error: 'Unknown sheet: ' + sheetName};

  // Build row array from rowData
  var newRow = headers.map(function(h) {
    var val = rowData[h];
    // JSON stringify objects/arrays
    if (val !== null && val !== undefined && typeof val === 'object') return JSON.stringify(val);
    return val !== null && val !== undefined ? val : '';
  });

  // Find existing row by key fields
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var match = true;
    for (var k = 0; k < keyFields.length; k++) {
      var colIdx = headers.indexOf(keyFields[k]);
      if (colIdx < 0 || String(data[i][colIdx]) !== String(rowData[keyFields[k]])) {
        match = false; break;
      }
    }
    if (match) {
      // Update existing row
      sheet.getRange(i + 1, 1, 1, newRow.length).setValues([newRow]);
      return {status: 'ok', action: 'updated', row: i + 1};
    }
  }
  // Append new row
  sheet.appendRow(newRow);
  return {status: 'ok', action: 'inserted', row: sheet.getLastRow()};
}

// ─── pushDailySales: 자정 자동 푸시 (upsert by date+branch) ───
function pushDailySales(data) {
  if (!data.date || !data.branch) return {error: 'date and branch required'};
  var row = {
    date:               data.date,
    branch:             data.branch,
    branchName:         data.branchName || data.branch,
    totalOrders:        data.totalOrders || 0,
    main_cashTotal:     data.main ? data.main.cashTotal : 0,
    main_cardTotal:     data.main ? data.main.cardTotal : 0,
    main_grandTotal:    data.main ? data.main.grandTotal : 0,
    main_vatTotal:      data.main ? data.main.vatTotal : 0,
    main_vatBreakdown:  data.main ? data.main.vatBreakdown : {},
    sub_cashPct:        data.sub ? data.sub.cashPct : 100,
    sub_cashTotal:      data.sub ? data.sub.cashTotal : 0,
    sub_cardTotal:      data.sub ? data.sub.cardTotal : 0,
    sub_grandTotal:     data.sub ? data.sub.grandTotal : 0,
    sub_vatTotal:       data.sub ? data.sub.vatTotal : 0,
    sub_vatBreakdown:   data.sub ? data.sub.vatBreakdown : {},
    cashCount:          data.cashCount || 0,
    cardCount:          data.cardCount || 0,
    itemBreakdown:      data.itemBreakdown || [],
    pushedAt:           new Date().toISOString()
  };
  return _upsertSalesRow('DailySales', ['date', 'branch'], row);
}

// ─── pushLiveSales: 매시간 라이브 푸시 (upsert by date+branch) ───
function pushLiveSales(data) {
  if (!data.date || !data.branch) return {error: 'date and branch required'};
  // ★ 자동 리셋: 오늘이 아닌 이전 날짜 데이터 삭제
  _cleanOldLiveSales(data.date);
  var row = {
    date:            data.date,
    branch:          data.branch,
    branchName:      data.branchName || data.branch,
    main_grandTotal: data.main_grandTotal || 0,
    main_vatTotal:   data.main_vatTotal || 0,
    sub_grandTotal:  data.sub_grandTotal || 0,
    sub_vatTotal:    data.sub_vatTotal || 0,
    totalOrders:     data.totalOrders || 0,
    cashCount:       data.cashCount || 0,
    cardCount:       data.cardCount || 0,
    lastUpdated:     new Date().toISOString()
  };
  return _upsertSalesRow('LiveSales', ['date', 'branch'], row);
}

// ★ LiveSales 자동 리셋 — 오늘 날짜가 아닌 행 삭제
function _cleanOldLiveSales(todayStr) {
  try {
    var sheetName = 'LiveSales';
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(sheetName);
    if (!sh) return;
    var lastRow = sh.getLastRow();
    if (lastRow <= 1) return; // header only
    var dates = sh.getRange(2, 1, lastRow - 1, 1).getValues(); // column A = date
    var rowsToDelete = [];
    for (var i = 0; i < dates.length; i++) {
      var cellDate = String(dates[i][0]).trim();
      if (cellDate && cellDate !== todayStr) {
        rowsToDelete.push(i + 2); // sheet row (1-indexed, skip header)
      }
    }
    // 역순으로 삭제 (아래→위, 인덱스 변동 방지)
    for (var j = rowsToDelete.length - 1; j >= 0; j--) {
      sh.deleteRow(rowsToDelete[j]);
    }
    if (rowsToDelete.length > 0) {
      Logger.log('[LiveSales] Cleaned ' + rowsToDelete.length + ' old rows (not ' + todayStr + ')');
    }
  } catch (e) {
    Logger.log('[LiveSales] Clean error: ' + e.message);
  }
}

// ─── pushEndSales: END Sales 이벤트 푸시 (append) ───
function pushEndSales(data) {
  if (!data.id || !data.branch) return {error: 'id and branch required'};
  var row = {
    id:              data.id,
    branch:          data.branch,
    branchName:      data.branchName || data.branch,
    periodFrom:      data.periodFrom || '',
    periodTo:        data.periodTo || '',
    totalOrders:     data.totalOrders || 0,
    main_cashTotal:  data.main ? data.main.cashTotal : 0,
    main_cardTotal:  data.main ? data.main.cardTotal : 0,
    main_grandTotal: data.main ? data.main.grandTotal : 0,
    main_vatTotal:   data.main ? data.main.vatTotal : 0,
    sub_cashPct:     data.sub ? data.sub.cashPct : 100,
    sub_grandTotal:  data.sub ? data.sub.grandTotal : 0,
    sub_vatTotal:    data.sub ? data.sub.vatTotal : 0,
    itemBreakdown:   data.itemBreakdown || [],
    staff:           data.staff || '',
    pushedAt:        new Date().toISOString()
  };
  // Upsert by id to prevent duplicates
  return _upsertSalesRow('EndSales', ['id'], row);
}

// ─── getSalesReport: 날짜 범위 DailySales 조회 (TBMS 대시보드) ───
function getSalesReport(branch, from, to) {
  var rows = readSheet('DailySales');
  if (branch) rows = rows.filter(function(r) { return r.branch === branch; });
  if (from)   rows = rows.filter(function(r) { return r.date >= from; });
  if (to)     rows = rows.filter(function(r) { return r.date <= to; });
  // Parse JSON fields
  rows.forEach(function(r) {
    try { if (typeof r.main_vatBreakdown === 'string') r.main_vatBreakdown = JSON.parse(r.main_vatBreakdown); } catch(e) {}
    try { if (typeof r.sub_vatBreakdown === 'string') r.sub_vatBreakdown = JSON.parse(r.sub_vatBreakdown); } catch(e) {}
    try { if (typeof r.itemBreakdown === 'string') r.itemBreakdown = JSON.parse(r.itemBreakdown); } catch(e) {}
  });
  rows.sort(function(a,b) { return a.date < b.date ? 1 : a.date > b.date ? -1 : 0; });
  return {status: 'ok', data: rows, total: rows.length};
}

// ─── getLiveSales: 라이브 세일즈 전체 지점 (TBMS 대시보드) ───
// date 파라미터 없으면 오늘, 있으면 해당 날짜 조회
// 오늘 = LiveSales 시트, 과거 = DailySales 시트에서 조회
function getLiveSales(date) {
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var targetDate = date || today;
  var isToday = (targetDate === today);

  if (isToday) {
    // 오늘: LiveSales 시트에서 실시간 데이터
    var rows = readSheet('LiveSales');
    rows = rows.filter(function(r) { return r.date === targetDate; });
    return {status: 'ok', data: rows, date: targetDate, isToday: true};
  } else {
    // 과거: DailySales 시트에서 조회 → LiveSales 형식으로 변환
    var dailyRows = readSheet('DailySales');
    dailyRows = dailyRows.filter(function(r) { return r.date === targetDate; });
    var mapped = dailyRows.map(function(r) {
      return {
        date:            r.date,
        branch:          r.branch,
        branchName:      r.branchName || r.branch,
        main_grandTotal: r.main_grandTotal || 0,
        main_vatTotal:   r.main_vatTotal || 0,
        sub_grandTotal:  r.sub_grandTotal || 0,
        sub_vatTotal:    r.sub_vatTotal || 0,
        totalOrders:     r.totalOrders || 0,
        cashCount:       r.cashCount || 0,
        cardCount:       r.cardCount || 0,
        lastUpdated:     r.pushedAt || '',
        // DailySales 추가 정보
        main_cashTotal:  r.main_cashTotal || 0,
        main_cardTotal:  r.main_cardTotal || 0,
        sub_cashTotal:   r.sub_cashTotal || 0,
        sub_cardTotal:   r.sub_cardTotal || 0
      };
    });
    return {status: 'ok', data: mapped, date: targetDate, isToday: false};
  }
}

// ─── getEndSalesLog: END Sales 기록 조회 ───
function getEndSalesLog(branch, from, to) {
  var rows = readSheet('EndSales');
  if (branch) rows = rows.filter(function(r) { return r.branch === branch; });
  if (from)   rows = rows.filter(function(r) { return r.periodTo >= from; });
  if (to)     rows = rows.filter(function(r) { return r.periodTo <= (to.length === 10 ? to + 'T23:59:59' : to); });
  try { rows.forEach(function(r) {
    if (typeof r.itemBreakdown === 'string') r.itemBreakdown = JSON.parse(r.itemBreakdown);
  }); } catch(e) {}
  rows.sort(function(a,b) { return a.pushedAt < b.pushedAt ? 1 : a.pushedAt > b.pushedAt ? -1 : 0; });
  return {status: 'ok', data: rows, total: rows.length};
}
