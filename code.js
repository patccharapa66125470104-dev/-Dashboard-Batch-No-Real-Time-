/**
 * ================================================================
 *  Packing Material Stock — Executive Dashboard (Real-time)
 *  Code.gs
 * ------------------------------------------------------------------
 *  - Reads every sheet in the bound spreadsheet (each sheet = a
 *    ประเภท of packing material, e.g. "200L Plastic Drum LD85").
 *  - Headers are located dynamically (by matching header text),
 *    NOT by hard-coded column letters — so the dashboard keeps
 *    working even if a sheet gains/loses a column.
 *  - Each "batch" is a block of rows that starts wherever the
 *    Product Code / Batch No. cell is filled in, and continues
 *    until the next filled Product Code cell (this mirrors how the
 *    source sheets are laid out: one header row per batch, followed
 *    by one row per withdrawal/issue transaction).
 *  - Remain is always CALCULATED as QtyIn - SUM(QtyOut for the
 *    block). This is more reliable than trusting the sheet's own
 *    "Remain" column, which is sometimes left blank for a batch
 *    that has not been issued yet.
 * ================================================================
 */

// ---- Header text we search for (case-insensitive, partial match) ----
var HEADERS = {
  CODE: 'product code',
  NAME: 'product name',
  SUPPLIER: 'supplier',
  BATCH: 'batch no',
  MFGDATE: 'mfg',
  QTYIN: 'quantity in',
  QTYOUT: 'quantity out',
  RECORDDATE: 'record date',
  DELIVEREDDATE: 'delivered date'
};

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Packing Material Stock — Executive Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Main entry point called from the client (google.script.run).
 * Returns 100% fresh data straight from the spreadsheet every time.
 */
function getDashboardData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var result = {
    generatedAt: new Date().toISOString(),
    types: [],        // list of sheet/type names, in sheet order
    batches: [],      // flattened list of every batch, across all sheets
    transactions: []  // every individual in/out movement, with its own real date
  };

  sheets.forEach(function (sheet) {
    var typeName = sheet.getName().trim();
    var parsed = parseSheet(sheet);
    if (parsed.batches.length === 0) return;
    result.types.push(typeName);
    parsed.batches.forEach(function (b) {
      b.type = typeName;
      result.batches.push(b);
    });
    parsed.transactions.forEach(function (t) {
      t.type = typeName;
      result.transactions.push(t);
    });
  });

  return result;
}

/**
 * Locate the header row (the row containing "Batch No."), map every
 * needed column, then walk the sheet row-by-row grouping rows into
 * batch blocks.
 */
function parseSheet(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 2) return { batches: [], transactions: [] };

  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  // ---- 1. find the main header row (contains "Batch No.") ----
  var headerRowIdx = -1; // 0-based index into `values`
  for (var r = 0; r < Math.min(values.length, 10); r++) {
    for (var c = 0; c < lastCol; c++) {
      if (norm_(values[r][c]).indexOf(HEADERS.BATCH) !== -1) {
        headerRowIdx = r;
        break;
      }
    }
    if (headerRowIdx !== -1) break;
  }
  if (headerRowIdx === -1) return { batches: [], transactions: [] };

  var mainRow = values[headerRowIdx];
  var subRow = values[headerRowIdx + 1] || [];

  var col = {
    code: findCol_(mainRow, HEADERS.CODE),
    name: findCol_(mainRow, HEADERS.NAME),
    supplier: findCol_(mainRow, HEADERS.SUPPLIER),
    batch: findCol_(mainRow, HEADERS.BATCH),
    mfgDate: findCol_(mainRow, HEADERS.MFGDATE),
    qtyIn: findCol_(mainRow, HEADERS.QTYIN),
    qtyOut: findCol_(mainRow, HEADERS.QTYOUT),                  // first occurrence
    recordDate: findCol_(subRow, HEADERS.RECORDDATE),           // date the batch was received in (วันรับเข้า)
    deliveredDate: findCol_(subRow, HEADERS.DELIVEREDDATE)      // date each withdrawal row was issued out (วันจ่ายออก)
  };

  // bail out if the essentials are missing
  if (col.code === -1 || col.batch === -1 || col.qtyIn === -1 || col.qtyOut === -1) {
    return { batches: [], transactions: [] };
  }

  var dataStartIdx = headerRowIdx + 2; // skip main header + sub-header row

  // ---- 2. find where every batch block starts ----
  var starts = [];
  for (var i = dataStartIdx; i < values.length; i++) {
    var row = values[i];
    if (isFilled_(row[col.code]) && isFilled_(row[col.batch])) {
      starts.push(i);
    }
  }

  // ---- 3. build one record per batch block, plus one transaction per
  //         real movement (using each row's own actual date, exactly
  //         as it appears in the sheet) ----
  var batches = [];
  var transactions = [];
  for (var s = 0; s < starts.length; s++) {
    var from = starts[s];
    var to = (s + 1 < starts.length) ? starts[s + 1] : values.length;
    var block = values.slice(from, to);
    var batchNo = String(block[0][col.batch] || '').trim();

    var qtyIn = toNumber_(block[0][col.qtyIn]);
    var recordDate = col.recordDate !== -1 ? parseFlexibleDate_(block[0][col.recordDate]) : null;
    if (qtyIn > 0) {
      transactions.push({ batchNo: batchNo, direction: 'in', date: recordDate, qty: qtyIn });
    }

    var qtyOutSum = 0;
    for (var k = 0; k < block.length; k++) {
      var qtyOutVal = toNumber_(block[k][col.qtyOut]);
      qtyOutSum += qtyOutVal;
      if (qtyOutVal > 0) {
        var deliveredDate = col.deliveredDate !== -1 ? parseFlexibleDate_(block[k][col.deliveredDate]) : null;
        transactions.push({ batchNo: batchNo, direction: 'out', date: deliveredDate, qty: qtyOutVal });
      }
    }
    var remain = qtyIn - qtyOutSum;

    batches.push({
      productCode: String(block[0][col.code] || '').trim(),
      productName: col.name !== -1 ? String(block[0][col.name] || '').trim() : '',
      supplier: col.supplier !== -1 ? String(block[0][col.supplier] || '').trim() : '',
      batchNo: batchNo,
      mfgDate: parseFlexibleDate_(block[0][col.mfgDate]),
      recordDate: recordDate,
      qtyIn: qtyIn,
      qtyOut: qtyOutSum,
      remain: remain
    });
  }

  return { batches: batches, transactions: transactions };
}

// ---------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------
function norm_(v) {
  return String(v == null ? '' : v).trim().toLowerCase();
}

function findCol_(headerRow, needle) {
  for (var c = 0; c < headerRow.length; c++) {
    if (norm_(headerRow[c]).indexOf(needle) !== -1) return c;
  }
  return -1;
}

function isFilled_(v) {
  return v !== null && v !== undefined && String(v).trim() !== '';
}

function toNumber_(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (v instanceof Date) return 0;
  var n = Number(v);
  return isNaN(n) ? 0 : n;
}

/**
 * Sheet cells are sometimes real Date objects (normal case) and
 * sometimes messy text ("-", "08/17/2021", "11 /11/2025", blank).
 * Returns an ISO date string ("yyyy-MM-dd") or null.
 */
function parseFlexibleDate_(v) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(v).trim();
  if (s === '' || s === '-') return null;
  s = s.replace(/\s+/g, ''); // "11 /11/2025" -> "11/11/2025"

  // dd/mm/yyyy  or  mm/dd/yyyy (only one of day/month can be >12)
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    var a = parseInt(m[1], 10), b = parseInt(m[2], 10), y = parseInt(m[3], 10);
    var day, month;
    if (a > 12) { day = a; month = b; }        // dd/mm
    else if (b > 12) { day = b; month = a; }    // mm/dd
    else { day = a; month = b; }                // ambiguous -> assume dd/mm
    var d = new Date(y, month - 1, day);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
  }

  // dd-Mon-yy  (e.g. 24-Dec-25)
  var parsed = new Date(s);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return null;
}
