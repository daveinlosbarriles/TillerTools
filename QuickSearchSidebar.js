// Quick Search sidebar: client-side filter via helper columns + Basic Filter.
// All sheet and column dependencies are listed below.

// --- Sheet names ---
var QUICK_SEARCH_SHEET_NAME = "Transactions";  // Filter operates only on this sheet
var ACCOUNTS_SHEET_NAME = "Accounts";
var CATEGORIES_SHEET_NAME = "Categories";

// --- Transactions sheet: column header labels (row 1, used to find columns) ---
var TRANSACTIONS_HEADER_DATE = "Date";
var TRANSACTIONS_HEADER_DESCRIPTION = "Description";
var TRANSACTIONS_HEADER_AMOUNT = "Amount";
var TRANSACTIONS_HEADER_ACCOUNT = "Account";
var TRANSACTIONS_HEADER_CATEGORY = "Category";
var TRANSACTIONS_HEADER_CATEGORIES = "Categories";  // alternate header for category column

// --- Accounts sheet: column indices (1-based) and filter value ---
var ACCOUNTS_COL_ACCOUNT = 10;   // Column J
var ACCOUNTS_COL_HIDE = 17;      // Column Q
var ACCOUNTS_HIDE_VALUE = "Hide";

// --- Categories sheet: column indices (1-based). Data starts row 2. ---
var CATEGORIES_COL_CATEGORY = 1;
var CATEGORIES_COL_GROUP = 2;

// --- Accounts sheet: 0-based index within the 2-column range (ACCOUNTS_COL_ACCOUNT..ACCOUNTS_COL_HIDE) ---
var ACCOUNTS_RANGE_INDEX_ACCOUNT = 0;
var ACCOUNTS_RANGE_INDEX_HIDE = 1;

// --- Common: row and column positions (1-based) ---
var ROW_HEADER = 1;
var ROW_DATA_FIRST = 2;   // First data row (row 1 = headers)
var COL_FIRST = 1;

// --- Quick Search helper columns: default positions when Transactions sheet is empty ---
var HELPER_COL_MATCH_DEFAULT = 1;
var HELPER_COL_CRITERIA_DEFAULT = 2;

// --- 0-based indices when reading a 2-column range (e.g. last two columns for helper headers) ---
var RANGE_INDEX_FIRST = 0;
var RANGE_INDEX_SECOND = 1;

// --- Quick Search helper column headers ---
var QUICK_SEARCH_MATCH_HEADER = "QuickSearch";
var QUICK_SEARCH_CRITERIA_HEADER = "QuickCriteria";
// Criteria: single string in row 1 only. Format: D:dateFrom,dateTo|X:description|C:cat1,cat2|A:min,max|Q:acc1,acc2

/**
 * Row 1 header name -> 1-based column index (matches Tiller header labels).
 * Defined here so Quick Search does not depend on any other script file; safe to use in a standalone extension.
 */
function getTillerColumnMap(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    var h = headers[i];
    if (h) {
      map[String(h).trim()] = i + 1;
    }
  }
  return map;
}

/**
 * Opens the Quick Search sidebar.
 */
function openQuickSearchSidebar() {
  var html = HtmlService.createHtmlOutputFromFile("QuickSearch");
  html.setTitle("Tiller Quick Search");
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Returns category and account options. Accounts from Accounts sheet column J; rows with Q = "Hide" are excluded.
 */
function getQuickSearchOptions() {
  var ss = SpreadsheetApp.getActive();
  var categoriesSheet = ss.getSheetByName(CATEGORIES_SHEET_NAME);
  var accountsSheet = ss.getSheetByName(ACCOUNTS_SHEET_NAME);

  var categories = [];
  if (categoriesSheet && categoriesSheet.getLastRow() >= ROW_DATA_FIRST) {
    var values = categoriesSheet.getRange(ROW_DATA_FIRST, CATEGORIES_COL_CATEGORY, categoriesSheet.getLastRow(), CATEGORIES_COL_GROUP).getDisplayValues();
    values.forEach(function(row) {
      var category = row[CATEGORIES_COL_CATEGORY - 1];
      var group = row[CATEGORIES_COL_GROUP - 1];
      if (category) {
        var label = group ? String(category) + ": " + String(group) : String(category);
        categories.push({ category: String(category), group: group ? String(group) : "", label: label });
      }
    });
    categories.unshift({ category: "(Blank)", group: "", label: "(Blank)" });
  }

  var accounts = [];
  if (accountsSheet && accountsSheet.getLastRow() >= ROW_DATA_FIRST) {
    var lastRow = accountsSheet.getLastRow();
    var vals = accountsSheet.getRange(ROW_DATA_FIRST, ACCOUNTS_COL_ACCOUNT, lastRow, ACCOUNTS_COL_HIDE).getDisplayValues();
    var seen = {};
    vals.forEach(function(r) {
      var accountVal = (r[ACCOUNTS_RANGE_INDEX_ACCOUNT] || "").trim();
      var hideVal = (r[ACCOUNTS_RANGE_INDEX_HIDE] || "").trim();
      if (hideVal !== ACCOUNTS_HIDE_VALUE && accountVal && !seen[accountVal]) {
        seen[accountVal] = true;
        accounts.push(accountVal);
      }
    });
  }

  return { categories: categories, accounts: accounts };
}

/** Max transaction rows to load for client-side search (avoids timeout). */
var QUICK_SEARCH_MAX_ROWS = 15000;

/**
 * Returns options (categories, accounts) plus transaction rows for client-side filtering.
 * Called once when the sidebar loads. No server calls needed when user clicks Search.
 */
function getQuickSearchData() {
  var opts = getQuickSearchOptions();
  var ss = SpreadsheetApp.getActive();
  var sheet = getQuickSearchSheet(ss);
  var rows = [];
  if (sheet) {
    var map = getTillerColumnMap(sheet);
    var dateCol = map[TRANSACTIONS_HEADER_DATE];
    var descCol = map[TRANSACTIONS_HEADER_DESCRIPTION];
    var amountCol = map[TRANSACTIONS_HEADER_AMOUNT];
    var accountCol = map[TRANSACTIONS_HEADER_ACCOUNT];
    var catCol = map[TRANSACTIONS_HEADER_CATEGORY] || map[TRANSACTIONS_HEADER_CATEGORIES];
    if (!catCol && sheet.getLastColumn() >= COL_FIRST) {
      var headers = sheet.getRange(ROW_HEADER, COL_FIRST, ROW_HEADER, sheet.getLastColumn()).getDisplayValues()[0];
      var catLower = TRANSACTIONS_HEADER_CATEGORY.toLowerCase();
      var catsLower = TRANSACTIONS_HEADER_CATEGORIES.toLowerCase();
      for (var i = 0; i < headers.length; i++) {
        var h = (headers[i] && String(headers[i]).trim()).toLowerCase();
        if (h === catLower || h === catsLower) { catCol = i + 1; break; }
      }
    }
    if (dateCol && descCol && amountCol && accountCol && catCol) {
      var lastRow = Math.min(sheet.getLastRow(), 1 + QUICK_SEARCH_MAX_ROWS);
      if (lastRow >= ROW_DATA_FIRST) {
        var startCol = Math.min(dateCol, descCol, amountCol, accountCol, catCol);
        var endCol = Math.max(dateCol, descCol, amountCol, accountCol, catCol);
        var grid = sheet.getRange(ROW_DATA_FIRST, startCol, lastRow, endCol).getDisplayValues();
        for (var r = 0; r < grid.length; r++) {
          var row = grid[r];
          rows.push({
            date: (row[dateCol - startCol] != null) ? String(row[dateCol - startCol]).trim() : "",
            description: (row[descCol - startCol] != null) ? String(row[descCol - startCol]).trim() : "",
            amount: (row[amountCol - startCol] != null) ? String(row[amountCol - startCol]).trim() : "",
            account: (row[accountCol - startCol] != null) ? String(row[accountCol - startCol]).trim() : "",
            category: (row[catCol - startCol] != null) ? String(row[catCol - startCol]).trim() : ""
          });
        }
      }
    }
  }
  return { categories: opts.categories, accounts: opts.accounts, rows: rows };
}

/**
 * Converts 1-based column index to letter(s).
 */
function quickSearchColIndexToLetter(index) {
  var letter = "";
  var n = index;
  while (n > 0) {
    var r = (n - 1) % 26;
    letter = String.fromCharCode(65 + r) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter || "A";
}

/**
 * Escapes a string for use inside a regex (literal contains match).
 */
function quickSearchEscapeRegex(s) {
  if (s == null || s === "") return "";
  return String(s).replace(/[\\^$.*+?()[\]{}|]/g, "\\$&");
}

/**
 * Ensures the Transactions sheet has two helper columns at the end (visible).
 * Match column: row 1 shows "QuickSearch" (formula output for header row).
 * Criteria column: always the column immediately after Match; row 1 holds the criteria string (so header is overwritten).
 * We find Match by "QuickSearch"; Criteria = Match + 1. Never add columns if we already find Match.
 */
function ensureQuickSearchHelperColumns() {
  var ss = SpreadsheetApp.getActive();
  var sheet = getQuickSearchSheet(ss);
  if (!sheet) return null;

  var lastCol = sheet.getLastColumn();
  if (lastCol < COL_FIRST) {
    sheet.getRange(ROW_HEADER, HELPER_COL_MATCH_DEFAULT, ROW_HEADER, HELPER_COL_CRITERIA_DEFAULT).setValues([["", ""]]);
    sheet.getRange(ROW_HEADER, HELPER_COL_MATCH_DEFAULT).setValue(QUICK_SEARCH_MATCH_HEADER);
    sheet.getRange(ROW_HEADER, HELPER_COL_CRITERIA_DEFAULT).setValue(QUICK_SEARCH_CRITERIA_HEADER);
    return { matchColIndex: HELPER_COL_MATCH_DEFAULT, criteriaColIndex: HELPER_COL_CRITERIA_DEFAULT };
  }
  var startCol = Math.max(COL_FIRST, lastCol - 1);
  var row1 = sheet.getRange(ROW_HEADER, startCol, ROW_HEADER, lastCol).getDisplayValues()[0];
  var h0 = (row1[RANGE_INDEX_FIRST] && String(row1[RANGE_INDEX_FIRST]).trim()) || "";
  var h1 = (row1[RANGE_INDEX_SECOND] != null ? String(row1[RANGE_INDEX_SECOND]).trim() : "") || "";
  if (h0 === QUICK_SEARCH_MATCH_HEADER) {
    return { matchColIndex: startCol, criteriaColIndex: startCol + 1 };
  }
  if (h1 === QUICK_SEARCH_MATCH_HEADER) {
    sheet.insertColumnAfter(lastCol);
    sheet.getRange(ROW_HEADER, lastCol + 1).setValue(QUICK_SEARCH_CRITERIA_HEADER);
    return { matchColIndex: lastCol, criteriaColIndex: lastCol + 1 };
  }
  if (h0 === QUICK_SEARCH_CRITERIA_HEADER) {
    return { matchColIndex: startCol - 1, criteriaColIndex: startCol };
  }
  if (h1 === QUICK_SEARCH_CRITERIA_HEADER) {
    return { matchColIndex: lastCol - 1, criteriaColIndex: lastCol };
  }

  // No existing helper columns: append two columns
  var matchColIndex, criteriaColIndex;
  if (lastCol === 0) {
    sheet.getRange(ROW_HEADER, HELPER_COL_MATCH_DEFAULT, ROW_HEADER, HELPER_COL_CRITERIA_DEFAULT).setValues([["", ""]]);
    sheet.getRange(ROW_HEADER, HELPER_COL_MATCH_DEFAULT).setValue(QUICK_SEARCH_MATCH_HEADER);
    sheet.getRange(ROW_HEADER, HELPER_COL_CRITERIA_DEFAULT).setValue(QUICK_SEARCH_CRITERIA_HEADER);
    matchColIndex = HELPER_COL_MATCH_DEFAULT;
    criteriaColIndex = HELPER_COL_CRITERIA_DEFAULT;
  } else {
    sheet.insertColumnAfter(lastCol);
    sheet.insertColumnAfter(lastCol + 1);
    matchColIndex = lastCol + 1;
    criteriaColIndex = lastCol + 2;
    sheet.getRange(ROW_HEADER, matchColIndex).setValue(QUICK_SEARCH_MATCH_HEADER);
    sheet.getRange(ROW_HEADER, criteriaColIndex).setValue(QUICK_SEARCH_CRITERIA_HEADER);
  }

  return { matchColIndex: matchColIndex, criteriaColIndex: criteriaColIndex };
}

/**
 * Builds the ARRAYFORMULA for the Match column (row 1). Criteria in criteria cell as single string:
 * D:dateFrom,dateTo|X:description|C:cat1,cat2|A:min,max|Q:acc1,acc2
 * Uses LET with d_split/a_split+VALUE, c_esc for category only, LEN checks for empty date/amount, blank row => FALSE.
 */
function buildQuickSearchFormula(tillerCols, criteriaColLetter, dateCol, descCol, amountCol, accountCol, categoryCol) {
  var dateLetter = quickSearchColIndexToLetter(dateCol);
  var descLetter = quickSearchColIndexToLetter(descCol);
  var amountLetter = quickSearchColIndexToLetter(amountCol);
  var accountLetter = quickSearchColIndexToLetter(accountCol);
  var categoryLetter = quickSearchColIndexToLetter(categoryCol);
  var w = criteriaColLetter;
  var reEsc = "([$^*+?.()|\\[\\]{}])";
  var reRep = "\\\\$1";
  var cBlankPat = "\\\\{\\\\\\{B\\\\\\}\\\\}";  // regex in formula to match literal {{B}}
  return "=ARRAYFORMULA(IF(ROW(" + dateLetter + ":" + dateLetter + ")=1,\"" + QUICK_SEARCH_MATCH_HEADER + "\",LET(" +
    "input," + w + "$1,"
    + "d_raw,IFERROR(REGEXEXTRACT(input,\"D:([^|]*)\")),"
    + "x_raw,IFERROR(REGEXEXTRACT(input,\"X:([^|]*)\")),"
    + "c_raw,IFERROR(REGEXEXTRACT(input,\"C:([^|]*)\")),"
    + "a_raw,IFERROR(REGEXEXTRACT(input,\"A:([^|]*)\")),"
    + "q_raw,IFERROR(REGEXEXTRACT(input,\"Q:([^|]*)\")),"
    + "d_split,IFERROR(SPLIT(d_raw,\",\")),"
    + "d_start,IFERROR(VALUE(INDEX(d_split,1,1)),0),"
    + "d_end,IFERROR(VALUE(INDEX(d_split,1,2)),99999),"
    + "a_split,IFERROR(SPLIT(a_raw,\",\")),"
    + "a_min,IFERROR(VALUE(INDEX(a_split,1,1)),-999999),"
    + "a_max,IFERROR(VALUE(INDEX(a_split,1,2)),999999),"
    + "c_no_blank,SUBSTITUTE(c_raw,\"{{B}}\",\"\"),"
    + "c_has_blank,IFERROR(REGEXMATCH(c_raw,\"" + cBlankPat + "\"),FALSE),"
    + "c_esc,REGEXREPLACE(c_no_blank,\"" + reEsc + "\",\"" + reRep + "\"),"
    + "c_match,IF(c_has_blank,\"(?i)^(\"&\"|\"&SUBSTITUTE(c_esc,\",\",\"|\")&\")$\",\"(?i)^(\"&SUBSTITUTE(c_esc,\",\",\"|\")&\")$\"),"
    + "q_match,\"(?i)^(\"&SUBSTITUTE(q_raw,\",\",\"|\")&\")$\","
    + "x_include_raw,TRIM(IFERROR(INDEX(SPLIT(x_raw,\"{{NOT}}\"),1,1),x_raw)),"
    + "x_exclude_raw,IFERROR(TRIM(INDEX(SPLIT(x_raw,\"{{NOT}}\"),1,2)),\"\"),"
    + "x_match_include,\"(?i)\"&TRIM(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(x_include_raw,\" {{PIPE}} \",\"|\"),\" {{PIPE}}\",\"|\"),\"{{PIPE}} \",\"|\"),\"{{PIPE}}\",\"|\")),"
    + "x_match_exclude,\"(?i)\"&TRIM(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(x_exclude_raw,\" {{PIPE}} \",\"|\"),\" {{PIPE}}\",\"|\"),\"{{PIPE}} \",\"|\"),\"{{PIPE}}\",\"|\")),"
    + "d_check,IF(LEN(d_raw)<3,1,(N(" + dateLetter + ":" + dateLetter + ")>=d_start)*(N(" + dateLetter + ":" + dateLetter + ")<=d_end)),"
    + "x_check,IF(x_raw=\"\",1,IF(x_exclude_raw=\"\",REGEXMATCH(TO_TEXT(" + descLetter + ":" + descLetter + "),x_match_include),(REGEXMATCH(TO_TEXT(" + descLetter + ":" + descLetter + "),x_match_include))*(1-(REGEXMATCH(TO_TEXT(" + descLetter + ":" + descLetter + "),x_match_exclude))))),"
    + "c_check,IF(c_raw=\"\",1,REGEXMATCH(TO_TEXT(" + categoryLetter + ":" + categoryLetter + "),c_match)),"
    + "a_check,IF(LEN(a_raw)<2,1,(N(" + amountLetter + ":" + amountLetter + ")>=a_min)*(N(" + amountLetter + ":" + amountLetter + ")<=a_max)),"
    + "q_check,IF(q_raw=\"\",1,REGEXMATCH(TO_TEXT(" + accountLetter + ":" + accountLetter + "),q_match)),"
    + "IF(" + dateLetter + ":" + dateLetter + "=\"\",FALSE,(d_check*x_check*c_check*a_check*q_check)>0))))";
}

/**
 * One-time setup: ensures helper columns exist, sets the array formula if missing, applies the basic filter.
 * After this, only writing criteria cells is needed to update results.
 */
function ensureQuickSearchSetup() {
  var ss = SpreadsheetApp.getActive();
  var sheet = getQuickSearchSheet(ss);
  if (!sheet) return { ok: false, message: "Transactions sheet not found." };

  var cols = ensureQuickSearchHelperColumns();
  if (!cols) return { ok: false, message: "Could not ensure helper columns." };

  var tillerCols = getTillerColumnMap(sheet);
  var dateCol = tillerCols[TRANSACTIONS_HEADER_DATE];
  var descCol = tillerCols[TRANSACTIONS_HEADER_DESCRIPTION];
  var amountCol = tillerCols[TRANSACTIONS_HEADER_AMOUNT];
  var accountCol = tillerCols[TRANSACTIONS_HEADER_ACCOUNT];
  var catCol = tillerCols[TRANSACTIONS_HEADER_CATEGORY] || tillerCols[TRANSACTIONS_HEADER_CATEGORIES];
  if (!catCol) {
    var headers = sheet.getRange(ROW_HEADER, COL_FIRST, ROW_HEADER, sheet.getLastColumn()).getDisplayValues()[0];
    var catLower = TRANSACTIONS_HEADER_CATEGORY.toLowerCase();
    var catsLower = TRANSACTIONS_HEADER_CATEGORIES.toLowerCase();
    for (var i = 0; i < headers.length; i++) {
      var h = (headers[i] && String(headers[i]).trim()).toLowerCase();
      if (h === catLower || h === catsLower) { catCol = i + 1; break; }
    }
  }
  if (!dateCol || !descCol || !amountCol || !accountCol || !catCol) {
    return { ok: false, message: "Required columns (" + TRANSACTIONS_HEADER_DATE + ", " + TRANSACTIONS_HEADER_DESCRIPTION + ", " + TRANSACTIONS_HEADER_AMOUNT + ", " + TRANSACTIONS_HEADER_ACCOUNT + ", " + TRANSACTIONS_HEADER_CATEGORY + ") not found." };
  }

  var criteriaColLetter = quickSearchColIndexToLetter(cols.criteriaColIndex);
  var matchCell = sheet.getRange(ROW_HEADER, cols.matchColIndex);
  var formula = buildQuickSearchFormula(tillerCols, criteriaColLetter, dateCol, descCol, amountCol, accountCol, catCol);
  matchCell.setFormula(formula);

  var lastRow = Math.max(sheet.getLastRow(), ROW_DATA_FIRST);
  applyQuickSearchBasicFilter(sheet, cols.criteriaColIndex, lastRow);
  // Cache criteria column index so write/clear don't use getLastColumn() (which can point to Match when criteria is empty).
  try {
    PropertiesService.getDocumentProperties().setProperty("quickSearchCriteriaCol", String(cols.criteriaColIndex));
  } catch (e) { /* ignore */ }
  return { ok: true };
}

/**
 * Converts a date string (e.g. yyyy-MM-dd from input type="date") to MM/dd/yyyy for the sheet formula (DATEVALUE in US locale).
 */
function quickSearchFormatDateForCriteria(dateStr) {
  if (dateStr == null || String(dateStr).trim() === "") return "";
  var s = String(dateStr).trim();
  var d = new Date(s + (s.indexOf("T") === -1 ? "T12:00:00" : ""));
  if (isNaN(d.getTime())) return s;
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "MM/dd/yyyy");
}

/**
 * Google Sheets REGEXMATCH uses RE2, which does not support lookahead/lookbehind.
 * Returns an error message if desc contains unsupported constructs, or null if ok.
 * If desc contains " but not ", validates both include and exclude parts.
 */
function quickSearchDescriptionRegexUnsupported(desc) {
  if (!desc || typeof desc !== "string") return null;
  var s = desc.trim();
  if (!s.length) return null;
  var parts = s.indexOf(" but not ") !== -1 ? s.split(" but not ") : [s];
  for (var i = 0; i < parts.length; i++) {
    var p = (parts[i] || "").trim();
    if (!p.length) continue;
    if (/\(\?[=!]/.test(p)) return "Lookahead ((?=...) or (?!...)) is not supported in Google Sheets. Use a simpler pattern (e.g. Alaska to match rows containing Alaska).";
    if (/\(\?[<>]/.test(p)) return "Lookbehind ((?<=...) or (?<!...)) is not supported in Google Sheets. Use a simpler pattern.";
  }
  return null;
}

/**
 * Builds the single criteria string for W1. Format: D:dateFrom,dateTo|X:description|C:cat1,cat2|A:min,max|Q:acc1,acc2
 * Dates are sent as MM/dd/yyyy so DATEVALUE in the formula matches the sheet's date format.
 */
function buildQuickSearchCriteriaString(criteria) {
  function safe(s) {
    return (s == null) ? "" : String(s).replace(/\|/g, " ").trim();
  }
  var parts = [];
  var dateFromVal = safe(quickSearchFormatDateForCriteria(criteria.dateFrom));
  var dateToVal = safe(quickSearchFormatDateForCriteria(criteria.dateTo));
  if (dateFromVal || dateToVal) {
    parts.push("D:" + dateFromVal + "," + dateToVal);
  } else {
    parts.push("D:");
  }
  // Description supports regex. " but not " splits into include/exclude (stored as include{{NOT}}exclude). Encode | as {{PIPE}}.
  var descInput = (criteria.description == null ? "" : String(criteria.description).trim());
  var descVal;
  if (descInput.indexOf(" but not ") !== -1) {
    var descParts = descInput.split(" but not ");
    var includePart = (descParts[0] || "").trim().replace(/\|/g, "{{PIPE}}");
    var excludePart = (descParts[1] || "").trim().replace(/\|/g, "{{PIPE}}");
    descVal = includePart + "{{NOT}}" + excludePart;
  } else {
    descVal = descInput.replace(/\|/g, "{{PIPE}}");
  }
  parts.push("X:" + descVal);
  var categoryList = Array.isArray(criteria.categories) ? criteria.categories : [];
  var categoryVal = categoryList.map(function(c) {
    var s = (c == null) ? "" : String(c).trim();
    if (s === "(Blank)") return "{{B}}";  // placeholder so formula can match empty cell
    return quickSearchEscapeRegex(safe(c));
  }).filter(function(x) { return x !== undefined && x !== null; }).join(",");
  parts.push("C:" + categoryVal);
  var amtFrom = criteria.amountFrom;
  var amtTo = criteria.amountTo;
  var amountFromVal = (amtFrom !== "" && amtFrom != null && !isNaN(Number(String(amtFrom).replace(/,/g, "")))) ? String(amtFrom).replace(/,/g, "") : "";
  var amountToVal = (amtTo !== "" && amtTo != null && !isNaN(Number(String(amtTo).replace(/,/g, "")))) ? String(amtTo).replace(/,/g, "") : "";
  parts.push("A:" + amountFromVal + "," + amountToVal);
  var accountList = Array.isArray(criteria.accounts) ? criteria.accounts : [];
  var accountVal = accountList.map(function(a) { return quickSearchEscapeRegex(safe(a)); }).filter(Boolean).join(",");
  parts.push("Q:" + accountVal);
  return parts.join("|");
}

/**
 * Returns { matchColIndex, criteriaColIndex } for existing Quick Search columns, or null.
 * Match column: row 1 displays "QuickSearch". Criteria column: next column.
 * Only reads the last 2 columns (where helpers live) to avoid slow full-row read on wide sheets.
 */
function getQuickSearchColumnIndices(sheet) {
  if (!sheet) return null;
  var lastCol = sheet.getLastColumn();
  if (lastCol < COL_FIRST) return null;
  var startCol = Math.max(COL_FIRST, lastCol - 1);
  var row1 = sheet.getRange(ROW_HEADER, startCol, ROW_HEADER, lastCol).getDisplayValues()[0];
  var h0 = (row1[RANGE_INDEX_FIRST] && String(row1[RANGE_INDEX_FIRST]).trim()) || "";
  var h1 = (row1[RANGE_INDEX_SECOND] != null ? String(row1[RANGE_INDEX_SECOND]).trim() : "") || "";
  if (h0 === QUICK_SEARCH_MATCH_HEADER) {
    return { matchColIndex: startCol, criteriaColIndex: startCol + 1 };
  }
  if (h1 === QUICK_SEARCH_MATCH_HEADER) {
    return { matchColIndex: lastCol, criteriaColIndex: lastCol + 1 };
  }
  if (h0 === QUICK_SEARCH_CRITERIA_HEADER) {
    return { matchColIndex: startCol - 1, criteriaColIndex: startCol };
  }
  if (h1 === QUICK_SEARCH_CRITERIA_HEADER) {
    return { matchColIndex: lastCol - 1, criteriaColIndex: lastCol };
  }
  return null;
}

/** Returns cached criteria column index (1-based) from Document Properties, or null if not set. */
function getQuickSearchCriteriaColCached() {
  try {
    var s = PropertiesService.getDocumentProperties().getProperty("quickSearchCriteriaCol");
    if (s == null || s === "") return null;
    var n = parseInt(s, 10);
    return isNaN(n) || n < 1 ? null : n;
  } catch (e) {
    return null;
  }
}

/** Returns cached sheet ID (numeric) for the Transactions sheet, or null if not set. */
function getQuickSearchSheetIdCached() {
  try {
    var s = PropertiesService.getDocumentProperties().getProperty("quickSearchSheetId");
    if (s == null || s === "") return null;
    var n = parseInt(s, 10);
    return isNaN(n) ? null : n;
  } catch (e) {
    return null;
  }
}

/**
 * Reads all Quick Search cached values from Document Properties in one call (fewer round-trips).
 * Returns { criteriaColIndex: number|null, sheetIdCached: number|null }.
 */
function getQuickSearchCachedProps() {
  try {
    var all = PropertiesService.getDocumentProperties().getProperties();
    var col = all.quickSearchCriteriaCol;
    var criteriaColIndex = (col != null && col !== "") ? parseInt(String(col), 10) : null;
    if (criteriaColIndex == null || isNaN(criteriaColIndex) || criteriaColIndex < 1) criteriaColIndex = null;
    var sid = all.quickSearchSheetId;
    var sheetIdCached = (sid != null && sid !== "") ? parseInt(String(sid), 10) : null;
    if (sheetIdCached != null && isNaN(sheetIdCached)) sheetIdCached = null;
    return { criteriaColIndex: criteriaColIndex, sheetIdCached: sheetIdCached };
  } catch (e) {
    return { criteriaColIndex: null, sheetIdCached: null };
  }
}

function setQuickSearchSheetIdCached(id) {
  try {
    if (id != null && id >= 0) PropertiesService.getDocumentProperties().setProperty("quickSearchSheetId", String(id));
  } catch (e) { /* ignore */ }
}

/**
 * Returns the Transactions sheet, using cached sheet ID when available (faster than getSheetByName on every search).
 * @param {Spreadsheet} ss - Optional; if omitted, uses SpreadsheetApp.getActive().
 * @param {number|null} cachedSheetId - Optional; if provided, avoids a property read (use from getQuickSearchCachedProps()).
 */
function getQuickSearchSheet(ss, cachedSheetId) {
  if (!ss) ss = SpreadsheetApp.getActive();
  var sheetId = cachedSheetId != null ? cachedSheetId : getQuickSearchSheetIdCached();
  var sheet = sheetId != null ? ss.getSheetById(sheetId) : ss.getSheetByName(QUICK_SEARCH_SHEET_NAME);
  if (sheet) setQuickSearchSheetIdCached(sheet.getSheetId());
  return sheet || null;
}

/**
 * Applies Quick Search to the sheet's basic Filter (show only rows where Match column = TRUE).
 * Uses filter.setColumnFilterCriteria so the visible filter refreshes immediately.
 * Returns { ok: boolean, timing: { getRangeMs, getFilterMs, createFilterMs?, buildCriteriaMs, setColumnFilterCriteriaMs } } for granular logging.
 */
function applyQuickSearchBasicFilter(sheet, criteriaColIndex, lastRow) {
  var timing = { getRangeMs: 0, getFilterMs: 0, createFilterMs: 0, buildCriteriaMs: 0, setColumnFilterCriteriaMs: 0 };
  if (!sheet || criteriaColIndex == null || criteriaColIndex < 2) return { ok: false, timing: timing };
  var matchCol1Based = criteriaColIndex - 1;

  var t0 = Date.now();
  var filterRange = sheet.getRange(ROW_HEADER, COL_FIRST, lastRow || Math.max(sheet.getLastRow(), ROW_DATA_FIRST), criteriaColIndex);
  timing.getRangeMs = Date.now() - t0;

  var t1 = Date.now();
  var filter = sheet.getFilter();
  timing.getFilterMs = Date.now() - t1;

  if (filter) {
    var existingRange = filter.getRange();
    var dataLastRow = Math.max(sheet.getLastRow(), ROW_DATA_FIRST);
    if (existingRange.getLastColumn() < criteriaColIndex || existingRange.getLastRow() < dataLastRow) {
      try { filter.remove(); } catch (e) { /* ignore */ }
      filter = null;
    }
  }
  if (!filter) {
    var t1b = Date.now();
    try { filter = filterRange.createFilter(); } catch (e) { return { ok: false, timing: timing }; }
    timing.createFilterMs = Date.now() - t1b;
  }
  if (!filter) return { ok: false, timing: timing };

  var t2 = Date.now();
  var criteria = SpreadsheetApp.newFilterCriteria().setHiddenValues(["FALSE", ""]).build();
  timing.buildCriteriaMs = Date.now() - t2;

  var t3 = Date.now();
  try {
    filter.setColumnFilterCriteria(matchCol1Based, criteria);
    timing.setColumnFilterCriteriaMs = Date.now() - t3;
    return { ok: true, timing: timing };
  } catch (e) {
    timing.setColumnFilterCriteriaMs = Date.now() - t3;
    return { ok: false, timing: timing };
  }
}

/**
 * Refreshes the filter. When the filter already exists (common case), only getFilter + setColumnFilterCriteria
 * (~550 ms). When no filter exists, creates it using API row count and createFilter().
 * @param {Sheet} sheet - Transactions sheet.
 * @param {number} criteriaColIndex - Criteria column index (1-based).
 * @param {string} spreadsheetId - Spreadsheet ID.
 * @param {number} sheetId - Sheet ID.
 * @returns {{ ok: boolean, sheetId?: number, timing: object }}
 */
function refreshQuickSearchFilterView(sheet, criteriaColIndex, spreadsheetId, sheetId) {
  var timing = { getLastRowMs: 0, getFilterViewIdMs: 0, batchUpdateMs: 0, applyBasicFilterMs: 0, getFilterMs: 0, createFilterMs: 0, setColumnFilterCriteriaMs: 0 };
  if (!sheet || criteriaColIndex == null || criteriaColIndex < 2) return { ok: false, timing: timing };

  var matchCol1Based = criteriaColIndex - 1;
  var criteria = SpreadsheetApp.newFilterCriteria().setHiddenValues(["FALSE", ""]).build();

  var t0 = Date.now();
  var filter = sheet.getFilter();
  timing.getFilterMs = Date.now() - t0;

  var dataLastRow = Math.max(sheet.getLastRow(), ROW_DATA_FIRST);
  if (filter) {
    try {
      var fr = filter.getRange();
      if (fr.getLastRow() < dataLastRow || fr.getLastColumn() < criteriaColIndex) {
        try {
          filter.remove();
        } catch (eRm) { /* ignore */ }
        filter = null;
      }
    } catch (eR) { /* ignore */ }
  }

  if (filter) {
    var t1 = Date.now();
    try { filter.setColumnFilterCriteria(matchCol1Based, criteria); } catch (e) { /* ignore */ }
    timing.setColumnFilterCriteriaMs = Date.now() - t1;
    timing.applyBasicFilterBreakdown = {
      setBasicFilterApiMs: 0,
      getFilterMs: timing.getFilterMs,
      createFilterMs: 0,
      setColumnFilterCriteriaMs: timing.setColumnFilterCriteriaMs
    };
    return { ok: true, sheetId: sheetId, timing: timing };
  }

  var t2 = Date.now();
  var lastRow = getQuickSearchSheetGridRowCount(spreadsheetId, sheetId, sheet);
  timing.getLastRowMs = Date.now() - t2;

  var filterRange = sheet.getRange(ROW_HEADER, COL_FIRST, lastRow, criteriaColIndex);
  var t3 = Date.now();
  try { filter = filterRange.createFilter(); } catch (e) { /* ignore */ }
  timing.createFilterMs = Date.now() - t3;
  filter = sheet.getFilter();
  if (filter) {
    var t4 = Date.now();
    try { filter.setColumnFilterCriteria(matchCol1Based, criteria); } catch (e) { /* ignore */ }
    timing.setColumnFilterCriteriaMs = Date.now() - t4;
  }
  timing.applyBasicFilterMs = timing.getFilterMs + timing.createFilterMs + timing.setColumnFilterCriteriaMs;
  timing.applyBasicFilterBreakdown = {
    setBasicFilterApiMs: 0,
    getFilterMs: timing.getFilterMs,
    createFilterMs: timing.createFilterMs,
    setColumnFilterCriteriaMs: timing.setColumnFilterCriteriaMs
  };
  return { ok: true, sheetId: sheetId, timing: timing };
}

/**
 * Writes the single criteria string to the criteria cell (row 1). Uses cached criteria column index from setup
 * (getLastColumn() would point to Match column when criteria cell is empty, overwriting the formula).
 * If refreshView is true, refreshes the basic filter so the sheet UI updates immediately.
 * Returns timingBreakdown: { getSheetMs, getColMs, buildCriteriaMs, setValueMs, refreshTiming } for performance logging.
 */
function writeQuickSearchCriteria(criteriaJson, refreshView) {
  var tInvocation = Date.now();
  var t0 = Date.now();
  var timing = { invocationToFirstOpMs: t0 - tInvocation, getSheetMs: 0, getColMs: 0, buildCriteriaMs: 0, setValueMs: 0, refresh: null };
  var criteria = JSON.parse(criteriaJson || "{}");
  var t1 = Date.now();
  var cached = getQuickSearchCachedProps();
  var ss = SpreadsheetApp.getActive();
  var sheet = getQuickSearchSheet(ss, cached.sheetIdCached);
  timing.getSheetMs = Date.now() - t1;
  timing.getColMs = 0;
  if (!sheet) return { ok: false, message: "Transactions sheet not found.", serverMs: Date.now() - t0, timingBreakdown: timing };

  var criteriaColIndex = cached.criteriaColIndex;
  if (criteriaColIndex == null || criteriaColIndex < 1) {
    return { ok: false, message: "Quick Search not set up. Run setup first.", serverMs: Date.now() - t0, timingBreakdown: timing };
  }
  // When criteria cell is empty, getLastColumn() can be Match column (one less); only treat as not set up if criteria column is beyond that.
  if (criteriaColIndex > sheet.getLastColumn() + 1) {
    return { ok: false, message: "Quick Search not set up. Run setup first.", serverMs: Date.now() - t0, timingBreakdown: timing };
  }

  var descError = quickSearchDescriptionRegexUnsupported(criteria.description);
  if (descError) return { ok: false, message: descError, serverMs: Date.now() - t0, timingBreakdown: timing };

  var t3 = Date.now();
  var criteriaString = buildQuickSearchCriteriaString(criteria);
  timing.buildCriteriaMs = Date.now() - t3;

  var t4 = Date.now();
  sheet.getRange(ROW_HEADER, criteriaColIndex).setValue(criteriaString);
  timing.setValueMs = Date.now() - t4;

  var out = { ok: true, message: "Criteria updated.", serverMs: Date.now() - t0, timingBreakdown: timing };
  if (refreshView) {
    SpreadsheetApp.flush();
    var ref = refreshQuickSearchFilterView(sheet, criteriaColIndex, ss.getId(), sheet.getSheetId());
    if (ref && ref.timing) out.timingBreakdown.refresh = ref.timing;
    if (ref && ref.sheetId != null) out.sheetId = ref.sheetId;
  }
  return out;
}

/**
 * Applies Quick Search: ensures setup once (formula + basic filter), then writes criteria and refreshes the filter.
 */
function applyQuickSearch(criteriaJson) {
  var setup = ensureQuickSearchSetup();
  if (!setup.ok) return setup;
  return writeQuickSearchCriteria(criteriaJson, true);
}

/**
 * Timing test: measures how long it takes to (1) get criteria column (cached), (2) write a string to criteria cell, (3) clear it.
 */
function testQuickSearchCellTiming() {
  var ss = SpreadsheetApp.getActive();
  var sheet = getQuickSearchSheet(ss);
  if (!sheet) return { ok: false, message: "Transactions sheet not found.", getColumnMs: 0, writeMs: 0, clearMs: 0 };

  var t0 = Date.now();
  var criteriaColIndex = getQuickSearchCriteriaColCached();
  var t1 = Date.now();
  if (criteriaColIndex == null || criteriaColIndex < 1) {
    return { ok: false, message: "Quick Search not set up (no cached column). Run setup first.", getColumnMs: t1 - t0, writeMs: 0, clearMs: 0 };
  }

  var cell = sheet.getRange(ROW_HEADER, criteriaColIndex);
  cell.setValue("timing test");
  var t2 = Date.now();
  cell.setValue("");
  var t3 = Date.now();

  return {
    ok: true,
    getColumnMs: t1 - t0,
    writeMs: t2 - t1,
    clearMs: t3 - t2,
    message: "getLastColumn: " + (t1 - t0) + " ms, write: " + (t2 - t1) + " ms, clear: " + (t3 - t2) + " ms"
  };
}

/**
 * Clears Quick Search: clears the criteria cell so all rows match, then refreshes the filter (same as Search button).
 * Returns timingBreakdown for performance logging.
 */
function clearQuickSearch() {
  var t0 = Date.now();
  var timing = { getSheetMs: 0, getColMs: 0, setValueMs: 0, refresh: null };
  var t1 = Date.now();
  var cached = getQuickSearchCachedProps();
  var ss = SpreadsheetApp.getActive();
  var sheet = getQuickSearchSheet(ss, cached.sheetIdCached);
  timing.getSheetMs = Date.now() - t1;
  timing.getColMs = 0;
  if (!sheet) return { ok: true, serverMs: Date.now() - t0, timingBreakdown: timing };

  var criteriaColIndex = cached.criteriaColIndex;
  if (criteriaColIndex != null && criteriaColIndex >= 1) {
    if (criteriaColIndex <= sheet.getLastColumn() + 1) {
      var t3 = Date.now();
      sheet.getRange(ROW_HEADER, criteriaColIndex).setValue("");
      timing.setValueMs = Date.now() - t3;
      SpreadsheetApp.flush();
      var ref = refreshQuickSearchFilterView(sheet, criteriaColIndex, ss.getId(), sheet.getSheetId());
      if (ref && ref.timing) timing.refresh = ref.timing;
    } else {
      try { PropertiesService.getDocumentProperties().deleteProperty("quickSearchCriteriaCol"); } catch (e) { /* ignore */ }
    }
  }
  return { ok: true, serverMs: Date.now() - t0, timingBreakdown: timing };
}

/**
 * Row count for the filter data range: max of Sheets API grid rowCount and sheet.getLastRow().
 * Avoids returning 2 on API failure (that used to create a 2-row basic filter with no effect on real data).
 * @param {string} spreadsheetId
 * @param {number} sheetId
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function getQuickSearchSheetGridRowCount(spreadsheetId, sheetId, sheet) {
  var fallback = sheet ? Math.max(sheet.getLastRow(), ROW_DATA_FIRST) : ROW_DATA_FIRST;
  try {
    var spread = Sheets.Spreadsheets.get(spreadsheetId, { fields: "sheets(properties(sheetId,gridProperties(rowCount)))" });
    var sheets = spread.sheets || [];
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].properties && Number(sheets[i].properties.sheetId) === Number(sheetId)) {
        var rc = sheets[i].properties.gridProperties ? sheets[i].properties.gridProperties.rowCount : undefined;
        var gridRows = Math.max(rc != null ? Number(rc) : 0, ROW_DATA_FIRST);
        return Math.max(gridRows, fallback);
      }
    }
  } catch (e) { /* ignore */ }
  return fallback;
}
