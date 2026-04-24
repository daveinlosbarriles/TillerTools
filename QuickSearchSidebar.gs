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

// --- Scan row 1 for helper headers: must cover columns to the RIGHT of Match/Criteria (e.g. "Priority") ---
var QUICK_SEARCH_HEADER_SCAN_MAX_COLS = 120;

// --- Quick Search helper column headers ---
var QUICK_SEARCH_MATCH_HEADER = "QuickSearch";
var QUICK_SEARCH_CRITERIA_HEADER = "QuickCriteria";
// Criteria: single string in row 1 only. Format: D:dateFrom,dateTo|X:description|C:cat1,cat2|A:min,max|Q:acc1,acc2

// Literal fragments inside buildQuickSearchFormula — used to tell our match ARRAYFORMULA from any other ARRAYFORMULA on the sheet.
var QUICK_SEARCH_AF_ROW1_SNIPPET = ')=1,"' + QUICK_SEARCH_MATCH_HEADER + '",LET(';
var QUICK_SEARCH_AF_D_BLOCK_SNIPPET = 'REGEXEXTRACT(input,"D:([^|]*)"';

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
  html.setTitle("Tiller™ Quick Search");
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
 * Reads legacy last-two-column row-1 labels (used for debug + edge cases).
 * @returns {{ legacyStart: number, h0: string, h1: string }}
 */
function quickSearchLegacyLastTwoHeaders_(sheet, lastCol) {
  var legacyStart = Math.max(COL_FIRST, lastCol - 1);
  var row1 = sheet.getRange(ROW_HEADER, legacyStart, ROW_HEADER, lastCol).getDisplayValues()[0];
  var h0 = (row1[RANGE_INDEX_FIRST] && String(row1[RANGE_INDEX_FIRST]).trim()) || "";
  var h1 = (row1[RANGE_INDEX_SECOND] != null ? String(row1[RANGE_INDEX_SECOND]).trim() : "") || "";
  return { legacyStart: legacyStart, h0: h0, h1: h1 };
}

/**
 * Finds Quick Search helper columns by scanning row 1 from the right (Match shows "QuickSearch" from ARRAYFORMULA).
 * Wider than the last two columns so a column after Criteria (e.g. "Priority") does not hide the helpers from detection.
 * @returns {{ matchColIndex: number|null, criteriaColIndex: number|null, needCriteriaInsert: boolean, debug: object }}
 */
function quickSearchResolveHelperColumnsFromSheet_(sheet) {
  var lastCol = sheet.getLastColumn();
  var leg = quickSearchLegacyLastTwoHeaders_(sheet, lastCol);
  var debug = {
    hypothesisId: "H1",
    lastCol: lastCol,
    legacyH0: leg.h0,
    legacyH1: leg.h1,
    scanWidth: 0,
    scanStart: 0,
    path: "none"
  };
  if (lastCol < COL_FIRST) {
    debug.path = "empty_sheet";
    return { matchColIndex: null, criteriaColIndex: null, needCriteriaInsert: false, debug: debug };
  }

  var scanWidth = Math.min(QUICK_SEARCH_HEADER_SCAN_MAX_COLS, lastCol);
  var scanStart = Math.max(COL_FIRST, lastCol - scanWidth + 1);
  debug.scanWidth = scanWidth;
  debug.scanStart = scanStart;
  var row1 = sheet.getRange(ROW_HEADER, scanStart, ROW_HEADER, lastCol).getDisplayValues()[0];

  var i;
  for (i = row1.length - 1; i >= 0; i--) {
    var cell = (row1[i] && String(row1[i]).trim()) || "";
    if (cell === QUICK_SEARCH_MATCH_HEADER) {
      var matchColIndex = scanStart + i;
      if (matchColIndex === lastCol) {
        debug.path = "wide_scan_match_last_needs_criteria_col";
        return { matchColIndex: matchColIndex, criteriaColIndex: null, needCriteriaInsert: true, debug: debug };
      }
      debug.path = "wide_scan_match";
      debug.resolvedMatchCol = matchColIndex;
      debug.resolvedCriteriaCol = matchColIndex + 1;
      return { matchColIndex: matchColIndex, criteriaColIndex: matchColIndex + 1, needCriteriaInsert: false, debug: debug };
    }
  }
  for (i = row1.length - 1; i >= 0; i--) {
    var c2 = (row1[i] && String(row1[i]).trim()) || "";
    if (c2 === QUICK_SEARCH_CRITERIA_HEADER) {
      var critCol = scanStart + i;
      if (critCol > COL_FIRST) {
        debug.path = "wide_scan_criteria_literal";
        debug.resolvedMatchCol = critCol - 1;
        debug.resolvedCriteriaCol = critCol;
        return { matchColIndex: critCol - 1, criteriaColIndex: critCol, needCriteriaInsert: false, debug: debug };
      }
    }
  }

  if (leg.h0 === QUICK_SEARCH_MATCH_HEADER) {
    debug.path = "legacy_last2_h0_match";
    return { matchColIndex: leg.legacyStart, criteriaColIndex: leg.legacyStart + 1, needCriteriaInsert: false, debug: debug };
  }
  if (leg.h1 === QUICK_SEARCH_MATCH_HEADER) {
    debug.path = "legacy_last2_h1_match";
    return { matchColIndex: lastCol, criteriaColIndex: null, needCriteriaInsert: true, debug: debug };
  }
  if (leg.h0 === QUICK_SEARCH_CRITERIA_HEADER && leg.legacyStart > COL_FIRST) {
    debug.path = "legacy_last2_h0_criteria_header";
    return { matchColIndex: leg.legacyStart - 1, criteriaColIndex: leg.legacyStart, needCriteriaInsert: false, debug: debug };
  }
  if (leg.h1 === QUICK_SEARCH_CRITERIA_HEADER) {
    debug.path = "legacy_last2_h1_criteria_header";
    return { matchColIndex: lastCol - 1, criteriaColIndex: lastCol, needCriteriaInsert: false, debug: debug };
  }

  debug.path = "append_new_pair";
  debug.matchHeaderHitsInScan = quickSearchCountMatchHeadersInRow1_(row1, scanStart);
  return { matchColIndex: null, criteriaColIndex: null, needCriteriaInsert: false, debug: debug };
}

/**
 * Counts row-1 cells in the scanned range whose display value equals QuickSearch (diagnostics only).
 */
function quickSearchCountMatchHeadersInRow1_(row1, scanStart) {
  var hits = [];
  var k;
  for (k = 0; k < row1.length; k++) {
    var cell = (row1[k] && String(row1[k]).trim()) || "";
    if (cell === QUICK_SEARCH_MATCH_HEADER) {
      hits.push(scanStart + k);
    }
  }
  return { count: hits.length, cols: hits };
}

/**
 * Last N columns of row 1: 1-based index, column letter, truncated display (for logs).
 */
function quickSearchRow1TailSamples_(sheet, lastCol, tailN) {
  if (!sheet || lastCol < 1) return [];
  var n = Math.min(tailN, lastCol);
  var c0 = Math.max(COL_FIRST, lastCol - n + 1);
  var row = sheet.getRange(ROW_HEADER, c0, ROW_HEADER, lastCol).getDisplayValues()[0];
  var out = [];
  var j;
  for (j = 0; j < row.length; j++) {
    var colIdx = c0 + j;
    var cv = row[j] != null ? String(row[j]).trim() : "";
    out.push({ col: colIdx, letter: quickSearchColIndexToLetter(colIdx), disp: cv.substring(0, 56) });
  }
  return out;
}

/**
 * Match column is immediately left of criteria (row 1). It must contain **this** add-on's ARRAYFORMULA
 * (same strings as buildQuickSearchFormula), not some other ARRAYFORMULA elsewhere on the sheet.
 */
function quickSearchMatchColumnHasOurArrayFormula_(sheet, criteriaColIndex) {
  if (!sheet || criteriaColIndex == null || criteriaColIndex < 2) return false;
  var mCol = criteriaColIndex - 1;
  try {
    var f = sheet.getRange(ROW_HEADER, mCol).getFormula();
    var s = f == null ? "" : String(f).trim();
    if (!s) return false;
    var u = s.toUpperCase();
    if (u.indexOf("ARRAYFORMULA") === -1) return false;
    if (s.indexOf(QUICK_SEARCH_AF_ROW1_SNIPPET) === -1) return false;
    if (s.indexOf(QUICK_SEARCH_AF_D_BLOCK_SNIPPET) === -1) return false;
    return true;
  } catch (e) {
    return false;
  }
}

/**
 * Ensures the Transactions sheet has two helper columns at the end (visible).
 * Match column: row 1 shows "QuickSearch" (formula output for header row).
 * Criteria column: always the column immediately after Match; row 1 holds the criteria string (so header is overwritten).
 * We find Match by scanning row 1 from the right for "QuickSearch" (not only the last two columns), so extra columns
 * after Criteria do not prevent detection. Criteria = Match + 1. Never add columns if we already find Match.
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
    return {
      matchColIndex: HELPER_COL_MATCH_DEFAULT,
      criteriaColIndex: HELPER_COL_CRITERIA_DEFAULT,
      _agentDebug: { hypothesisId: "H1", path: "init_empty_grid", lastCol: lastCol }
    };
  }

  var resolved = quickSearchResolveHelperColumnsFromSheet_(sheet);
  var dbg = resolved.debug || {};
  dbg.row1Tail = quickSearchRow1TailSamples_(sheet, lastCol, 10);
  dbg.phase = "after_resolve";

  if (resolved.matchColIndex != null) {
    var badMatch = resolved.matchColIndex < COL_FIRST;
    var badCrit = !resolved.needCriteriaInsert && (resolved.criteriaColIndex == null || resolved.criteriaColIndex < COL_FIRST);
    if (badMatch || badCrit) {
      dbg.phase = "invalid_resolved_indices_fallback_append";
      dbg.badMatch = badMatch ? resolved.matchColIndex : null;
      dbg.badCrit = badCrit ? resolved.criteriaColIndex : null;
      resolved = { matchColIndex: null, criteriaColIndex: null, needCriteriaInsert: false, debug: dbg };
    }
  }

  if (resolved.matchColIndex != null) {
    if (resolved.needCriteriaInsert) {
      try {
        sheet.insertColumnAfter(lastCol);
        sheet.getRange(ROW_HEADER, lastCol + 1).setValue(QUICK_SEARCH_CRITERIA_HEADER);
      } catch (eIns) {
        dbg.insertCriteriaError = String(eIns && eIns.message ? eIns.message : eIns);
        dbg.phase = "insert_criteria_failed";
        return { matchColIndex: null, criteriaColIndex: null, _agentDebug: dbg };
      }
      dbg.path = (dbg.path || "") + "_inserted_criteria";
      dbg.phase = "inserted_criteria_only";
      dbg.lastColAfter = sheet.getLastColumn();
      return {
        matchColIndex: lastCol,
        criteriaColIndex: lastCol + 1,
        _agentDebug: dbg
      };
    }
    if (!quickSearchMatchColumnHasOurArrayFormula_(sheet, resolved.criteriaColIndex)) {
      dbg.phase = "reused_row1_quicksearch_but_match_has_not_our_arrayformula_append";
      dbg.staleReuseMatchCol = resolved.matchColIndex;
      dbg.staleReuseCriteriaCol = resolved.criteriaColIndex;
      resolved = { matchColIndex: null, criteriaColIndex: null, needCriteriaInsert: false, debug: dbg };
    } else {
      dbg.phase = "reused_existing_helpers";
      return {
        matchColIndex: resolved.matchColIndex,
        criteriaColIndex: resolved.criteriaColIndex,
        _agentDebug: dbg
      };
    }
  }

  // No existing helper columns: append two columns
  var matchColIndex;
  var criteriaColIndex;
  if (lastCol === 0) {
    sheet.getRange(ROW_HEADER, HELPER_COL_MATCH_DEFAULT, ROW_HEADER, HELPER_COL_CRITERIA_DEFAULT).setValues([["", ""]]);
    sheet.getRange(ROW_HEADER, HELPER_COL_MATCH_DEFAULT).setValue(QUICK_SEARCH_MATCH_HEADER);
    sheet.getRange(ROW_HEADER, HELPER_COL_CRITERIA_DEFAULT).setValue(QUICK_SEARCH_CRITERIA_HEADER);
    matchColIndex = HELPER_COL_MATCH_DEFAULT;
    criteriaColIndex = HELPER_COL_CRITERIA_DEFAULT;
    dbg.phase = "init_pair_at_default_AB";
  } else {
    var lastColBeforeInsert = lastCol;
    dbg.lastColBeforeInsert = lastColBeforeInsert;
    try {
      sheet.insertColumnAfter(lastCol);
      sheet.insertColumnAfter(lastCol + 1);
    } catch (eApp) {
      dbg.appendInsertError = String(eApp && eApp.message ? eApp.message : eApp);
      dbg.phase = "append_two_columns_failed";
      return { matchColIndex: null, criteriaColIndex: null, _agentDebug: dbg };
    }
    matchColIndex = lastCol + 1;
    criteriaColIndex = lastCol + 2;
    sheet.getRange(ROW_HEADER, matchColIndex).setValue(QUICK_SEARCH_MATCH_HEADER);
    sheet.getRange(ROW_HEADER, criteriaColIndex).setValue(QUICK_SEARCH_CRITERIA_HEADER);
    dbg.phase = "appended_two_columns";
    dbg.lastColAfterInsert = sheet.getLastColumn();
  }

  dbg.appendedMatchCol = matchColIndex;
  dbg.appendedCriteriaCol = criteriaColIndex;
  return { matchColIndex: matchColIndex, criteriaColIndex: criteriaColIndex, _agentDebug: dbg };
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
    // SPLIT must use FALSE: Sheets defaults split_by_each=TRUE, so delimiter "{{NOT}}" splits on each char including N — breaking plain "MN" into "M" only.
    + "x_include_raw,TRIM(IFERROR(INDEX(SPLIT(x_raw,\"{{NOT}}\",FALSE),1,1),x_raw)),"
    + "x_exclude_raw,IFERROR(TRIM(INDEX(SPLIT(x_raw,\"{{NOT}}\",FALSE),1,2)),\"\"),"
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
  if (!cols) {
    return { ok: false, message: "Could not ensure helper columns.", agentDebug: { sessionId: "8efe09", hypothesisId: "H6", location: "ensureQuickSearchSetup", phase: "ensure_helpers_returned_null" } };
  }
  if (cols.matchColIndex == null || cols.criteriaColIndex == null) {
    return {
      ok: false,
      message: "Could not ensure helper columns (invalid indices after insert).",
      agentDebug: Object.assign({ sessionId: "8efe09", hypothesisId: "H6", location: "ensureQuickSearchSetup", phase: "ensure_helpers_invalid_indices" }, cols._agentDebug || {})
    };
  }

  var setupAgentDebug = cols._agentDebug || null;
  var matchColUse = cols.matchColIndex;
  var criteriaColUse = cols.criteriaColIndex;

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
    return {
      ok: false,
      message: "Required columns (" + TRANSACTIONS_HEADER_DATE + ", " + TRANSACTIONS_HEADER_DESCRIPTION + ", " + TRANSACTIONS_HEADER_AMOUNT + ", " + TRANSACTIONS_HEADER_ACCOUNT + ", " + TRANSACTIONS_HEADER_CATEGORY + ") not found.",
      agentDebug: Object.assign(
        {
          sessionId: "8efe09",
          hypothesisId: "H5",
          location: "ensureQuickSearchSetup",
          phase: "missing_tiller_headers",
          hasDate: !!dateCol,
          hasDesc: !!descCol,
          hasAmount: !!amountCol,
          hasAccount: !!accountCol,
          hasCategory: !!catCol,
          sheetLastCol: sheet.getLastColumn()
        },
        setupAgentDebug || {}
      )
    };
  }

  var criteriaColLetter = quickSearchColIndexToLetter(criteriaColUse);
  var matchCell = sheet.getRange(ROW_HEADER, matchColUse);
  var formula = buildQuickSearchFormula(tillerCols, criteriaColLetter, dateCol, descCol, amountCol, accountCol, catCol);
  try {
    matchCell.setFormula(formula);
  } catch (eForm) {
    return {
      ok: false,
      message: "Could not set Quick Search formula: " + String(eForm && eForm.message ? eForm.message : eForm),
      agentDebug: Object.assign(
        {
          sessionId: "8efe09",
          hypothesisId: "H7",
          location: "ensureQuickSearchSetup",
          phase: "setFormula_threw",
          matchColUse: matchColUse,
          criteriaColUse: criteriaColUse,
          criteriaColLetter: criteriaColLetter,
          formulaLen: formula ? String(formula).length : 0,
          setFormulaError: String(eForm && eForm.message ? eForm.message : eForm)
        },
        setupAgentDebug || {}
      )
    };
  }

  var lastRow = Math.max(sheet.getLastRow(), ROW_DATA_FIRST);
  applyQuickSearchBasicFilter(sheet, criteriaColUse, lastRow);
  // Cache criteria column index so write/clear don't use getLastColumn() (which can point to Match when criteria is empty).
  try {
    PropertiesService.getDocumentProperties().setProperty("quickSearchCriteriaCol", String(criteriaColUse));
  } catch (e) { /* ignore */ }
  var out = { ok: true };
  if (setupAgentDebug) {
    out.agentDebug = Object.assign(
      {
        sessionId: "8efe09",
        location: "ensureQuickSearchSetup",
        phase: "setup_ok",
        cachedCriteriaAfter: criteriaColUse,
        matchColAfter: matchColUse,
        criteriaColLetter: criteriaColLetter,
        formulaLen: formula ? String(formula).length : 0
      },
      setupAgentDebug
    );
  }
  return out;
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
 * Scans a trailing window of row 1 (same as ensureQuickSearchHelperColumns) so columns after Criteria do not hide helpers.
 */
function getQuickSearchColumnIndices(sheet) {
  if (!sheet) return null;
  var lastCol = sheet.getLastColumn();
  if (lastCol < COL_FIRST) return null;
  var resolved = quickSearchResolveHelperColumnsFromSheet_(sheet);
  if (resolved.matchColIndex == null) return null;
  if (resolved.needCriteriaInsert) return null;
  return { matchColIndex: resolved.matchColIndex, criteriaColIndex: resolved.criteriaColIndex };
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
  var sheet = null;
  if (sheetId != null) {
    try {
      sheet = ss.getSheetById(sheetId);
    } catch (e) {
      sheet = null;
    }
  }
  // Stale cache: sheet was deleted/recreated or id no longer valid — fall back to name.
  if (!sheet) {
    sheet = ss.getSheetByName(QUICK_SEARCH_SHEET_NAME);
  }
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

  var resolvedLastRow = lastRow || Math.max(sheet.getLastRow(), ROW_DATA_FIRST);
  var maxRows = 0;
  try {
    maxRows = sheet.getMaxRows();
  } catch (e) { /* ignore */ }
  if (maxRows > 0) {
    resolvedLastRow = Math.min(resolvedLastRow, maxRows);
  }

  var t0 = Date.now();
  var filterRange = sheet.getRange(ROW_HEADER, COL_FIRST, resolvedLastRow, criteriaColIndex);
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
    amzEnsureSheetGridCovers(sheet, resolvedLastRow, criteriaColIndex);
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
 * (~550 ms). When no filter exists, creates it using the same getLastRow/getMaxRows + grid nudge pattern as
 * Amazon CSV import (amzApplyTransactionsSortAndFilterCore_), then createFilter().
 * @param {Sheet} sheet - Transactions sheet.
 * @param {number} criteriaColIndex - Criteria column index (1-based).
 * @param {string} spreadsheetId - Unused; kept for call-site compatibility.
 * @param {number} sheetId - Passed through to the return value for client compatibility.
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
  var lastRow = getQuickSearchSheetGridRowCount(sheet);
  timing.getLastRowMs = Date.now() - t2;

  amzEnsureSheetGridCovers(sheet, lastRow, criteriaColIndex);
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
  if (!sheet) {
    return {
      ok: false,
      message: "Transactions sheet not found.",
      serverMs: Date.now() - t0,
      timingBreakdown: timing,
      agentDebug: { sessionId: "8efe09", hypothesisId: "H8", location: "writeQuickSearchCriteria", phase: "no_transactions_sheet" }
    };
  }

  var criteriaColIndex = cached.criteriaColIndex;
  if (criteriaColIndex == null || criteriaColIndex < 1) {
    return {
      ok: false,
      message: "Quick Search not set up. Run setup first.",
      serverMs: Date.now() - t0,
      timingBreakdown: timing,
      agentDebug: {
        sessionId: "8efe09",
        hypothesisId: "H2",
        location: "writeQuickSearchCriteria",
        phase: "criteria_col_not_cached",
        cachedCriteriaCol: criteriaColIndex,
        sheetLastCol: sheet.getLastColumn()
      }
    };
  }
  // When criteria cell is empty, getLastColumn() can be Match column (one less); only treat as not set up if criteria column is beyond that.
  if (criteriaColIndex > sheet.getLastColumn() + 1) {
    return {
      ok: false,
      message: "Quick Search not set up. Run setup first.",
      serverMs: Date.now() - t0,
      timingBreakdown: timing,
      agentDebug: {
        sessionId: "8efe09",
        hypothesisId: "H2",
        location: "writeQuickSearchCriteria",
        phase: "criteria_col_index_out_of_range",
        cachedCriteriaCol: criteriaColIndex,
        sheetLastCol: sheet.getLastColumn()
      }
    };
  }

  var descError = quickSearchDescriptionRegexUnsupported(criteria.description);
  if (descError) {
    return {
      ok: false,
      message: descError,
      serverMs: Date.now() - t0,
      timingBreakdown: timing,
      agentDebug: { sessionId: "8efe09", hypothesisId: "H9", location: "writeQuickSearchCriteria", phase: "description_regex_blocked" }
    };
  }

  if (!quickSearchMatchColumnHasOurArrayFormula_(sheet, criteriaColIndex)) {
    var matchFp = "";
    try {
      matchFp = (sheet.getRange(ROW_HEADER, criteriaColIndex - 1).getFormula() || "").substring(0, 80);
    } catch (eM) {
      matchFp = "(read_error)";
    }
    return {
      ok: false,
      message: "Quick Search not set up. Run setup first.",
      serverMs: Date.now() - t0,
      timingBreakdown: timing,
      agentDebug: {
        sessionId: "8efe09",
        hypothesisId: "H11",
        location: "writeQuickSearchCriteria",
        phase: "match_column_missing_our_quicksearch_formula",
        cachedCriteriaCol: criteriaColIndex,
        expectedMatchCol: criteriaColIndex - 1,
        matchColFormulaPrefix: matchFp,
        sheetLastCol: sheet.getLastColumn()
      }
    };
  }

  var t3 = Date.now();
  var criteriaString = buildQuickSearchCriteriaString(criteria);
  timing.buildCriteriaMs = Date.now() - t3;

  var t4 = Date.now();
  sheet.getRange(ROW_HEADER, criteriaColIndex).setValue(criteriaString);
  timing.setValueMs = Date.now() - t4;

  var out = { ok: true, message: "Criteria updated.", serverMs: Date.now() - t0, timingBreakdown: timing };
  // #region agent log
  try {
    var r1Snippet = "";
    try {
      var dv = sheet.getRange(ROW_HEADER, criteriaColIndex).getDisplayValue();
      r1Snippet = dv != null ? String(dv).substring(0, 80) : "";
    } catch (eR1) { r1Snippet = "(read_error)"; }
    out.agentDebug = {
      sessionId: "8efe09",
      hypothesisId: "H2",
      location: "writeQuickSearchCriteria",
      phase: "write_ok",
      cachedCriteriaCol: criteriaColIndex,
      sheetLastCol: sheet.getLastColumn(),
      row1CriteriaDisplayPrefix: r1Snippet,
      row1Tail: quickSearchRow1TailSamples_(sheet, sheet.getLastColumn(), 8)
    };
  } catch (eDbg) { /* ignore */ }
  // #endregion
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
  var w = writeQuickSearchCriteria(criteriaJson, true);
  if (setup.agentDebug) {
    w.setupAgentDebug = setup.agentDebug;
  }
  return w;
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
 * Last row index for the Quick Search basic filter range (1-based), matching Amazon import metadata filter logic:
 * max(getLastRow(), ROW_DATA_FIRST), then clamp to getMaxRows() when the grid reports a positive max.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {number}
 */
function getQuickSearchSheetGridRowCount(sheet) {
  if (!sheet) return ROW_DATA_FIRST;
  var lastRowRaw = 0;
  try {
    lastRowRaw = sheet.getLastRow();
  } catch (e) { /* ignore */ }
  var lastRow = Math.max(lastRowRaw, ROW_DATA_FIRST);
  var maxRows = 0;
  try {
    maxRows = sheet.getMaxRows();
  } catch (e) { /* ignore */ }
  if (maxRows > 0) {
    lastRow = Math.min(lastRow, maxRows);
  }
  return lastRow;
}
