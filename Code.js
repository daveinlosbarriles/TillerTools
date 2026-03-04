// Adds Tiller Tools menu to Google Sheets
// Allows launching the Amazon Orders import dialog and Quick Search sidebar
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Tiller Tools")
    .addItem("Amazon Orders Import", "importAmazonCSV_LocalUpload")
    .addItem("Quick Search", "openQuickSearchSidebar")
    .addSeparator()
    .addItem("Quick Search - Run setup", "runQuickSearchSetup")
    .addItem("Quick Search - Cell timing test", "runQuickSearchCellTimingTest")
    .addToUi();
}

/** Runs Quick Search setup (helper columns, formula, filter view). Use once to fix or update the formula. */
function runQuickSearchSetup() {
  var result = ensureQuickSearchSetup();
  var msg = (result && result.ok) ? "Setup complete." : ((result && result.message) || "Setup failed.");
  SpreadsheetApp.getUi().alert("Quick Search\n\n" + msg);
}

/** Runs the Quick Search cell write/clear timing test and shows the result in an alert. */
function runQuickSearchCellTimingTest() {
  var result = testQuickSearchCellTiming();
  var msg = result.ok
    ? result.message
    : (result.message || "Error") + "\ngetColumnMs: " + (result.getColumnMs || 0);
  SpreadsheetApp.getUi().alert("Quick Search cell timing\n\n" + msg);
}





// Defines expected Amazon CSV column names
// Defines target Tiller sheet and static account values
const AMAZON_CONFIG = {
  COLUMNS: {
    ORDER_DATE: "Order Date",
    ORDER_ID: "Order ID",
    PRODUCT_NAME: "Product Name",
    TOTAL_AMOUNT: "Total Amount",
    ASIN: "ASIN"
  }
};

const TILLER_CONFIG = {
  SHEET_NAME: "Transactions",
  COLUMNS: {
    DATE: "Date",
    DESCRIPTION: "Description",
    AMOUNT: "Amount",
    TRANSACTION_ID: "Transaction ID",
    FULL_DESCRIPTION: "Full Description",
    DATE_ADDED: "Date Added",
    MONTH: "Month",
    WEEK: "Week",
    ACCOUNT: "Account",
    ACCOUNT_NUMBER: "Account #",
    INSTITUTION: "Institution",
    ACCOUNT_ID: "Account ID",
    METADATA: "Metadata"
  },
  STATIC_VALUES: {
    ACCOUNT: "Chase Amazon Visa",
    ACCOUNT_NUMBER: "xxxx8534",
    INSTITUTION: "Chase",
    ACCOUNT_ID: "636838acde7b2a0033ff46d5"
  }
};


// Generates unique Transaction IDs for Tiller
// Ensures Amazon Order IDs are not reused as IDs
function generateGuid() {
  return Utilities.getUuid();
}

// Returns Sunday start date for a given date
// Required because Tiller Week must be a real date
function getWeekStartDate(date) {
  const d = new Date(date);
  const day = d.getDay();
  d.setDate(d.getDate() - day);
  d.setHours(0,0,0,0);
  return d;
}


// Maps Tiller column headers to positions
// Allows header-based mapping instead of column position
function getTillerColumnMap(sheet){
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0];
  const map = {};
  headers.forEach((h,i)=>{
    if(h) map[h.trim()] = i + 1;
  });
  return map;
}



// Opens CSV upload dialog
// Allows optional months lookback
function importAmazonCSV_LocalUpload() {

  const html = HtmlService.createHtmlOutput(`
    <style>
      body {
        font-family: Arial, sans-serif;
        font-size: 14px;
        margin: 10px;
      }
      .field-row {
        display: flex;
        align-items: center;
        gap: 8px;
        margin-bottom: 8px;
      }
      label,
      input,
      button,
      a {
        font-family: inherit;
        font-size: inherit;
        font-weight: normal;
      }
      #log {
        font-family: Consolas, Menlo, monospace;
        font-size: 11px;
        border: 1px solid #ccc;
        padding: 6px;
        max-height: 320px;
        overflow-y: hidden;
        background: #fafafa;
        white-space: pre-wrap;
      }
      input {
        margin-top: 4px;
      }
      #months {
        width: 40px;
      }
    </style>
    <div class="field-row">
      <label for="months">Months lookback (leave blank for all):</label>
      <input type="number" id="months" value="6">
      <span style="display:inline-block; width: 8ch;"></span>
      <a id="helpLink" href="https://docs.google.com/document/d/1Mx38hFE2tKHGmD8hKKC9u4uFgOYyC_FcjwUH5Gops84/edit?usp=sharing"
         target="_blank">
        Amazon Orders Import Help
      </a>
    </div>
    <div style="margin-top:16px; margin-bottom:10px;">
      <span>Request Amazon Order History: </span>
      <a href="https://www.amazon.com/hz/privacy-central/data-requests/preview.html"
         target="_blank">
        Amazon's "Request My Data" portal
      </a>
    </div>
    <div style="margin:16px 0 12px 0;">
      <button type="button" id="fileBtn">Choose "Order History.csv" File</button>
      <span id="fileName" style="margin-left:8px;color:#555;">No file chosen</span>
      <input type="file" id="file" style="display:none">
    </div>
    <div style="margin-bottom:4px;">Import status:</div>
    <pre id="log"></pre>
    <button id="closeBtn" style="margin-top:8px; display:none;">Close</button>

    <script>
      function log(msg){
        const logEl = document.getElementById('log');
        logEl.textContent += msg + "\\n";
        logEl.scrollTop = logEl.scrollHeight;
      }

      function markComplete(){
        log("");
        log("=== Import complete ===");
        document.getElementById('closeBtn').style.display = 'inline-block';
      }

      document.getElementById('closeBtn').addEventListener('click', function(){
        google.script.host.close();
      });

      document.getElementById('fileBtn').addEventListener('click', () => {
        document.getElementById('file').click();
      });

      document.getElementById('file').addEventListener('change', e => {

        const file = e.target.files[0];
        const logEl = document.getElementById('log');
        logEl.textContent = "";

        if (!file) {
          log("No file selected.");
          return;
        }

        const start = Date.now();
        log("File selected: " + file.name + " (" + file.size + " bytes)");
        document.getElementById('fileName').textContent = file.name;

        const monthsInput = document.getElementById('months').value;
        const months = monthsInput ? parseInt(monthsInput) : null;

        const reader = new FileReader();

        reader.onload = function(evt){
          const text = evt.target.result || "";
          const charCount = text.length;
          const kb = (charCount / 1024).toFixed(1);
          log("Client: file read into memory (" + charCount + " characters, ~" + kb + " KB).");
          log("Sending data to Apps Script for processing... This step may take 10–20 seconds depending on CSV size and sheet rows.");

          const tServerStart = Date.now();
          google.script.run
            .withSuccessHandler(res => {
              const totalSec = ((Date.now() - start) / 1000).toFixed(2);
              const serverSec = ((Date.now() - tServerStart) / 1000).toFixed(2);
              log("Server finished. Estimated server time: " + serverSec + " s");
              log("Total elapsed time (client + server): " + totalSec + " s");
              const lines = String(res).split(/\\r?\\n/);
              lines.forEach(line => {
                if (line) log(line);
              });
              markComplete();
            })
            .importAmazonRecent(text, months);
        };

        log("Client: starting file read...");
        reader.readAsText(file);
      });
    </script>
  `);

  html.setWidth(600).setHeight(440);
  SpreadsheetApp.getUi().showModalDialog(html, "Import Amazon Orders");
}



// Imports Amazon transactions into Tiller
// Uses exact Full Description match for duplicate detection
function importAmazonRecent(csvText, months){

  const t0 = Date.now();
  const timing = [];

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(TILLER_CONFIG.SHEET_NAME);

  const tillerCols = getTillerColumnMap(sheet);

  const tParseStart = Date.now();
  const csv = Utilities.parseCsv(csvText);
  const tParseEnd = Date.now();
  timing.push("Server: parse CSV: " + ((tParseEnd - tParseStart) / 1000).toFixed(2) + " s");

  const headers = csv[0];
  const col = {};
  headers.forEach((h,i)=> col[h.trim()] = i);

  let cutoff = null;
  if(months){
    cutoff = new Date();
    cutoff.setMonth(cutoff.getMonth() - months);
  }

  const lastRow = sheet.getLastRow();

  // Builds a lookup Set of existing Full Descriptions
  // Used for exact duplicate detection
  const existingFullDescSet = new Set();
  const tDupStart = Date.now();
  if(lastRow > 1){

    const fullDescs = sheet.getRange(
      2,
      tillerCols[TILLER_CONFIG.COLUMNS.FULL_DESCRIPTION],
      lastRow - 1,
      1
    ).getValues();

    for(let i=0; i<fullDescs.length; i++){
      const val = fullDescs[i][0];
      if(val){
        existingFullDescSet.add(String(val));
      }
    }
  }
  const tDupEnd = Date.now();
  timing.push(
    "Server: read Full Description column + build duplicate set (" +
    existingFullDescSet.size + " entries): " +
    ((tDupEnd - tDupStart) / 1000).toFixed(2) + " s"
  );

  const numCols = sheet.getLastColumn();
  const output = [];
  let totalImported = 0;

  const tLoopStart = Date.now();
  // MAIN CSV LOOP
  for(let i = 1; i < csv.length; i++){

    const r = csv[i];
    let orderDate = new Date(r[col[AMAZON_CONFIG.COLUMNS.ORDER_DATE]]);
    orderDate.setHours(0,0,0,0);
    if(cutoff && orderDate < cutoff) continue;

    const orderID = r[col[AMAZON_CONFIG.COLUMNS.ORDER_ID]];
    const productName = r[col[AMAZON_CONFIG.COLUMNS.PRODUCT_NAME]];

    const asin = r[col[AMAZON_CONFIG.COLUMNS.ASIN]];

    const expectedFullDesc =
      "Amazon Order ID " + orderID + ": " +
      productName + " (" + asin + ")";

    if(existingFullDescSet.has(expectedFullDesc)) continue;

    let amount = parseFloat(r[col[AMAZON_CONFIG.COLUMNS.TOTAL_AMOUNT]]) * -1;
    totalImported += amount;

    const now = new Date();
    const month = Utilities.formatDate(orderDate, Session.getScriptTimeZone(), "yyyy-MM");
    const week = getWeekStartDate(orderDate);

    const descriptionText = "[AMZ] " + productName;
    const fullDesc = expectedFullDesc;

    const row = new Array(numCols).fill("");

    row[tillerCols[TILLER_CONFIG.COLUMNS.DATE]-1] = orderDate;
    row[tillerCols[TILLER_CONFIG.COLUMNS.DESCRIPTION]-1] = descriptionText;
    row[tillerCols[TILLER_CONFIG.COLUMNS.FULL_DESCRIPTION]-1] = fullDesc;
    row[tillerCols[TILLER_CONFIG.COLUMNS.AMOUNT]-1] = amount;
    row[tillerCols[TILLER_CONFIG.COLUMNS.TRANSACTION_ID]-1] = generateGuid();
    row[tillerCols[TILLER_CONFIG.COLUMNS.DATE_ADDED]-1] = now;
    row[tillerCols[TILLER_CONFIG.COLUMNS.MONTH]-1] = month;
    row[tillerCols[TILLER_CONFIG.COLUMNS.WEEK]-1] = week;
    row[tillerCols[TILLER_CONFIG.COLUMNS.ACCOUNT]-1] = TILLER_CONFIG.STATIC_VALUES.ACCOUNT;
    row[tillerCols[TILLER_CONFIG.COLUMNS.ACCOUNT_NUMBER]-1] = TILLER_CONFIG.STATIC_VALUES.ACCOUNT_NUMBER;
    row[tillerCols[TILLER_CONFIG.COLUMNS.INSTITUTION]-1] = TILLER_CONFIG.STATIC_VALUES.INSTITUTION;
    row[tillerCols[TILLER_CONFIG.COLUMNS.ACCOUNT_ID]-1] = TILLER_CONFIG.STATIC_VALUES.ACCOUNT_ID;
    row[tillerCols[TILLER_CONFIG.COLUMNS.METADATA]-1] =
      "Imported by AmazonCSVImporter on " + now;

    output.push(row);
  }
  const tLoopEnd = Date.now();
  timing.push(
    "Server: main loop over CSV (" + (csv.length - 1) + " data rows, " +
    output.length + " new rows): " +
    ((tLoopEnd - tLoopStart) / 1000).toFixed(2) + " s"
  );

  if(!output.length) return "No new transactions found";

  const tWriteNewStart = Date.now();
  sheet.getRange(sheet.getLastRow()+1,1,output.length,numCols).setValues(output);
  const tWriteNewEnd = Date.now();
  timing.push(
    "Server: write new rows to sheet (" + output.length + " rows): " +
    ((tWriteNewEnd - tWriteNewStart) / 1000).toFixed(2) + " s"
  );

  // Adds balancing offset entry for imported Amazon transactions
  // Includes timestamp and uses GUID as Transaction ID
  if(totalImported !== 0){

    const now = new Date();
    const month = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM");
    const week = getWeekStartDate(now);

    const offset = new Array(numCols).fill("");

    const desc = "Amazon purchase offset for " +
      Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

    let offsetDate = new Date(now);
    offsetDate.setHours(0,0,0,0);

    offset[tillerCols[TILLER_CONFIG.COLUMNS.DATE]-1] = offsetDate;
    offset[tillerCols[TILLER_CONFIG.COLUMNS.DESCRIPTION]-1] = desc;
    offset[tillerCols[TILLER_CONFIG.COLUMNS.FULL_DESCRIPTION]-1] = desc;
    offset[tillerCols[TILLER_CONFIG.COLUMNS.AMOUNT]-1] = Math.abs(totalImported);
    offset[tillerCols[TILLER_CONFIG.COLUMNS.TRANSACTION_ID]-1] = generateGuid();
    offset[tillerCols[TILLER_CONFIG.COLUMNS.DATE_ADDED]-1] = now;
    offset[tillerCols[TILLER_CONFIG.COLUMNS.MONTH]-1] = month;
    offset[tillerCols[TILLER_CONFIG.COLUMNS.WEEK]-1] = week;
    offset[tillerCols[TILLER_CONFIG.COLUMNS.ACCOUNT]-1] = TILLER_CONFIG.STATIC_VALUES.ACCOUNT;
    offset[tillerCols[TILLER_CONFIG.COLUMNS.ACCOUNT_NUMBER]-1] = TILLER_CONFIG.STATIC_VALUES.ACCOUNT_NUMBER;
    offset[tillerCols[TILLER_CONFIG.COLUMNS.INSTITUTION]-1] = TILLER_CONFIG.STATIC_VALUES.INSTITUTION;
    offset[tillerCols[TILLER_CONFIG.COLUMNS.ACCOUNT_ID]-1] = TILLER_CONFIG.STATIC_VALUES.ACCOUNT_ID;
    offset[tillerCols[TILLER_CONFIG.COLUMNS.METADATA]-1] =
      "Imported by AmazonCSVImporter on " + now;

    const tWriteOffsetStart = Date.now();
    sheet.getRange(sheet.getLastRow()+1,1,1,numCols).setValues([offset]);
    const tWriteOffsetEnd = Date.now();
    timing.push(
      "Server: write offset row to sheet: " +
      ((tWriteOffsetEnd - tWriteOffsetStart) / 1000).toFixed(2) + " s"
    );
  }

  const tSortStart = Date.now();
  sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn())
    .sort({column: tillerCols[TILLER_CONFIG.COLUMNS.DATE], ascending:false});
  const tSortEnd = Date.now();
  timing.push(
    "Server: sort sheet by Date: " +
    ((tSortEnd - tSortStart) / 1000).toFixed(2) + " s"
  );

  const tEnd = Date.now();
  timing.push(
    "Server: TOTAL importAmazonRecent time: " +
    ((tEnd - t0) / 1000).toFixed(2) + " s"
  );

  const summary = output.length + " transactions imported";
  timing.unshift(summary);

  // Join with real newline characters so the client can
  // display each entry on its own line.
  return timing.join("\n");
}


function cursorTest() {
  Logger.log("Cursor is connected");
}