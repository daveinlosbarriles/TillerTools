// Adds Amazon Tools menu to Google Sheets
// Allows launching the CSV import dialog
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Amazon Tools")
    .addItem("Import Amazon CSV", "importAmazonCSV_LocalUpload")
    .addToUi();
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
  SHEET_NAME: "Transactions Test",
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
    <label>Months lookback (leave blank for all):</label><br>
    <input type="number" id="months"><br><br>
    <input type="file" id="file"><br><br>
    <pre id="log"></pre>

    <script>
      function log(msg){
        document.getElementById('log').textContent += msg + "\\n";
      }

      document.getElementById('file').addEventListener('change', e => {

        const monthsInput = document.getElementById('months').value;
        const months = monthsInput ? parseInt(monthsInput) : null;

        const reader = new FileReader();

        reader.onload = function(evt){
          log("Processing...");
          google.script.run
            .withSuccessHandler(res => log(res))
            .importAmazonRecent(evt.target.result, months);
        };

        reader.readAsText(e.target.files[0]);
      });
    </script>
  `);

  SpreadsheetApp.getUi().showModalDialog(html, "Import Amazon CSV");
}



// Imports Amazon transactions into Tiller
// Uses exact Full Description match for duplicate detection
function importAmazonRecent(csvText, months){

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(TILLER_CONFIG.SHEET_NAME);

  const tillerCols = getTillerColumnMap(sheet);

  const csv = Utilities.parseCsv(csvText);
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

  if(lastRow > 1){

    const fullDescs = sheet.getRange(
      2,
      tillerCols[TILLER_CONFIG.COLUMNS.FULL_DESCRIPTION],
      lastRow - 1
    ).getValues();

    for(let i=0; i<fullDescs.length; i++){
      const val = fullDescs[i][0];
      if(val){
        existingFullDescSet.add(String(val));
      }
    }
  }

  const numCols = sheet.getLastColumn();
  const output = [];
  let totalImported = 0;

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

  if(!output.length) return "No new transactions found";

  sheet.getRange(sheet.getLastRow()+1,1,output.length,numCols).setValues(output);

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

    sheet.getRange(sheet.getLastRow()+1,1,1,numCols).setValues([offset]);
  }

  sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn())
    .sort({column: tillerCols[TILLER_CONFIG.COLUMNS.DATE], ascending:false});
  

  return output.length + " transactions imported";
}


function cursorTest() {
  Logger.log("Cursor is connected");
}