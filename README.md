# Tiller Tools

One GitHub repo and Apps Script project for **Tiller** (https://tiller.com/) spreadsheets:

- **Tiller Amazon Import** — sidebar wizard: upload Amazon’s **orders data ZIP**, map payment types on the **AMZ Import** tab, and append purchase/return rows to your **Transactions** sheet with balancing **offsets** where needed.
- **Tiller Quick Search** — sidebar to filter **Transactions** by date, amount, description, account, and category (basic filter + helper columns).

Repository: [github.com/daveinlosbarriles/TillerTools](https://github.com/daveinlosbarriles/TillerTools).

This software is produced by a Tiller user and is not affiliated with Tiller LLC.

## Privacy, terms, and contact

| Document | Link |
|----------|------|
| Privacy Policy | [PRIVACY.md](PRIVACY.md) · [on GitHub](https://github.com/daveinlosbarriles/TillerTools/blob/master/PRIVACY.md) |
| Terms of Service | [TERMS.md](TERMS.md) · [on GitHub](https://github.com/daveinlosbarriles/TillerTools/blob/master/TERMS.md) |

Questions: **tillertoolsbydave@gmail.com**

---

## Install (supported path: bound script, copy/paste)

**This is the documented install:** bind the project to your Tiller spreadsheet and paste the files in the Apps Script editor. A future **Google Workspace Marketplace** add-on is *not* part of these instructions; the repo still contains `appsscript.json` add-on metadata for developers who use **clasp**.

### Step 1: Open Apps Script and add files

1. Open your **Tiller Google Sheet**.
2. Go to **Extensions** → **Apps Script**. A new tab opens with the Apps Script editor.
3. Remove or replace the default `Code.gs` sample if present. Add the files below (names matter for HTML includes).

   **Creating files**

   - **Script (`.gs`):** click **+** next to *Files* → **Script** → type the name (no extension in the dialog) → paste repo content → save.
   - **HTML:** **+** → **HTML** → type the name **without** `.html` in the dialog → paste → save.

   **Files to add**

   | Paste from repo | Apps Script name | Notes |
   |-----------------|------------------|--------|
   | [`Code.js`](Code.js) | `Code` (shows as **Code.gs**) | Menu **Tiller Tools** → Amazon Import + Quick Search; includes add-on homepage hook for future listing |
   | [`QuickSearchSidebar.js`](QuickSearchSidebar.js) | `QuickSearchSidebar` | Saves as **QuickSearchSidebar.gs** |
   | [`QuickSearch.html`](QuickSearch.html) | `QuickSearch` | Sidebar UI |
   | [`amazonorders.gs`](amazonorders.gs) | `amazonorders` | All Amazon CSV pipelines |
   | [`AmazonOrdersSidebar.html`](AmazonOrdersSidebar.html) | `AmazonOrdersSidebar` | Import sidebar (matches `HtmlService.createHtmlOutputFromFile("AmazonOrdersSidebar")`) |
   | [`AmazonOrdersDialog.html`](AmazonOrdersDialog.html) | `AmazonOrdersDialog` | Optional / legacy; menu uses the sidebar |

   **Quick Search only:** omit the three Amazon files and remove the **Tiller Amazon Import** line from `Code.gs` (see comments in `Code.js`).

4. **Save** the project (Ctrl+S) and give it a name (e.g. “Tiller Tools”) if prompted.

### Step 2: Enable Google Sheets API (Quick Search)

Quick Search uses the **Sheets API**.

1. In the Apps Script editor, open **Services** (or **Advanced Google Services** in older UIs).
2. Turn **Google Sheets API** **On**.
3. If prompted, open **Google Cloud Console** for the same project and ensure **Google Sheets API** is enabled.

### Step 3: Authorize and open the tools

1. Reload the **spreadsheet** tab (F5).
2. Use the **Tiller Tools** menu on the menu bar (**Extensions** may also list the same items depending on your workspace).
3. Choose **Tiller Amazon Import** or **Tiller Quick Search**. The first run opens Google **authorization** (Advanced → continue to your project → Allow). The script only accesses the current spreadsheet.

### For developers (optional): clasp

From a clone of this repo, `clasp login`, link or clone the Apps Script project, then `clasp push` to sync `.gs` / `.html` / `appsscript.json`. See [`.claspignore`](.claspignore) (e.g. ignores a legacy local `TillerAmazonOrdersCSVImport/` folder if present).

**OAuth scopes** declared in [`appsscript.json`](appsscript.json) include Spreadsheets (current only) and `script.container.ui`. Match these on the Cloud **OAuth consent screen** if you publish an add-on later.

---

## What you need in the spreadsheet

- **Transactions** — Tiller column headers in row 1 (Date, Description, Amount, Account, etc.). Amazon import reads target column names from the **AMZ Import** sheet (Table 4).
- **Categories** — Quick Search builds the category list from column A (category) and B (group) from row 2 down.
- **Accounts** — Quick Search reads account names from column **J**, row 2 down; rows with **Hide** in column **Q** are skipped.
- **AMZ Import** — created automatically the first time you run Amazon import, with default payment tables and Tiller label mapping. Edit payment **Type → Account** mappings here before importing if the wizard flags unknown methods.

---

## Tiller Amazon Import — how it works

Detailed behavior and **AMZ Import** tables below are aligned with the original write-up in **[TillerAmazonOrdersCSVImport](https://github.com/daveinlosbarriles/TillerAmazonOrdersCSVImport/blob/main/README.md)**. That repo described a **single-CSV dialog** flow (`AmazonOrders.gs`, **Amazon Orders Import**). **This** repo merges the same ideas with **Quick Search**, a **ZIP** sidebar wizard (`amazonorders.gs`, **Tiller Amazon Import**), **chunked** server imports, and extra pipelines (**Orders returns**, **Digital returns**). For copy-from-GitHub-**Raw** tutorials and older **screenshots**, the legacy README is still the right reference.

### What this importer does

- Converts Amazon **ZIP** export data into Tiller **Transactions** rows (orders, digital items, physical refunds, digital returns — according to what you select and what files Amazon included).
- Prevents **duplicate** imports using stable keys in each row’s **Metadata** JSON (so dedup still works if you edit **Full Description** text).
- Prefixes lines **[AMZ]** (physical / Order History–style) or **[AMZD]** (digital) on descriptions where applicable; **Account** fields come from **AMZ Import** Table 1.
- Writes **balancing offset** rows so you are not double-counting the same spend against both a card charge and Amazon line items (see wizard **Category for offsets** help).
- After a full ZIP run, **sorts** Transactions and applies a **Metadata** filter for this import’s timestamp so you can focus on rows from that run.

### Get your data from Amazon

Use Amazon’s **privacy / data request** flow to download a ZIP of your orders (the sidebar links to [Amazon’s Request My Data portal](https://www.amazon.com/hz/privacy-central/data-requests/preview.html); choose **Your Orders** / submit as Amazon directs). When the email arrives, download the ZIP and use **Choose orders ZIP file** in the wizard.

### How to use it (this repo)

1. Open your Tiller spreadsheet and choose **Tiller Tools** → **Tiller Amazon Import**.
2. Pick the **ZIP**, enable the order types you want, set **Ignore orders older than** if you need a cutoff (similar intent to the legacy “months lookback”; use a date that overlaps a bit with your last import).
3. Complete **payment method** review if the wizard shows it; fix **AMZ Import** Table 1 if anything is **Unknown**.
4. Click **Import** and leave the sidebar open until the log shows completion.

**Tips:** Copy the spreadsheet and try there first. To undo later, filter **Metadata** with **contains** `Imported by AmazonCSVImporter` and delete those rows.

### What the ZIP can contain

The importer looks for these names (case-insensitive paths inside the ZIP):

| File | Used for |
|------|-----------|
| **Order History.csv** | Physical / Whole Foods / Amazon Fresh **orders** (line items), when **Orders** and/or **Whole Foods / Amazon Fresh** is checked |
| **Refund Details.csv** or **Returns.csv** | **Orders returns** (refund amounts by Order ID) |
| **Digital Content Orders.csv** | Digital purchases |
| **Digital Returns.csv** | Digital returns |

Not every export includes every file; the wizard lists which ones it found.

### Order History vs Digital Content Orders (CSV detection)

For **Order History.csv** and **Digital Content Orders.csv**, file type is detected from the **header row** (same rules as the [legacy README](https://github.com/daveinlosbarriles/TillerAmazonOrdersCSVImport/blob/main/README.md#4-amazon-csv-columns-used)):

- If **`Digital Order Item ID`** is present → **Digital Content Orders** (wins if both markers appear).
- Else if **`Carrier Name & Tracking Number`** is present → **Order History** (standard).
- If **neither** is present → error.

**Which CSV header maps to which field** is configured on **AMZ Import** Tables 2 and 3. Every non-empty mapped column for that file type must exist in the CSV. For digital imports, **exactly one** Table 1 row must have **Use for Digital orders?** = **Yes**. If Amazon renames columns, update the sheet — no code change.

### Tiller columns written (typical)

| Column | Content |
|--------|--------|
| **Date** | Transaction date for that pipeline |
| **Description** | `[AMZ]` / `[AMZD]` (and similar) + short text / product title per pipeline |
| **Full Description** | Longer line (often Order ID, title, ASIN where applicable) |
| **Amount** | Line amount |
| **Transaction ID** | Unique id |
| **Date Added** | Import run time |
| **Account** | From **AMZ Import** by payment type or digital row |
| **Account #** / **Institution** / **Account ID** | Same routing |
| **Metadata** | `Imported by AmazonCSVImporter on <date/time>` plus JSON `{ "amazon": … }` — filter **contains** `Imported by AmazonCSVImporter` to find or remove imports |

Refund / return rows use pipeline-specific description patterns; metadata always carries a stable `type` for dedup.

### AMZ Import sheet (key settings)

The **AMZ Import** tab holds all routing. **First run** creates it with defaults.

- **Table 1 – Payment method → Tiller account**  
  **Payment Type** | **Account** | **Account #** | **Institution** | **Account ID** | **Use for Digital orders?**  
  Each Order History payment string (e.g. from **Payment Method Type**) gets a row. **Exactly one** row must be **Yes** on **Use for Digital orders?**; that row’s account fields are used for **Digital Content Orders**. Offsets are created per payment type / pipeline rules so accounts net correctly.

- **Table 2 – Core field mapping**  
  **Amazon CSV column name** | **Digital Orders CSV column name** | **Name in Code** — maps logical fields to real headers. Older two-column sheets may still work.

- **Table 3 – Metadata JSON mapping**  
  **Amazon CSV column name** | **Digital Orders CSV column name** | **Metadata field name** — builds the `amazon` object. Older two-column layout supported.

- **Table 4 – Tiller column labels**  
  **Name in Code** → **Tiller label** for your **Transactions** header row (Date, Metadata, etc.).

If a payment type is missing, add a Table 1 row with the **exact** Amazon string.

### Wizard steps (summary)

1. **ZIP** — Select the file; **Next** reads the archive in the browser (JSZip).
2. **Options** — **Orders**, **Orders Returns**, **Digital orders**, **Digital returns**, **Whole Foods / Amazon Fresh** (website filter on Order History).
3. **Cutoff** — **Ignore orders older than** (rows on that calendar day are included).
4. **Category for offsets** — Optional Tiller category for offset rows.
5. **Check for new payment methods** — Pause to align unknown types on **AMZ Import** before **Import**.
6. **Import** — Chunks run on the server (one round trip per main CSV + **finalize** sort/filter). **Do not close the sidebar** until finished.

### After import

- New rows on **Transactions**; status may report *“N duplicate transactions were not imported.”* when Metadata keys already existed.
- **Show debug messages** — full server timings and diagnostics; off by default for a shorter log.

### Behavior notes

- **Orders returns (Refund Details):** Date from **Refund Date** when parseable; otherwise **Creation Date** when present. **$0** refund totals are skipped with no extra messages. CSV needs **Order ID**, **Refund Amount**, **Website**, and at least one of **Refund Date** or **Creation Date**.
- **Offsets:** Purchase/refund pipelines add offset rows per in-code rules; use **AMZ Import** + wizard hints so accounts match your Tiller **Accounts** tab.

### Processing outline

1. Read **AMZ Import** config (Tables 1–4).
2. For each selected pipeline, parse the relevant CSV (from ZIP text), detect standard vs digital where applicable, validate mapped headers.
3. Resolve payment / digital account rows; build rows and Metadata; skip duplicates via existing **Metadata** keys.
4. Append rows; add offsets as needed.
5. **Finalize:** sort by date (newest first) and set Metadata filter for this bundle’s import timestamp.

---

## Tiller Quick Search — key features

**Search by:**

- **Date** – range or exact From / To
- **Amount** – min and max
- **Description** – “contains” text, with options like:
  - `Amazon | Walmart` – rows containing Amazon or Walmart
  - `^Vanguard` – name starts with Vanguard
  - `1234$` – name ends with 1234
  - `^Gas$` – exactly “Gas” (no “Gasoline”)
  - `Amazon.*Return|Return.*Amazon` – contains both Amazon and Return (any order)
  - `gas but not chevron` – contains “gas” but not “chevron”
- **Account** – choose one or more
- **Category** – one or more categories, sorted by group
- See all of your search criteria at once in the sidebar

<img width="3840" height="2063" alt="Tiller Tools screenshot" src="https://github.com/user-attachments/assets/80d36566-54d4-40f0-b86d-35040a49a8e1" />

### First-time use: two new columns

The first time you search, the script may add **QuickSearch** and **QuickCriteria** on the right of **Transactions**. You can delete them; they are recreated on the next search if missing.

### How the filter works

Quick Search applies a **basic filter** so only rows where **QuickSearch** is TRUE are shown. You can combine with other filter dropdowns as usual.

---

## Usage tips

- **Quick Search:** set criteria, then **Search**.
- **Amazon import:** complete payment review if prompted, then **Import**. Use **Start over** to pick a new ZIP.
- **Description** regex and **` but not `** syntax: see examples above and the sidebar help on **Description**.

---

## Repo contents (reference)

| File in repo | In Apps Script | Purpose |
|--------------|----------------|---------|
| `Code.js` | `Code.gs` | Menu, `onInstall` / `onOpen`, optional add-on homepage |
| `QuickSearchSidebar.js` | `QuickSearchSidebar.gs` | Quick Search logic |
| `QuickSearch.html` | `QuickSearch.html` | Quick Search sidebar UI |
| `amazonorders.gs` | `amazonorders.gs` | Amazon CSV pipelines, chunked bundle + finalize |
| `AmazonOrdersSidebar.html` | `AmazonOrdersSidebar.html` | Amazon ZIP wizard UI |
| `AmazonOrdersDialog.html` | `AmazonOrdersDialog.html` | Legacy dialog HTML |
| `appsscript.json` | project settings | Time zone, scopes, optional `addOns` block |
| `assets/tiller-tools-logo.png` | — | Icon URL referenced for a future listing |

**Quick Search** and **Amazon import** do not call each other; they share only the menu in `Code.js`.

**Source of truth:** [github.com/daveinlosbarriles/TillerTools](https://github.com/daveinlosbarriles/TillerTools).

**Earlier standalone importer (archived pattern):** [TillerAmazonOrdersCSVImport](https://github.com/daveinlosbarriles/TillerAmazonOrdersCSVImport) — same **AMZ Import** concepts; different file names (`AmazonOrders.gs`) and **dialog** UI.

---

## To uninstall

1. **Extensions** → **Apps Script** — remove or edit **Code.gs** so the menu is gone; delete **QuickSearchSidebar.gs**, **QuickSearch.html**, `amazonorders.gs`, **AmazonOrdersSidebar.html**, **AmazonOrdersDialog.html** as needed.
2. Optionally delete **QuickSearch** / **QuickCriteria** columns on **Transactions**.
3. Optionally **Data** → **Turn off filter** on **Transactions**.
