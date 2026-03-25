# Tiller Tools

One GitHub repo and Apps Script project for **Tiller** (https://tiller.com/) spreadsheets:

- **Tiller Amazon Import** — sidebar wizard: upload Amazon’s **orders data ZIP**, and insert your Amazon orders, returns, digital orders, and digital returns in to Tiller.
- **Tiller Quick Search** — sidebar to filter **Transactions** by date, amount, description, account, and category (basic filter + helper columns).  See your search criteria for all key fields at all times, use date ranges, boolean operators in the description, and easily modify. 

This software is produced by a Tiller user and is not affiliated with Tiller LLC.

## Tiller Amazon Import Process

1. Request your orders.zip data file from Amazon https://www.amazon.com/hz/privacy-central/data-requests/preview.html
2. Click **Tiller Tools** / **Tiller Amazon Import**.  The Amazon Import sidebar appears.
3. **Select your zip file**, and click **Next**
4. The contents of the zip file are examined - these types of transactions can be imported:
   - Orders (including Whole Foods groceries)
   - Digital Orders
   - Order returns
   - Digital returns
6. **Uncheck** any transaction types you don't want to import
7. Ignore orders older than a certain date - default is 6 months, which is usually fine.  Duplicates are never created.
8. **Set a category for offset transactions** - normally Amazon or Transfer
9. If this is your first import, or if you have new credit cards in Amazon, **Check for new payment methods**
10. Click the **Import** button
11. Wait 30-50 seconds
12. New transactions appear on the Transactions tab, pre-filtered to show just those.

## Tiller Quick Search Process
This is as simple as it sounds.  
1. Click **Tiller Tools** / **Tiller Amazon Import**.  The QuickSearch sidebar appears.
2. Enter in your criteria, and hit **enter** key or **Search** button at any time to see filtered results.
3. Click on the Description link in the sidebar to see examples of more complex searches such as Amazon | Walmart NOT Costco

---

## Install

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

4. **Save** the project (Ctrl+S) and give it a name (e.g. “Tiller Tools”) if prompted.

### Step 1b: Project manifest (`appsscript.json`)

The manifest is **not** added with **+ → Script / HTML**. It controls **time zone**, **V8 runtime**, **OAuth scopes**, the **Sheets advanced service** (Quick Search), and optional **add-on** metadata.

1. In the Apps Script editor, open **Project Settings** (gear icon).
2. Under **General settings**, enable **Show “appsscript.json” manifest file in editor** (wording may vary slightly by UI).
3. Back in the editor file list, open **`appsscript.json`**.
4. Replace the entire file with the contents of the repo’s [`appsscript.json`](appsscript.json) (copy/paste the whole JSON).
5. **Save.**

For **menu-only / personal** use you can delete the optional `"addOns": { ... }` block afterward to keep the manifest smaller; `Code.js` and the sidebars do not require it. Do **not** remove `oauthScopes`, `runtimeVersion`, or the `dependencies.enabledAdvancedServices` entry for **Sheets** if you use Quick Search.

### Step 2: Enable Google Sheets API (Quick Search)

Quick Search uses the **Sheets API**. Pasting the repo `appsscript.json` already declares the **Sheets** advanced service; still confirm it is active.

1. In the Apps Script editor, open **Services** (or **Advanced Google Services** in older UIs).
2. Turn **Google Sheets API** **On**.
3. If prompted, open **Google Cloud Console** for the same project and ensure **Google Sheets API** is enabled.

### Step 3: Authorize and open the tools

1. Reload the **spreadsheet** tab (F5).
2. Use the **Tiller Tools** menu on the menu bar (**Extensions** may also list the same items depending on your workspace).
3. Choose **Tiller Amazon Import** or **Tiller Quick Search**. The first run opens Google **authorization** (Advanced → continue to your project → Allow). The script only accesses the current spreadsheet.

---

## What you need in Tiller

- **Transactions tab** — Tiller column headers in row 1 (Date, Description, Amount, Account, etc.). Amazon import reads target column names from the **AMZ Import** sheet (**Tiller column labels** section).
- **Categories tab** — Both Amazon Import and Quick Search builds the category list from column A (category) and B (group) from row 2 down.
- **Accounts tab** — Quick Search reads account names from column **J**, row 2 down; rows with **Hide** in column **Q** are skipped.
- **AMZ Import** — This tab is created automatically the first time you run Amazon import, with default payment tables and Tiller label mapping. Edit payment **Amazon Payment Type → Account** mappings here before importing if the wizard flags unknown credit cards.

---

### What this importer does

- Converts Amazon **ZIP** export data into Tiller **Transactions** rows (orders, digital items, physical refunds, digital returns — according to what you select and what files Amazon included).
- Prevents **duplicate** imports using stable keys in each row’s **Metadata** JSON (so dedup still works if you edit **Full Description** text).
- Backwardly comptible with the older Amazon order import
- Prefixes lines **[AMZ]** (physical / Order History–style) or **[AMZD]** (digital) on descriptions where applicable; **Account** fields come from **AMZ Import** (payment method → account table).
- Writes **balancing offset** rows so you are not double-counting the same spend against both a card charge and Amazon line items (see wizard **Category for offsets** help).
- After the import completes, **sorts** Transactions and applies a filter on the **Metadata** column for this import’s timestamp so you can focus on rows from that run.

### Amazon Zip File Details

The importer looks for these names (case-insensitive paths inside the ZIP):

| File | Used for |
|------|-----------|
| **Order History.csv** | Physical / Whole Foods / Amazon Fresh **orders** (line items), when **Orders** and/or **Whole Foods / Amazon Fresh** is checked |
| **Refund Details.csv** or **Returns.csv** | **Orders returns** (refund amounts by Order ID) |
| **Digital Content Orders.csv** | Digital purchases |
| **Digital Returns.csv** | Digital returns |

### Order History vs Digital Content Orders (CSV detection)

For **Order History.csv** and **Digital Content Orders.csv**, file type is detected from the **header row** using marker columns (defaults: **`Digital Order Item ID`** for digital, **`Carrier Name & Tracking Number`** for standard). Digital wins if both appear. On the **AMZ Import** tab, the first rows under **Source file** `_file_detection` let you change those marker header names if Amazon renames them.

**Which CSV header maps to which logical field and metadata key** is configured in one **CSV column map** table: **Source file** | **Header** | **Name in code** | **Metadata field name**. Refund Details, Returns, and Digital Returns rows in that same table drive the returns pipelines. **AMZ Import** must include that table (first run seeds it). Every non-empty mapped column for that file type must exist in the CSV. For digital imports, **exactly one** payment row must have **Use for Digital orders?** = **Yes**. If Amazon renames columns, update the sheet — no code change.

### Tiller columns written (typical)

| Column | Content |
|--------|--------|
| **Date** | Transaction date for that pipeline |
| **Description** | `[AMZ]` / `[AMZD]` (and similar) + short text / product title per pipeline |
| **Full Description** | Longer line (often Order ID, title, ASIN where applicable) |
| **Amount** | Line amount |
| **Transaction ID** | New unique id |
| **Date Added** | Import run time |
| **Account** | From **AMZ Import** by payment type or digital row |
| **Account #** / **Institution** / **Account ID** | Same routing |
| **Metadata** | `Imported by AmazonCSVImporter on <date/time>` plus JSON `{ "amazon": … }` — filter **contains** `Imported by AmazonCSVImporter` to find or remove imports |

Refund / return rows use pipeline-specific description patterns; metadata always carries a stable `type` for dedup.

### AMZ Import sheet (key settings)

The **AMZ Import** tab holds all routing. **First run** creates it with defaults.

- **Payment method → Tiller account**  
  **Payment Type** | **Account** | **Account #** | **Institution** | **Account ID** | **Use for Digital orders?**  
  Each Order History payment string (e.g. from **Payment Method Type**) gets a row. **Exactly one** row must be **Yes** on **Use for Digital orders?**; that row’s account fields are used for **Digital Content Orders**. Offsets are created per payment type / pipeline rules so accounts net correctly.

- **CSV column map (single table)**  
  **Source file** | **Header** | **Name in code** | **Metadata field name** — for each Amazon export file, maps the real CSV header to the importer’s logical field name and (when filled) the `amazon` metadata key. Optional `_file_detection` rows set standard vs digital marker column names.

- **Tiller column labels**  
  **Name in Code** | **Tiller label** — must match your **Transactions** header row (Date, Metadata, etc.).

If a payment type is missing, add a payment row with the **exact** Amazon string.

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

The first time you search, the script will add **QuickSearch** and **QuickCriteria** columns on the right side of the **Transactions** tab. You can delete them; they are recreated on the next search if missing.

### How the filter works

Quick Search applies a **basic filter** so only rows where **QuickSearch** is TRUE are shown. You can combine with other filter dropdowns as usual.

---

## Repo contents (reference)

| File in repo | In Apps Script | Purpose |
|--------------|----------------|---------|
| `Code.js` | `Code.gs` | Menu, `onInstall` / `onOpen`, optional add-on homepage |
| `QuickSearchSidebar.js` | `QuickSearchSidebar.gs` | Quick Search logic |
| `QuickSearch.html` | `QuickSearch.html` | Quick Search sidebar UI |
| `amazonorders.gs` | `amazonorders.gs` | Amazon CSV pipelines, chunked bundle + finalize |
| `AmazonOrdersSidebar.html` | `AmazonOrdersSidebar.html` | Amazon ZIP wizard UI |
| `appsscript.json` | project settings | Time zone, scopes, optional `addOns` block |
| `assets/tiller-tools-logo.png` | — | Icon URL referenced for a future listing |
| [docs/screenshots/](docs/screenshots/) | — | Images for this README (not pushed via clasp) |

**Quick Search** and **Amazon import** do not call each other; they share only the menu in `Code.js`.

---

## To uninstall

1. **Extensions** → **Apps Script** — remove or edit **Code.gs** so the menu is gone; delete **QuickSearchSidebar.gs**, **QuickSearch.html**, **amazonorders.gs**, **AmazonOrdersSidebar.html** .
2. Delete **QuickSearch** / **QuickCriteria** columns on **Transactions**.
3. Delete **AMZ Import** tab 

## Privacy, terms, and contact

| Document | Link |
|----------|------|
| Privacy Policy | [PRIVACY.md](PRIVACY.md) |
| Terms of Service | [TERMS.md](TERMS.md) |

Questions: **tillertoolsbydave@gmail.com**
