// Tiller Tools menu. Quick Search and Amazon are independent codebases.
// Use createMenu (not createAddonMenu): createAddonMenu only shows entries when the script runs as an
// installed Editor add-on—easy to end up with no menu during clasp/standalone dev. createMenu works for
// bound spreadsheets and for installed add-ons; for published add-ons Google moves this under Extensions.
// Standalone add-on: still install via Deploy → Test deployments → Google Workspace add-on, then reopen Sheets.
// For Quick Search only: remove the Amazon menu item and omit amazonorders.gs + Amazon HTML files.

/** @param {GoogleAppsScript.Events.SheetsOnOpen | GoogleAppsScript.Events.SheetsOnInstall} [e] */
function onInstall(e) {
  onOpen(e);
}

/** @param {GoogleAppsScript.Events.SheetsOnOpen | GoogleAppsScript.Events.SheetsOnInstall} [e] */
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu("Tiller Tools")
    .addItem("Tiller Amazon Import", "openAmazonOrdersSidebar")
    .addItem("Tiller Quick Search", "openQuickSearchSidebar")
    .addToUi();
}

/**
 * Google Workspace add-on homepage (side panel). HtmlService workflows live under Extensions → Tiller Tools.
 * Replace logoUrl in appsscript.json with your own icon before Marketplace submission.
 * @param {Object} e - Add-on event object
 */
function tillerToolsOnHomepage(e) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("Tiller Tools"))
    .addSection(
      CardService.newCardSection()
        .addWidget(
          CardService.newTextParagraph().setText(
            "Open Tiller Tools → Tiller Amazon Import or Tiller Quick Search. (When installed as an add-on, the same items may appear under Extensions → Tiller Tools.)"
          )
        )
    )
    .build();
}
