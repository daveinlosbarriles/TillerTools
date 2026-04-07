// Tiller Tools menu. Quick Search and Amazon are independent codebases.
// onOpen uses createAddonMenu() — standard for Marketplace / Editor add-ons (entries under Extensions).
// For bound-script dev without add-on install, use createMenu() temporarily if the add-on menu is empty.
// Standalone add-on: install via Deploy → Test deployments → Google Workspace add-on, then reopen Sheets.
// For Quick Search only: remove the Amazon menu item and omit amazonorders.gs + Amazon HTML files.

/** @param {GoogleAppsScript.Events.SheetsOnOpen | GoogleAppsScript.Events.SheetsOnInstall} [e] */
function onInstall(e) {
  onOpen(e);
}

/** @param {GoogleAppsScript.Events.SheetsOnOpen | GoogleAppsScript.Events.SheetsOnInstall} [e] */
function onOpen(e) {
  // createAddonMenu() is the standard for Marketplace submission
  SpreadsheetApp.getUi()
    .createAddonMenu() 
    .addItem("Tiller™ Amazon Import", "openAmazonOrdersSidebar")
    .addItem("Tiller™ Quick Search", "openQuickSearchSidebar")
    .addToUi();
}

/**
 * Google Workspace add-on homepage (side panel). HtmlService workflows live under Extensions → Tiller Tools.
 * Replace logoUrl in appsscript.json with your own icon before Marketplace submission.
 * @param {Object} e - Add-on event object
 */
function tillerToolsOnHomepage(e) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("Tiller™ Tools"))
    .addSection(
      CardService.newCardSection()
        .addWidget(
          CardService.newTextParagraph().setText(
            "Open Tiller™ Tools → Tiller™ Amazon Import or Tiller™ Quick Search. (When installed as an add-on, the same items may appear under Extensions → Tiller™ Tools.)"
          )
        )
    )
    .build();
}
