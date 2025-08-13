
/**
 * Adds a custom menu to the Google Sheet when the document is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Drive Checker')
    .addItem('Analyze Drive Folder', 'showAnalyzeDialog')
    .addToUi();
}

/**
 * Opens the modal dialog for folder analysis.
 */
function showAnalyzeDialog() {
  var html = HtmlService.createHtmlOutputFromFile('AnalyzeDialog')
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Analyze Google Drive Folder');
}
