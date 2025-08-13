/*
  * Google Apps Script to analyze a Google Drive folder and display the results in a Google Sheet.
  * The script collects attributes of image files and logs them in a specified sheet.
  * It also provides a custom menu to trigger the analysis.
*/

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
    .setWidth(1024)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Analyze Google Drive Folder');
}
