/*
  * Google Apps Script to analyze a Google Drive folder and display the results in a Google Sheet.
  * The script collects attributes of image files and logs them in a specified sheet.
  * It also provides a custom menu to trigger the analysis.
*/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Drive Checker')
    .addItem('Analyze Model Answer Folder', 'showAnalyzeDialog')
    .addItem('Analyze Student Folders', 'showStudentFoldersDialog')
    .addItem('Analyze Student Folders in Drive', 'showStudentTaskDialog')
    .addToUi();
}

/**
 * Shows the Student Task Submission Form modal.
 */
function showStudentTaskDialog() {
  var html = HtmlService.createHtmlOutputFromFile('StudentTaskSubmissionForm')
    .setWidth(1024)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Student Task Submission');
}
/**
 * Shows the modal dialog for student folders analysis.
 */
function showStudentFoldersDialog() {
  var html = HtmlService.createHtmlOutputFromFile('StudentFoldersDialog')
    .setWidth(1024)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Analyze Student Folders');
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
