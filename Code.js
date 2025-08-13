function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu("Drive Image Validator")
    .addItem("Scan Answer Key", "scanAnswerKey")
    .addItem("Compare Student Folders", "compareStudentFolders")
    .addToUi();
}
