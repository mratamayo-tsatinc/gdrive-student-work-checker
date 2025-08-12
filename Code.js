function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Functions')
    .addItem('Fill SNO and Enrollment Status', 'fillSNOAndEnrollmentStatus')
    .addToUi();

  ui.createMenu("Drive Image Validator")
    .addItem("Scan Answer Key", "scanAnswerKey")
    .addItem("Compare Student Folders", "compareStudentFolders")
    .addToUi();
}
