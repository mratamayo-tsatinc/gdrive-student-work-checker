/**
 * Evaluates a specific student folder by comparing its files to the model answer.
 * Writes results to _STUDENT_RESULTS and updates the score in _SCORE.
 * @param {string} folderId The ID of the student Google Drive folder.
 * @return {string} Status message.
 */
function evaluateStudentFolder(folderId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Get config: total score
    var configSheet = ss.getSheetByName('_CONFIG');
    var totalScore = 100;
    if (configSheet && configSheet.getRange('B1').getValue()) {
      totalScore = Number(configSheet.getRange('B1').getValue());
    }

    // Get model answers
    var modelSheet = ss.getSheetByName('_MODEL_ANSWER');
    if (!modelSheet) throw new Error('No _MODEL_ANSWER sheet found.');
    var modelData = modelSheet.getDataRange().getValues();
    var modelHeaders = modelData[0];
    var modelFiles = {};
    for (var i = 1; i < modelData.length; i++) {
      var row = modelData[i];
      var fileName = row[0];
      modelFiles[fileName] = {
        path: row[2],
        type: row[1],
        width: row[4],
        height: row[5]
      };
    }

    // Get student answers
    var studentSheet = ss.getSheetByName('_STUDENT_ANSWERS');
    if (!studentSheet) throw new Error('No _STUDENT_ANSWERS sheet found.');
    var studentData = studentSheet.getDataRange().getValues();
    var studentHeaders = studentData[0];
    var studentRows = [];
    var studentName = '';
    for (var i = 1; i < studentData.length; i++) {
      var row = studentData[i];
      if (row[1] == folderId) {
        studentRows.push(row);
        studentName = row[0];
      }
    }
    if (studentRows.length === 0) {
      return 'No files found for this student in _STUDENT_ANSWER.';
    }
    Logger.log('Evaluating student: ' + studentName);
    Logger.log(`Files found: ${studentRows.length}`);
    // Prepare results sheet
    var resultsSheetName = '_STUDENT_RESULTS';
    var resultsSheet = ss.getSheetByName(resultsSheetName);
    if (!resultsSheet) {
      resultsSheet = ss.insertSheet(resultsSheetName);
      resultsSheet.appendRow([
        'Student Folder Name', 'Student Folder ID', 'Student File Name',
        'Path Match', 'Type Match', 'Width Match', 'Height Match', 'Awarded Score'
      ]);
    }
    // Remove previous results for this student
    var resultsData = resultsSheet.getDataRange().getValues();
    var keepRows = [resultsData[0]];
    for (var i = 1; i < resultsData.length; i++) {
      if (resultsData[i][1] != folderId) {
        keepRows.push(resultsData[i]);
      }
    }
    resultsSheet.clear();
    resultsSheet.getRange(1, 1, keepRows.length, keepRows[0].length).setValues(keepRows);

    // Evaluate each student file
    var rawScore = 0;
    var results = [];
    studentRows.forEach(function(row) {
      var fileName = row[2];
      var studentPath = row[4];
      var studentType = row[3];
      var studentWidth = row[6];
      var studentHeight = row[7];
      var model = modelFiles[fileName];
      var pathMatch = false, typeMatch = false, widthMatch = false, heightMatch = false, awarded = 0;
      if (model) {
        pathMatch = (String(studentPath) === String(model.path));
        typeMatch = (String(studentType) === String(model.type));
        widthMatch = (String(studentWidth) === String(model.width));
        heightMatch = (String(studentHeight) === String(model.height));
        awarded = (pathMatch && typeMatch && widthMatch && heightMatch) ? 1 : 0;
      }
      results.push([
        row[0], row[1], fileName,
        pathMatch ? '✔' : '', typeMatch ? '✔' : '', widthMatch ? '✔' : '', heightMatch ? '✔' : '', awarded
      ]);
      rawScore += awarded;
    });

    // Write results
    if (results.length > 0) {
      resultsSheet.getRange(resultsSheet.getLastRow() + 1, 1, results.length, results[0].length).setValues(results);
    }

    // Scale score
    var finalScore = 0;
    if (results.length > 0) {
      finalScore = Math.round(rawScore * totalScore / results.length);
    }

    // Update _SCORE sheet
    var scoreSheet = ss.getSheetByName('_SCORE');
    if (!scoreSheet) {
      scoreSheet = ss.insertSheet('_SCORE');
      scoreSheet.appendRow(['Folder ID', 'Score']);
    }
    var scoreData = scoreSheet.getDataRange().getValues();
    var found = false;
    for (var i = 1; i < scoreData.length; i++) {
      if (scoreData[i][0] == folderId) {
        scoreSheet.getRange(i + 1, 2).setValue(finalScore);
        found = true;
        break;
      }
    }
    if (!found) {
      scoreSheet.appendRow([folderId, finalScore]);
    }

    return 'Evaluation complete. Raw score: ' + rawScore + '/' + results.length +
      ', Final score: ' + finalScore + ' (scaled to ' + totalScore + '). See _STUDENT_RESULTS and _SCORE sheets.';
  } catch (e) {
    return 'Error: ' + e.message;
  }
}

/**
 * Analyzes a specific student folder and logs file attributes to the "_STUDENT_ANSWERS" sheet.
 * @param {string} folderId The ID of the student Google Drive folder.
 * @return {string} Status message.
 */
function analyzeStudentFolder(folderId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var foldersSheet = ss.getSheetByName('_STUDENT_FOLDERS');
    var studentFolders = foldersSheet ? foldersSheet.getDataRange().getValues() : [];
    var studentName = '';
    for (var i = 1; i < studentFolders.length; i++) {
      if (studentFolders[i][1] == folderId) {
        studentName = studentFolders[i][0];
        break;
      }
    }

    var sheetName = '_STUDENT_ANSWERS';
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow([
        'Student Folder Name', 'Student Folder ID', 'File Name', 'File Type', 'Path',
        'Size (bytes)', 'Width', 'Height', 'Aspect Ratio', 'Orientation'
      ]);
    }

    // Remove previous entries for this student folder
    var data = sheet.getDataRange().getValues();
    var rowsToKeep = [data[0]]; // keep header
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] != folderId) {
        rowsToKeep.push(data[i]);
      }
    }
    sheet.clear();
    sheet.getRange(1, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);

    var folder = DriveApp.getFolderById(folderId);
    var results = [];
    processStudentFolder(folder, '', results, studentName, folderId);

    if (results.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, results.length, results[0].length).setValues(results);
      return 'Analysis complete. ' + results.length + ' file(s) found for "' + studentName + '". See the "_STUDENT_ANSWER" sheet.';
    } else {
      return 'No files found in the student folder.';
    }
  } catch (e) {
    return 'Error: ' + e.message;
  }
}

/**
 * Recursively processes a student folder and collects file attributes.
 * @param {Folder} folder The DriveApp Folder.
 * @param {string} path The relative path from the student folder root.
 * @param {Array} results The array to store results.
 * @param {string} studentName The name of the student.
 * @param {string} studentId The folder ID of the student.
 */
function processStudentFolder(folder, path, results, studentName, studentId) {
  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    var mime = file.getMimeType();
    var name = file.getName();
    var relPath = path ? path + '/': '/';
    var size = file.getSize();
    var width = '', height = '', aspect = '', orientation = '';
    if (mime.startsWith('image/')) {
      try {
        var fileId = file.getId();
        var meta = Drive.Files.get(fileId, {fields: "imageMediaMetadata"});
        var imgMeta = meta.imageMediaMetadata || {};
        width = imgMeta.width || '';
        height = imgMeta.height || '';
        if (width && height) {
          aspect = getAspectRatio(width, height);
          orientation = width > height ? 'Landscape' : width < height ? 'Portrait' : 'Square';
        }
      } catch (err) {
        Logger.log("Error processing image " + name + ": " + err.message);
      }
    }
    results.push([studentName, studentId, name, mime, relPath, size, width, height, aspect, orientation]);
  }
  var folders = folder.getFolders();
  while (folders.hasNext()) {
    var sub = folders.next();
    var subPath = path ? path + '/' + sub.getName() : sub.getName();
    processStudentFolder(sub, subPath, results, studentName, studentId);
  }
}

/**
 * Returns student folder records and scores for the modal UI.
 * Ensures _STUDENT_FOLDERS and _SCORE sheets exist.
 * @return {Object} {folders: [{name, id}], scores: {folderId: score}}
 */
function getStudentFoldersAndScores() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Ensure _STUDENT_FOLDERS exists
  var foldersSheet = ss.getSheetByName('_STUDENT_FOLDERS');
  if (!foldersSheet) {
    foldersSheet = ss.insertSheet('_STUDENT_FOLDERS');
    foldersSheet.appendRow(['Student Folder Name', 'Folder ID']);
  }
  var foldersData = foldersSheet.getDataRange().getValues();
  var folders = [];
  for (var i = 1; i < foldersData.length; i++) {
    folders.push({ name: foldersData[i][0], id: foldersData[i][1] });
  }

  // Ensure _SCORE exists
  var scoreSheet = ss.getSheetByName('_SCORE');
  if (!scoreSheet) {
    scoreSheet = ss.insertSheet('_SCORE');
    scoreSheet.appendRow(['Folder ID', 'Score']);
  }
  var scoreData = scoreSheet.getDataRange().getValues();
  var scores = {};
  for (var j = 1; j < scoreData.length; j++) {
    scores[scoreData[j][0]] = scoreData[j][1];
  }

  return { folders: folders, scores: scores };
}

/**
 * Updates the score for a student folder.
 * @param {string} folderId
 * @param {number} score
 * @return {string} Status message.
 */
function updateStudentScore(folderId, score) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scoreSheet = ss.getSheetByName('_SCORE');
  if (!scoreSheet) {
    scoreSheet = ss.insertSheet('_SCORE');
    scoreSheet.appendRow(['Folder ID', 'Score']);
  }
  var data = scoreSheet.getDataRange().getValues();
  var found = false;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == folderId) {
      scoreSheet.getRange(i + 1, 2).setValue(score);
      found = true;
      break;
    }
  }
  if (!found) {
    scoreSheet.appendRow([folderId, score]);
  }
  return 'Score updated.';
}

/**
 * Analyzes a Google Drive folder and records all immediate subfolders (student folders)
 * in a sheet named "_STUDENT_FOLDERS".
 * @param {string} folderId The ID of the parent Google Drive folder.
 * @return {string} Status message.
 */
function analyzeStudentFolders(folderId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = '_STUDENT_FOLDERS';
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    } else {
      sheet.clear();
    }
    sheet.appendRow(['Student Folder Name', 'Folder ID']);
    var folder = DriveApp.getFolderById(folderId);
    var folders = folder.getFolders();
    var count = 0;
    while (folders.hasNext()) {
      var sub = folders.next();
      sheet.appendRow([sub.getName(), sub.getId()]);
      count++;
    }
    if (count > 0) {
      return 'Found ' + count + ' student folder(s). See the "_STUDENT_FOLDERS" sheet.';
    } else {
      return 'No student folders found in the specified folder.';
    }
  } catch (e) {
    return 'Error: ' + e.message;
  }
}

/**
 * Analyzes a Google Drive folder and logs image file attributes to the "_MODEL_ANSWER" sheet.
 * @param {string} folderId The ID of the Google Drive folder.
 * @return {string} Status message.
 */
function analyzeDriveFolder(folderId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = '_MODEL_ANSWER';
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    } else {
      sheet.clear();
    }
    // Write headers
    sheet.appendRow([
      'File Name', 'File Type', 'Path', 'Size (bytes)', 
      'Width', 'Height', 'Aspect Ratio', 'Orientation'
    ]);
    var folder = DriveApp.getFolderById(folderId);
    var results = [];
    processFolder(folder, '', results);

    // Write results to sheet
    if (results.length > 0) {
      sheet.getRange(2, 1, results.length, results[0].length).setValues(results);
      return 'Analysis complete. ' + results.length + ' image file(s) found. See the "_MODEL_ANSWER" sheet.';
    } else {
      return 'No image files found in the folder.';
    }
  } catch (e) {
    return 'Error: ' + e.message;
  }
}

/**
 * Recursively processes a folder and collects image file attributes.
 * @param {Folder} folder The DriveApp Folder.
 * @param {string} path The relative path from the root folder.
 * @param {Array} results The array to store results.
 */
function processFolder(folder, path, results) {
  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    var mime = file.getMimeType();
    if (mime.startsWith('image/')) {
      var name = file.getName();
      var relPath = path ? path + '/' : '/';
      var size = file.getSize();
      var width = '', height = '', aspect = '', orientation = '';
      try {
        var fileId = file.getId();
        var meta = Drive.Files.get(fileId, {fields: "imageMediaMetadata"});
        var imgMeta = meta.imageMediaMetadata || {};
        width = imgMeta.width || '';
        height = imgMeta.height || '';
        if (width && height) {
          aspect = getAspectRatio(width, height);
          orientation = width > height ? 'Landscape' : width < height ? 'Portrait' : 'Square';
        }
      } catch (err) {
        Logger.log("Error processing image " + name + ": " + err.message);
      }
      results.push([name, mime, relPath, size, width, height, aspect, orientation]);
    }
  }
  var folders = folder.getFolders();
  while (folders.hasNext()) {
    var sub = folders.next();
    var subPath = path ? path + '/' + sub.getName() : sub.getName();
    processFolder(sub, subPath, results);
  }
}

/**
 * Returns the aspect ratio as a string (e.g., "16:9", "4:3", etc).
 * @param {number} width 
 * @param {number} height 
 * @return {string}
 */
function getAspectRatio(width, height) {
  var gcd = function(a, b) {
    return b == 0 ? a : gcd(b, a % b);
  };
  var divisor = gcd(width, height);
  var w = Math.round(width / divisor);
  var h = Math.round(height / divisor);
  var commonRatios = {
    '1:1': [1, 1],
    '4:3': [4, 3],
    '3:2': [3, 2],
    '5:4': [5, 4],
    '16:9': [16, 9],
    '16:10': [16, 10],
    '21:9': [21, 9]
  };
  for (var key in commonRatios) {
    var ratio = commonRatios[key];
    if (Math.abs(w / h - ratio[0] / ratio[1]) < 0.05) {
      return key;
    }
  }
  return w + ':' + h;
}

