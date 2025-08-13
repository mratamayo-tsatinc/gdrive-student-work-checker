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
      var relPath = path ? path + '/' : name;
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