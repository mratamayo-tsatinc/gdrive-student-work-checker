/**
 * Google Apps Script context for Google Sheet extension:
 * 
 * Platform: Google Sheet
 * Language: Google Apps Script
 * 
 * Objective:
 * - Adds a custom menu to Google Sheets to launch a web UI through a pop-up window.
 * - The UI allows input of a Google Drive folder ID (model answer folder).
 * - Analyze the folder, a script that checks the folder structure and inspects image files for:
 *   - File name
 *   - File type
 *   - Path (relative to root folder)
 *   - Size in bytes
 *   - Width and height in pixels
 *   - Aspect ratio (standard forms: 1:1, 5:4, 16:9, etc)
 *   - Orientation (landscape or portrait)
 * - Results are recorded in a designated sheet in the spreadsheet.
 * 
 * UI/UX:
 * - Custom menu item in Google Sheet launches a modal (pop-up web content).
 * - Web UI contains headers, an input box for folder ID, and an "Analyze" button.
 * - On clicking "Analyze", the script processes the folder and writes results to the sheet.
 * 
 * Usage:
 * - User opens the Google Sheet.
 * - User clicks the custom menu and opens the modal window.
 * - User enters the folder ID and clicks "Analyze".
 * - The script analyzes the folder and logs image file attributes in the sheet.
 * 
 * Additional Feature:
 * - Adds a submenu under the custom menu to launch a second modal dialog.
 * - The modal uses a Google Material theme with a white overlay and a spinner during processing.
 * - This modal presents a form for the user to enter a Google Drive folder ID or URL containing student submission folders.
 * - On clicking "Analyze", the script extracts all individual student subfolders within the specified folder.
 * - The names and IDs of these student subfolders are recorded in a dedicated sheet named "_STUDENT_FOLDERS".


Additional Feature: Student Task Submission Form
- Adds a submenu to launch a modal form with a Google Material-based web UI.
- The modal displays records from the `_STUDENT_FOLDERS` and `_SCORE` sheets (creating them if they do not exist).
- Presents a list with three columns: **Item**, **Score**, and **Action**.
    - **Item**: Shows the folder name (prominently) and folder ID (less prominent). The folder name is clickable and opens the folder in a new tab.
    - **Score**: Displays the score from the `_SCORE` sheet, using the folder ID as the lookup key. Defaults to 0 if no score is found.
    - **Action**: Contains an "Analyze", "Evaluate" button for each row to trigger analysis for the corresponding student folder.
- Analyze button
When the "Analyze" button is pressed for a student folder, the script will analyze the contents of that specific student folder and extract the following information for each file. The extracted data will be recorded in a designated sheet named `_STUDENT_ANSWERS`:
- Student folder name
- Student folder ID
- File name
- File type
- Path (relative to the student folder root)
- Size in bytes
- Width and height in pixels
- Aspect ratio (standard forms: 1:1, 5:4, 16:9, etc.)
- Orientation (landscape or portrait)

Each row in the `_STUDENT_ANSWERS` sheet will represent a file found within the student folder, with all the above attributes captured for review and further processing.


Evaluate Button Functionality
- Uses the `_MODEL_ANSWER`, `_STUDENT_ANSWERS`, and `_CONFIG` sheets.
- For each student (identified by folder ID), extracts all files listed in the `_STUDENT_ANSWERS` sheet.
- For each student file, finds the corresponding model answer file (matching by file name).
- Compares the following attributes between student and model answer files:
    - **Path**: Must match exactly.
    - **File Type**: Must match exactly.
    - **Width**: Must match exactly.
    - **Height**: Must match exactly.
- For each attribute, records a checkbox (checked if matched, unchecked if not) in the `_STUDENT_RESULTS` sheet, along with:
    - Student folder name
    - Student folder ID
    - Student file name
    - Match status for file type, path, width, and height
    - Awarded score for the file (1 if all attributes match, 0 otherwise)
- Sums the awarded scores for all files to get the student's total raw score.
- Adjusts the total raw score to the activity total score specified in `_CONFIG` sheet cell B1 (scaling proportionally).
- Records the final score (folder ID and score) in the `_SCORE` sheet.
- Displays a white overlay with a spinner during processing for user feedback.
- Update the score column in the student task submission list

Additional Feature: Point Deduction for Extra Files
- Compares the number of files submitted by each student (from `_STUDENT_ANSWERS`) to the number required by the model answer (from `_MODEL_ANSWER`).
- If the student's file count is less than or equal to the model answer file count, no deduction is applied.
- If the student's file count exceeds the model answer file count, deduct one point from the student's total raw score for each extra file.
- The deduction is applied before scaling the raw score to the activity total score in `_CONFIG` B1.
- The final score, after deduction and scaling, is recorded in the `_SCORE` sheet.
- The deduction logic is reflected in the student task submission list and any relevant UI feedback.
 */