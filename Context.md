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
 */