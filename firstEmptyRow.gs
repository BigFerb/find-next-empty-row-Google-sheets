/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */

/**
 * Creates a menu entry in the Google Sheets UI when the spreadsheet is opened.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version. (I think)
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
      .addItem('Find next empty row', 'findEmptyRow')
      .addToUi();
}

/**
 * Runs when the add-on is installed.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version. (I think)
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Gets the current, active range.  
 * 
 * Finds the first empty row (999 columns) 
 * at or below the last row of the current range. 
 *
 * Sets the cell in that empty row as the new
 * active range.
 *
 * @return {Array.<string>} The selected text.
 */
function findEmptyRow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getActiveRange();
  var startRow = range.getLastRow();
  var startCol = range.getColumn();
  
  // find first blank row
  var i = startRow;
  while (!sheet.getRange(i,1,1,999).isBlank()) {
    i++;
  }
  
  // set active cell
  var nextCell = sheet.getRange(i,startCol);
  sheet.setActiveRange(nextCell);
}
