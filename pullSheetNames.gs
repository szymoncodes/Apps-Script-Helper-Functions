function pullSheetNames() {
  /**
   * Pulls the names of all the sheets in the active document 
   *
   * @return {array} sheetNameArray   An array of the names of all the sheets in the active document
   */

  var sheetNameArray = [];

  // Get the name of each sheet and add them to the array
  var allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  sheetNameArray = allSheets.map(function (sheet) {
    return [sheet.getName()];
  });

  return sheetNameArray;
};
