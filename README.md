# Apps Script Helper Functions
---

## pullSheetNames
---
A simple function that allows you to load the names of all the sheets in the active document. 

```JavaScript
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
```

## queryMultipleSheets
---
A function that allows you to set a more complex QUERY formula using all the specified sheets.

```JavaScript
function queryMultipleSheets(sheet, setQuery, columnNotNull, sheetNameArray, dataRange) {
  /**
   * Pulls the data from sheetNameArray, using the dataRange.
   * Then sets the QUERY formula in the setQuery cell on the sheet.
   *
   * @param {Sheet}  sheet            Sheet to set the QUERY to
   * @param {String} setQuery         Cell to set the QUERY function to
   * @param {String} columnNotNull    Set the column that is required to have data in QUERY
   * @param {Array}  sheetNameArray   Array of sheet names
   * @param {String} dataRange        Range of cells to pull
   */

  // Picks the range and edits it so that it can be used with QUERY
  var sheetRange = sheetNameArray.map(function (value) {
    value = value.concat("'" + value + "'!" + dataRange);
    return value.splice(1, 1);
  });
  sheetRange = sheetRange.join("; ")

  // Create the QUERY formula to pull the data from all sheets
  var queryFormula = "=QUERY({" + sheetRange + "},\"select * where " + columnNotNull + " is not null\", 0)";
  var cell = sheet.getRange(setQuery);
  cell.setFormula(queryFormula);
  Logger.log("Data from " + sheetRange + " added.");
};
````
