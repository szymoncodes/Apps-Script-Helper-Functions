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

## setMargins
---
A function that allows you to set specified columns into margins.

```JavaScript
function setMargins(sheet, columnWidth, ...args) {
  /**
  * Turns columns into margins by changing their width
  * 
  * @param {Sheet} sheet          Sheet to apply the script to
  * @param {number} columnWidth   Width of the margins
  * @param {number} ...args       Column numbers
  */

  if (args.length > 0) {
    args.forEach(x => sheet.setColumnWidth(x, columnWidth))
  } else {
    return
  };

  Logger.log("Margins set.")
};
```

## setTable
---
A function that allows you to create a simple table header, together with subtitles and a colored background. 

```JavaScript
function setTable(sheet, headerRange, headerColor, titleColor, title, subTitlesRange, subTitles) {
  /**
   * Create a table with a header and Sub Titles.
   *
   * @param {Sheet}   sheet            Sheet to create the table in
   * @param {String}  headerRange      Range of cells that the header will cover
   * @param {String}  headerColor      Color for the header in HEX format
   * @param {String}  titleColor       Color for the title in HEX format
   * @param {String}  title            Title for the table
   * @param {String}  subTitlesRange   Range of cells requiring Sub Titles
   * @param {Array}   subTitles        1D array of Sub Titles to be used
   */

  // Color the headerRange
  var range = sheet.getRange(headerRange);
  range.setBackgroundColor(headerColor);

  // Set the title
  var header = sheet.getRange(headerRange);
  header.merge();
  header.setHorizontalAlignment("center");
  header.setValue(title);

  // Edit the title font
  range.setFontColor(titleColor);
  range.setFontWeight("bold");
  range.setWrap(true);

  // Set the Sub Titles
  var subTitlesArray = [subTitles];
  range = sheet.getRange(subTitlesRange);
  range.setValues(subTitlesArray);

  Logger.log("Table " + title + " set.")
};
```

## simplePullData
---
A function that allows you to pull data through simple formulas like SUM, MAX, MIN etc.

```JavaScript
function simplePullData(row, column, dataRange, formulaName) {
  /**
   * This function allows you to pull data through simple formulas like SUM, MAX, MIN etc.
   *
   * @param {Integer} row          Number of row where you want the data to pull to
   * @param {Integer} column       Number of column where you want the data to pull to
   * @param {String}  dataRange    Range of cells to pull
   * @param {String}  formulaName  The name of the formula you want to use
   */

  var sheetRange = sheetNameArray.map(function (value) {
    value = value.concat("=" + formulaName +"('" + value + "'!" + dataRange + ")");
    return value.splice(1, 1);
  });

  var cell = sheet.getRange(row, column, sheetNameArray.length, 1);
  cell.setFormulas(sheetRange);
  Logger.log(sheetRange + " set.");
};
```
