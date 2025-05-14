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
