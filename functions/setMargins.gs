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
