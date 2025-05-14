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
