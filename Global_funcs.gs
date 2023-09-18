/**
 * If no row number is specified, returns a 2d array of all JSON-parsed data in a sheet. 
 * Otherwise, returns a 2d array containing the JSON-parsed data of the specified row in [0]. 
 * This method is global for access by all classes. 
 * 
 * @param {string} sheetName The name of the sheet to parse
 * @param {number} [rowNum] The number of the row to extract, if applicable
 * @return A 2d array of displayed values, with valid JSON strings converted to objects
 */
function parsedValues(sheetName, rowNum){
  const sheet = APP_CONSTS.SS.getSheetByName(sheetName);
  const range = rowNum ? sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()) : sheet.getDataRange();
  const values = range.getDisplayValues();

  const parsedValues = values.map(value => {
    return value.map(cell => {
      try {
        return JSON.parse(cell);
      } catch(e) {
        return cell;
      }
    });
  });
  return parsedValues;
}