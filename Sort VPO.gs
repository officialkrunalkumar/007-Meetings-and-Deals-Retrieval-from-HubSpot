function sortVPO() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('VPO');
  
  // Define the range you want to sort, excluding the header row
  var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  
  // Sort the range by column D (which is column 4 in zero-indexed)
  range.sort({column: 4, ascending: false});
}
