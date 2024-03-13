function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Tools')
      .addItem('Multi-Select Dropdown', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Page')
      .setTitle('Multi-Select Dropdown')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getItems() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  var activeColumn = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell().getColumn();
  var range = sheet.getRange(1, activeColumn, sheet.getMaxRows(), 1); 
  var values = range.getValues();
  var items = values.flat().filter(String); 
  return items;
}

function writeToSheet(selectedItems) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var cell = sheet.getActiveCell();
  cell.setValue(selectedItems.join(", "));
}
