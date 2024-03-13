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
  // Replace 'Data' with the name of your sheet and adjust range accordingly
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  var range = sheet.getRange("A1:A"); // Assuming your items are in column A
  var values = range.getValues();
  var items = values.flat().filter(String); // Remove empty strings
  return items;
}

function writeToSheet(selectedItems) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var cell = sheet.getActiveCell();
  cell.setValue(selectedItems.join(", ")); // Join array items with comma and space
}
