# googlesheets-multiple-dropdown-select

Creating a multiple selections drop-down menu in a Google Sheet cell without using external integrations is a challenge because Google Sheets does not natively support multi-select dropdowns. However, you can simulate this functionality by using Google Apps Script to create a custom sidebar or dialog that allows multiple selections, which are then concatenated into a single cell. Here's an example approach using a custom sidebar with checkboxes:

## Step 1: Create the Google Apps Script and HTML file:"

Create a new script file in the Apps Script editor:

1. Open your Google Sheet.
2. Go to Extensions > Apps Script.
3. Delete any code in the script editor and paste in the following script:

This is the file Google-Apps-Script-Multiple-Dropdown-Select.js

```
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

```

Please note you will need to change 

```
  // Replace 'Data' with the name of your sheet and adjust range accordingly
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  var range = sheet.getRange("A1:A"); // Assuming your items are in column A
```

Create a new HTML file in the Apps Script editor:

1. Click on File > New > Html file.
2. Name it Page.
3. Paste in the following HTML and JavaScript:

```
<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
  </head>
  <body>
    <div id="checkboxes"></div>
    <button onclick="submit()">Submit</button>
    
    <script>
      // Load items from the server-side function getItems
      google.script.run.withSuccessHandler(function(items) {
        var container = document.getElementById('checkboxes');
        items.forEach(function(item) {
          var checkbox = document.createElement('input');
          checkbox.type = 'checkbox';
          checkbox.id = item;
          checkbox.value = item;
          
          var label = document.createElement('label');
          label.htmlFor = item;
          label.appendChild(document.createTextNode(item));
          
          container.appendChild(checkbox);
          container.appendChild(label);
          container.appendChild(document.createElement('br'));
        });
      }).getItems();
      
      function submit() {
        var selectedItems = [];
        var checkboxes = document.getElementById('checkboxes').querySelectorAll('input[type=checkbox]:checked');
        checkboxes.forEach(function(checkbox) {
          selectedItems.push(checkbox.value);
        });
        
        // Write the selected items back to the sheet
        google.script.run.writeToSheet(selectedItems);
        
        // Optionally close the sidebar
        google.script.host.close();
      }
    </script>
  </body>
</html>
```
This is the file Google-Apps-Script-Multiple-Dropdown-Select.html


## Step 2: Use the Multi-Select Dropdown

1. After saving your script and HTML file, reload your Google Sheet.
2. You'll see a new menu item "Custom Tools" with the option "Multi-Select Dropdown". Clicking on this will open a sidebar with checkboxes.
3. Select multiple items and click "Submit" to write the selected values to the active cell, separated by commas.

This method utilizes a sidebar for selections, combining selected items into a single string for the active cell. Adjust the script as needed for your specific use case, such as modifying the range where dropdown items are sourced or the way selections are formatted.
