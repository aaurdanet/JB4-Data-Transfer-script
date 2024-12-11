function resetTableWithFormattedSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Define the name of the sheet with the desired format
  var formattedSheetName = "SOURCE_FORMAT"; // Replace with the name of your template sheet
  
  // Get the formatted sheet
  var formattedSheet = spreadsheet.getSheetByName(formattedSheetName);
  if (!formattedSheet) {
    SpreadsheetApp.getUi().alert("Formatted sheet '" + formattedSheetName + "' not found!");
    return;
  }
  
  // Define the range to copy
  var rangeToCopy = formattedSheet.getRange("A1:G29");
  var copiedValues = rangeToCopy.getValues();
  var copiedBackgrounds = rangeToCopy.getBackgrounds();
  var copiedFontColors = rangeToCopy.getFontColors();
  var copiedFontWeights = rangeToCopy.getFontWeights();
  var copiedNumberFormats = rangeToCopy.getNumberFormats();
  
  // Get the active sheet
  var activeSheet = spreadsheet.getActiveSheet();
  
  // Clear the active sheet
  activeSheet.clear();
  
  // Paste the copied range into the active sheet
  var destinationRange = activeSheet.getRange("A1:G29");
  destinationRange.setValues(copiedValues); // Paste values
  destinationRange.setBackgrounds(copiedBackgrounds); // Paste backgrounds
  destinationRange.setFontColors(copiedFontColors); // Paste font colors
  destinationRange.setFontWeights(copiedFontWeights); // Paste font weights
  destinationRange.setNumberFormats(copiedNumberFormats); // Paste number formats
  
  // Set font to Arial for the range
  destinationRange.setFontFamily("Arial");
  
  // Merge cells F1:G1
  activeSheet.getRange("F1:G1").merge();
  
  // Optionally, adjust column widths for the specified range
  for (var i = 1; i <= 7; i++) { // Columns A to G
    activeSheet.setColumnWidth(i, formattedSheet.getColumnWidth(i));
  }
  
  Logger.log("CLEARED ALL FIELDS")
}
