// Function to extract the value of a specific cell
function getDriverID() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the value from cell A2
  const inputValue = sheet.getRange("A2").getValue();
  
  // Extract the integer using regex
  const extractedValue = inputValue.match(/\d+/)[0]; // Matches the first integer in the string
  
  return getFromLoadsReport(extractedValue);
}


function formatDateToMMDDYYYY(dateString) {
  const dateObject = new Date(dateString);

  // Check for invalid date
  if (isNaN(dateObject)) {
    throw new Error("Invalid date format");
  }

  const month = String(dateObject.getMonth() + 1).padStart(2, '0'); // Months are zero-based
  const day = String(dateObject.getDate()).padStart(2, '0');
  const year = dateObject.getFullYear();

  return `${month}/${day}/${year}`;
}


function getFromLoadsReport(driverID) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange("B2").setValue(driverID);
  driverID = driverID.toString().trim();
  
  const sourceSpreadsheet = SpreadsheetApp.openById('1ZUQUjjNpyX9XnXHSGSX571WD4kL4G9MBfud54-7blPU');
  const otherSheet = sourceSpreadsheet.getSheetByName("Copy of RICARDO A.");
  
  const DRIVER_ID_COLUMN = 11; // Column K (1-based index)
  const PAID_TO_DRIVER_FLAG_COLUMN = 20; // Column B (1-based index) contains the specific value to fetch
  const LOAD_ID_COLUMN = 1; // Column A
  const CITY_FROM_COLUMN = 2; // Column B
  const STATE_FROM_COLUMN = 3; // Column C
  const CITY_TO_COLUMN = 4; // Column D
  const DELIVERY_DATE_COLUMN = 9; // Column I
  const GROSS_RATE_COLUMN = 7; // Column G
  
  // Get all the data from the sheet
  const data = otherSheet.getDataRange().getValues();
  
 // Initialize an array to store results
  const results = [];
  
  // Find and store all occurrences of the driver ID with specific value
  for (let i = 1; i < data.length; i++) { // Start from 1 to skip the header row
    if (data[i][DRIVER_ID_COLUMN - 1].toString().trim() === driverID.toString().trim()) {
      const specificValue = data[i][PAID_TO_DRIVER_FLAG_COLUMN - 1].toString().trim();
      if (specificValue === "Yes") { // Only include items where specific value is "Yes"
        // const loadID = data[i][LOAD_ID_COLUMN - 1].toString().trim();
        results.push({
          row: i + 1, // 1-based row index
          column: DRIVER_ID_COLUMN, // The column of the Driver ID
          loadID: data[i][LOAD_ID_COLUMN - 1].toString().trim(),
          cityFrom: data[i][CITY_FROM_COLUMN - 1].toString().trim(),
          stateFrom: data[i][STATE_FROM_COLUMN - 1].toString().trim(),
          cityTo: data[i][CITY_TO_COLUMN - 1].toString().trim(),
          deliveryDate: formatDateToMMDDYYYY(data[i][DELIVERY_DATE_COLUMN - 1]),
          grossRate: data[i][GROSS_RATE_COLUMN - 1]
        });
      }
    }
  }
  
// Log the results
if (results.length > 0) {
  results.forEach(result => {
    Logger.log(`Driver ID ${driverID} found at Row: ${result.row}, Column: ${result.column}, LoadID Value: ${result.loadID} City From: ${result.cityFrom}, State From: ${result.stateFrom}, City To: ${result.cityTo}, Delivery Date: ${result.deliveryDate}, Gross Rate: ${result.grossRate}`);

  });
} else {
  Logger.log(`Driver ID ${driverID} not found in the data with Specific Value "Yes".`);
}
  // Return results
  return results;
}






