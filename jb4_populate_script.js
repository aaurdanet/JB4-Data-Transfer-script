// Function to extract the value of a specific cell
function getDriverID() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the value from cell A2
  const inputValue = sheet.getRange("A2").getValue();
  
  // Extract the integer using regex
  const extractedValue = inputValue.match(/\d+/)[0]; // Matches the first integer in the string
  // Ensure inputValue is a string
  const inputString = String(extractedValue);
  getFromFuelPurchases(inputString);

  return getFromLoadsReport(inputString);
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

  driverID = driverID.toString().trim();
  
  const sourceSpreadsheet = SpreadsheetApp.openById('1ZUQUjjNpyX9XnXHSGSX571WD4kL4G9MBfud54-7blPU');
  // CHANGE LINE BELLOW TO SOURCE SHEET NAME
  const otherSheet = sourceSpreadsheet.getSheetByName("DANIEL A. BACKUP");
  
  const DRIVER_ID_COLUMN = 11; // Column K (1-based index)
  const PAID_TO_DRIVER_FLAG_COLUMN = 20; // Column t (1-based index) contains the specific value to fetch
  const BILLED_STATUS_FLAG_COLUMN = 19; // Column t (1-based index) contains the specific value to fetch
  const LOAD_ID_COLUMN = 1; // Column A
  const CITY_FROM_COLUMN = 2; // Column B
  const STATE_FROM_COLUMN = 3; // Column C
  const CITY_TO_COLUMN = 4; // Column D
  const STATE_TO_COLUMN = 5;// Column E
  const DELIVERY_DATE_COLUMN = 9; // Column I
  const GROSS_RATE_COLUMN = 7; // Column G
  
  // Get all the data from the sheet
  const data = otherSheet.getDataRange().getValues();
  
 // Initialize an array to store results
  const results = [];
  
  // Find and store all occurrences of the driver ID with specific value
  for (let i = 1; i < data.length; i++) { // Start from 1 to skip the header row
    if (data[i][DRIVER_ID_COLUMN - 1].toString().trim() === driverID.toString().trim()) {
      const Y_N_Flag = data[i][PAID_TO_DRIVER_FLAG_COLUMN - 1].toString().trim();
      const Billed_Flag = data[i][BILLED_STATUS_FLAG_COLUMN - 1].toString().trim();
      Logger.log(Billed_Flag)
      if (Y_N_Flag === "No" && Billed_Flag ==="Billed") { // Only include items where specific value is "Yes"
        // const loadID = data[i][LOAD_ID_COLUMN - 1].toString().trim();
        results.push({
          // row: i + 1, // 1-based row index <----- uncomment to debug rows and columns
          // column: DRIVER_ID_COLUMN, // The column of the Driver ID
          loadID: data[i][LOAD_ID_COLUMN - 1].toString().trim(),
          cityFrom: data[i][CITY_FROM_COLUMN - 1].toString().trim(),
          stateFrom: data[i][STATE_FROM_COLUMN - 1].toString().trim(),
          cityTo: data[i][CITY_TO_COLUMN - 1].toString().trim(),
          stateTo:data[i][STATE_TO_COLUMN - 1].toString().trim(),
          deliveryDate: formatDateToMMDDYYYY(data[i][DELIVERY_DATE_COLUMN - 1]),
          grossRate: parseFloat(data[i][GROSS_RATE_COLUMN - 1])
        });
      }
    }
  }
  
// Log the results
// if (results.length > 0) {
//   results.forEach(result => {
//     Logger.log(`Driver ID ${driverID} found at Row: ${result.row}, Column: ${result.column}, LoadID Value: ${result.loadID} City From: ${result.cityFrom}, State From: ${result.stateFrom}, City To: ${result.cityTo}, State To: ${result.stateTo} Delivery Date: ${result.deliveryDate}, Gross Rate: ${result.grossRate}`);

//   });
// } else {
//   Logger.log(`Driver ID ${driverID} not found in the data with Specific Value "No".`);
//     return
// }
  // Return results
  return insertIntoLoadID(results);
  
}

function getFromFuelPurchases(driverID){
  driverID = driverID.toString().trim();

  
  const sourceSpreadsheet = SpreadsheetApp.openById('18L-u_NAfp6mVs8t2cmY_T8RwR963Ygl02E5NcxvSY2g');
  const otherSheet = sourceSpreadsheet.getSheetByName("Fuel Purchases");
  
  const DRIVER_ID_COLUMN = 1; // Column A (1-based index)
  const PURCHASE_CITY_COLUMN = 6; // Column F (1-based index)
  const PURCHASE_STATE_COLUMN = 7; // Column G (1-based index) 
  const PURCHASE_DATE_COLUMN = 8; // Column H (1-based index)
  const TO_BE_CHARGED_COLUMN = 13; // Column M (1-based index) 
  const STATUS_FLAG_COLUMN = 14; // Column N (1-based index)
 
  
  // Get all the data from the sheet
  const data = otherSheet.getDataRange().getValues();
  
 // Initialize an array to store results
  const results = [];
  
  // Find and store all occurrences of the driver ID with specific value
  
  for (let i = 1; i < data.length; i++) { // Start from 1 to skip the header row
  
    if (data[i][DRIVER_ID_COLUMN - 1].toString().trim() === driverID.toString().trim()) {
      const specificValue = data[i][STATUS_FLAG_COLUMN - 1].toString().trim();
      
      if (specificValue === "Pending") { // Only include items where specific value is "Yes"
        results.push({
          // row: i + 1, // 1-based row index <----- uncomment to debug rows and columns
          // column: DRIVER_ID_COLUMN, // The column of the Driver ID
          purchaseCity: data[i][PURCHASE_CITY_COLUMN - 1].toString().trim(),
          purchaseState: data[i][PURCHASE_STATE_COLUMN - 1].toString().trim(),
          purchaseDate: formatDateToMMDDYYYY(data[i][PURCHASE_DATE_COLUMN - 1].toString().trim()),
          amount: data[i][TO_BE_CHARGED_COLUMN - 1].toString().trim()
        });
      }
    }
  }
  
// Log the results
if (results.length > 0) {
  results.forEach(result => {
    Logger.log(`Driver ID ${driverID} purchaseCity: ${result.purchaseCity} purchaseState: ${result.purchaseState}, purchaseDate: ${result.purchaseDate}, amount: ${result.amount}`);

  });
} else {
  Logger.log(`Driver ID ${driverID} not found in the data with Specific Value "Pending Fuel Purchases".`);
    return
}
  insertIntoFuelPurchases(results)

}






function insertIntoLoadID(data) {
try {
  // Logger.log('Starting insertIntoSettlement function');
  
  const destinationSpreadsheet = SpreadsheetApp.openById('18L-u_NAfp6mVs8t2cmY_T8RwR963Ygl02E5NcxvSY2g');
  // Logger.log('Spreadsheet opened successfully');
  
  const settlementSheet = destinationSpreadsheet.getSheetByName("Settlement Format");
  if (!settlementSheet) {
    Logger.log('Sheet not found');
    throw new Error('Sheet "Settlement Format" does not exist');
  }
  // Logger.log('Sheet name: ' + settlementSheet.getName());

  const startRow = 5; // Row number where rows will be inserted
  const numberOfRows  = data.length

  const keys = Object.keys(data[0]);
  const numberOfColumns = keys.length;
  
  // Insert rows after the last row
  settlementSheet.insertRows(startRow, numberOfRows);
  // Logger.log(`Inserted ${numberOfRows} rows starting at row ${startRow}`);

  const dataArray = data.map(obj => keys.map(key => obj[key]));


  const range = settlementSheet.getRange(startRow, 1, numberOfRows, numberOfColumns);
  range.setValues(dataArray);
  
  // Change the font of the range
  range.setFontFamily("Arial"); // You can specify your preferred font here
  
} catch (error) {
  Logger.log('Error: could not insert in insert into loadID table');
}

}

function insertIntoFuelPurchases(data) {
try {
  // Logger.log('Starting insertIntoSettlement function');
  
  const destinationSpreadsheet = SpreadsheetApp.openById('18L-u_NAfp6mVs8t2cmY_T8RwR963Ygl02E5NcxvSY2g');
  // Logger.log('Spreadsheet opened successfully');
  
  const settlementSheet = destinationSpreadsheet.getSheetByName("Settlement Format");
  if (!settlementSheet) {
    Logger.log('Sheet not found');
    throw new Error('Sheet "Settlement Format" does not exist');
  }
  // Logger.log('Sheet name: ' + settlementSheet.getName());
  const startColumn = 4; // Start at column D
  const startRow = 22; // Row number where rows will be inserted
  const numberOfRows  = data.length

  const keys = Object.keys(data[0]);
  const numberOfColumns = keys.length;
  
  // Insert rows after the last row
  settlementSheet.insertRows(startRow, numberOfRows);
  // Logger.log(`Inserted ${numberOfRows} rows starting at row ${startRow}`);

  const dataArray = data.map(obj => keys.map(key => obj[key]));


  const range = settlementSheet.getRange(startRow, startColumn, numberOfRows, numberOfColumns);
  range.setValues(dataArray);
  
  // Change the font of the range
  range.setFontFamily("Arial"); // You can specify your preferred font here
  
} catch (error) {
  Logger.log('Error: could not insert in insert into settlement table');
}
}




















