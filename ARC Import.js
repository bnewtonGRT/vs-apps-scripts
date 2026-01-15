/**
 * Prepares the active spreadsheet for ARC processing.
 * - Inserts two columns at the beginning.
 * - Retrieves data from the active sheet.
 * - Processes the data step by step, removing asterisks, handling missing fare and ticket numbers.
 * - Prepares the final data.
 * - Clears the sheet and sets it with the final data.
 * - Formats the sheet.
 */
function prepareARCsheet() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();

  // Insert two columns at the beginning
  sheet.insertColumnsBefore(sheet.getRange('A1').getColumn(), 2);

  // Get the data array from the sheet
  var dataArray = sheet.getDataRange().getValues();

  // Process the data step by step
  var removeAsterisksData = removeAsterisks(dataArray);
  var missingFareData = missingFare(removeAsterisksData);
  var missingTicketNumberData = missingTicketNumber(missingFareData);

  // Prepare the final data
  var finalData = prepareFinalData(missingTicketNumberData);

  // Clear the sheet and set the final values
  sheet.clearContents();

  var numRows = finalData.length;
  var numColumns = finalData[0].length;
  sheet.getRange(1, 1, numRows, numColumns).setValues(finalData);

  // Format the sheet
  addPaymentTypeColumn(sheet);

  // Format the sheet
  formatArcSheet(sheet);
}

/**
 * Removes asterisks from each cell in a 2D data array.
 *
 * @param {any[][]} dataArray - The 2D data array to process.
 * @returns {any[][]} A modified 2D data array with asterisks removed.
 */
function removeAsterisks(dataArray) {
  // Define the character to replace and the replacement string
  var toReplace = /\*/g; // Using a regular expression to replace all occurrences of '*'
  var replaceWith = '';

  // Use map to process each cell in the array
  return dataArray.map(row => row.map(cell => cell.toString().replace(toReplace, replaceWith)));
}

/**
 * Filters rows in a 2D data array to include only those with a non-empty fare value.
 *
 * @param {any[][]} dataArray - The 2D data array to filter.
 * @returns {any[][]} A new 2D data array containing rows with non-empty fare values.
 */
function missingFare(dataArray) {
  const arcArray = [];

  for (var i = 0; i < dataArray.length; i++) {
    var row = dataArray[i];
    if (row[7].length !== 0 || row[5] === "CX") {
      arcArray.push(row);
    }
  }

  return arcArray;
}

/**
 * Filters rows in a 2D data array to include only those with a non-empty ticket number.
 *
 * @param {any[][]} dataArray - The 2D data array to filter.
 * @returns {any[][]} A new 2D data array containing rows with non-empty ticket numbers.
 */
function missingTicketNumber(dataArray) {
  const arcArray = [];

  for (var i = 0; i < dataArray.length; i++) {
    var row = dataArray[i];
    if (row[4] !== "") {
      arcArray.push(row);
    }
  }

  return arcArray;
}

/**
 * Prepares the final data for ARC processing from rows with missing ticket numbers.
 *
 * @param {any[][]} missingTicketNumber - The 2D data array with rows containing missing ticket numbers.
 * @returns {any[][]} A new 2D data array with the final data prepared for ARC processing.
 */
function prepareFinalData(missingTicketNumber) {
  return missingTicketNumber.map(row => {
    const singleRow = [
      row[4],
      row[7],
      row[8],
      row[10],
      "ARC",
      row[5],
      row[6],
      row[2]
    ];

    // Process columns 1 to 3 for trailing minus
    for (let r = 1; r < 4; r++) {
      const cellValue = singleRow[r].toString();
      if (cellValue.endsWith("-")) {
        singleRow[r] = `-${cellValue.slice(0, -1)}`;
      }
    }

    return singleRow;
  });
}

function setValues(sheet, data) {
  var numRows = data.length;
  var numColumns = data[0].length;
  sheet.getRange(1, 1, numRows, numColumns).setValues(data);
}

function formatArcSheet(sheet) {
  const lastRow = sheet.getLastRow();

  // Delete extra columns and rows
  deleteExtraRowsAndColumns(sheet)

  // Format cells
  sheet.getRange(1, 1, lastRow, 9)
    .clearFormat()
    .setBackground('#b4a7d6')
    .setHorizontalAlignment("center")
    .setFontFamily("Arial")
    .setFontSize(10)
    .setNumberFormat("@");

  // Set number format for columns 2, 3, and 4
  sheet.getRange(1, 2, lastRow, 4)
    .setNumberFormat('#,##0.00'); // Adjust the format as needed

  // Add an additional row at the end of the sheet
  sheet.insertRowAfter(lastRow);
}



/**
 * Adds payment type column
 *
 * 
 */
function addPaymentTypeColumn(sheet) {
  const data = sheet.getDataRange().getValues();
  const updatedArray = [];

  for (var i = 1; i < data.length; i++) {

    let paymentType = data[i][6];
    const creditCardArray = ["AX", "CA", "VI", "DS", "TP"];
    if (creditCardArray.includes(data[i][6])) { paymentType = "CREDIT CARD" }
    if (data[i][6] === "") { paymentType = "CASH" }

    updatedArray.push([...data[i], paymentType])
  }

  sheet.clearContents();
  sheet.getRange(1, 1, updatedArray.length, updatedArray[0].length).setValues(updatedArray);
}