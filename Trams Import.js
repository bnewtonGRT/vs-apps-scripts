function tramsFormat() {
  const sheet = SpreadsheetApp.getActiveSheet();

  deleteExtraColumns(sheet);

  addSourceInfo(sheet);

  convertNetDueToNetRemit(sheet);

  addPaymentTypeColumnTrams(sheet);

  formatTramsSheet(sheet);
}

/**
 * Deletes extra columns and rows beyond a specific range in a Google Sheets document.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The target sheet where columns and rows will be deleted.
 */
function deleteExtraColumns(sheet) {
  const lastColumn = sheet.getMaxColumns();
  sheet.deleteColumns(8, lastColumn - 7);
}

/**
 * Adds source information ("TRAMS") to a specific range in a Google Sheets document.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The target sheet where source information will be added.
 */
function addSourceInfo(sheet) {
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(2, 5, lastRow - 1, 1);
  range.setValue("TRAMS");
}

/**
 * Converts the "Net Due" column values to "Net Remit" by negating them in a Google Sheets document.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The target sheet where the conversion will be performed.
 */
function convertNetDueToNetRemit(sheet) {
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    data[i][3] = -data[i][3];
  }

  sheet.clearContents();
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
}

/**
 * Formats a specific range of cells in a Google Sheets document.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The target sheet where cell formatting will be applied.
 */
function formatTramsSheet(sheet) {
  const dataRange = sheet.getDataRange();
  const lastRow = dataRange.getLastRow();
  const lastColumn = dataRange.getLastColumn();

  sheet.getRange(1, 1, lastRow, lastColumn)
    .clearFormat()
    .setBackground('#fff2cc')
    .setHorizontalAlignment("center")
    .setFontFamily("Arial")
    .setFontSize(10)
    .setNumberFormat("@");

  // Set number format for columns 2, 3, and 4
  sheet.getRange(1, 2, lastRow, 4)
    .setNumberFormat('#,##0.00'); // Adjust the format as needed

}

/**
 * Adds payment type column
 *
 * 
 */
function addPaymentTypeColumnTrams(sheet) {
  const data = sheet.getDataRange().getValues();
  const updatedArray = [];

  for (let i = 1; i < data.length; i++) {

    const tramsPayment = data[i][6];
    let paymentType = data[i][6];
    let tramsGroup = data[i][5];


    const cashPaymentMethods = ["Check", "", "ACH", "Other", "EFT"];

    if (cashPaymentMethods.includes(tramsPayment)) {
      paymentType = "CASH";
    }

    if (tramsGroup === "CCGRT") { paymentType = "CREDIT CARD" }

    if (tramsPayment === "C/C" && tramsGroup !== "CCGRT") {
      paymentType = "CREDIT CARD"
    } else if (tramsPayment === "C/C" && tramsGroup === "CCGRT") { paymentType = "UNCLEARED" }

    updatedArray.push([...data[i], "", paymentType])
  }

  sheet.clear();
  sheet.getRange(1, 1, updatedArray.length, updatedArray[0].length).setValues(updatedArray);
}