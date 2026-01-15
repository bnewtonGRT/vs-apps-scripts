function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('GRT')
    .addItem('Reconcile ARC/Trams', 'reconcileARCTRAMS')
    .addItem('ARC Auto-Formatting', 'prepareARCsheet')
    .addItem('Trams Auto-Formatting', 'tramsFormat')
    .addItem('Delete Extra Rows/Columns', 'cleanUpSheet')
    .addToUi();
}

/** @OnlyCurrentDoc */

/**
 * Removes duplicate rows from the current sheet.
 */

function reconcileARCTRAMS() {

  //const sheet = SpreadsheetApp.getActiveSheet();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const tramsImport = spreadsheet.getSheetByName("Trams Import");
  const arcImport = spreadsheet.getSheetByName("ARC Import");

  const tramsData = tramsImport.getDataRange().getValues();
  const arcData = arcImport.getDataRange().getValues();
  const dataArray = [...arcData, ...tramsData];

  console.log(dataArray.length)

  let ticketObject = {
    dataArray: dataArray,
    newDataArray: [],
    dupDataArray: [],
    voidedTicketHopperArray: [],
    voidedTicketFinalArray: [],
    finalDataArray: [],
    reconciledDataArray: []
  }

  //This function compares each ticket in the dataArray to each ticket in the newDataArray and marks them as as duplicate (true or false).
  ticketObject = reconcileTickets(ticketObject, "firstPhase");

  //This next for loop now compares each ticket in the newDataArray to each ticket in the dupDataArray and marks them as as duplicate (true or false).
  //This will remove the second half of each duplicate "pair."
  //This will also remove any voided ARC tickets that do not show as non-voided in ARC
  ticketObject = reconcileTickets(ticketObject, "secondPhase");

  //finally, this clears the sheet and pastes in the tickets that need to be manually reconciled.
  const reconciliation = spreadsheet.getSheetByName("Reconciliation");
  reconciliation.clearContents();
  reconciliation.getRange(1, 1, ticketObject.finalDataArray.length, ticketObject.finalDataArray[0].length).setValues(ticketObject.finalDataArray);

  displayArrays(ticketObject);


  formatSheet();

}

function reconcileTickets(ticketObject, phase) {

  const { firstArray, secondArray } = assignArrays(ticketObject, phase);

  /*this iterates over each row, labels as duplicate (true or false)*/
  for (let i = 0; i < firstArray.length; i++) {

    let firstArrayRow = firstArray[i];
    let secondArrayRow = [];
    const firstRowObject = createRowObject(firstArrayRow);

    let duplicate = false;

    for (let j = 0; j < secondArray.length; j++) {

      secondArrayRow = secondArray[j];

      const secondRowObject = createRowObject(secondArrayRow);

      const isDuplicate = reconcileRow(firstRowObject, secondRowObject)

      duplicate = isDuplicate ? isDuplicate : duplicate;

      /*
      if (phase === "secondPhase") {
        console.log(`${firstArrayRow[0]} ==? ${secondArrayRow[0]} : ${duplicate}`);
      }
      */

    }

    if (phase === "firstPhase") {
      firstPhase(ticketObject, firstArrayRow, duplicate)
    }

    if (phase === "secondPhase") {
      secondPhase(ticketObject, firstArrayRow, secondArrayRow, duplicate)
    }
  }
  return ticketObject;
}

function assignArrays(ticketObject, phase) {

  let firstArray
  let secondArray

  const { dataArray, newDataArray, dupDataArray } = ticketObject;

  switch (phase) {
    case 'firstPhase':
      firstArray = dataArray;
      secondArray = newDataArray;
      break;
    case 'secondPhase':
      firstArray = newDataArray;
      secondArray = dupDataArray;
      break;
    default:
      console.log(`Error: No phase set.`);
  }

  return { firstArray, secondArray }
}

function reconcileRow(firstRow, secondRow) {

  let duplicate = false
  /*
    //Checks if payment type matches
    if (
      firstRow.ticketNumber == secondRow.ticketNumber &&
      firstRow.paymentType == secondRow.paymentType
    ) {
      duplicate = true;
    }
    */


  //Checks if ticket number, total fare, commission, NET FARE, and payment type
  //These are the normal tickets, that match exactly in ARC and Trams
  if (
    firstRow.ticketNumber == secondRow.ticketNumber &&
    firstRow.totalFare == secondRow.totalFare &&
    firstRow.commission == secondRow.commission &&
    firstRow.netRemit == secondRow.netRemit &&
    firstRow.paymentType == secondRow.paymentType
  ) {
    duplicate = true;
  }

  //Checks if ticket number, total fare, commission, and payment type are the same, but net remit is different
  //These are tickets that have different forms of payment in ARC and Trams
  //(usually tickets issued on a GRT CC)

  if (
    firstRow.ticketNumber == secondRow.ticketNumber &&
    firstRow.totalFare == secondRow.totalFare &&
    firstRow.commission == secondRow.commission &&
    firstRow.paymentType == secondRow.paymentType &&
    firstRow.netRemit != secondRow.netRemit &&
    (firstRow.group === "CCGRT" || secondRow.group === "CCGRT")
  ) {
    duplicate = true;
  }


  //Checks if ticket number, net remit, and payment type match, but commission and total fare are different.
  //This is to reconcile net fare check payment tickets (usually marked up with RM*FV lines).
  if (
    firstRow.ticketNumber == secondRow.ticketNumber &&
    firstRow.totalFare != secondRow.totalFare &&
    firstRow.commission != secondRow.commission &&
    firstRow.netRemit == secondRow.netRemit &&
    firstRow.paymentType == secondRow.paymentType
  ) {
    duplicate = true;
  }


  /*
    //Checks if ticket number and net remit match, but commission and total fare are different.
    //This is to reconcile tickets issued on a GRT CC and marked up with RM*FV
    if (
      firstRow.ticketNumber == secondRow.ticketNumber &&
      ((firstRow.totalFare == secondRow.netRemit) || (firstRow.netRemit == secondRow.totalFare)) &&
      ((firstRow.group == "CCGRT") || (secondRow.group == "CCGRT"))
  
    ) {
      duplicate = true;
    }
  */


  /*
    //Checks if ticket number and net remit match, but commission and total fare are different.
    //This is to reconcile tickets issued on a GRT CC and marked up with RM*FV
    if (
      firstRow.ticketNumber == secondRow.ticketNumber &&
      (firstRow.totalFare - firstRow.commission == secondRow.totalFare - secondRow.commission) &&
      ((firstRow.group == "CCGRT") || (secondRow.group == "CCGRT")) &&
      firstRow.paymentType == secondRow.paymentType
  
    ) {
      duplicate = true;
    }
    */


  //Checks if total fare is zero dollars and commissions match; if so, then payment type is not important; elso checks if ticket is in error
  if (
    firstRow.ticketNumber == secondRow.ticketNumber &&
    Number(firstRow.totalFare) === 0 &&
    Number(secondRow.totalFare) === 0 &&
    Number(secondRow.commission) === Number(firstRow.commission) &&
    (firstRow.status !== "E" && secondRow.status !== "E")
  ) {
    duplicate = true;
  }


  //Checks if total fare is greater than or equal to zero but commission is less than zero; if so, then this should not be marked as reconciled
  if (
    firstRow.ticketNumber == secondRow.ticketNumber &&
    Number(firstRow.totalFare) >= 0 &&
    Number(secondRow.totalFare) >= 0 &&
    Number(secondRow.commission) < 0 &&
    Number(firstRow.commission) < 0
  ) {
    duplicate = false;
  }

  return duplicate;

}

function createRowObject(row) {

  const rowObject = {
    ticketNumber: row[0],
    totalFare: row[1],
    commission: row[2],
    netRemit: row[3],
    source: row[4],
    group: row[5],
    status: row[7],
    paymentType: row[8]

  };

  return rowObject;
}

function firstPhase(ticketObject, row, duplicate) {

  const source = row[4];

  const {
    voidedTicketHopperArray,
    newDataArray,
    dupDataArray
  } = ticketObject;

  if (!duplicate && source == "TRAMS") {
    voidedTicketHopperArray.push(row[0]);
  }

  //If a ticket is marked as duplicate, it gets pushed into the dupDataArray.
  //If a tikcet is NOT marked as duplicate, it gets pushed into the newDataArray.
  //The newDataArray will then contain all unique tickets; but only half of each duplicate "pair" has been removed. We must run a second loop to remove the seond half of each duplicate "pair."
  if (!duplicate) {
    newDataArray.push(row);
  } else {
    dupDataArray.push(row);
  }

}

function secondPhase(ticketObject, firstRow, secondRow, duplicate) {

  const ticketNumber = firstRow[0];
  const source = firstRow[4];
  const status = firstRow[7];

  const {
    voidedTicketHopperArray,
    voidedTicketFinalArray,
    finalDataArray,
    reconciledDataArray,
    dupDataArray
  } = ticketObject;

  //This checks whether the current ticket number is included in the voidedTicketHopperArray and returns a boolean (true or false)
  let isInVoidedHopper = voidedTicketHopperArray.includes(ticketNumber);

  //If the current ticket is not marked as "duplicate" (aka Reconciled) and if the ticket 1) is ARC and marked VOIDED and 2) Does not have a TRAMS counterpart, then mark the ticket as reconciled.
  if (!duplicate) {
    if (!isInVoidedHopper && source == "ARC" && status == "V") {
      //console.log(ticketNumber + " is voided: " + isInVoidedHopper)
      voidedTicketFinalArray.push(firstRow);
      duplicate = true;
    }
  }

  //If a ticket is marked as duplicate, this means that the ARC ticket and the TRAMS ticket match and do not need to be manually reconciled, it is pushed into the "reconciledDataArray".
  //If a tikcet is NOT marked as duplicate, it means that ARC and TRAMS do NOT match and the tickets need to be manually reconciled. These tickets get pushed into the finalDayaArray.
  if (!duplicate) {
    finalDataArray.push(firstRow);
  } else {
    reconciledDataArray.push(firstRow, secondRow);
  }

}

function formatSheet() {

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  //const sheet = ss.getActiveSheet();
  const sheet = spreadsheet.getSheetByName("Reconciliation");
  const dataRange = sheet.getDataRange();
  const finalData = dataRange.getValues();
  const lastRow = dataRange.getLastRow();
  const lastColumn = dataRange.getLastColumn();

  //set background to purple
  sheet.getRange(1, 1, lastRow, lastColumn)
    .clearFormat()
    .setBackground('#b4a7d6')
    .setHorizontalAlignment("center")
    .setFontFamily("Arial")
    .setFontSize(10);


  //Hight XD charges, debit memos, MCOs, etc. according to memo form code ranges
  formCodeId(finalData, sheet, lastColumn);

  //Set TRAMS background to yellow
  for (let i = 1; i < finalData.length; i++) {
    const row = finalData[i];
    if (row[4] == "TRAMS") {
      sheet.getRange([i + 1], 1, 1, lastColumn).setBackground('#fff2cc');
    };
  };

  //Sort by ticket number
  const sortRange = sheet.getRange(1, 1, lastRow, lastColumn);
  sortRange.sort(1);

  deleteExtraRowsAndColumns(sheet);

}

function formCodeId(finalData, sheet, lastColumn) {

  //Set TRAMS background to Red
  for (let i = 0; i < finalData.length; i++) {

    const row = finalData[i];

    const formCode = Number(row[0].toString().padStart(10, '0').slice(0, 4));
    const ticketType = row[5];

    if (ticketType === "RF") {
      sheet.getRange([i + 1], 1, 1, lastColumn).setBackground('#c27ba0');
    }

    if (formCode >= 500 && formCode <= 999) {
      sheet.getRange([i + 1], 1, 1, lastColumn).setBackground('#ADD8E6');
      sheet.getRange([i + 1], lastColumn).setValue('XD Charge');
      sheet.getRange([i + 1], 1).setNumberFormat('@').setValue(row[0].toString().padStart(10, '0'));
    }

    if (formCode > 8300 && formCode <= 8959) {
      sheet.getRange([i + 1], 1, 1, lastColumn).setBackground('#ADD8E6');
      sheet.getRange([i + 1], lastColumn).setValue('EMD');
    }

    if (formCode >= 8960 && formCode <= 8969) {
      sheet.getRange([i + 1], 1, 1, lastColumn).setBackground('#FFA500');
      sheet.getRange([i + 1], lastColumn).setValue('Debit Memo');
    }

    if (formCode >= 8970 && formCode <= 8976) {
      sheet.getRange([i + 1], 1, 1, lastColumn).setBackground('#FFA500');
      sheet.getRange([i + 1], lastColumn).setValue('Credit Memo');
    }

    if (formCode >= 8977 && formCode <= 8979) {
      sheet.getRange([i + 1], 1, 1, lastColumn).setBackground('#FFA500');
      sheet.getRange([i + 1], lastColumn).setValue('Automated Agent Deduction');
    }

    if (formCode >= 8980 && formCode <= 8989) {
      sheet.getRange([i + 1], 1, 1, lastColumn).setBackground('#FFA500');
      sheet.getRange([i + 1], lastColumn).setValue('Commission Recall');
    }

    if (formCode >= 8990 && formCode <= 8999) {
      sheet.getRange([i + 1], 1, 1, lastColumn).setBackground('#FFA500');
      sheet.getRange([i + 1], lastColumn).setValue('MCO');
    }

  }

}

function displayArrays(ticketObject) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  for (const [key, value] of Object.entries(ticketObject)) {
    pasteArrayIntoSheet(spreadsheet, key, value)
  }



};

function pasteArrayIntoSheet(spreadsheet, arrayName, ticketArray) {

  if (arrayName === "voidedTicketHopperArray") {
    ticketArray = voidedTicketHopperHelper(ticketArray)
  }

  console.log(`${arrayName}: ${ticketArray.length}, ${ticketArray[0] ? ticketArray[0].length : 0}`)

  const sheet = spreadsheet.getSheetByName(arrayName);
  sheet.clearContents();

  if (ticketArray.length > 0) {
    sheet.getRange(1, 1, ticketArray.length, ticketArray[0].length).setValues(ticketArray);
  }

}

function voidedTicketHopperHelper(array) {

  const twoDimensionalArray = array.map((x) => [x]);

  return twoDimensionalArray;

}

function cleanUpSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  deleteExtraRowsAndColumns(sheet)
}


function deleteExtraRowsAndColumns(sheet) {

  const dataRange = sheet.getDataRange();
  const lastRow = dataRange.getLastRow();
  const lastColumn = dataRange.getLastColumn();
  const maxRows = sheet.getMaxRows();
  const maxColumns = sheet.getMaxColumns();

  // Delete extra rows
  if (maxRows > lastRow) {
    sheet.deleteRows(lastRow + 1, maxRows - lastRow);
  }

  // Delete extra columns
  if (maxColumns > lastColumn) {
    sheet.deleteColumns(lastColumn + 1, maxColumns - lastColumn);
  }
}