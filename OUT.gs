function outgoingStock() {

  var t1 = new Date().getTime();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const getTblUniqueINIDLastRow = ss.getSheetByName("tblUniqueINID").getLastRow();

  const getOutgoingLastRow = ss.getSheetByName("OUTGOING").getLastRow();
  const getOutgoingLastCol = ss.getSheetByName("OUTGOING").getLastColumn();
  const getTblStockOUTLastRow = ss.getSheetByName("tblStockOUT").getLastRow();
  const getOUTListLastRow = ss.getSheetByName("OUT LIST").getLastRow();
  let getTblUniqueINIDarray = ss.getSheetByName("tblUniqueINID").getRange(2, 1, getTblUniqueINIDLastRow - 1, 5).getValues();
  let getMasterList = ss.getSheetByName("MasterL").getRange(2, 1, getMasterListLastRow - 1, 9).getValues();

  let getOutgoingList = ss.getSheetByName("OUTGOING").getRange(3, 1, getOutgoingLastRow, 2).getValues();

  //console.log(getOutgoingList);

  // To collect selected items for checkout
  // UniqueID, Item name, Type, Location, Lot number, Expiry date
  let listOfChoosenItems = [];
  let listOfChoosenUniqueIDs = [];

  for (let i = 0; i < getOutgoingList.length; i++) {
    if (getOutgoingList[i][0] === true) {
      listOfChoosenItems.push(getOutgoingList[i]);
      listOfChoosenUniqueIDs.push(getOutgoingList[i][1]);

    }
  }
  //console.log(listOfChoosenItems);
  //console.log(listOfChoosenUniqueIDs);

  // Timestamp
  let stampUpAllItemsList = [];

  for (var n = 0; n < listOfChoosenItems.length; n++) {
    const today = new Date();

    const dateEachItemDay = today.getDate();
    const dateEachItemMonth = today.getMonth() + 1;
    const dateEachItemYear = today.getFullYear();
    const dateEachItemHours = today.getHours();
    const dateEachItemMinutes = today.getMinutes();
    const dateEachItemSeconds = today.getSeconds();

    let time = `${dateEachItemMonth}/${dateEachItemDay}/${dateEachItemYear} ${dateEachItemHours}:${dateEachItemMinutes}:${dateEachItemSeconds}`;

    stampUpAllItemsList.push(time);
  }
  //console.log(stampUpAllItemsList);

  // Use Unique ID to extract Item code, name, type, location, lot number, and expiry date

  let extractedItemInfoList = [];
  let matchedItemNameValue = '';
  let matchedItemTypeValue = '';
  let matchedItemLocationValue = '';
  let matchedItemCodeValue = '';
  let matchedLotNumberValue = '';
  let matchedExpDateValue = '';

  for (let r = 0; r < listOfChoosenUniqueIDs.length; r++) {
    for (let s = 0; s < getTblUniqueINIDarray.length; s++) {

      if (listOfChoosenUniqueIDs[r] === getTblUniqueINIDarray[s][0]) {

        for (let t = 0; t < getMasterList.length; t++) {
          if (getTblUniqueINIDarray[s][2] === getMasterList[t][0]) {
            matchedItemNameValue = getMasterList[t][1];
            matchedItemTypeValue = getMasterList[t][2];
            matchedItemLocationValue = getMasterList[t][3];
            break; // New
          }
        }
        matchedItemCodeValue = getTblUniqueINIDarray[s][2];
        matchedLotNumberValue = getTblUniqueINIDarray[s][3];
        matchedExpDateValue = getTblUniqueINIDarray[s][4];
        break; // New
      }
    }
    extractedItemInfoList.push([listOfChoosenUniqueIDs[r], // UniqueProdLotExpID
      matchedItemCodeValue,         // Item Code
      matchedItemNameValue,         // Item Name
      matchedItemTypeValue,         // Item Type
      matchedItemLocationValue,     // Location
      matchedLotNumberValue,        // Lot Number
      matchedExpDateValue           // Expiry Date
    ]);
  }

  // console.log('Length of array =', extractedItemInfoList.length); // 3

  let blankUpCol = [];

  for (var s = 0; s < listOfChoosenItems.length; s++) {
    const blankValue = 'MANUAL OUTGOING';

    blankUpCol.push([blankValue]);

  }


  // Quantity
  let countOutgoingItems = [];

  for (var m = 0; m < listOfChoosenItems.length; m++) {
    const countOutgoingVal = 1;


    countOutgoingItems.push([countOutgoingVal]);

  }
  //console.log(countOutgoingItems);

  //  11) Collect all arrays and put into one
  //let col1TransactionId = 

  // Column One - Date and Time
  let col1Timestamp = stampUpAllItemsList;
  // Column Two to Eight - UniqueID, code, name, type, location, lot number, exp date
  let col2FromUniqueID = extractedItemInfoList;
  // Column Nine - Quantity in the smallest UOM
  let col9Quantity = countOutgoingItems;
  //let col11QRcode = 


  // For OUT LIST
  let resArrayListForOutgoingStock = [];
  // For tblStockOUT
  let resArrayListToUpdateTblStockOut = [];

  for (r = 0; r < col1Timestamp.length; r++) {
    resArrayListForOutgoingStock.push(appendArrays(
      col1Timestamp[r],                // Timestamp
      col2FromUniqueID[r][0],          // UniqueProdLotExpID
      col2FromUniqueID[r][1],          // Item Code
      col2FromUniqueID[r][2],          // Item Name
      col2FromUniqueID[r][3],          // Item Type
      col2FromUniqueID[r][4],          // Location
      "'" + col2FromUniqueID[r][5],      // Lot Number
      col2FromUniqueID[r][6],          // Exp Date
      col9Quantity[r][0]               // Quantity out
    ));

    resArrayListToUpdateTblStockOut.push(appendArrays(
      col1Timestamp[r],                // Timestamp
      col2FromUniqueID[r][0],          // UniqueProdLotExpID
      col9Quantity[r][0]               // Quantity out
    ));

  }

  // console.log(resArrayListForOutgoingStock);

  // To paste the array to OUT LIST sheet
  ss.getSheetByName("OUT LIST").getRange(getOUTListLastRow + 1, 1, resArrayListForOutgoingStock.length, resArrayListForOutgoingStock[0].length).setValues(resArrayListForOutgoingStock);

  //console.log(resArrayListToUpdateTblStockOut);
  // To paste the array to tblStockOUT sheet
  // To auto add to existing date; Note: To add 1 so to append to new row without the deleting the previous record
  ss.getSheetByName("tblStockOUT").getRange(getTblStockOUTLastRow + 1, 1, resArrayListToUpdateTblStockOut.length, resArrayListToUpdateTblStockOut[0].length).setValues(resArrayListToUpdateTblStockOut);

  ss.getSheetByName("OUTGOING").getRange(1, 4, 1, 1).setValue('Please click the update button!');
  let getOutgoingListLastRow = ss.getSheetByName("OUTGOING").getLastRow();

  if (getOutgoingListLastRow === 1) {
  } else {
    ss.getSheetByName("OUTGOING").getRange(3, 1, getOutgoingListLastRow - 1, 7).clearContent().removeCheckboxes();
    // Remove the filter set by user
    ss.getSheetByName("OUTGOING").getRange(3, 1, getOutgoingListLastRow - 1, 7).getFilter().remove();
  }
  ss.getSheetByName("OUTGOING").getRange(3, 3, 1, 1).setValue('Empty List');
  ss.getSheetByName("OUTGOING").getRange(2, 1, 2, 7).createFilter();

  //updateQOHList();
  //updateOUTGOINGpaste();

  // var t2 = new Date().getTime();
  // var timeDiff = t2 - t2;
  // console.log(timeDiff); // 0 ms before and after adding breaks.

}
