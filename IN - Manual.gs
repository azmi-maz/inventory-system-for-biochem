 function manualIncomingStock() {
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const getIncomingListLastRow = ss.getSheetByName("MANUAL").getLastRow();
  const getIncomingListLastCol = ss.getSheetByName("MANUAL").getLastColumn();
  const getTblINLISTLastRow = ss.getSheetByName("IN LIST").getLastRow();
  const getTblStockINLastRow = ss.getSheetByName("tblStockIN").getLastRow();
  const getTblUniqueINIDLastRow = ss.getSheetByName("tblUniqueINID").getLastRow();

  let getIncomingList = ss.getSheetByName("MANUAL").getRange(2,1,getIncomingListLastRow-1,getIncomingListLastCol).getValues();
  let getListCodeToMasterCode = ss.getSheetByName("ItemCodeL").getRange(2,1,getListCodeLastRow-1,2).getValues();
  let getMasterList = ss.getSheetByName("MasterL").getRange(2,1,getMasterListLastRow-1,9).getValues();


  //console.log(copyToIncomingSheet);
  //console.log(getIncomingList);
  //console.log(getListCodeToMasterCode);
  //console.log(getIncomingListLastRow);
  //console.log(getIncomingListLastCol);
  
//Goes through each array for incoming items
let trimmedNameList = [];
  for (var i = 0; i < getIncomingList.length; i++){
//  1) Trim up
    if (getIncomingList[i][0] != '' &&
        getIncomingList[i][3] != '' ||
        getIncomingList[i][2] != '' ||
        getIncomingList[i][3] != ''){
    firstValue = getIncomingList[i][0].toString();
    cutFirstValue = firstValue.trim();
    getQuantity = getIncomingList[i][3];
// Add in the [] to the cutFirstValue to make it into proper array
    trimmedNameList.push([cutFirstValue, getQuantity]);

  }
}
//console.log(getIncomingList);
//console.log(trimmedNameList);

// 3) Transaction id

// 4.1) Timestamp
let stampUpAllItemsList = [];

    for (var n = 0; n < getIncomingList.length; n++){
    if (getIncomingList[n][0] != '' &&
        getIncomingList[n][3] != '' ||
        getIncomingList[n][2] != '' ||
        getIncomingList[n][3] != ''){
    
      const today = new Date();

      const dateEachItemDay = today.getDate();
      const dateEachItemMonth = today.getMonth()+1;
      const dateEachItemYear = today.getFullYear();
      const dateEachItemHours = today.getHours();
      const dateEachItemMinutes = today.getMinutes();
      const dateEachItemSeconds = today.getSeconds();

      let time = `${dateEachItemMonth}/${dateEachItemDay}/${dateEachItemYear} ${dateEachItemHours}:${dateEachItemMinutes}:${dateEachItemSeconds}`;

      stampUpAllItemsList.push(time);
      }
    }
//console.log(stampUpAllItemsList);

// 4.2) Blank column
  let blankUpCol = [];

    for (var s = 0; s < getIncomingList.length; s++){
    if (getIncomingList[s][0] != '' &&
        getIncomingList[s][3] != '' ||
        getIncomingList[s][2] != '' ||
        getIncomingList[s][3] != ''){
      const blankValue = 'MANUAL';

      blankUpCol.push([blankValue]);
      }
    }
//console.log(blankUpCol);

// 5) Vesalius Code

// 6) Item Code
  let refItemCodeList = getListCodeToMasterCode;
  let matchedMasterCode = [];  

  for (var j = 0; j < trimmedNameList.length; j++){  
      for (var k = 0; k < refItemCodeList.length; k++){
          if (trimmedNameList[j][0] === refItemCodeList[k][0]){
              const matchedValue = refItemCodeList[k][1];
              //Can break k here once found?
              matchedMasterCode.push([matchedValue]);
    } 
  }
}
//console.log(matchedMasterCode);

// 7) Item Type, Location, Extract Multicounts
    let matchedItemType = [];
    let matchedMultiCountsList = [];

    for (var k = 0; k < matchedMasterCode.length; k++){
        for (var l = 0; l < getMasterList.length; l++){
      
          if(matchedMasterCode[k][0] === getMasterList[l][0]){
              let matchedItemTypeValue = getMasterList[l][2];
              let matchedItemLocationValue = getMasterList[l][3];
              let matchedMultiCountsValue = getMasterList[l][7];

                matchedItemType.push([matchedItemTypeValue,matchedItemLocationValue]);
                matchedMultiCountsList.push([matchedMultiCountsValue]);

      }
    }
  }

//console.log(matchedItemType);
//console.log(matchedMultiCountsList);

// 8) Modify Lot Number
  let modifiedLotNumberList = [];
  for (var p = 0; p < getIncomingList.length; p++){
    if (getIncomingList[p][0] != '' &&
        getIncomingList[p][3] != '' ||
        getIncomingList[p][2] != '' ||
        getIncomingList[p][3] != ''){

    const upperLotNumberValue = getIncomingList[p][1];
    
    // If purely numerical, dont use upper case
    let modifiedLotNumberValue = '';
    
    if(typeof(upperLotNumberValue) === 'string'){

      modifiedLotNumberValue = "\'" + upperLotNumberValue;

    } else if (typeof(upperLotNumberValue) === 'number'){

      modifiedLotNumberValue = upperLotNumberValue;

    }
    //upperLotNumberValue.toUpperCase();
    
    if(getIncomingList[p][1] === ''){

       modifiedLotNumberValue = "'"+'BLANK';
    }
    modifiedLotNumberList.push([modifiedLotNumberValue]);
    }
  }

//console.log(modifiedLotNumberList);

// 9) Expiry date extraction
  let extractedExpDate = [];
  for (var q = 0; q < getIncomingList.length; q++){
    if (getIncomingList[q][0] != '' &&
        getIncomingList[q][3] != '' ||
        getIncomingList[q][2] != '' ||
        getIncomingList[q][3] != ''){

    let expDateValue = getIncomingList[q][2];
    
    if(getIncomingList[q][2] === ''){

       expDateValue = "'"+'BLANK';
    }
    extractedExpDate.push([expDateValue]);
    }
  }
//console.log(extractedExpDate);

// 10) Quantity
let correctedCountPerItems = [];

  for (var m = 0; m < matchedMultiCountsList.length; m++){
    
    multiplyMulticountValue = trimmedNameList[m][1] * matchedMultiCountsList[m][0];
    correctedCountPerItems.push([multiplyMulticountValue]);
    
  }
//console.log(correctedCountPerItems);

// QC Code

//  11) Collect all arrays and put into one

    // Column One - Date and Time
    let col2Timestamp = stampUpAllItemsList;
    // Column Two - Just Blanks
    let colBlank = blankUpCol;
    // Column Two - Item name
    let col3ItemName = trimmedNameList;
    // Column Three - Master code
    let col4ItemCode = matchedMasterCode;
    // Column Four/Five/Six - Type and location
    let col5ItemType = matchedItemType;
    // Column Seven - Lot number
    let col7LotNumber = modifiedLotNumberList;
    // Column Eight - Expiry date
    let col8ExpDate = extractedExpDate;
    // Column Nine - Quantity in the smallest UOM
    let col10Quantity = correctedCountPerItems;
    //let col11QRcode = 
    

    let resArrayListForManualIncomingStock = [];
for (r = 0; r < col2Timestamp.length; r++) {
  resArrayListForManualIncomingStock.push(appendArrays(
    col2Timestamp[r],       // Timestamp
    colBlank[0],            // MANUAL stamp
    col4ItemCode[r][0],     // Item Code
    col3ItemName[r][0],     // Item Name
    col5ItemType[r][0],     // Item Type
    col5ItemType[r][1],     // Location
    col7LotNumber[r][0],    // Lot Number
    col8ExpDate[r][0],      // Exp Date
    col10Quantity[r][0]     // Quantity in smallest UOM
    ));
}

//console.log(resArrayListForManualIncomingStock);

// Expand each row based on multicount to make each transaction unique
// Autogenerate id for each row of transaction

//console.log(resArrayListForManualIncomingStock);
    let addNewRowsPerCount = col10Quantity;
    //let expandRows = [];
    let tblUniqueINIDList = [];
    let tblStockINList = [];
    let newLengthCount = [];
    let countEachRow = 1;
    let fromIncomingSheet = 'MANUAL';

    for (var y = 0; y < addNewRowsPerCount.length; y++){
      newLengthCount.push([addNewRowsPerCount[y][0]]);

        for (z = 0; z < newLengthCount[y]; z++){

        let resID = createUniqueID(resArrayListForManualIncomingStock,fromIncomingSheet);

          tblUniqueINIDList.push([
                           resID,                                         // Transaction ID
                           resArrayListForManualIncomingStock[y][0],      // Timestamp
                           resArrayListForManualIncomingStock[y][2],      // Item code
                           resArrayListForManualIncomingStock[y][6],      // Lot Number
                           resArrayListForManualIncomingStock[y][7]       // Exp date
                           ]);

          tblStockINList.push([
                           resArrayListForManualIncomingStock[y][0],      // Timestamp
                           resID,                                         // Transaction ID
                           countEachRow                                   // Count as one
                           ]);
        }
    }
    

    //console.log(tblUniqueINIDList);
    //console.log(resArrayListForManualIncomingStock);

    // For tblUniqueINID - to make a list of all unique IDs
    ss.getSheetByName("tblUniqueINID").getRange(getTblUniqueINIDLastRow+1,1,tblUniqueINIDList.length,tblUniqueINIDList[0].length).setValues(tblUniqueINIDList);
    
    // For tblStockIN - for incoming counter
    ss.getSheetByName("tblStockIN").getRange(getTblStockINLastRow+1,1,tblStockINList.length,tblStockINList[0].length).setValues(tblStockINList);

    // For visual display only/ or check for errors - IN LIST
    ss.getSheetByName("IN LIST").getRange(getTblINLISTLastRow+1,1,resArrayListForManualIncomingStock.length,resArrayListForManualIncomingStock[0].length).setValues(resArrayListForManualIncomingStock);

    // Clear off MANUAL Sheet the lot numbers, exp date and quantity
    ss.getSheetByName("MANUAL").getRange(2,2,getIncomingListLastRow-1,3).clearContent();


    updateQOHList();
    updateOUTGOINGpaste();
  


}







