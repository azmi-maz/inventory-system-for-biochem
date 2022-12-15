function incomingStock() {
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const getIncomingListLastRow = ss.getSheetByName("INCOMING").getLastRow();
  const getIncomingListLastCol = ss.getSheetByName("INCOMING").getLastColumn();
  const getTblINLISTLastRow = ss.getSheetByName("IN LIST").getLastRow();
  const getTblStockINLastRow = ss.getSheetByName("tblStockIN").getLastRow();
  const getTblUniqueINIDLastRow = ss.getSheetByName("tblUniqueINID").getLastRow();
  const getTblVerificationLastRow = ss.getSheetByName("Verification").getLastRow();


  let getIncomingList = ss.getSheetByName("INCOMING").getRange(1,1,getIncomingListLastRow,getIncomingListLastCol).getValues();
  let getListCodeToMasterCode = ss.getSheetByName("ItemCodeL").getRange(2,1,getListCodeLastRow-1,2).getValues();
  let getMasterList = ss.getSheetByName("MasterL").getRange(2,1,getMasterListLastRow-1,9).getValues();


  let getFullVerificationArray = ss.getSheetByName("Verification").getRange(2,1,getTblVerificationLastRow-1,12).getValues();

  //console.log(copyToIncomingSheet);
  //console.log(getIncomingList);
  //console.log(getListCodeToMasterCode);
  
//Goes through each array for incoming items
let cleanedUpBarcodeList = [];
  for (var i = 0; i < getIncomingList.length; i++){
//  1) Trim up OR remove blank spaces AND Substitute char(45) = '-' AND Substitute char(29) = ''
    const firstValue = getIncomingList[i][0].toString();
    const cutFirstValue = firstValue.replace(/\s+/g,"").replace("-","").replace("","").trim();
// Add in the [] to the cutFirstValue to make it into proper array
    cleanedUpBarcodeList.push([cutFirstValue]);

  }
//console.log(getIncomingList);
//console.log(cleanedUpBarcodeList);

//  2.1) Convert List Code
  let extractedListCode = [];
  let lookForFinalListCode = '';
  let finalListCode = '';
    for (var i = 0; i < cleanedUpBarcodeList.length; i++){

    const catchListCode = cleanedUpBarcodeList[i][0];
    const extractListCodeLength = catchListCode.length;

// For CSF variables
    const find240InCSF = catchListCode.indexOf(240);
    const countCF10InString = occurrences(catchListCode,'CF10');

// For 1P0603 - transplant tubes
    const lookFor1P0603 = catchListCode.substring(catchListCode.length-6);

// For BNP Cal
    const lookForBNP = catchListCode.substring(catchListCode.length-3);

// For 9P4240 - TACRO WBT PPT
    const lookFor9P4240 = catchListCode.substring(catchListCode.length-6);

// For CG8+ Cartridges
    const lookForCg8Cart = catchListCode.substring(9,16);

// For UROMETER strips
    const lookForUrometer = catchListCode.substring(26,30).toUpperCase();
    //console.log(lookForUrometer)

// For Biorad Control Kit
    const lookForBiorad = catchListCode.substring(2,16);

// For Techno WBT
    const lookForTechnoWBT = catchListCode.substring(20,26);

// Check SIEMENS items if they are scannable

// SPACE FOR MORE VARIABLES TO HANDLE DIFFERENT BARCODE TYPES

// Else look for 240 with 1/2/3 occurrences, extract 6 length
    const count240InString = occurrences(catchListCode,240);
    const find240WithOneOccurrences = catchListCode.indexOf(240)+4;
    const find240WithTwoOccurrences = catchListCode.indexOf(240,find240WithOneOccurrences+1)+4;
    const find240WithThreeOccurrences = catchListCode.indexOf(240,find240WithTwoOccurrences+1)+4;


// 1st condition to handle first occurrence of 240 & NOT CSF, transplant tube, BNP Cal, and TACRO WBT PPT 
    if (count240InString === 1 && countCF10InString === 0 && lookFor1P0603 != '1P0603' && lookForBNP != 'BNP' && lookFor9P4240 != '9P4240' && lookForUrometer != 'UC11'){
      
      finalListCode = catchListCode.substring(find240WithOneOccurrences,find240WithOneOccurrences+6).toUpperCase();

// 2nd condition to handle CSF only
    } else if (countCF10InString > 0 && extractListCodeLength === 43){
      
      finalListCode = catchListCode.substring(find240InCSF+3,find240InCSF+9).toUpperCase();

// 3rd condition to handle transplant tubes
    } else if (lookFor1P0603 === '1P0603'){
      
      finalListCode = '1P0603';

// 4th condition to handle BNP Cal
    } else if (lookForBNP === 'BNP'){

      finalListCode = catchListCode.substring(39, 45).toUpperCase();

// 5th condition to handle TACRO WBT PPT
    } else if (lookFor9P4240 === '9P4240'){

      finalListCode = '9P4240';

// 5.1th condition to handle Techno WBT items with 2 occurrences of 240
    } else if (lookForTechnoWBT === '4S1610' & count240InString === 2){

                finalListCode = lookForTechnoWBT.toUpperCase();

// 6th condition to handle second occurrences of 240
    } else if (count240InString === 2){

      finalListCode = catchListCode.substring(find240WithTwoOccurrences,find240WithTwoOccurrences+6).toUpperCase();

// 7th condition to handle third occurrences of 240
    } else if (count240InString === 3){

      finalListCode = catchListCode.substring(find240WithThreeOccurrences).toUpperCase();

// 8th condition to handle CG8+ items
    } else if (lookForCg8Cart === '9000166' ||
               lookForCg8Cart === '9000463' ||
               lookForCg8Cart === '9000647' ||
               lookForCg8Cart === '9000661'){

                finalListCode = lookForCg8Cart;

// 9th condition to handle Urometer items
// NEED TO ADD THE MMD UROMETER HERE
    } else if (lookForUrometer === 'U05K' ||
               lookForUrometer === 'UC11' ||
               lookForUrometer === 'U100'){
                 
                finalListCode = lookForUrometer.toUpperCase();

// 10th condition to handle Biorad control items
    } else if (lookForBiorad === '06950996001649'){

                finalListCode = lookForBiorad.toUpperCase();

    }


  let checkWithItemCodeList = getListCodeToMasterCode;
  let a = false;

    for (var t = 0; t < checkWithItemCodeList.length; t++){
        if (finalListCode === checkWithItemCodeList[t][0]){

          lookForFinalListCode = finalListCode;
  let a = finalListCode === checkWithItemCodeList[t][0];
              if (a === true){
                  break;

               }
        } else {

          lookForFinalListCode = 'ERROR';

        }
    }
      extractedListCode.push([lookForFinalListCode]);
}

//console.log(finalListCode);
//console.log(extractedListCode);

//  2.2) Convert List Code to Master Code
  let refItemCodeList = getListCodeToMasterCode;
  let matchedMasterCode = [];  

  for (var z = 0; z < extractedListCode.length; z++){  
      for (var j = 0; j < refItemCodeList.length; j++){
          if (extractedListCode[z][0] === refItemCodeList[j][0]){
              const matchedValue = refItemCodeList[j][1];
      
              matchedMasterCode.push([matchedValue]);
    } 
  }
}
//console.log(extractedListCode);
//console.log(matchedMasterCode);

//  3) Item Name, Item Type, Location, Lot Number length to extract, Extract Multicounts
    let matchedItemName = [];
    let matchedLotNumberLength = [];
    let matchedMultiCountsList = [];

    for (var k = 0; k < matchedMasterCode.length; k++){
        for (var l = 0; l < getMasterList.length; l++){
      
          if(matchedMasterCode[k][0] === getMasterList[l][0]){
              let matchedItemNameValue = getMasterList[l][1];
              let matchedItemTypeValue = getMasterList[l][2];
              let matchedItemLocationValue = getMasterList[l][3];
              let matchedLotNumberLengthValue = getMasterList[l][6];
              let matchedMultiCountsValue = getMasterList[l][7];

              
                matchedItemName.push([matchedItemNameValue,matchedItemTypeValue,matchedItemLocationValue]);
                matchedLotNumberLength.push([matchedLotNumberLengthValue]);
                matchedMultiCountsList.push([matchedMultiCountsValue]);

      }
    }
  }

//console.log(matchedItemName);
//console.log(matchedLotNumberLength);
//console.log(matchedMultiCountsList);

//  4) Find Lot Number Length and Expiry Date in DateValue
    let resLotNumberList = [];
    let resExpDateList = [];
    let cutLengthForLotNumber = '';
    let toUpperCaseCutLengthForLotNumber = '';

    for (var m = 0; m < cleanedUpBarcodeList.length; m++){

/* Need to handle Lot Number extraction specifically for
A152-A153 = TECHNO CSF
A149-A151 = TECHNO CC
A154/157/156 = TECHNO IA/WBT/U
*/
    const lengthOfEachBarcode = cleanedUpBarcodeList[m][0].length;
    if (matchedMasterCode[m][0] === 'A152' || 
        matchedMasterCode[m][0] === 'A153' || 
        matchedMasterCode[m][0] === 'A149' || 
        matchedMasterCode[m][0] === 'A150' || 
        matchedMasterCode[m][0] === 'A151' || 
        matchedMasterCode[m][0] === 'A154' || 
        matchedMasterCode[m][0] === 'A157' || 
        matchedMasterCode[m][0] === 'A156'){

      cutLengthForLotNumber = "\'" + cleanedUpBarcodeList[m][0].substring(lengthOfEachBarcode-matchedLotNumberLength[m][0]);
      toUpperCaseCutLengthForLotNumber = cutLengthForLotNumber.toUpperCase();

    } else {    

// Note "\'" was put in front of Lot Number so that it paste correctly as text format
      
      cutLengthForLotNumber = "\'" + cleanedUpBarcodeList[m][0].substring(26,26+matchedLotNumberLength[m][0]);
      toUpperCaseCutLengthForLotNumber = cutLengthForLotNumber.toUpperCase();

    }
/* To extract expiry dates at different length for specific items
A152,A153 - TECHNO CSF - extract at 27
A148 - TECHNO A1C - extract at 18
A157,A151,A150,A149,A154,A156 - WBT,CC,IA & U - extract at 28
A031 - CAL CNTL CAP - exp date is empty
For other items, extract at 18
*/
    let lengthForExpDateToStart = 0;
    if (matchedMasterCode[m][0] === 'A152' ||
        matchedMasterCode[m][0] === 'A153'){
        
                lengthForExpDateToStart = 27;
    
  } else if (matchedMasterCode[m][0] === 'A148'){

                lengthForExpDateToStart = 18;

  } else if ( matchedMasterCode[m][0] === 'A157' ||
              matchedMasterCode[m][0] === 'A151' ||
              matchedMasterCode[m][0] === 'A150' ||
              matchedMasterCode[m][0] === 'A149' ||
              matchedMasterCode[m][0] === 'A154' ||
              matchedMasterCode[m][0] === 'A156'){

                lengthForExpDateToStart = 28;
  } else {

                lengthForExpDateToStart = 18;

  }
                  
      const cutLengthForYear = cleanedUpBarcodeList[m][0].substring(lengthForExpDateToStart,lengthForExpDateToStart+6).substring(0,2);
      const cutLengthForMonth = cleanedUpBarcodeList[m][0].substring(lengthForExpDateToStart,lengthForExpDateToStart+6).substring(2,4);
      const cutLengthForDay = cleanedUpBarcodeList[m][0].substring(lengthForExpDateToStart,lengthForExpDateToStart+6).substring(4,7);
      
      // Try to keep this in mm/dd/yyyy format, then column format stays as datevalue?
      // To check if can convert to datevalue here (note: valueOf or milliseconds required conversion)
      const reformatForExpDate = cutLengthForMonth + "/" +cutLengthForDay + "/" + "20" + cutLengthForYear;
      let newDateValue = reformatForExpDate.toString();
   
   if (matchedMasterCode[m][0] === 'A031'){
                
                newDateValue = '';

  }

      resLotNumberList.push([toUpperCaseCutLengthForLotNumber]);
      resExpDateList.push([newDateValue]);

    }

    //console.log(cleanedUpBarcodeList);
    //console.log(resLotNumberList);
    //console.log(resExpDateList);


//  5) Quantity as Lowest UOM = Important when outgoing stock
    let correctedCountPerItems = matchedMultiCountsList;

//console.log(correctedCountPerItems);

//  6) QR code, paste as image OR generate as needed?
    //=IMAGE("https://api.qrserver.com/v1/create-qr-code/?size=150x150&data="&ENCODEURL(ENTER VALUE HERE))


//  8) Timestamp for all transactions
//  Check that these dates are pasted in the correct format

    let stampUpAllItemsList = [];

    for (var n = 0; n < cleanedUpBarcodeList.length; n++){
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
    //console.log(stampUpAllItemsList);
  
//  9) Collect all arrays and put into one
    // Column One - Date and Time
    let col2Timestamp = stampUpAllItemsList;
    // Column Two - Scanned barcode
    let col3Barcode = cleanedUpBarcodeList;
    // Column Three - Master code
    let col4ItemCode = matchedMasterCode;
    // Column Four/Five/Six - Item name, type, location
    let col5ItemName = matchedItemName;
    // Column Seven - Lot number
    let col7LotNumber = resLotNumberList;
    // Column Eight - Expiry date
    let col8ExpDate = resExpDateList;
    // Column Nine - Quantity in the smallest UOM
    let col10Quantity = correctedCountPerItems;
    

    // for array below
    let resArrayListForIncomingStock = [];
    // For tblComb1 INPUT
    let resArrayListINtblComb1 = [];

for (p = 0; p < col2Timestamp.length; p++) {
  resArrayListForIncomingStock.push(appendArrays(
    // Time stamp
    col2Timestamp[p],
    // Scanned barcode
    col3Barcode[p][0], 
    // Mastercode
    col4ItemCode[p][0], 
    // Item name
    col5ItemName[p][0], 
    // Item type
    col5ItemName[p][1], 
    // Item location
    col5ItemName[p][2], 
    // Lot number
    col7LotNumber[p][0], 
    // Expiry date
    col8ExpDate[p][0], 
    // Quantity
    col10Quantity[p][0]
    ));

}

// To display in kits
let tblIncomingDisplayList = []
let countOfEachKit = 1;
for (p = 0; p < col2Timestamp.length; p++) {
  tblIncomingDisplayList.push(appendArrays(
    // Time stamp
    col2Timestamp[p],
    // Scanned barcode
    col3Barcode[p][0], 
    // Mastercode
    col4ItemCode[p][0], 
    // Item name
    col5ItemName[p][0], 
    // Item type
    col5ItemName[p][1], 
    // Item location
    col5ItemName[p][2], 
    // Lot number
    col7LotNumber[p][0], 
    // Expiry date
    col8ExpDate[p][0], 
    // Quantity in kit
    countOfEachKit
    ));

}

//console.log(resArrayListForIncomingStock[0][9]);
//console.log(resArrayListForIncomingStock);


// Expand each row based on multicount to make each transaction unique
// Autogenerate id for each row of transaction
    let addNewRowsPerCount = col10Quantity;
    //let expandRows = [];
    let tblUniqueINIDList = [];
    let tblStockINList = [];
    let newLengthCount = [];
    let countEachRow = 1;
    let fromIncomingSheet = 'INCOMING';

    for (var y = 0; y < addNewRowsPerCount.length; y++){
      newLengthCount.push([addNewRowsPerCount[y][0]]);

        for (z = 0; z < newLengthCount[y]; z++){

        let resID = createUniqueID(resArrayListForIncomingStock,fromIncomingSheet);

          tblUniqueINIDList.push([
                           resID,                                   // Transaction ID
                           resArrayListForIncomingStock[y][0],      // Timestamp
                           resArrayListForIncomingStock[y][2],      // Item code
                           resArrayListForIncomingStock[y][6],      // Lot Number
                           resArrayListForIncomingStock[y][7]       // Exp date
                           ]);

          tblStockINList.push([
                           resArrayListForIncomingStock[y][0],      // Timestamp
                           resID,                                   // Transaction ID
                           countEachRow                             // Count as one
                           ]);
    }
  }

  // Loop for new lots from incoming and check with existing lot numbers in verification list
  let getCurrentLotNumbers = [];
  for (a = 0; a < getFullVerificationArray.length; a++){
    getCurrentLotNumbers.push("\'"+getFullVerificationArray[a][2].toUpperCase());
  }
  //console.log(getCurrentLotNumbers);
  let findUniqueLotNumbers = countArrayElem(getCurrentLotNumbers);

  let countedUniqueLots = findUniqueLotNumbers[0];

  let getNewLotsFromIncoming = [];
  for (b = 0; b < tblIncomingDisplayList.length; b++){
    getNewLotsFromIncoming.push(tblIncomingDisplayList[b][6]);
  }
  //console.log(getNewLotsFromIncoming)

  let removeSameLots = findRemainingUniqueID(getNewLotsFromIncoming.sort(),countedUniqueLots.sort());

  //console.log(removeSameLots)

  let findUniqueLotFromSameLotRemoval = countArrayElem(removeSameLots);

  let countedUniqueLotForPaste = findUniqueLotFromSameLotRemoval[0];

  //console.log(countedUniqueLotForPaste);

  // Clean up the incoming list with only reagents
  let cleanedUpIncomingForReagentsOnly = [];
  for (e = 0; e < tblIncomingDisplayList.length; e++){
    if (tblIncomingDisplayList[e][4] === "REAGENT"){
      cleanedUpIncomingForReagentsOnly.push([
              tblIncomingDisplayList[e][0],   // Date
              tblIncomingDisplayList[e][3],                                         // Item Name
              tblIncomingDisplayList[e][6],                                         // Lot Number
              tblIncomingDisplayList[e][7],   // Exp Date
              tblIncomingDisplayList[e][2],                                         // Item Code
              ]);
    }
  }
  //console.log(cleanedUpIncomingForReagentsOnly)

  let findUniqueIncomingItems = countArrayElem(cleanedUpIncomingForReagentsOnly);

  let countedCleanedUpIncoming = findUniqueIncomingItems[0];

  //console.log(countedCleanedUpIncoming[0])

  // Loop through original incoming list to expand
  let finalArrayForVerification = [];
  for (c = 0; c < countedUniqueLotForPaste.length; c++){
    for (d = 0; d < countedCleanedUpIncoming.length; d++){
      if (countedUniqueLotForPaste[c] === countedCleanedUpIncoming[d].split(",")[2]){
              
              dateVal = countedCleanedUpIncoming[d].split(",")[3];
              
              finalArrayForVerification.push([
                          countedCleanedUpIncoming[d].split(",")[0],                                 // Date
                          countedCleanedUpIncoming[d].split(",")[1],                                 // Item Name
                          countedCleanedUpIncoming[d].split(",")[2],                                 // Lot Number
                          dateVal,                                                                   // Exp Date
                          "",                                                                        // Current Lot Number
                          "",                                                                        // Exp Date of current lot
                          "",                                                                        // Date of verification
                          "",                                                                        // Performed By
                          "Balum",                                                                   // Status default
                          "",                                                                        // Remarks
                          countedCleanedUpIncoming[d].split(",")[4],                                 // Item Code
                          ""                                                                         // Available Kits 
                          ]);
      }
    }
  }
  //console.log(finalArrayForVerification)
  //console.log(tblUniqueINIDList);
  //console.log(tblIncomingDisplayList);
  //console.log(tblUniqueINIDList);
  //console.log(tblIncomingDisplayList);

    // For tblUniqueINID - to make a list of all unique IDs
    ss.getSheetByName("tblUniqueINID").getRange(getTblUniqueINIDLastRow+1,1,tblUniqueINIDList.length,tblUniqueINIDList[0].length).setValues(tblUniqueINIDList);
    
    
    // For tblStockIN - for incoming counter
    ss.getSheetByName("tblStockIN").getRange(getTblStockINLastRow+1,1,tblStockINList.length,tblStockINList[0].length).setValues(tblStockINList);


    // For visual display only/ or check for errors - IN LIST
    ss.getSheetByName("IN LIST").getRange(getTblINLISTLastRow+1,1,tblIncomingDisplayList.length,tblIncomingDisplayList[0].length).setValues(tblIncomingDisplayList);



    // To error handle if there are problem items that gives off ERROR code, break all and dont delete incoming list

    // To auto delete INCOMING list after successful run
    ss.getSheetByName("INCOMING").getRange(1,1,getIncomingListLastRow,1).clearContent();

    updateVerification();
    updateQOHList();
    updateOUTGOINGpaste();



}

