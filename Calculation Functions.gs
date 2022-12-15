const globalSheet = SpreadsheetApp.getActiveSpreadsheet();
const getMasterListLastRow = globalSheet.getSheetByName("MasterL").getLastRow();

let remainingIDLeftForQOHList = []; // Global variable to make it accessible to updateBatchList
let finalStoreTransfer = []; // Global variable to make it for Store_Alinity
let newUpdatedQOHFullList = [] // For OUTGOING list paste

function updateOUTGOING() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const getTblStockINLastRow = ss.getSheetByName("tblStockIN").getLastRow();
    const getTblUniqueINIDLastRow = ss.getSheetByName("tblUniqueINID").getLastRow();
    const getTblStockOUTLastRow = ss.getSheetByName("tblStockOUT").getLastRow();
    let getOutgoingListLastRow = ss.getSheetByName("OUTGOING").getLastRow();

    let getTblStockINarray = ss.getSheetByName("tblStockIN").getRange(2,1,getTblStockINLastRow-1,3).getValues();
    let getTblUniqueINIDarray = ss.getSheetByName("tblUniqueINID").getRange(2,1,getTblUniqueINIDLastRow-1,5).getValues();
    let getTblStockOUTarray = [];
    if(getTblStockOUTLastRow === 1){
      getTblStockOUTarray = [];
    } else {
      getTblStockOUTarray = ss.getSheetByName("tblStockOUT").getRange(2,1,getTblStockOUTLastRow-1,3).getValues();
    }  
    let getMasterList = ss.getSheetByName("MasterL").getRange(2,1,getMasterListLastRow-1,9).getValues();

    //console.log(getTblStockINarray);
    //console.log(getTblUniqueINIDarray);

    let collectAllIncomingUniqueID = [];

    for (i = 0; i < getTblStockINarray.length; i++){
      collectAllIncomingUniqueID.push(getTblStockINarray[i][1]);
    }

    //console.log(collectAllIncomingUniqueID);
    //console.log(collectAllIncomingUniqueID.length);

    let collectAllOutgoingUniqueID = [];

    for (j = 0; j < getTblStockOUTarray.length; j++){
      collectAllOutgoingUniqueID.push(getTblStockOUTarray[j][1]);
    }

    //console.log(collectAllOutgoingUniqueID);

    // Calculate the QOH using custom function
    remainingIDLeftForQOHList = findRemainingUniqueID(collectAllIncomingUniqueID.sort(), collectAllOutgoingUniqueID.sort());
    //console.log(remainingIDLeftForQOHList[500]);
    //console.log(collectAllIncomingUniqueID);
    //console.log(collectAllOutgoingUniqueID.length);
    //console.log(collectAllIncomingUniqueID[500]);

    let newUpdatedQOHuniqueIDList = collectAllIncomingUniqueID;
    let getItemNameValue = '';
    let getItemTypeValue = '';
    let getItemLocationValue = '';

    for (let j = 0; j < newUpdatedQOHuniqueIDList.length; j++){
      for (let k = 0; k < getTblUniqueINIDarray.length; k++){
        
        if (newUpdatedQOHuniqueIDList[j] === getTblUniqueINIDarray[k][0]){
          
          for (let l = 0; l < getMasterList.length; l++){
            
            if (getTblUniqueINIDarray[k][2] === getMasterList[l][0]){
              
              getItemNameValue = getMasterList[l][1];
              getItemTypeValue = getMasterList[l][2];
              getItemLocationValue = getMasterList[l][3];
            }
          }
         
          newUpdatedQOHFullList.push([false,
                                     getTblUniqueINIDarray[k][0],   // UniqueProdLotExpID
                                     getItemNameValue,              // Item Name
                                     getItemTypeValue,              // Item Type
                                     getItemLocationValue,          // Item Location
                                     "'"+getTblUniqueINIDarray[k][3],   // Lot Number
                                     getTblUniqueINIDarray[k][4]    // Exp date
                                     ])
        }
      }
    }

    // Get a new list to update OUT_Alinity by the number of kits
    let newUpdatedQOHForStoreAlinity = [];
    let itemNameValue = '';
    let itemTypeValue = '';
    let itemLocationValue = '';
        for (let j = 0; j < newUpdatedQOHuniqueIDList.length; j++){
      for (let k = 0; k < getTblUniqueINIDarray.length; k++){
        
        if (newUpdatedQOHuniqueIDList[j] === getTblUniqueINIDarray[k][0]){
          
          for (let l = 0; l < getMasterList.length; l++){
            
            if (getTblUniqueINIDarray[k][2] === getMasterList[l][0]){
              
              itemNameValue = getMasterList[l][1];
              itemTypeValue = getMasterList[l][2];
              itemLocationValue = getMasterList[l][3];
              itemSubLocation = getMasterList[l][4];
              itemMultiCount = getMasterList[l][7];
              itemPRtype = getMasterList[l][5];
            }
          }
         
          newUpdatedQOHForStoreAlinity.push([false,
                                     itemNameValue,                                                      // Item Name
                                     itemTypeValue,                                                      // Item Type
                                     itemLocationValue,                                                  // Item Location
                                     "'"+getTblUniqueINIDarray[k][3],                                    // Lot Number
                                     new Date(getTblUniqueINIDarray[k][4]).toLocaleDateString("en-UK"),  // Exp date
                                     itemSubLocation,
                                     itemMultiCount,
                                     itemPRtype,
                                     getTblUniqueINIDarray[k][2]                                         // Item Code
                                     ])
        }
      }
    }
    //console.log(newUpdatedQOHForStoreAlinity)

    // Filter out the items based on the sublocations
    let storeAlinityTransferPending = [];

    // Transfer to Section loop
    for (a = 0; a < newUpdatedQOHForStoreAlinity.length; a++){
      if (newUpdatedQOHForStoreAlinity[a][6] === "Transfer to Section" &&
          newUpdatedQOHForStoreAlinity[a][8] === "PR"){
        storeAlinityTransferPending.push([
          newUpdatedQOHForStoreAlinity[a][0],   // False
          newUpdatedQOHForStoreAlinity[a][1],   // Item Name
          newUpdatedQOHForStoreAlinity[a][2],   // Item Type
          newUpdatedQOHForStoreAlinity[a][3],   // Item Location
          newUpdatedQOHForStoreAlinity[a][4],   // Lot Number
          newUpdatedQOHForStoreAlinity[a][5],   // Exp Date
          newUpdatedQOHForStoreAlinity[a][7],   // Multicount
          newUpdatedQOHForStoreAlinity[a][9]    // Item Code
        ])
      }
    }

    // Count the sublocations
    let countUpStoreTransfer = countArrayElem(storeAlinityTransferPending);

    let resCountStoreTransfer = countUpStoreTransfer[2];

    // Expand back these three arrays

    for (a = 0; a < resCountStoreTransfer.length; a++){
        countPerBox = resCountStoreTransfer[a][1] / resCountStoreTransfer[a][0].split(",")[6];
        finalStoreTransfer.push([
                resCountStoreTransfer[a][0].split(",")[0],      // False
                resCountStoreTransfer[a][0].split(",")[7],      // Item Code
                resCountStoreTransfer[a][0].split(",")[1],      // Item Name
                resCountStoreTransfer[a][0].split(",")[4],      // Lot Number
                resCountStoreTransfer[a][0].split(",")[5],      // Exp Date
                countPerBox,                                    // Count in kits
                resCountStoreTransfer[a][1],                    // Count in cartridges
                ]);

    }
    
    //console.log(finalStoreTransfer)
    //console.log(finalStoreOutside)
    //console.log(finalStoreSection)

}

function updateOUTGOINGpaste(){

    updateOUTGOING();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let getOutgoingListLastRow = ss.getSheetByName("OUTGOING").getLastRow();

    // Clear content to paste the updated list later
    if (getOutgoingListLastRow === 1){
    } else { 
    ss.getSheetByName("OUTGOING").getRange(3,1,getOutgoingListLastRow-1,7).clearContent().removeCheckboxes();
    // Remove the filter set by user
    ss.getSheetByName("OUTGOING").getRange(3,1,getOutgoingListLastRow-1,7).getFilter().remove();
    }

    if (newUpdatedQOHFullList.length === 0){
    } else { 
    ss.getSheetByName("OUTGOING").getRange(3,1,newUpdatedQOHFullList.length,newUpdatedQOHFullList[0].length).setValues(newUpdatedQOHFullList).sort([{column: 3, ascending: true},{column: 7, ascending: true}]);
    ss.getSheetByName("OUTGOING").getRange(3,1,newUpdatedQOHFullList.length,1).insertCheckboxes();

    // Get the second column in OUTGOING list
    const range = ss.getSheetByName("OUTGOING").getRange("B2");
    ss.getSheetByName("OUTGOING").hideColumn(range);
    // Set filter to new array sheet
    ss.getSheetByName("OUTGOING").getRange(2,1,newUpdatedQOHFullList.length+1,7).createFilter();
    ss.getSheetByName("OUTGOING").getRange(3,7,newUpdatedQOHFullList.length+1,1).setNumberFormat("DD/MM/YYYY");
    }

    ss.getSheetByName("OUTGOING").getRange(1,4,1,1).setValue('List is updated!');

}

let readyArrayFortblQOHPRpaste = []; // To make it available to Medicorp
let readyArrayFortblQOHFOCpaste = []; // To make it available to Medicorp


function updateQOHList(){

    updateOUTGOING();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const getListOfCurrentINIDLastRow = ss.getSheetByName("OUTGOING").getLastRow();
    const getTblUniqueINIDLastRow = ss.getSheetByName("tblUniqueINID").getLastRow();
    const getLotNumberInUniqueListLastRow = ss.getSheetByName("tblUniqueINID").getLastRow();
    const getPRLISTLastRow = ss.getSheetByName("QOH PR").getLastRow();
    const getFOCLISTLastRow = ss.getSheetByName("QOH FOC").getLastRow();
    const getEXPLastRow = ss.getSheetByName("EXPIRED").getLastRow();

    let getListOfCurrentINIDarray = remainingIDLeftForQOHList; //ss.getSheetByName("OUTGOING").getRange(2,2,getListOfCurrentINIDLastRow-1,1).getValues(); // Change to remainingIDLeftForQOHList
    let getTblUniqueINIDarray = ss.getSheetByName("tblUniqueINID").getRange(2,1,getTblUniqueINIDLastRow-1,5).getValues();
    let getTblUniqueINIDLotNumbersarray = ss.getSheetByName("tblUniqueINID").getRange(2,3,getLotNumberInUniqueListLastRow-1,3).getValues();
    let getMasterList = ss.getSheetByName("MasterL").getRange(2,1,getMasterListLastRow-1,9).getValues();
    let getCheckboxPRFilter = ss.getSheetByName("QOH PR").getRange(1,8,1,1).getValue();
    let getPRFilterValueFrom = ss.getSheetByName("QOH PR").getRange(3,8,1,1).getValue();
    let getPRFilterValueTo = ss.getSheetByName("QOH PR").getRange(3,9,1,1).getValue();
    let getCheckboxFOCFilter = ss.getSheetByName("QOH FOC").getRange(1,8,1,1).getValue();
    let getFOCFilterValueFrom = ss.getSheetByName("QOH FOC").getRange(3,8,1,1).getValue();
    let getFOCFilterValueTo = ss.getSheetByName("QOH FOC").getRange(3,9,1,1).getValue();

    //console.log(getListOfCurrentINIDarray);
    //console.log(getTblUniqueINIDLotNumbersarray);

    // Expand on the CurrentID list
    let expandedListOfCurrentID = [];
    let listOfItemCodes = [];
    let listOfLotNumbers = [];
    let listOfItemCodesPlusLot =[];

    let getItemCodeValue = '';
    let getItemNameValue = '';
    let getItemTypeValue = '';
    let getItemLocationValue = '';
    let getItemPRType = '';

    for (let j = 0; j < getListOfCurrentINIDarray.length; j++){
      for (let k = 0; k < getTblUniqueINIDarray.length; k++){
        
        if (getListOfCurrentINIDarray[j] === getTblUniqueINIDarray[k][0]){
          
          for (let l = 0; l < getMasterList.length; l++){
            
            if (getTblUniqueINIDarray[k][2] === getMasterList[l][0]){
              
              
              getItemCodeValue = getMasterList[l][0];
              getItemNameValue = getMasterList[l][1];
              getItemTypeValue = getMasterList[l][2];
              getItemLocationValue = getMasterList[l][3];
              getItemPRType = getMasterList[l][5];
            }
          }
         
          expandedListOfCurrentID.push([
                                     getItemCodeValue,                // Item Code
                                     getItemNameValue,                // Item Name
                                     getItemTypeValue,                // Item Type
                                     getItemLocationValue,            // Item Location
                                     "'"+getTblUniqueINIDarray[k][3], // Lot Number
                                     getTblUniqueINIDarray[k][4],     // Exp date
                                     getItemPRType                    // PR/FOC
                                     ])
          listOfItemCodes.push([getItemCodeValue]);
          listOfItemCodesPlusLot.push([getItemCodeValue,
                                       "'"+getTblUniqueINIDarray[k][3]]);
          listOfLotNumbers.push(["'"+getTblUniqueINIDarray[k][3]]);

        }
      }
    }
    //console.log(expandedListOfCurrentID);
    //console.log(listOfItemCodes)
    //console.log(listOfItemCodesPlusLot);

    // To calculate the QOH based on current OUTGOING LIST
    // Basically, showing the QOH by cartridges in the QOH display column

    let sentToCountEachItemCode = countArrayElem(listOfItemCodesPlusLot);
    //let sentToCountEachLotNumber = countArrayElem(listOfItemCodesPlusLot);
    
    let newArrayOfQOHListCodes = sentToCountEachItemCode[0];
    let newArrayOfQOHListCodesCount = sentToCountEachItemCode[1];
    let newArrayOfQOHLotNumbers = sentToCountEachItemCode[0];

    //console.log(newArrayOfQOHListCodes)
    //console.log(newArrayOfQOHListCodesCount)
    //console.log(newArrayOfQOHLotNumbers)

    let getItemNameFromMasterList = '';
    let getLocationFromMasterList = '';
    let getMultiCountFromMasterList = '';
    let getItemTypeFromMasterList = '';
    let getPRTypeFromMasterList = '';
    
    let captureNewArray = [];


    for (let j = 0; j < newArrayOfQOHListCodes.length; j++){
        for (let l = 0; l < getMasterList.length; l++){
            
            if (newArrayOfQOHListCodes[j].split(",")[0] === getMasterList[l][0]){
              
              
              getItemNameFromMasterList = getMasterList[l][1];
              getLocationFromMasterList = getMasterList[l][3];
              getMultiCountFromMasterList = getMasterList[l][7];
              getItemTypeFromMasterList = getMasterList[l][2];
              getPRTypeFromMasterList = getMasterList[l][5];

            }
          }
         
          captureNewArray.push([
                                     getItemNameFromMasterList,               // Item Name
                                     newArrayOfQOHLotNumbers[j],              // Item Code + Lot Numbers
                                     newArrayOfQOHListCodesCount[j],          // Quantity per cartridge
                                     getMultiCountFromMasterList,             // MultiCount
                                     getLocationFromMasterList,               // Item Location
                                     getItemTypeFromMasterList,               // Item Type
                                     getPRTypeFromMasterList                  // PR/FOC
                                     ])

        }
        //console.log(captureNewArray);
        //console.log(captureNewArray[0][1]);
        //console.log(captureNewArray[0][1]);
        //console.log(getTblUniqueINIDLotNumbersarray[0][0]);

    // To look for exp date and calculate quantity in boxes
    
    let valueToMatchLotNumber = '';
    let extractedLotNumberValue = '';
    let valueToMatchQuantityPerCartridge = 0;
    let valueToMatchMulticount = 0;
    let valueToMatchItemName = '';
    let valueToMatchItemType = '';
    let valueToMatchItemLocation = '';
    let lookForExpDateValue = '';
    let countForQuantityInBoxes = 0;


    //let readyArrayFortblQOHPRpaste = []; // Made to global variable

    for (let m = 0; m < captureNewArray.length; m++){
      if (captureNewArray[m][6] === 'PR'){
        valueToMatchLotNumber = captureNewArray[m][1]                                                           // Item Code + Lot number
        extractedLotNumberValue = "'" + captureNewArray[m][1].substring(6,captureNewArray[m][1].length+6);      // Extracted Lot Number
        valueToMatchQuantityPerCartridge = captureNewArray[m][2]                                                // Quantity per cartridge
        valueToMatchMulticount = captureNewArray[m][3]                                                          // MultiCount
        valueToMatchItemName = captureNewArray[m][0]                                                            // Item Name
        valueToMatchItemType = captureNewArray[m][5]                                                            // Item Type
        valueToMatchItemLocation = captureNewArray[m][4]                                                        // Item Location
        
        countForQuantityInBoxes = valueToMatchQuantityPerCartridge/valueToMatchMulticount;

        //console.log(valueToMatchLotNumber);
        //console.log(extractedLotNumberValue);

      for (let n = 0; n < getTblUniqueINIDLotNumbersarray.length; n++){
        let searchInTblUniqueItemCodeAndLot = getTblUniqueINIDLotNumbersarray[n][0]+",'"+getTblUniqueINIDLotNumbersarray[n][1];
        //console.log(searchInTblUniqueItemCodeAndLot);
        if (valueToMatchLotNumber === searchInTblUniqueItemCodeAndLot){
          lookForExpDateValue = getTblUniqueINIDLotNumbersarray[n][2];
        }
      }

      readyArrayFortblQOHPRpaste.push([
                                     valueToMatchItemName,                    // Item Name
                                     valueToMatchItemType,                    // Item Type
                                     extractedLotNumberValue,                 // Lot Numbers
                                     lookForExpDateValue,                     // Exp Date
                                     valueToMatchQuantityPerCartridge,        // Quantity per cartridge
                                     countForQuantityInBoxes,                 // Quantity in boxes
                                     valueToMatchItemLocation,                // Item Location
      ])
    }
  }

  //console.log(readyArrayFortblQOHPRpaste)

  // Search for expired items
  let prepForExpiredItems = [];
  let lookForExpiredItems = [];


  for (let p = 0; p < captureNewArray.length; p++){
        valueToMatchLotNumber = captureNewArray[p][1]                                                           // Item Code + Lot number
        extractedLotNumberValue = "'" + captureNewArray[p][1].substring(6,captureNewArray[p][1].length+6);      // Extracted Lot Number
        valueToMatchQuantityPerCartridge = captureNewArray[p][2]                                                // Quantity per cartridge
        valueToMatchMulticount = captureNewArray[p][3]                                                          // MultiCount
        valueToMatchItemName = captureNewArray[p][0]                                                            // Item Name
        valueToMatchItemType = captureNewArray[p][5]                                                            // Item Type
        valueToMatchItemLocation = captureNewArray[p][4]                                                        // Item Location
        
        countForQuantityInBoxes = valueToMatchQuantityPerCartridge/valueToMatchMulticount;

        //console.log(valueToMatchLotNumber);
        //console.log(extractedLotNumberValue);

      for (let n = 0; n < getTblUniqueINIDLotNumbersarray.length; n++){
        let searchInTblUniqueItemCodeAndLot = getTblUniqueINIDLotNumbersarray[n][0]+",'"+getTblUniqueINIDLotNumbersarray[n][1];
        //console.log(searchInTblUniqueItemCodeAndLot);
        if (valueToMatchLotNumber === searchInTblUniqueItemCodeAndLot){
          lookForExpDateValue = getTblUniqueINIDLotNumbersarray[n][2];
        }
      }

      prepForExpiredItems.push([
                                     valueToMatchItemName,                    // Item Name
                                     valueToMatchItemType,                    // Item Type
                                     extractedLotNumberValue,                 // Lot Numbers
                                     lookForExpDateValue,                     // Exp Date
                                     valueToMatchQuantityPerCartridge,        // Quantity per cartridge
                                     countForQuantityInBoxes,                 // Quantity in boxes
                                     valueToMatchItemLocation,                // Item Location
      ])
  }
  
  for (u = 0; u < prepForExpiredItems.length; u++){
    const today = new Date();
    if (prepForExpiredItems[u][3] < today){
      lookForExpiredItems.push([
              prepForExpiredItems[u][0],   // Item Name
              prepForExpiredItems[u][1],   // Item Type
              prepForExpiredItems[u][2],   // Lot Numbers
              prepForExpiredItems[u][3],   // Exp Date
              prepForExpiredItems[u][4],   // Quantity per cartridge
              prepForExpiredItems[u][5],   // Quantity in boxes
              prepForExpiredItems[u][6],   // Item Location
              ]);
    }
  }
  //console.log(lookForExpiredItems)


  let filteredQOHPRList = []; 

  for (let q = 0; q < readyArrayFortblQOHPRpaste.length; q++){
    if (getCheckboxPRFilter === false){
      filteredQOHPRList.push([
                              readyArrayFortblQOHPRpaste[q][0],
                              readyArrayFortblQOHPRpaste[q][1],
                              readyArrayFortblQOHPRpaste[q][2],
                              readyArrayFortblQOHPRpaste[q][3],
                              readyArrayFortblQOHPRpaste[q][4],
                              readyArrayFortblQOHPRpaste[q][5],
                              readyArrayFortblQOHPRpaste[q][6],
      ])
    } else if (getCheckboxPRFilter === true){


      if (readyArrayFortblQOHPRpaste[q][3] <= getPRFilterValueTo &&
          readyArrayFortblQOHPRpaste[q][3] >= getPRFilterValueFrom){

        
      filteredQOHPRList.push([
                              readyArrayFortblQOHPRpaste[q][0],
                              readyArrayFortblQOHPRpaste[q][1],
                              readyArrayFortblQOHPRpaste[q][2],
                              readyArrayFortblQOHPRpaste[q][3],
                              readyArrayFortblQOHPRpaste[q][4],
                              readyArrayFortblQOHPRpaste[q][5],
                              readyArrayFortblQOHPRpaste[q][6],
                              ])


    }
  }
}
    //console.log(readyArrayFortblQOHPRpaste);
    //console.log(filteredQOHPRList);

    // Clear off the sheet first
    if (getPRLISTLastRow-1 === 0){
    } else {
    ss.getSheetByName("QOH PR").getRange(2,1,getPRLISTLastRow-1,7).clearContent();
    ss.getSheetByName("QOH PR").getRange(2,1,getPRLISTLastRow-1,7).getFilter().remove();
    }

    // No Filter paste
    if (getCheckboxPRFilter === false){
    ss.getSheetByName("QOH PR").getRange(2,1,readyArrayFortblQOHPRpaste.length,readyArrayFortblQOHPRpaste[0].length).setValues(readyArrayFortblQOHPRpaste).sort([{column: 1,ascending: true},{column: 4,ascending: true}]);
    // Set filter for user
    ss.getSheetByName("QOH PR").getRange(1,1,readyArrayFortblQOHPRpaste.length+1,7).createFilter();
    ss.getSheetByName("QOH PR").getRange(2,4,readyArrayFortblQOHPRpaste.length+1,1).setNumberFormat("DD/MM/YYYY");

    // Filter paste
    } else {
    ss.getSheetByName("QOH PR").getRange(2,1,filteredQOHPRList.length,filteredQOHPRList[0].length).setValues(filteredQOHPRList).sort([{column: 1,ascending: true},{column: 4,ascending: true}]);
    // Set filter for user
    ss.getSheetByName("QOH PR").getRange(1,1,filteredQOHPRList.length+1,7).createFilter();
    ss.getSheetByName("QOH PR").getRange(2,4,readyArrayFortblQOHPRpaste.length+1,1).setNumberFormat("DD/MM/YYYY");
    }

    // Set the filter to false after every filter run
    ss.getSheetByName("QOH PR").getRange(1,8,1,1).setValue(false);
    // Set the datevaluefrom cell to today() formula
    ss.getSheetByName("QOH PR").getRange(3,8,1,1).setFormula("=today()");


    //let readyArrayFortblQOHFOCpaste = [];  // Make it to global variable

    for (let r = 0; r < captureNewArray.length; r++){
      if (captureNewArray[r][6] === 'FOC'){
        valueToMatchLotNumber = captureNewArray[r][1]                                                           // Item Code + Lot number
        extractedLotNumberValue = "'" + captureNewArray[r][1].substring(6,captureNewArray[r][1].length+6);      // Extracted Lot Number
        valueToMatchQuantityPerCartridge = captureNewArray[r][2]                                                // Quantity per cartridge
        valueToMatchMulticount = captureNewArray[r][3]                                                          // MultiCount
        valueToMatchItemName = captureNewArray[r][0]                                                            // Item Name
        valueToMatchItemType = captureNewArray[r][5]                                                            // Item Type
        valueToMatchItemLocation = captureNewArray[r][4]                                                        // Item Location
        
        countForQuantityInBoxes = valueToMatchQuantityPerCartridge/valueToMatchMulticount;

        //console.log(valueToMatchLotNumber);
        //console.log(extractedLotNumberValue);

      for (let n = 0; n < getTblUniqueINIDLotNumbersarray.length; n++){
        let searchInTblUniqueItemCodeAndLot = getTblUniqueINIDLotNumbersarray[n][0]+",'"+getTblUniqueINIDLotNumbersarray[n][1];
        //console.log(searchInTblUniqueItemCodeAndLot);
        if (valueToMatchLotNumber === searchInTblUniqueItemCodeAndLot){
          lookForExpDateValue = getTblUniqueINIDLotNumbersarray[n][2];
        }
      }

      readyArrayFortblQOHFOCpaste.push([
                                     valueToMatchItemName,                    // Item Name
                                     valueToMatchItemType,                    // Item Type
                                     extractedLotNumberValue,                 // Lot Numbers
                                     lookForExpDateValue,                     // Exp Date
                                     valueToMatchQuantityPerCartridge,        // Quantity per cartridge
                                     countForQuantityInBoxes,                 // Quantity in boxes
                                     valueToMatchItemLocation,                // Item Location
      ])
    }
  }

  let filteredQOHFOCList = []; 

  for (let t = 0; t < readyArrayFortblQOHFOCpaste.length; t++){
    if (getCheckboxFOCFilter === false){
      filteredQOHFOCList.push([
                              readyArrayFortblQOHFOCpaste[t][0],
                              readyArrayFortblQOHFOCpaste[t][1],
                              readyArrayFortblQOHFOCpaste[t][2],
                              readyArrayFortblQOHFOCpaste[t][3],
                              readyArrayFortblQOHFOCpaste[t][4],
                              readyArrayFortblQOHFOCpaste[t][5],
                              readyArrayFortblQOHFOCpaste[t][6],
      ])
    } else if (getCheckboxFOCFilter === true){


      if (readyArrayFortblQOHFOCpaste[t][3] <= getFOCFilterValueTo &&
          readyArrayFortblQOHFOCpaste[t][3] >= getFOCFilterValueFrom){

        
      filteredQOHFOCList.push([
                              readyArrayFortblQOHFOCpaste[t][0],
                              readyArrayFortblQOHFOCpaste[t][1],
                              readyArrayFortblQOHFOCpaste[t][2],
                              readyArrayFortblQOHFOCpaste[t][3],
                              readyArrayFortblQOHFOCpaste[t][4],
                              readyArrayFortblQOHFOCpaste[t][5],
                              readyArrayFortblQOHFOCpaste[t][6],
                              ])
      
      if (readyArrayFortblQOHFOCpaste[t][3] < getFOCFilterValueFrom &&
          readyArrayFortblQOHFOCpaste[t][3] != "'"+'BLANK'){

        
      lookForExpiredItems.push([
                              readyArrayFortblQOHFOCpaste[t][0],
                              readyArrayFortblQOHFOCpaste[t][1],
                              readyArrayFortblQOHFOCpaste[t][2],
                              readyArrayFortblQOHFOCpaste[t][3],
                              readyArrayFortblQOHFOCpaste[t][4],
                              readyArrayFortblQOHFOCpaste[t][5],
                              readyArrayFortblQOHFOCpaste[t][6],
                              ])
      }


    }
  }
}
    //console.log(filteredQOHFOCList);
    //console.log(readyArrayFortblQOHFOCpaste);
    
    // Clear off the sheet first
    if (getFOCLISTLastRow-1 === 0){
    } else {
    ss.getSheetByName("QOH FOC").getRange(2,1,getFOCLISTLastRow-1,7).clearContent();
    ss.getSheetByName("QOH FOC").getRange(2,1,getFOCLISTLastRow-1,7).getFilter().remove();
    }

    // No Filter paste
    if (getCheckboxFOCFilter === false){
    ss.getSheetByName("QOH FOC").getRange(2,1,readyArrayFortblQOHFOCpaste.length,readyArrayFortblQOHFOCpaste[0].length).setValues(readyArrayFortblQOHFOCpaste).sort([{column: 1,ascending: true},{column: 4,ascending: true}]);
    // Set filter for user
    ss.getSheetByName("QOH FOC").getRange(1,1,readyArrayFortblQOHFOCpaste.length+1,7).createFilter();
    ss.getSheetByName("QOH FOC").getRange(2,4,readyArrayFortblQOHFOCpaste.length+1,1).setNumberFormat("DD/MM/YYYY");

    // Filter paste
    } else {
    ss.getSheetByName("QOH FOC").getRange(2,1,filteredQOHFOCList.length,filteredQOHFOCList[0].length).setValues(filteredQOHFOCList).sort([{column: 1,ascending: true},{column: 4,ascending: true}]);
    // Set filter for user
    ss.getSheetByName("QOH FOC").getRange(1,1,filteredQOHFOCList.length+1,7).createFilter();
    ss.getSheetByName("QOH FOC").getRange(2,4,readyArrayFortblQOHFOCpaste.length+1,1).setNumberFormat("DD/MM/YYYY");
    }

    // Set the filter to false after every filter run
    ss.getSheetByName("QOH FOC").getRange(1,8,1,1).setValue(false);
    // Set the datevaluefrom cell to today() formula
    ss.getSheetByName("QOH FOC").getRange(3,8,1,1).setFormula("=today()");

    // Expired items
    // Clear off the expired items list first
    if (getEXPLastRow-1 === 0){
    } else {
    ss.getSheetByName("EXPIRED").getRange(2,1,getEXPLastRow-1,7).clearContent();
    ss.getSheetByName("EXPIRED").getRange(2,1,getEXPLastRow-1,7).getFilter().remove();

    }
    // Paste the new expired list to the sheet
    ss.getSheetByName("EXPIRED").getRange(2,1,lookForExpiredItems.length,lookForExpiredItems[0].length).setValues(lookForExpiredItems);
    ss.getSheetByName("EXPIRED").getRange(1,1,lookForExpiredItems.length+1,7).createFilter();
    ss.getSheetByName("EXPIRED").getRange(2,4,lookForExpiredItems.length+1,1).setNumberFormat("DD/MM/YYYY");


}

function updatePRList() {

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const getPRLISTLastRow = ss.getSheetByName("tblPR").getLastRow();
    const getDOLISTLastRow = ss.getSheetByName("tblDO").getLastRow();

    let getListOfCurrentPRarray = ss.getSheetByName("tblPR").getRange(2,1,getPRLISTLastRow-1,10).getValues();
    // TO handle zero DO
    let getListOfCurrentDOarray = ss.getSheetByName("tblDO").getRange(2,1,getDOLISTLastRow-1,9).getValues();

    //console.log(getListOfCurrentPRarray);
    //console.log(getListOfCurrentDOarray);

    // Get current PR list
    let cleanPRArray = []
    for (i = 0; i < getListOfCurrentPRarray.length; i++) {
      cleanPRArray.push(appendArrays(
                  getListOfCurrentPRarray[i][1], // PR Number
                  getListOfCurrentPRarray[i][2], // Item Code
                  getListOfCurrentPRarray[i][6]  // Quantity Ordered
      ))}
      //console.log(cleanPRArray.length)

    // Get current DO list
    let cleanDOArray = []
    for (j = 0; j < getListOfCurrentDOarray.length; j++) {
      cleanDOArray.push(appendArrays(
                  getListOfCurrentDOarray[j][3], // PR Number
                  getListOfCurrentDOarray[j][4], // Item Code
                  getListOfCurrentDOarray[j][7]  // Quantity Received
      ))}
      //console.log(cleanDOArray);

      // Count the remaining quantity
    let resultOfCalculatingRemainder = countRemainderFromPR(cleanPRArray, cleanDOArray);

      //console.log(resultOfCalculatingRemainder);
      //console.log(cleanPRArray.length)
      //console.log(cleanDOArray);
      //console.log(resultOfCalculatingRemainder.length)

      // Status update to check PR is partial or completed
    let statusUpdateList = [];
    let statusResult = '';
      for (k = 0; k < cleanPRArray.length; k++){
        if (cleanPRArray[k][2] === resultOfCalculatingRemainder[k][2]){
          // Pending PO/ PO approved check
          statusResult = 'No DO yet';
        
        } else if (cleanPRArray[k][2] > resultOfCalculatingRemainder[k][2] && resultOfCalculatingRemainder[k][2] !== 0){
          statusResult = 'Partial Receipt';
        
        } else if (resultOfCalculatingRemainder[k][2] === 0){
          statusResult = 'Completed';
        
        } statusUpdateList.push([statusResult]);
        }
        //console.log(statusUpdateList);


      // Combine the new array list to update tblPR
      let newArrayForTblPR = [];
      for (k = 0; k < getListOfCurrentPRarray.length; k++){
          newArrayForTblPR.push(appendArrays(
                      getListOfCurrentPRarray[k][0], // Timestamp
                      getListOfCurrentPRarray[k][1], // PR Number
                      getListOfCurrentPRarray[k][2], // Item Code
                      getListOfCurrentPRarray[k][3], // Item Name
                      getListOfCurrentPRarray[k][4], // Item Type
                      getListOfCurrentPRarray[k][5], // Vesalius Code
                      getListOfCurrentPRarray[k][6], // Quantity Ordered
                      resultOfCalculatingRemainder[k][2], // Quantity Left
                      statusUpdateList[k][0],             // Status
                      getListOfCurrentPRarray[k][9]       // Remarks
          ));
      }
      //console.log(newArrayForTblPR);

      // Paste the new tblPR with updated Quantity Left and status
      ss.getSheetByName("tblPR").getRange(2,1,newArrayForTblPR.length,newArrayForTblPR[0].length).setValues(newArrayForTblPR);


}

function updatePOEntry() {

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const getTblPRLastRow = ss.getSheetByName("tblPR").getLastRow();
    let getPOEntryLastRow = ss.getSheetByName("PO Entry").getLastRow();
    const getTblPOLastRow = ss.getSheetByName("tblPO").getLastRow();

    // Get all of MasterL - longer columns than the others 9 to 14
    let getMasterList = ss.getSheetByName("MasterL").getRange(2,1,getMasterListLastRow-1,14).getValues();
    let getActivePRList = ss.getSheetByName("tblPR").getRange(2,1,getTblPRLastRow-1,10).getValues();
    let getActivePOList = ss.getSheetByName("tblPO").getRange(2,1,getTblPOLastRow-1,8).getValues();

    //console.log(getPONumberEnteredValue);
    //console.log(getPORemarksValue);
    //console.log(getActivePRList);
    //console.log(getActivePOList);


    // Prep the Active PR list for PO Entry
    let listForPOEntry = [];
    let listOfPRNumbersInTblPR = [];
    let listOfPRNumbersInTblPO = [];
    let listOfRemainingPR = [];

  // Get all the PR Numbers from tblPR, assuming all are unique per items
    for (a = 0; a < getActivePRList.length; a++){
    listOfPRNumbersInTblPR.push(getActivePRList[a][1]);
  }
  //console.log(listOfPRNumbersInTblPR);

  // Get all the PR Numbers from tblPO, assuming all are unique per items
    for (b = 0; b < getActivePOList.length; b++){
    listOfPRNumbersInTblPO.push(getActivePOList[b][2]);
  }
  //console.log(listOfPRNumbersInTblPO);

    findRemainingUniqueID(listOfPRNumbersInTblPR.sort(),listOfPRNumbersInTblPO.sort());
    listOfRemainingPR = listOfPRNumbersInTblPR;

    //console.log(listOfRemainingPR);
    //console.log(listOfPRNumbersInTblPR);


    for (c = 0; c < getActivePRList.length; c++) {
      // Add if statement here to take only PR without PO or status changed
      for (d = 0; d < listOfRemainingPR.length; d++){

        if (getActivePRList[c][1] === listOfRemainingPR[d]){

      listForPOEntry.push([
                          false,
                          getActivePRList[c][1],  // PR Number
                          getActivePRList[c][2],  // Item Code
                          getActivePRList[c][3],  // Item Name
                          getActivePRList[c][4],  // Item Type
                          getActivePRList[c][6],  // Quantity Ordered
      ])}}};
      //console.log(listForPOEntry);
      //console.log(getPOEntryLastRow);

      // Clear off previous contents first
      //console.log(getPOEntryLastRow);
      if (listForPOEntry.length === 0){
        
        const promptForEmptyLength = SpreadsheetApp.getUi().alert("There are no more PR pending for PO approval.", SpreadsheetApp.getUi().ButtonSet.OK);
        SpreadsheetApp.getActive().toast(promptForEmptyLength);

      } else {

      if (getPOEntryLastRow-4 === 0){
      
      } else {
      
      ss.getSheetByName("PO Entry").getRange(5,1,getPOEntryLastRow-4,6).clearContent().removeCheckboxes();

      }

      // Paste the Active PR list
      ss.getSheetByName("PO Entry").getRange(5,1,listForPOEntry.length,listForPOEntry[0].length).setValues(listForPOEntry);
      // Insert checkboxes in the first column
      getPOEntryLastRow = ss.getSheetByName("PO Entry").getLastRow(); // To update the number of rows before inserting checkboxes
      ss.getSheetByName("PO Entry").getRange(5,1,getPOEntryLastRow-4,1).insertCheckboxes();
      }

}

function updateDOEntry() {

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const getTblPRLastRow = ss.getSheetByName("tblPR").getLastRow();
    let   getDOEntryLastRow = ss.getSheetByName("DO Entry").getLastRow();
    const getTblPOLastRow = ss.getSheetByName("tblPO").getLastRow();
    const getBackDateValue = ss.getSheetByName("DO Entry").getRange("H1").getValue();
    let checkBoxValueForPRitems = ss.getSheetByName("DO Entry").getRange(1,3,1,1).getValue();
    let checkBoxValueForFOCitems = ss.getSheetByName("DO Entry").getRange(2,3,1,1).getValue();
    let getActivePRList = ss.getSheetByName("tblPR").getRange(2,1,getTblPRLastRow-1,10).getValues();

    // Get all of MasterL - longer columns than the others 9 to 14
    let getMasterList = ss.getSheetByName("MasterL").getRange(2,1,getMasterListLastRow-1,14).getValues();
    let getActivePOListFull = ss.getSheetByName("tblPO").getRange(2,1,getTblPOLastRow-1,8).getValues();

    // Auto checks the checkbox PR if both empty
    if (checkBoxValueForPRitems === true &&
        checkBoxValueForFOCitems === false){
      checkBoxValueForFOCitems = false;
    } else if (checkBoxValueForPRitems === false &&
               checkBoxValueForFOCitems === true){
      checkBoxValueForFOCitems = true;
    } else if (checkBoxValueForPRitems === false &&
               checkBoxValueForFOCitems === false){
      checkBoxValueForPRitems = true;

    }
    
    // Time stamp all inputs
    let today = "";

    if (getBackDateValue != ""){
      today = new Date(getBackDateValue);
    } else {
      today = new Date();
    }

      const dateEachItemDay = today.getDate();
      const dateEachItemMonth = today.getMonth()+1;
      const dateEachItemYear = today.getFullYear();
      //console.log(dateEachItemDay)
      //console.log(dateEachItemMonth)
      //console.log(dateEachItemYear)
      let time = `${dateEachItemDay}/${dateEachItemMonth}/${dateEachItemYear}`;

    // Prep the Active PR PO list for tblDO
    let listOfNewPRPOArray = [];
    let getActivePOList = [];

    for (i = 0; i < getActivePOListFull.length; i++){
    if (checkBoxValueForPRitems === true &&
        getActivePOListFull[i][1] != 'FOC Order'){  // For PR Items
      getActivePOList.push(appendArrays(
                getActivePOListFull[i][0], // Timestamp
                getActivePOListFull[i][1], // PO Number
                getActivePOListFull[i][2], // PR Number
                getActivePOListFull[i][3], // Item Code
                getActivePOListFull[i][4], // Item Name
                getActivePOListFull[i][5], // Item Type
                getActivePOListFull[i][6], // Quantity Ordered
                getActivePOListFull[i][7]  // Remarks
      ))} else if (checkBoxValueForFOCitems === true &&
                   getActivePOListFull[i][1] === 'FOC Order'){ // For FOC items
      getActivePOList.push(appendArrays(
                getActivePOListFull[i][0], // Timestamp
                getActivePOListFull[i][1], // PO Number
                getActivePOListFull[i][2], // PR Number
                getActivePOListFull[i][3], // Item Code
                getActivePOListFull[i][4], // Item Name
                getActivePOListFull[i][5], // Item Type
                getActivePOListFull[i][6], // Quantity Ordered
                getActivePOListFull[i][7]  // Remarks
      ))}
    }
    //console.log(getActivePOList);

    // Get the current INCOMING items List
    captureDOitemsForPRupdate();
    let fullArrayOfINList = extractINLISTDayMonthYearWithFullList;

    //console.log(fullArrayOfINList);
    // Get the breakdown of Date today in DateMonthYear format to look for matches in the IN List
    const todayDateInDateMonthYear = `${dateEachItemDay}${dateEachItemMonth}${dateEachItemYear}`;
    
    //const todayDateInDateMonthYear = '1452022'; // For manual date entry if DO was not entered on the same day
    //console.log(todayDateInDateMonthYear);

    // Go through each IN List to find today's received items and reverse back any multicounts to per box/kit
    let modifiedArrayOfINList = [];
    let getActualQuantityFromINList = [];

    for (d = 0; d < fullArrayOfINList.length; d++){
      if (fullArrayOfINList[d][0] === todayDateInDateMonthYear){
        let getItemCodeValue = fullArrayOfINList[d][3];
        //let getMultiCountValue = 0;
        for (e = 0; e < getMasterList.length; e++){
          if (getItemCodeValue === getMasterList[e][0]){
            //getMultiCountValue = getMasterList[e][7];
}
}    
          modifiedArrayOfINList.push([
          fullArrayOfINList[d][3],
          //fullArrayOfINList[d][9]  //getMultiCountValue// Do not divide quantity by Multicount
        ]);

          getActualQuantityFromINList.push([
            fullArrayOfINList[d][3],
            fullArrayOfINList[d][9]
          ]);
}
}
  //console.log(modifiedArrayOfINList);

  // Calculate same item to be summed up
  let countModifiedArrayOfINList = countArrayElem(modifiedArrayOfINList);

  //console.log(justItemCodesOfINList);
  let countActualItemReceived = [];

  for (i = 0; i < countModifiedArrayOfINList[0].length; i++){
    let counterForMultipleItems = 0;
    for (j = 0; j < getActualQuantityFromINList.length; j++){
      if (countModifiedArrayOfINList[0][i] === getActualQuantityFromINList[j][0]){
      
        counterForMultipleItems += getActualQuantityFromINList[j][1];

      }
    }
        countActualItemReceived.push([
        countModifiedArrayOfINList[0][i],
        counterForMultipleItems
        ]);
  }

  //console.log(countActualItemReceived);

    // Cycle through PO list and find active PR
    for (a = 0; a < getActivePOList.length; a++) {
      // Look for the quantity left for PR Number and Item Code
      const matchPRNumberValue = getActivePOList[a][2];
      const matchItemCodeValue = getActivePOList[a][3];
      let matchQuantityReceivedValue = 0;

      for (b = 0; b < getActivePRList.length; b++){
        const quantityLeftValue = getActivePRList[b][7];
        // To look for matchvalue of PR and Item Code and Quantity Left is not equal to Zero
        if (quantityLeftValue !== 0 &&
            matchPRNumberValue === getActivePRList[b][1] &&
            matchItemCodeValue === getActivePRList[b][2]){

        // To look for IN LIST items to match for DO entry
                for (f = 0; f < countActualItemReceived.length; f++){
                  if (matchItemCodeValue ===  countActualItemReceived[f][0]){
                    matchQuantityReceivedValue = countActualItemReceived[f][1];

                listOfNewPRPOArray.push(appendArrays(
                          false,                     // FALSE checkbox
                          getActivePOList[a][2],     // PR Number
                          getActivePOList[a][1],     // PO Number
                          getActivePOList[a][3],     // Item Code
                          getActivePOList[a][4],     // Item Name
                          getActivePOList[a][5],     // Item Type
                          getActivePOList[a][6],     // Quantity Ordered
                          //matchQuantityLeftValue,    // Quantity Left cannot calculate here!
                          matchQuantityReceivedValue // Quantity Received
                          
                 ))
           }
         }
       }
     }
   };
      //console.log(listOfNewPRPOArray);

      // Clear off previous contents first
      // If zero length, to stop script
      if (getDOEntryLastRow-4 === 0){
      } else {
      ss.getSheetByName("DO Entry").getRange(5,1,getDOEntryLastRow-4,8).clearContent().removeCheckboxes();
      }

      // Paste the Active PR PO list
      if(listOfNewPRPOArray.length === 0){

          const promptForEmptyLength = SpreadsheetApp.getUi().alert(`There are no received items found. Please back date if items weren't received today, ${time}.`, SpreadsheetApp.getUi().ButtonSet.OK);
        SpreadsheetApp.getActive().toast(promptForEmptyLength);

      } else {
      ss.getSheetByName("DO Entry").getRange(5,1,listOfNewPRPOArray.length,listOfNewPRPOArray[0].length).setValues(listOfNewPRPOArray);
      getDOEntryLastRow = ss.getSheetByName("DO Entry").getLastRow();
      ss.getSheetByName("DO Entry").getRange(5,1,getDOEntryLastRow-4,1).insertCheckboxes();
      }
}

let extractINLISTDayMonthYearWithFullList = []; // Array ready for DO entry

function captureDOitemsForPRupdate() {

    // Capture DO items received from the IN LIST
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const getINListLastRow = ss.getSheetByName("IN LIST").getLastRow();

    let getINListarray = ss.getSheetByName("IN LIST").getRange(2,1,getINListLastRow-1,9).getValues();

    // Go through each row to get match Dates
    let concatDayMthYear = '';
    
    for (i = 0; i < getINListarray.length; i++){
          let dateValue = new Date(getINListarray[i][0]);

          const extDateValue = dateValue.getDate();
          const extMthValue = dateValue.getMonth()+1;
          const extYearValue = dateValue.getFullYear();
          concatDayMthYear = `${extDateValue}${extMthValue}${extYearValue}`;

          extractINLISTDayMonthYearWithFullList.push([
            concatDayMthYear,      // Concatenated Day+Month+Year to match with DO entry
            getINListarray[i][0],  // Timestamp
            getINListarray[i][1],  // Barcode
            getINListarray[i][2],  // Item Code
            getINListarray[i][3],  // Item Name
            getINListarray[i][4],  // Item Type
            getINListarray[i][5],  // Location
            getINListarray[i][6],  // Lot Number
            getINListarray[i][7],  // Expiry Date
            getINListarray[i][8],  // Quantity
            ]);
    }
    //console.log(extractINLISTDayMonthYearWithFullList);
}

function updateBestExp() {

    const ss = SpreadsheetApp.getActiveSpreadsheet(); 
    const getPRLISTLastRow = ss.getSheetByName("tblPR").getLastRow();
    const getPOLISTLastRow = ss.getSheetByName("tblPO").getLastRow();
    const getBestExpLastRow = ss.getSheetByName("BestExp").getLastRow();

    updatePRList();

    let getListOfCurrentPRarray = ss.getSheetByName("tblPR").getRange(2,1,getPRLISTLastRow-1,9).getValues();
    let getListOfCurrentPOarray = ss.getSheetByName("tblPO").getRange(2,1,getPOLISTLastRow-1,8).getValues();

    //console.log(getListOfCurrentPRarray);
    //console.log(getListOfCurrentPOarray);

    // Collect PR arrays to fill up BestExp form
    let onlyQuantityLeftPRlist = [];
    let carryThePOMatchValue = '';
    for (i = 0; i < getListOfCurrentPRarray.length; i++){
      if (getListOfCurrentPRarray[i][7] !== 0){
    for (j = 0; j < getListOfCurrentPOarray.length; j++){
      if (getListOfCurrentPRarray[i][1] === getListOfCurrentPOarray[j][2]){
          carryThePOMatchValue = getListOfCurrentPOarray[j][1];
}
}
          onlyQuantityLeftPRlist.push([
                  getListOfCurrentPRarray[i][1],  // PR Number
                  carryThePOMatchValue,           // PO Number
                  getListOfCurrentPRarray[i][2],  // Item Code
                  getListOfCurrentPRarray[i][3],  // Item Name
                  getListOfCurrentPRarray[i][4],  // Item Type
                  getListOfCurrentPRarray[i][6],  // Quantity Ordered
                  getListOfCurrentPRarray[i][7]   // Quantity Remaining
                  ])
}
}
//console.log(onlyQuantityLeftPRlist);

// Need to clear up BestExp table to handle different length of Incomplete PR list
if (getBestExpLastRow-3 === 0){
} else {ss.getSheetByName("BestExp").getRange(4,1,getBestExpLastRow-3,8).clearContent();
}

// Paste the updated list of Incomplete PR to BestExp table
ss.getSheetByName("BestExp").getRange(4,1,onlyQuantityLeftPRlist.length,onlyQuantityLeftPRlist[0].length).setValues(onlyQuantityLeftPRlist);

}

function updateBatchNumber() {

  // To update the validation drop down list for Batch Numbers from tblBestExp
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const getBestExpListLastRow = ss.getSheetByName("tblBestExp").getLastRow();
    const getListofBatchNumber = ss.getSheetByName("tblBestExp").getRange(2,2,getBestExpListLastRow-1,1).getValues();
    let getUniqueListofBatchNumberFull = countArrayElem(getListofBatchNumber);
    let getUniqueListofBatchNumber = SpreadsheetApp.newDataValidation().requireValueInList(getUniqueListofBatchNumberFull[0]).build();
    ss.getSheetByName("Batch List").getRange(1,2,1,1).setDataValidation(getUniqueListofBatchNumber);
}


function updateBatchList() {

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const getBestExpListLastRow = ss.getSheetByName("tblBestExp").getLastRow();
    const getTblUniqueINIDLastRow = ss.getSheetByName("tblUniqueINID").getLastRow();
    let getBatchListLastRow = ss.getSheetByName("Batch List").getLastRow();
    let cellOfBatchNoList = ss.getSheetByName("Batch List").getRange(1,2,1,1).getValue();

    let getListOfBestExparray = ss.getSheetByName("tblBestExp").getRange(2,1,getBestExpListLastRow-1,11).getValues();
    let getListOfUniqueID = ss.getSheetByName("tblUniqueINID").getRange(2,1,getTblUniqueINIDLastRow-1,5).getValues();
    let getMasterList = ss.getSheetByName("MasterL").getRange(2,1,getMasterListLastRow-1,17).getValues();
    let getMasterListJustCodes = ss.getSheetByName("MasterL").getRange(2,1,getMasterListLastRow-1,1).getValues();

    let cleanedgetMasterListJustCodes = [];
    for (a= 0; a < getMasterListJustCodes.length; a++){
      cleanedgetMasterListJustCodes.push(getMasterListJustCodes[a].toString());
    }

    //console.log(getListOfBestExparray);
    //console.log(getListofBatchNumber);
    //console.log(getUniqueListofBatchNumber);
    //console.log(getListOfUniqueID)
    //console.log(cellOfBatchNoList)
    
    // To update the validation drop down list for Batch Numbers from tblBestExp
    //ss.getSheetByName("Batch List").getRange(1,2,1,1).setDataValidation(getUniqueListofBatchNumber);
    
    // Just to see what is the update Batch Number List
    //let a = ss.getSheetByName("Batch List").getRange(1,2,1,1).getDataValidation().getCriteriaValues();
    //console.log(a);

    // To Update List of Batch Items
        updateOUTGOING();
    let listOfQOHSingleCount = remainingIDLeftForQOHList;
    let getItemCodeValueForLookUp = '';
    let getAPMValue = 0;
    let getTestPerKitValue = 0;
    let extractListOfQOH = [];
    let extractListOfQOHListForCount = [];
    let getAcceptableQuantity = 0;
    let calculateAcceptableArray = [];

    // Count current QOH and look for APM and Test Per Kit
    for (t = 0; t < listOfQOHSingleCount.length; t++){
    for (u = 0; u < getListOfUniqueID.length; u++){
        
        if (listOfQOHSingleCount[t] === getListOfUniqueID[u][0] )
        
        getItemCodeValueForLookUp = getListOfUniqueID[u][2];
        
    for (v = 0; v < getMasterList.length; v++){
        
        if (getItemCodeValueForLookUp === getMasterList[v][0]){
            getAPMValue = getMasterList[v][8];
            getTestPerKitValue = getMasterList[v][12];
            getMultiCountValue = getMasterList[v][7];
            
          }
        }
      }
            extractListOfQOH.push([
            getItemCodeValueForLookUp,
            getAPMValue,
            getTestPerKitValue,
            getMultiCountValue
            ])

            extractListOfQOHListForCount.push([
            getItemCodeValueForLookUp
            ])
    }
//console.log(extractListOfQOH)

    // Take extractListOfQOH and count each item
    let countExtractListOfQOHElem = countArrayElem(extractListOfQOHListForCount);
    let countExtractListOfQOH = countExtractListOfQOHElem[2];
    let countExtractListOfQOHWithAPMandTestPerKit = extractListOfQOH;
    
    //console.log(countExtractListOfQOH);
    //console.log(countExtractListOfQOHWithAPMandTestPerKit);

    // Calculate the QOHLAST of each current QOH items
    let countedQOHListWithQOHlastarray = [];
    let countForItemCodesOnly = [];
    let getItemCodeValue = '';
    let getQOHValue = 0;
    let getAPMValueFromArray = 0;
    let getTestPerKitFromArray = 0;
    let getQOHLastValue = 0;
    
    for (y = 0; y < countExtractListOfQOH.length; y++){
    for (x = 0; x < countExtractListOfQOHWithAPMandTestPerKit.length; x++){
      if (countExtractListOfQOH[y][0] === countExtractListOfQOHWithAPMandTestPerKit[x][0]){
         getItemCodeValue = countExtractListOfQOH[y][0];
         getQOHValue = countExtractListOfQOH[y][1];
         getAPMValueFromArray = countExtractListOfQOHWithAPMandTestPerKit[x][1];
         getTestPerKitFromArray = countExtractListOfQOHWithAPMandTestPerKit[x][2];
         getMultiCountFromArray = countExtractListOfQOHWithAPMandTestPerKit[x][3];
         getQOHLastValue = getQOHValue / getMultiCountFromArray / getAPMValueFromArray;
    }
    }
    countedQOHListWithQOHlastarray.push([
      getItemCodeValue,         // Item Code of current QOH
      getQOHValue,              // QOH of current QOH
      getQOHLastValue,          // QOH Last of current QOH
      getAPMValueFromArray,     // APM of current QOH
      getTestPerKitFromArray    // TestPerKit of current QOH
    ]);
    countForItemCodesOnly.push(getItemCodeValue);
    }
    //console.log(countedQOHListWithQOHlastarray);
    //console.log(cleanedgetMasterListJustCodes)
    //console.log(countForItemCodesOnly)

    let remainingItemCodeNotCounted = findRemainingUniqueID(cleanedgetMasterListJustCodes.sort(),countForItemCodesOnly.sort());

    //console.log(remainingItemCodeNotCounted[0])
    // Expand the remaining zero QOH
    let toBeConcatWithMainQOH = [];
    for (a = 0 ; a < remainingItemCodeNotCounted.length; a++){
      getAPMZero = getMasterList[cleanedgetMasterListJustCodes.indexOf(remainingItemCodeNotCounted[a])][8];
      getTestPerKitZero = getMasterList[cleanedgetMasterListJustCodes.indexOf(remainingItemCodeNotCounted[a])][12];
      toBeConcatWithMainQOH.push([
        remainingItemCodeNotCounted[a],
        0,
        0,
        getAPMZero,
        getTestPerKitZero
      ])
    }
//console.log(toBeConcatWithMainQOH)

    let newArrayForCounting = countedQOHListWithQOHlastarray.concat(toBeConcatWithMainQOH);

let getTheBatchListFromChoosenBatchNumber = [];
let getInfoArrayForNewCalculation = [];


    for (i = 0; i < getListOfBestExparray.length; i++){

    for (j = 0; j < newArrayForCounting.length; j++){
          
      if (getListOfBestExparray[i][1] === cellOfBatchNoList &&
          getListOfBestExparray[i][4] === newArrayForCounting[j][0]){
         
         getTestPerKitFromQOH = newArrayForCounting[j][4];
         getAPMFromQOH = newArrayForCounting[j][3];
         getQOHLastFromQOH = newArrayForCounting[j][2];
         getBestExpValue = getListOfBestExparray[i][9];
         qtyRemaining = getListOfBestExparray[i][8];


         getPRNumber = getListOfBestExparray[i][2];
         getPONumber = getListOfBestExparray[i][3];
         getItemCode = getListOfBestExparray[i][4];
         getItemName = getListOfBestExparray[i][5];
         getItemType = getListOfBestExparray[i][6];
         getQtyOrdered = getListOfBestExparray[i][7];
         getQtyRemaining = getListOfBestExparray[i][8];
         getExpDateOffered = getListOfBestExparray[i][9];
      

         getInfoArrayForNewCalculation.push([
                  getPRNumber,            // 0 PR Number
                  getPONumber,            // 1 PO Number
                  getItemCode,            // 2 Item Code
                  getTestPerKitFromQOH,   // 3
                  getAPMFromQOH,          // 4
                  getQOHLastFromQOH,      // 5
                  getBestExpValue,        // 6
                  qtyRemaining            // 7
         ]);
  
          getTheBatchListFromChoosenBatchNumber.push([
                  getPRNumber,            // PR Number
                  getPONumber,            // PO Number
                  getItemCode,            // Item Code
                  getItemName,            // Item Name
                  getItemType,            // Item Type
                  getQtyOrdered,          // Quantity Ordered
                  getQtyRemaining,        // Quantity Remaining
                  getExpDateOffered,      // Exp Date Offered
                  '',                     // Acceptable Quantity Count to get it from another array
                  false,                  // To auto check box
                  '',                     // For correction value, if different from the acceptable quantity (Minus or Plus kept from future APM review?)
                  '',                     // Auto comment on receiving Best Exp
                  '',                     // Remarks
                  getQOHLastFromQOH,      // 13th column
                  getAPMFromQOH,          // 14th column
                  getTestPerKitFromQOH    // 15th column
        ]);
    }
  }
}
//console.log(getInfoArrayForNewCalculation);
//console.log('QOH Last: '+getQOHLastFromQOH);
//console.log('Get APM: '+getAPMFromQOH);
//console.log(getTheBatchListFromChoosenBatchNumber);

// Go through each item to calculate the acceptable quantity
let resOfAcceptableQuantArray = [];
for (k = 0; k < getInfoArrayForNewCalculation.length; k++){

    arrayOfListToCountAcceptableKit = makeArrayListOfNewOrderBasedOnTest(getInfoArrayForNewCalculation[k][3],  // Test per kit
                                                                         getInfoArrayForNewCalculation[k][4],  // APM
                                                                         getInfoArrayForNewCalculation[k][5],  // QOH last
                                                                         getInfoArrayForNewCalculation[k][6],  // Best Exp
                                                                         getInfoArrayForNewCalculation[k][7]); // Qty remaining
    resOfAcceptableQuant = arrayOfListToCountAcceptableKit[1];

    resOfAcceptableQuantArray.push([
                getInfoArrayForNewCalculation[k][0], // PR Number
                getInfoArrayForNewCalculation[k][1], // PO Number
                getInfoArrayForNewCalculation[k][2], // Item Code
                resOfAcceptableQuant                 // Acceptable Quantity
                ]);
}
//console.log(resOfAcceptableQuantArray);

// Combine all these arrays, try with simple push first without matching PR, PO and Item - assuming all rows are the same

const today = new Date();
const convertToSeconds = 4*7*24*60*60*1000;
//newQOHLastDate = new Date(today.getTime()+(newPRoverAPMplusQOHlast*convertToSeconds));


let finalArrayOfBatchList = [];
for (l = 0; l < getTheBatchListFromChoosenBatchNumber.length; l++){

// New estimated date for QOH + Acceptable Quantity
newQOHLast = getTheBatchListFromChoosenBatchNumber[l][13] * convertToSeconds;
newQOHplusBOLast = (getTheBatchListFromChoosenBatchNumber[l][13] + (resOfAcceptableQuantArray[l][3]/getTheBatchListFromChoosenBatchNumber[l][14])) * convertToSeconds;
//console.log(newQOHLast)
//console.log(newQOHplusBOLast)

  finalArrayOfBatchList.push([
    getTheBatchListFromChoosenBatchNumber[l][0],  // PR Number
    getTheBatchListFromChoosenBatchNumber[l][1],  // PO Number
    getTheBatchListFromChoosenBatchNumber[l][2],  // Item Code
    getTheBatchListFromChoosenBatchNumber[l][3],  // Item Name
    getTheBatchListFromChoosenBatchNumber[l][4],  // Item Type
    getTheBatchListFromChoosenBatchNumber[l][5],  // Quantity Ordered
    getTheBatchListFromChoosenBatchNumber[l][6],  // Quantity Remaining
    getTheBatchListFromChoosenBatchNumber[l][7],  // Exp Date Offered
    resOfAcceptableQuantArray[l][3],              // Acceptable Quantity
    newQOHLastDate = new Date(today.getTime()+newQOHLast), // QOH Estimated Date
    newQOHplusBOLastDate = new Date(today.getTime()+newQOHplusBOLast),    // New Estimated Date


    getTheBatchListFromChoosenBatchNumber[l][9],  // FALSE
    getTheBatchListFromChoosenBatchNumber[l][10], // Empty for correction value
    getTheBatchListFromChoosenBatchNumber[l][11], // Auto comment
    getTheBatchListFromChoosenBatchNumber[l][12]  // Manual Remarks
    ]);

}
//console.log(finalArrayOfBatchList)

// Clear Batch List first if Batch Number changes
if (getBatchListLastRow-3 === 0){
} else {ss.getSheetByName("Batch List").getRange(4,1,getBatchListLastRow-3,15).clearContent().removeCheckboxes();
}

// Paste the Batch List Based on the choosen Batch Number
ss.getSheetByName("Batch List").getRange(4,1,finalArrayOfBatchList.length,finalArrayOfBatchList[0].length).setValues(finalArrayOfBatchList);

getBatchListLastRow = ss.getSheetByName("Batch List").getLastRow(); // To update the length of new Batch List

ss.getSheetByName("Batch List").getRange(4,12,getBatchListLastRow-3,1).insertCheckboxes();



}
