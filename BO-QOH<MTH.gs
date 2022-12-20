function updateNameSheet() {

  // var t1 = new Date().getTime();

  updatePRList();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const getTblPRLastRow = ss.getSheetByName("tblPR").getLastRow();
  const getMonthToFilterValue = ss.getSheetByName("BO-QOH<MTH").getRange("J1").getValue();

  let getTblPRListArray = ss.getSheetByName("tblPR").getRange(2,1,getTblPRLastRow-1,10).getValues();
  let getTblBOQOHMTHLastRow = ss.getSheetByName("BO-QOH<MTH").getLastRow();

  // Start copy of code from updateBatchList
  //-------------------------------------------------------------------------------------------------------------------
  // To Update List of QOH
    const getTblUniqueINIDLastRow = ss.getSheetByName("tblUniqueINID").getLastRow();
    let getListOfUniqueID = ss.getSheetByName("tblUniqueINID").getRange(2,1,getTblUniqueINIDLastRow-1,5).getValues();
    let getMasterList = ss.getSheetByName("MasterL").getRange(2,1,getMasterListLastRow-1,17).getValues();

        updateOUTGOING();
    let listOfQOHSingleCount = remainingIDLeftForQOHList;
    let getItemCodeValueForLookUp = '';
    let getAPMValue = 0;
    let getTestPerKitValue = 0;
    let getMultiCountValue = 0;
    let getItemType = '';
    let extractListOfQOH = [];
    let extractListOfQOHListForCount = [];

    // Count current QOH and look for APM and Test Per Kit
    for (t = 0; t < listOfQOHSingleCount.length; t++){
    for (u = 0; u < getListOfUniqueID.length; u++){
        
        if (listOfQOHSingleCount[t] === getListOfUniqueID[u][0]){
        
        getItemCodeValueForLookUp = getListOfUniqueID[u][2];
        // Do not break here
        
    for (v = 0; v < getMasterList.length; v++){
        
        if (getItemCodeValueForLookUp === getMasterList[v][0]){
            getAPMValue = getMasterList[v][8];
            getTestPerKitValue = getMasterList[v][12];
            getMultiCountValue = getMasterList[v][7];
            getItemType = getMasterList[v][2];
            break; // New
            
          }
        }
      }
    }
            extractListOfQOH.push([
            getItemCodeValueForLookUp,
            getAPMValue,
            getTestPerKitValue,
            getMultiCountValue,
            getItemType
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
         getItemTypeFromArray = countExtractListOfQOHWithAPMandTestPerKit[x][4];
         getQOHinBoxes = getQOHValue / getMultiCountFromArray;
         getQOHLastValue = getQOHValue / getMultiCountFromArray / getAPMValueFromArray;
         break; // New
    }
    }
    countedQOHListWithQOHlastarray.push([
      getItemCodeValue,         // Item Code of current QOH
      getItemTypeFromArray,     // Item Type
      getQOHValue,              // QOH in cartridges
      getQOHinBoxes,            // QOH in boxes
      getQOHLastValue,          // QOH Last of current QOH
      getAPMValueFromArray,     // APM of current QOH
      getTestPerKitFromArray    // TestPerKit of current QOH
    ])
    }
    //console.log(countedQOHListWithQOHlastarray);

  // End copy of code from updateBatchList
  //-------------------------------------------------------------------------------------------------------------------

  // Get all the active PR from tblPR
  let findAllActivePRList = [];
  for (a = 0; a < getTblPRListArray.length; a++){
    if (getTblPRListArray[a][7] != 0){
      findAllActivePRList.push([
                    getTblPRListArray[a][5],    // Vesalius Code
                    getTblPRListArray[a][2],    // Item Code
                    getTblPRListArray[a][3],    // Item Name
                    getTblPRListArray[a][4],    // Item Type
                    getTblPRListArray[a][7]     // Pending BO in Boxes
                    ]);
    }
  }
  //console.log(findAllActivePRList)

  // Count all the Item codes with their pending BO if there are multiple active PR
  let countAllPendingBOBasedOnItemCodes = countUpItemFromPR(findAllActivePRList);
  //console.log(countAllPendingBOBasedOnItemCodes)
  let resArrayForPaste = [];

  for (b = 0; b < countAllPendingBOBasedOnItemCodes.length; b++){
    for (c = 0; c < countedQOHListWithQOHlastarray.length; c++){

      today = new Date()
      getMonthValue = getMonthToFilterValue;
      maxDateValue = new Date(today.getTime()+(getMonthValue*4*7*24*60*60*1000));
      qohLastValue = countedQOHListWithQOHlastarray[c][4];
      qohLastInDateFormat = new Date(today.getTime()+(qohLastValue*4*7*24*60*60*1000));

    if (countAllPendingBOBasedOnItemCodes[b][1] === countedQOHListWithQOHlastarray[c][0] &&
        qohLastInDateFormat <= maxDateValue){

      resArrayForPaste.push([
                    countAllPendingBOBasedOnItemCodes[b][0],   // Vesalius Code
                    countAllPendingBOBasedOnItemCodes[b][1],   // Item Code
                    countAllPendingBOBasedOnItemCodes[b][2],   // Item Name
                    countedQOHListWithQOHlastarray[c][1],      // Item Type
                    countedQOHListWithQOHlastarray[c][2],      // Qty in cartridges
                    countedQOHListWithQOHlastarray[c][3],      // Qty in boxes
                    countAllPendingBOBasedOnItemCodes[b][3],   // Pending BO in Kits
                    qohLastInDateFormat                        // QOH last in date
                    ]);
                    break; // New


    }
  }
  }
  // console.log(resArrayForPaste)

  // Clear off the sheet first
  if (getTblBOQOHMTHLastRow-1 === 0){
  } else {
  ss.getSheetByName("BO-QOH<MTH").getRange(2,1,getTblBOQOHMTHLastRow-1,8).clearContent();
  ss.getSheetByName("BO-QOH<MTH").getRange(2,1,getTblBOQOHMTHLastRow-1,8).getFilter().remove();
  }

  // Paste the new array
  ss.getSheetByName("BO-QOH<MTH").getRange(2,1,resArrayForPaste.length,resArrayForPaste[0].length).setValues(resArrayForPaste);
  ss.getSheetByName("BO-QOH<MTH").getRange(1,1,resArrayForPaste.length+1,8).createFilter().sort(8,true);


  // var t2 = new Date().getTime();
  // var timeDiff = t2 - t1;
  // console.log(timeDiff); // 68840 ms before update, reduced to 9621 ms after added breaks

}
