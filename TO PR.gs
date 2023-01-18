let matchItemCodeForQOH = []; // For finding last PO value for QOH items
let matchItemCodeForQOHBO = []; // For finding last PO value for QOH + BO items
let matchItemCodeForZeroQOH = []; // For finding last PO value for zero QOH

function lookForItemsBelowPar() {

  // var t1 = new Date().getTime();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const getMasterListLastCol = ss.getSheetByName("MasterL").getLastColumn();
  const getToPRLastRow = ss.getSheetByName("To PR").getLastRow();

  let getMasterListArray = ss.getSheetByName("MasterL").getRange(2, 1, getMasterListLastRow - 1, getMasterListLastCol).getValues();

  // Start copy of code from BO-QOH<MTH
  //-------------------------------------------------------------------------------------------------------------------
  // To Update List of QOH
  const getTblUniqueINIDLastRow = ss.getSheetByName("tblUniqueINID").getLastRow();
  const getTblPRLastRow = ss.getSheetByName("tblPR").getLastRow();

  let getTblPRListArray = ss.getSheetByName("tblPR").getRange(2, 1, getTblPRLastRow - 1, 10).getValues();
  let getListOfUniqueID = ss.getSheetByName("tblUniqueINID").getRange(2, 1, getTblUniqueINIDLastRow - 1, 5).getValues();
  let getMasterList = ss.getSheetByName("MasterL").getRange(2, 1, getMasterListLastRow - 1, 17).getValues();
  // console.log(getListOfUniqueID.length);

  updateOUTGOING();
  //updatePRList here
  let listOfQOHSingleCount = remainingIDLeftForQOHList;
  let getItemCodeValueForLookUp = '';
  let getAPMValue = 0;
  let getTestPerKitValue = 0;
  let extractListOfQOH = [];
  let extractListOfQOHListForCount = [];
  // console.log(listOfQOHSingleCount.length)

  // Count current QOH and look for APM and Test Per Kit
  for (t = 0; t < listOfQOHSingleCount.length; t++) {
    for (u = 0; u < getListOfUniqueID.length; u++) {
      if (listOfQOHSingleCount[t] === getListOfUniqueID[u][0]) {

        getItemCodeValueForLookUp = getListOfUniqueID[u][2];
        // Add break here?

        for (v = 0; v < getMasterList.length; v++) {

          if (getItemCodeValueForLookUp === getMasterList[v][0]) {
            getAPMValue = getMasterList[v][8];
            getTestPerKitValue = getMasterList[v][12];
            getMultiCountValue = getMasterList[v][7];
            getItemType = getMasterList[v][2];
            break; // New

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
    }
  }
  // console.log(extractListOfQOH.length)
  // console.log(remainingIDLeftForQOHList)
  // console.log(extractListOfQOHListForCount.length)

  // Take extractListOfQOH and count each item
  let countExtractListOfQOHElem = countArrayElem(extractListOfQOHListForCount);
  let countExtractListOfQOH = countExtractListOfQOHElem[2];
  let countExtractListOfQOHWithAPMandTestPerKit = extractListOfQOH;

  // console.log(countExtractListOfQOH);
  //console.log(countExtractListOfQOHWithAPMandTestPerKit);

  // Calculate the QOHLAST of each current QOH items
  let countedQOHListWithQOHlastarray = [];
  let getItemCodeValue = '';
  let getQOHValue = 0;
  let getAPMValueFromArray = 0;
  let getTestPerKitFromArray = 0;
  let getQOHLastValue = 0;

  for (y = 0; y < countExtractListOfQOH.length; y++) {
    for (x = 0; x < countExtractListOfQOHWithAPMandTestPerKit.length; x++) {
      if (countExtractListOfQOH[y][0] === countExtractListOfQOHWithAPMandTestPerKit[x][0]) {
        getItemCodeValue = countExtractListOfQOH[y][0];
        getQOHValue = countExtractListOfQOH[y][1];
        getAPMValueFromArray = countExtractListOfQOHWithAPMandTestPerKit[x][1];
        getTestPerKitFromArray = countExtractListOfQOHWithAPMandTestPerKit[x][2];
        getMultiCountFromArray = countExtractListOfQOHWithAPMandTestPerKit[x][3];
        getItemTypeFromArray = countExtractListOfQOHWithAPMandTestPerKit[x][4];
        getQOHinBoxes = getQOHValue / getMultiCountFromArray;
        getQOHLastValue = getQOHValue / getMultiCountFromArray / getAPMValueFromArray;
        //break; // New
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

  // Get all the active PR from tblPR
  let findAllActivePRList = [];
  for (a = 0; a < getTblPRListArray.length; a++) {
    if (getTblPRListArray[a][7] != 0) {
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
  let resArrayForPendingBO = [];
  let resArrayForPendingBOJustItemCodes = [];
  let countAllPendingBOBasedOnItemCodesJustCodes = [];

  for (d = 0; d < countAllPendingBOBasedOnItemCodes.length; d++) {
    countAllPendingBOBasedOnItemCodesJustCodes.push(countAllPendingBOBasedOnItemCodes[d][1]);
  }

  for (b = 0; b < countAllPendingBOBasedOnItemCodes.length; b++) {
    for (c = 0; c < countedQOHListWithQOHlastarray.length; c++) {

      const today = new Date();
      qohLastValue = countedQOHListWithQOHlastarray[c][4];
      qohLastInDateFormat = new Date(today.getTime() + (qohLastValue * 4 * 7 * 24 * 60 * 60 * 1000));

      qohPlusBOInKits = countedQOHListWithQOHlastarray[c][3] + countAllPendingBOBasedOnItemCodes[b][3];
      qohPlusBOLastInMonths = qohPlusBOInKits / countedQOHListWithQOHlastarray[c][5];
      qohPlusBOLastInDateFormat = new Date(today.getTime() + (qohPlusBOLastInMonths * 4 * 7 * 24 * 60 * 60 * 1000));

      if (countAllPendingBOBasedOnItemCodes[b][1] === countedQOHListWithQOHlastarray[c][0]) {

        resArrayForPendingBO.push([
          countAllPendingBOBasedOnItemCodes[b][0],   // Vesalius Code
          countAllPendingBOBasedOnItemCodes[b][1],   // Item Code
          countAllPendingBOBasedOnItemCodes[b][2],   // Item Name
          countedQOHListWithQOHlastarray[c][1],      // Item Type
          countedQOHListWithQOHlastarray[c][2],      // Qty in cartridges
          countedQOHListWithQOHlastarray[c][3],      // Qty in boxes
          countAllPendingBOBasedOnItemCodes[b][3],   // Pending BO in Kits
          qohPlusBOInKits,                           // QOH + Bo in kits
          qohLastValue,                              // QOH last in months
          qohPlusBOLastInMonths,                     // QOH + BO in months
          qohLastInDateFormat,                       // QOH last in date
          qohPlusBOLastInDateFormat                  // QOH + BO in months date format
        ]);
        resArrayForPendingBOJustItemCodes.push(countAllPendingBOBasedOnItemCodes[b][1]);
        //break; // New


      }
    }
  }
  //console.log(resArrayForPendingBO)
  //console.log(resArrayForPendingBO.length)
  //console.log(typeof countAllPendingBOBasedOnItemCodesJustCodes[0])
  //console.log(typeof countAllPendingBOBasedOnItemCodes[0][1])

  findRemainingUniqueID(countAllPendingBOBasedOnItemCodesJustCodes.sort(), resArrayForPendingBOJustItemCodes.sort());

  let newBOArray = [];
  for (a = 0; a < countAllPendingBOBasedOnItemCodesJustCodes.length; a++) {
    for (b = 0; b < countAllPendingBOBasedOnItemCodes.length; b++) {

      if (countAllPendingBOBasedOnItemCodesJustCodes[a] === countAllPendingBOBasedOnItemCodes[b][1]) {
        newBOArray.push([
          countAllPendingBOBasedOnItemCodes[b][0],    // Vesalius Code
          countAllPendingBOBasedOnItemCodes[b][1],    // Item Code
          countAllPendingBOBasedOnItemCodes[b][2],    // Item Name
          countAllPendingBOBasedOnItemCodes[b][3]    // Pending BO in Boxes
        ]);
        //break; // New
      }
    }
  }
  //console.log(newBOArray)

  let toAddOnArray = [];

  for (a = 0; a < newBOArray.length; a++) {
    for (b = 0; b < getMasterList.length; b++) {

      const today3 = new Date();
      qohLastValue = 0;
      qohLastInDateFormat = new Date(today3.getTime() + (qohLastValue * 4 * 7 * 24 * 60 * 60 * 1000));

      if (newBOArray[a][1] === getMasterList[b][0]) {

        qohPlusBOInKits = newBOArray[a][3];
        qohPlusBOLastInMonths = qohPlusBOInKits / getMasterList[b][8];
        qohPlusBOLastInDateFormat = new Date(today3.getTime() + (qohPlusBOLastInMonths * 4 * 7 * 24 * 60 * 60 * 1000));

        toAddOnArray.push([
          newBOArray[a][0],                                   // Vesalius Code
          newBOArray[a][1],                                   // Item Code
          newBOArray[a][2],                                   // Item Name
          getMasterList[b][2],                                // Item Type
          0,                                                  // Qty in cartridges
          0,                                                  // Qty in boxes
          newBOArray[a][3],                                   // Pending BO in Kits
          qohPlusBOInKits,                                    // QOH + Bo in kits
          qohLastValue,                                       // QOH last in months
          qohPlusBOLastInMonths,                              // QOH + BO in months
          qohLastInDateFormat,                                // QOH last in date
          qohPlusBOLastInDateFormat                           // QOH + BO in months date format
        ]);
        //break; // New
      }
    }
  }
  //console.log(toAddOnArray)

  let concatResArrayForPendingBO = resArrayForPendingBO.concat(toAddOnArray);

  // End copy of code from BO-QOH<MTH
  //-------------------------------------------------------------------------------------------------------------------


  // Only look for PR items array and find their QOH, par level
  let findPRItemsOnly = [];
  let getItemCodesOnly = [];
  for (d = 0; d < getMasterListArray.length; d++) {
    if (getMasterListArray[d][5] === 'PR') {

      getItemCodesOnly.push(
        getMasterListArray[d][0]       // Item Code
      );

      findPRItemsOnly.push([
        getMasterListArray[d][0],       // Item Code
        getMasterListArray[d][1],       // Item Name
        getMasterListArray[d][2],       // Item Type
        getMasterListArray[d][8],       // APM
        getMasterListArray[d][9],       // Vesalius Code
        getMasterListArray[d][10],      // Item Full Description
        getMasterListArray[d][11],      // List Code
        getMasterListArray[d][12],      // Test Per Kit
        getMasterListArray[d][13],      // UOM
        getMasterListArray[d][16],      // Company
        getMasterListArray[d][18],      // Par Level
      ]);

    }
  }
  //console.log(findPRItemsOnly)


  // Multiple loops of filters start here

  // Use to filter out
  let fullArrayQOHandBO = [];
  let fullArrayQOH = [];
  let filteredQOH = [];

  // Look for items to order
  let needToOrderQOHandBO = [];
  let needToOrderQOH = [];
  let ordQOHBOItemCodeLessPar = [];
  let ordQOHBOItemCodeMorePar = [];
  let ordQOHItemCodesLessPar = [];
  let ordQOHItemCodesMorePar = [];
  //let needToOrderLeftFirst = [];
  //let needToOrderLeftSecond = [];
  let reorderQty = 0;


  // QOH+BO < Par Level
  for (a = 0; a < findPRItemsOnly.length; a++) {
    for (b = 0; b < concatResArrayForPendingBO.length; b++) {

      // For filter purpose
      if (findPRItemsOnly[a][0] === concatResArrayForPendingBO[b][1]) {
        parlevel = findPRItemsOnly[a][3] * findPRItemsOnly[a][10];

        fullArrayQOHandBO.push(
          concatResArrayForPendingBO[b][1]
        );
        //break; // New
      }

      if (findPRItemsOnly[a][0] === concatResArrayForPendingBO[b][1] &&
        concatResArrayForPendingBO[b][7] <= parlevel) {   //   ORIGINAL is <= parlevel

        const today = new Date();
        qohLastValue = concatResArrayForPendingBO[b][8];
        qohLastInDateFormat = new Date(today.getTime() + (qohLastValue * 4 * 7 * 24 * 60 * 60 * 1000));

        reorderQty = Math.round((findPRItemsOnly[a][3] * findPRItemsOnly[a][10]) - concatResArrayForPendingBO[b][7]);

        matchItemCodeForQOHBO.push(concatResArrayForPendingBO[b][1]);

        prLineFirst = `Reagent for Alinity Analysers. (Reagent lease scheme)` + "\r\n" +
          `1 ${findPRItemsOnly[a][8]} = ${findPRItemsOnly[a][7]} TESTS` + "\r\n" +
          `Item Code = ${findPRItemsOnly[a][6]}` + "\r\n" +
          `QOH = ${Math.round(concatResArrayForPendingBO[b][5])} ${findPRItemsOnly[a][8]}` + "\r\n" +
          `Kindly proceed with P.O as expiry date will be liaised with supplier directly before delivery arrangement.`
        //

        needToOrderQOHandBO.push([
          concatResArrayForPendingBO[b][0],    // Vesalius
          concatResArrayForPendingBO[b][1],    // Item Code
          concatResArrayForPendingBO[b][2],    // Item Name
          concatResArrayForPendingBO[b][3],    // Item Type
          concatResArrayForPendingBO[b][4],    // Qty in cartridges
          concatResArrayForPendingBO[b][5],    // Qty in boxes
          concatResArrayForPendingBO[b][6],    // Pending BO in Kits
          concatResArrayForPendingBO[b][7],    // QOH + Bo in kits
          concatResArrayForPendingBO[b][8],    // QOH last in months
          concatResArrayForPendingBO[b][9],    // QOH + BO in months
          concatResArrayForPendingBO[b][10],   // QOH last in date
          concatResArrayForPendingBO[b][11],   // QOH + BO in months date format
          findPRItemsOnly[a][6],               // List Code
          qohLastInDateFormat,                 // QOH last date
          reorderQty,                          // Reorder qty = (APM * Target) - QOH in kits - Pending BO in kits
          prLineFirst,                         // PR Remarks = Line + QOH in kits + List Code + Line
          // Need to append Last PO value from a different array, length must match
          findPRItemsOnly[a][9]                // Company
        ]);

        ordQOHBOItemCodeLessPar.push(concatResArrayForPendingBO[b][1]);
        //break; // New

      } else

        if (findPRItemsOnly[a][0] === concatResArrayForPendingBO[b][1] &&
          concatResArrayForPendingBO[b][7] > parlevel) {

          ordQOHBOItemCodeMorePar.push(concatResArrayForPendingBO[b][1]);
          //break; // New

        }
    }
  }
  //console.log(fullArrayQOHandBO)
  //console.log(ordQOHBOItemCodeLessPar);
  //console.log(ordQOHBOItemCodeMorePar);


  // QOH < Par Level
  for (a = 0; a < findPRItemsOnly.length; a++) {
    for (b = 0; b < countedQOHListWithQOHlastarray.length; b++) {

      if (findPRItemsOnly[a][0] === countedQOHListWithQOHlastarray[b][0]) {
        fullArrayQOH.push(
          countedQOHListWithQOHlastarray[b][0]
        );
        //break; // New
      }
    }
  }

  filteredQOH = findRemainingUniqueID(fullArrayQOH.sort(), fullArrayQOHandBO.sort());

  //console.log(filteredQOH)
  const today = new Date();
  let reorderQty2 = 0;
  //  let matchItemCode = []; // Make variable global

  for (a = 0; a < filteredQOH.length; a++) {
    for (b = 0; b < countedQOHListWithQOHlastarray.length; b++) {
      for (c = 0; c < findPRItemsOnly.length; c++) {

        if (filteredQOH[a] === countedQOHListWithQOHlastarray[b][0] &&
          filteredQOH[a] === findPRItemsOnly[c][0]) {
          parlevel = findPRItemsOnly[c][3] * findPRItemsOnly[c][10];

          if (countedQOHListWithQOHlastarray[b][3] <= parlevel) {  //   ORIGINAL is <= parlevel

            pendingBOVal = 0;

            qohLastValue = countedQOHListWithQOHlastarray[b][4];
            qohLastInDateFormat = new Date(today.getTime() + (qohLastValue * 4 * 7 * 24 * 60 * 60 * 1000));

            reorderQty2 = Math.round((countedQOHListWithQOHlastarray[b][5] * findPRItemsOnly[c][10]) - countedQOHListWithQOHlastarray[b][3]);

            matchItemCodeForQOH.push(countedQOHListWithQOHlastarray[b][0]);

            prLineFirst = `Reagent for Alinity Analysers. (Reagent lease scheme)` + "\r\n" +
              `1 ${findPRItemsOnly[c][8]} = ${findPRItemsOnly[c][7]} TESTS` + "\r\n" +
              `Item Code = ${findPRItemsOnly[c][6]}` + "\r\n" +
              `QOH = ${Math.round(countedQOHListWithQOHlastarray[b][3])} ${findPRItemsOnly[c][8]}` + "\r\n" +
              `Kindly proceed with P.O as expiry date will be liaised with supplier directly before delivery arrangement.`

            needToOrderQOH.push([
              countedQOHListWithQOHlastarray[b][0],    // Item Code of current QOH
              countedQOHListWithQOHlastarray[b][1],    // Item Type
              countedQOHListWithQOHlastarray[b][2],    // QOH in cartridges
              countedQOHListWithQOHlastarray[b][3],    // QOH in boxes
              countedQOHListWithQOHlastarray[b][4],    // QOH Last of current QOH
              countedQOHListWithQOHlastarray[b][5],    // APM of current QOH
              countedQOHListWithQOHlastarray[b][6],    // TestPerKit of current QOH
              findPRItemsOnly[c][1],                   // Item Name
              findPRItemsOnly[c][6],                   // List Code
              pendingBOVal,                            // Pending BO = 0
              qohLastInDateFormat,                     // QOH last date
              qohLastInDateFormat,                     // QOH + BO last date = QOH last date
              findPRItemsOnly[c][4],                   // Vesalius Code
              reorderQty2,                             // Reorder qty = (APM * Target) - QOH in kits
              prLineFirst,                             // PR Remarks = Line + QOH in kits + List Code + Line
              // Need to append Last PO value from a different array, length must match
              findPRItemsOnly[c][9]                    // Company
            ]);

            ordQOHItemCodesLessPar.push(countedQOHListWithQOHlastarray[b][0]);
            //break; // New

          } else if (countedQOHListWithQOHlastarray[b][3] > parlevel) {

            ordQOHItemCodesMorePar.push(countedQOHListWithQOHlastarray[b][0]);
            //break; // New
          }
        }
      }
    }
  }
  //console.log(needToOrderQOH)

  // To append to needToOrderQOH for last PO value
  let matchedLastPOvalueforQOHArray = findLastPOforQOH();
  //console.log(matchedLastPOvalueforQOHArray);

  // To append to needToOrderQOHBO for last PO value
  let matchedLastPOvalueforQOHBOArray = findLastPOforQOHBO();
  //console.log(matchedLastPOvalueforQOHBOArray.length);
  //console.log(needToOrderQOHandBO.length);

  // Append QOH+BO with their last PO values
  let firstAppendOfQOHBO = [];
  for (a = 0; a < needToOrderQOHandBO.length; a++) {

    firstAppendOfQOHBO.push([
      needToOrderQOHandBO[a][1],            // Item Code
      needToOrderQOHandBO[a][2],            // Item Name
      needToOrderQOHandBO[a][3],            // Item Type
      needToOrderQOHandBO[a][12],           // List Code
      needToOrderQOHandBO[a][4],            // Quantity in cartridges
      needToOrderQOHandBO[a][5],            // Quantity in boxes
      needToOrderQOHandBO[a][6],            // Pending BO
      needToOrderQOHandBO[a][13],           // QOH last date
      needToOrderQOHandBO[a][11],           // QOH + BO last date
      "",                                   // Leave blank for PR no.
      needToOrderQOHandBO[a][0],            // Vesalius Code
      needToOrderQOHandBO[a][14],           // Reorder Qty
      "",                                   // Leave blank for New Qty
      needToOrderQOHandBO[a][15],           // PR Remarks
      matchedLastPOvalueforQOHBOArray[a],   // Last PO
      needToOrderQOHandBO[a][16]            // Company
    ])
  }
  //console.log(firstAppendOfQOHBO)

  let firstAppendOfQOH = [];
  for (a = 0; a < needToOrderQOH.length; a++) {

    firstAppendOfQOH.push([
      needToOrderQOH[a][0],               // Item Code
      needToOrderQOH[a][7],               // Item Name
      needToOrderQOH[a][1],               // Item Type
      needToOrderQOH[a][8],               // List Code
      needToOrderQOH[a][2],               // Quantity in cartridges
      needToOrderQOH[a][3],               // Quantity in boxes
      needToOrderQOH[a][9],               // Pending BO
      needToOrderQOH[a][10],              // QOH last date
      needToOrderQOH[a][11],              // QOH + BO last date
      "",                                 // Leave blank for PR no.
      needToOrderQOH[a][12],              // Vesalius Code
      needToOrderQOH[a][13],              // Reorder Qty
      "",                                 // Leave blank for New Qty
      needToOrderQOH[a][14],              // PR Remarks
      matchedLastPOvalueforQOHArray[a],   // Last PO
      needToOrderQOH[a][15]               // Company
    ])
  }
  //console.log(firstAppendOfQOH)

  // Final loop to find true zero items

  findRemainingUniqueID(getItemCodesOnly.sort(), ordQOHBOItemCodeLessPar.sort());
  //console.log(getItemCodesOnly);
  findRemainingUniqueID(getItemCodesOnly.sort(), ordQOHBOItemCodeMorePar.sort());
  //console.log(getItemCodesOnly);
  findRemainingUniqueID(getItemCodesOnly.sort(), ordQOHItemCodesLessPar.sort());
  //console.log(getItemCodesOnly);
  findRemainingUniqueID(getItemCodesOnly.sort(), ordQOHItemCodesMorePar.sort());
  //console.log(getItemCodesOnly);


  let arrayOfEmptyQOH = [];


  const today2 = new Date();
  let reorderQty3 = 0;
  //  let matchItemCode = []; // Make variable global


  for (b = 0; b < getItemCodesOnly.length; b++) {
    for (c = 0; c < findPRItemsOnly.length; c++) {

      if (getItemCodesOnly[b] === findPRItemsOnly[c][0]) {
        parlevel = findPRItemsOnly[c][3] * findPRItemsOnly[c][10];

        pendingBOVal = 0;

        qohLastValue = 0;
        qohLastInDateFormat = new Date(today2.getTime() + (qohLastValue * 4 * 7 * 24 * 60 * 60 * 1000));

        reorderQty3 = Math.round(parlevel);

        matchItemCodeForZeroQOH.push(getItemCodesOnly[b]);

        prLineFirst = `Reagent for Alinity Analysers. (Reagent lease scheme)` + "\r\n" +
          `1 ${findPRItemsOnly[c][8]} = ${findPRItemsOnly[c][7]} TESTS` + "\r\n" +
          `Item Code = ${findPRItemsOnly[c][6]}` + "\r\n" +
          `QOH = 0 ${findPRItemsOnly[c][8]}` + "\r\n" +
          `Kindly proceed with P.O as expiry date will be liaised with supplier directly before delivery arrangement.`

        arrayOfEmptyQOH.push([
          getItemCodesOnly[b],                     // Item Code of current QOH
          findPRItemsOnly[c][2],                   // Item Type
          0,                                       // QOH in cartridges
          0,                                       // QOH in boxes
          qohLastValue,                            // QOH Last of current QOH
          0,                                       // APM of current QOH
          findPRItemsOnly[c][7],                   // TestPerKit of current QOH
          findPRItemsOnly[c][1],                   // Item Name
          findPRItemsOnly[c][6],                   // List Code
          0,                                       // Pending BO = 0
          qohLastInDateFormat,                     // QOH last date
          qohLastInDateFormat,                     // QOH + BO last date = QOH last date
          findPRItemsOnly[c][4],                   // Vesalius Code
          reorderQty3,                             // Reorder qty = (APM * Target) - QOH in kits
          prLineFirst,                             // PR Remarks = Line + QOH in kits + List Code + Line
          // Need to append Last PO value from a different array, length must match
          findPRItemsOnly[c][9]                    // Company
        ]);
        //break; // New

      }
    }
  }


  let matchedLastPOvalueforZeroQOH = findLastPOforZeroQOH();

  let firstAppendOfZeroQOH = [];
  for (a = 0; a < arrayOfEmptyQOH.length; a++) {

    firstAppendOfZeroQOH.push([
      arrayOfEmptyQOH[a][0],               // Item Code
      arrayOfEmptyQOH[a][7],               // Item Name
      arrayOfEmptyQOH[a][1],               // Item Type
      arrayOfEmptyQOH[a][8],               // List Code
      arrayOfEmptyQOH[a][2],               // Quantity in cartridges
      arrayOfEmptyQOH[a][3],               // Quantity in boxes
      arrayOfEmptyQOH[a][9],               // Pending BO
      arrayOfEmptyQOH[a][10],              // QOH last date
      arrayOfEmptyQOH[a][11],              // QOH + BO last date
      "",                                  // Leave blank for PR no.
      arrayOfEmptyQOH[a][12],              // Vesalius Code
      arrayOfEmptyQOH[a][13],              // Reorder Qty
      "",                                  // Leave blank for New Qty
      arrayOfEmptyQOH[a][14],              // PR Remarks
      matchedLastPOvalueforZeroQOH[a],     // Last PO
      arrayOfEmptyQOH[a][15]               // Company
    ])
  }
  //console.log(firstAppendOfZeroQOH)

  // Last combine the two arrays using concat
  let finalAppendedArrayList = firstAppendOfQOHBO.concat(firstAppendOfQOH).concat(firstAppendOfZeroQOH);

  // console.log(finalAppendedArrayList);

  // Clear off the sheet to prep for new array result
  if (getToPRLastRow - 1 === 0) {
  } else {
    ss.getSheetByName("To PR").getRange(3, 1, getToPRLastRow - 1, 16).clearContent();
  }

  // Paste the final array to the To PR sheet
  ss.getSheetByName("To PR").getRange(3, 1, finalAppendedArrayList.length, finalAppendedArrayList[0].length).setValues(finalAppendedArrayList).sort({ column: 2, ascending: true });

  // To reset the row heights to 21
  ss.getSheetByName("To PR").setRowHeightsForced(3, finalAppendedArrayList.length, 21);


  // var t2 = new Date().getTime();
  // var timeDiff = t2 - t1;
  // console.log(timeDiff); // 56598 ms before update, was reduced to 5285 ms after added breaks.

}



