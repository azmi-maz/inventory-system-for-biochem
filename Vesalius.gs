function collectVesalius() {
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const getVExcelLastRow = ss.getSheetByName("VExcel").getLastRow();

  let getVExcelListArray = ss.getSheetByName("VExcel").getRange(2,2,getVExcelLastRow-1,11).getValues();
  let getMasterListArray = ss.getSheetByName("MasterL").getRange(2,1,getMasterListLastRow-1,18).getValues();

    const getTblUniqueINIDLastRow = ss.getSheetByName("tblUniqueINID").getLastRow();
    let getListOfUniqueID = ss.getSheetByName("tblUniqueINID").getRange(2,1,getTblUniqueINIDLastRow-1,5).getValues();
    let getMasterList = ss.getSheetByName("MasterL").getRange(2,1,getMasterListLastRow-1,17).getValues();
    let getMasterJustCodeList = ss.getSheetByName("MasterL").getRange(2,1,getMasterListLastRow-1,1).getValues();

  //console.log(getMasterListArray.length)
  updateOUTGOING();
  let currentQOH = remainingIDLeftForQOHList;
  //console.log(currentQOH[0])

  let getItemCodeValueForLookUp = '';
    let extractListOfQOH = [];
    // Count current QOH and look for APM and Test Per Kit
    for (t = 0; t < currentQOH.length; t++){
    for (u = 0; u < getListOfUniqueID.length; u++){
        
        if (currentQOH[t] === getListOfUniqueID[u][0]){
        
        getItemCodeValueForLookUp = getListOfUniqueID[u][2];
        
    for (v = 0; v < getMasterList.length; v++){
        
        if (getItemCodeValueForLookUp === getMasterList[v][0]){
            getAPMValue = getMasterList[v][8];
            getTestPerKitValue = getMasterList[v][12];
            getMultiCountValue = getMasterList[v][7];
            break; // New
            
          }
        }
      }
    }
            extractListOfQOH.push([
            getItemCodeValueForLookUp,
            ])

    }
  //console.log(extractListOfQOH[0])

    let cleanedgetMasterListJustCodes = [];
    for (a= 0; a < getMasterJustCodeList.length; a++){
      cleanedgetMasterListJustCodes.push(getMasterJustCodeList[a].toString());
    }

    // Take extractListOfQOH and count each item
    let countExtractListOfQOHElem = countArrayElem(extractListOfQOH);

    let arrToRemove = countExtractListOfQOHElem[0];
    let missingItemCode = findRemainingUniqueID(cleanedgetMasterListJustCodes.sort(),arrToRemove.sort());
    
    let missingItemCodeList = [];
    for (a = 0; a < missingItemCode.length; a++){
      missingItemCodeList.push([
        missingItemCode[a],
        0
      ])
    }
    //console.log(missingItemCodeList[0])
    //console.log(countExtractListOfQOH[0])
    let countExtractListOfQOH = countExtractListOfQOHElem[2].concat(missingItemCodeList);

    //console.log(countExtractListOfQOH.length)

  let matchVesaliusCode = [];
  for (a = 0; a < getVExcelListArray.length; a++){
    for (b = 0; b < getMasterListArray.length; b++){
      if (getVExcelListArray[a][0] === getMasterListArray[b][9]){
        matchVesaliusCode.push([
          getVExcelListArray[a][0],  // Vesalius Code
          getMasterListArray[b][0],  // Item Code
          getVExcelListArray[a][1],  // Item Name
          getVExcelListArray[a][8],  // QOH from Vesalius
          getVExcelListArray[a][9],  // UOM
          getMasterListArray[b][7],  // Multicount
          b+2,                       // MasterL index for barcode
          getMasterListArray[b][13], // Master UOM
          getVExcelListArray[a][9], // VExcel UOM
          getMasterListArray[b][12]  // Master Test Per Kit
        ]);
        break; // New
      }
    }
  }
  //console.log(matchVesaliusCode)
  let calcVesaliusQOHwithInventoryQOH = [];
  let quant = 0;
  let flooredQOH = 0;
  for (a = 0; a < matchVesaliusCode.length; a++){
    for (b = 0; b < countExtractListOfQOH.length; b++){
      if (matchVesaliusCode[a][1] === countExtractListOfQOH[b][0]){


        if (matchVesaliusCode[a][7] === matchVesaliusCode[a][8]){
        
        quant = matchVesaliusCode[a][3] - Math.floor((countExtractListOfQOH[b][1]/matchVesaliusCode[a][5]));
        flooredQOH = Math.floor((countExtractListOfQOH[b][1]/matchVesaliusCode[a][5]));


        } else if (matchVesaliusCode[a][7] !== matchVesaliusCode[a][8]){
        
        quant = matchVesaliusCode[a][3] - Math.floor((countExtractListOfQOH[b][1]*matchVesaliusCode[a][9]));
        flooredQOH = Math.floor((countExtractListOfQOH[b][1]*matchVesaliusCode[a][9]));
        }
      
        if (quant >= 0 && matchVesaliusCode[a][3] > flooredQOH){
        calcVesaliusQOHwithInventoryQOH.push([
          matchVesaliusCode[a][2],  // Item Name
          "N/A",                    // No Lot Number Available  
          "N/A",                    // No exp date
          quant,                    // Quantity to correct
          matchVesaliusCode[a][4],  // UOM
          matchVesaliusCode[a][6]   // Index for barcode
        ]);
          
        }
      }
    }
  }
  //console.log(calcVesaliusQOHwithInventoryQOH);

  // To handle how many items there are in the array
  let countEachItemInArray = calcVesaliusQOHwithInventoryQOH.length;
  let length = countEachItemInArray;
  let alarmVal = '';
  //console.log(length)
  
  // To handle different lengths of array
  //let pageNumber = 0;
  switch (true){
    case (length <= 10):
        pageNumber = 1;
        break;
    case (length > 10 && length <= 20):
        pageNumber = 2;
        break;
    case (length > 20 && length <= 30):
        pageNumber = 3;
        break;
    case (length > 30 && length <= 40):
        pageNumber = 4;
        break;
    case (length > 40 && length <= 50):
        pageNumber = 5;
        break;
    case (length > 50 && length <= 60):
        pageNumber = 6;
        break;
    case (length > 60 && length <= 70):
        pageNumber = 7;
        break;
    case (length > 70 && length <= 80):
        pageNumber = 8;
        break;
    case (length > 80 && length <= 90):
        pageNumber = 9;
        break;
    case (length > 90 && length <= 100):
        pageNumber = 10;
        break;
    case (length > 100):
        pageNumber = '';
        alarmVal = "Exceeded more than 100";
        break;
    default:
        pageNumber = 0;
        alarmVal = '';
        break;
  }
  //console.log(pageNumber);
  //console.log(alarmVal);

  // Give the info for the user
  let updateInfo = '';
  if (alarmVal === "" && pageNumber < 11){
  updateInfo = `There are ${countEachItemInArray} items found.`+ "\r\n" +`${pageNumber} page(s) available.`;
  } else if (alarmVal != "" && pageNumber === "") {
  updateInfo = `There are ${countEachItemInArray} items found.`+ "\r\n" +`${alarmVal} items.`+ "\r\n" +`Please reduce to 100 items or less.`;

  // To prompt user of more than 100 items in array is selected
  const promptForExceededArray = SpreadsheetApp.getUi().alert(updateInfo, SpreadsheetApp.getUi().ButtonSet.OK);
  SpreadsheetApp.getActive().toast(promptForExceededArray);

  }
  //console.log(updateInfo);
  
  // Paste this info in the Store-PR sheet
  ss.getSheetByName("Ves-Correct").getRange(5,11,1,1).setValue(updateInfo);

  // Print the last date printed
  const printDate = new Date().toLocaleDateString("en-UK");
  
  ss.getSheetByName("Ves-Correct").getRange(11,11,1,1).setValue(`${printDate}`+ "\r\n" +`${countEachItemInArray} items printed`);

  // Clean up sheet for new arrays
  ss.getSheetByName("Ves-Correct").getRange(8,2,20,7).clearContent();

  // Look for choosen pages
  const choosenPageNumber = ss.getSheetByName("Ves-Correct").getRange(3,12,1,1).getValue();

  // Get the choosen array items from page number
  let arrayIndexFromPageNumber = [];
  if (choosenPageNumber === 1){
    arrayIndexFromPageNumber = [0,1,2,3,4,5,6,7,8,9];
  } else if (choosenPageNumber === 2){
    arrayIndexFromPageNumber = [10,11,12,13,14,15,16,17,18,19];
  } else if (choosenPageNumber === 3){
    arrayIndexFromPageNumber = [20,21,22,23,24,25,26,27,28,29];
  } else if (choosenPageNumber === 4){
    arrayIndexFromPageNumber = [30,31,32,33,34,35,36,37,38,39];
  } else if (choosenPageNumber === 5){
    arrayIndexFromPageNumber = [40,41,42,43,44,45,46,47,48,49];
  } else if (choosenPageNumber === 6){
    arrayIndexFromPageNumber = [50,51,52,53,54,55,56,57,58,59];
  } else if (choosenPageNumber === 7){
    arrayIndexFromPageNumber = [60,61,62,63,64,65,66,67,68,69];
  } else if (choosenPageNumber === 8){
    arrayIndexFromPageNumber = [70,71,72,73,74,75,76,77,78,79];
  } else if (choosenPageNumber === 9){
    arrayIndexFromPageNumber = [80,81,82,83,84,85,86,87,88,89];
  } else if (choosenPageNumber === 10){
    arrayIndexFromPageNumber = [90,91,92,93,94,95,96,97,98,99];
  }

  let pasteArrayToFormList = [];
  let item = '';

  for (a = 0; a < 10; a++){
    index = arrayIndexFromPageNumber[a];
    item = calcVesaliusQOHwithInventoryQOH[index];
    if (item != undefined){
    pasteArrayToFormList.push([
              item[0],  // Item Name
              item[1],  // Lot Number
              item[2],  // Exp Date
              item[3],  // Count as One
              item[4],  // UOM
              item[5]   // Range for image
              ]);
    }
  }
  //console.log(pasteArrayToFormList)

  // Go through each pasteArrayToFormList to each cells for reporting
  for (b = 0; b < pasteArrayToFormList.length; b++){
    jumpCelltwice = b * 2;

    let getItemNameCell = ss.getSheetByName("Ves-Correct").getRange(8+jumpCelltwice,2,1,1);
    let getBarcodeImage = ss.getSheetByName("Ves-Correct").getRange(9+jumpCelltwice,2,1,1);
    let getLotNumber = ss.getSheetByName("Ves-Correct").getRange(8+jumpCelltwice,4,1,1);
    let getExpDate = ss.getSheetByName("Ves-Correct").getRange(8+jumpCelltwice,6,1,1);
    let getUOM = ss.getSheetByName("Ves-Correct").getRange(8+jumpCelltwice,7,1,1);
    let getQuantityReq = ss.getSheetByName("Ves-Correct").getRange(9+jumpCelltwice,7,1,1);
    let getQuantityIssued = ss.getSheetByName("Ves-Correct").getRange(8+jumpCelltwice,8,1,1);
    
    // Add one row at a time for max 10 rows/items
    getItemNameCell.setValue(pasteArrayToFormList[b][0]);
    getBarcodeImage.setFormula(`=MasterL!R`+pasteArrayToFormList[b][5])
    getLotNumber.setValue(pasteArrayToFormList[b][1]);
    getExpDate.setValue(pasteArrayToFormList[b][2]);
    getUOM.setValue(pasteArrayToFormList[b][4]);
    getQuantityReq.setValue(pasteArrayToFormList[b][3]);
    getQuantityIssued.setValue(pasteArrayToFormList[b][3]);
  }

    // Set fixed values to form
    ss.getSheetByName("Ves-Correct").getRange("C30").setValue("Biochem");
    ss.getSheetByName("Ves-Correct").getRange("C33").setValue("Biochem");
    ss.getSheetByName("Ves-Correct").getRange("F30").setFormula(`=today()`);
    ss.getSheetByName("Ves-Correct").getRange("F32").setFormula(`=today()`);
    ss.getSheetByName("Ves-Correct").getRange("F33").setFormula(`=today()`);
    ss.getSheetByName("Ves-Correct").getRange("A4").setValue(true);
    ss.getSheetByName("Ves-Correct").getRange("A5").setValue(false);
    ss.getSheetByName("Ves-Correct").getRange("C4").setValue(false);
    ss.getSheetByName("Ves-Correct").getRange("C5").setValue(false);

    // Select whole page for printing
    styleStoreSheet();
    ss.getSheetByName("Ves-Correct").getRange('A1:H36').activate();


}
