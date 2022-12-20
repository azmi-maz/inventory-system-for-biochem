let countedFilteredStorePRItemsList = [];

function getStorePROutgoing() {
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const getTblStockOutLastRow = ss.getSheetByName("tblStockOut").getLastRow();
  const getTblUniqueINIDLastRow = ss.getSheetByName("tblUniqueINID").getLastRow();

  // Get all of MasterL
  let getMasterList = ss.getSheetByName("MasterL").getRange(2,1,getMasterListLastRow-1,14).getValues();
  let getUniqueINID = ss.getSheetByName("tblUniqueINID").getRange(2,1,getTblUniqueINIDLastRow-1,5).getValues();
  let getTblStockOut = ss.getSheetByName("tblStockOUT").getRange(2,1,getTblStockOutLastRow-1,3).getValues();


  // Go through each tblStockOUT to find the Store PR items
  let findStorePRItemsList = [];
  for (a = 0; a < getTblStockOut.length; a++){
    for (b = 0; b < getUniqueINID.length; b++){
      for (c = 0; c < getMasterList.length; c++){


        if (getTblStockOut[a][1] === getUniqueINID[b][0]){
          
          getItemCode = getUniqueINID[b][2];
          
          if (getItemCode === getMasterList[c][0] && 
              getMasterList[c][3] === 'Storeroom' &&
              getMasterList[c][5] === 'PR' &&
              getMasterList[c][4] != 'Transfer to Section'){

          getStockOutDate = getTblStockOut[a][0];
          getItemCode = getUniqueINID[b][2];
          getItemName = getMasterList[c][10];
          getLotNumber = "'"+getUniqueINID[b][3];
          getExpDate = getUniqueINID[b][4];
          countMultiCount = getMasterList[c][7];
          itemUOM = getMasterList[c][13];
          getNumRowForImage = c+2; // Added 2 to include in the header and index starts from 0

          findStorePRItemsList.push([
                    getStockOutDate,
                    getItemCode,
                    getItemName,
                    getLotNumber,
                    getExpDate,
                    countMultiCount,
                    itemUOM,
                    getNumRowForImage
                    ]);
                    break; // New
          }
        }
      }
    }
  }
  //console.log(findStorePRItemsList);

  // Go through each store PR items array and prepare for display array
  let startFilterDate = ss.getSheetByName("Store-PR").getRange(2,11,1,1).getValue();
  let endFilterDate = new Date(ss.getSheetByName("Store-PR").getRange(2,12,1,1).getValue().getTime()+23.9999*60*60*1000);

  // Filter by Date range selected from the Store-PR sheet
  let filteredStorePRItemsList = [];
  for (d = 0; d < findStorePRItemsList.length; d++){
    if (findStorePRItemsList[d][0] >= startFilterDate &&
        findStorePRItemsList[d][0] <= endFilterDate){
          
          dateVal = new Date(findStorePRItemsList[d][4]).toLocaleDateString("en-UK");
          if (dateVal != 'Invalid Date'){
            dateVal = new Date(findStorePRItemsList[d][4]).toLocaleDateString("en-UK");
          } else {
            dateVal = 'BLANK';
          }
          
          
          filteredStorePRItemsList.push([
                        findStorePRItemsList[d][1],   // Item Code
                        findStorePRItemsList[d][2],   // Item Name
                        findStorePRItemsList[d][3],   // Lot Number
                        dateVal,                      // Exp Date
                        findStorePRItemsList[d][5],   // Multicount
                        findStorePRItemsList[d][6],   // UOM
                        findStorePRItemsList[d][7],   // Range for image
                        ]);
    }
  }
  //console.log(filteredStorePRItemsList);

  let initialCount = countArrayElem(filteredStorePRItemsList);
  let countedArrayList = initialCount[2];

  //let countedFilteredStorePRItemsList = [];  //Made this variable global
  // Divide the total with multicount
  for (e = 0; e < countedArrayList.length; e++){
    resOfDivision = Math.ceil(countedArrayList[e][1]/countedArrayList[e][0].split(",")[4]);

    countedFilteredStorePRItemsList.push([
                    countedArrayList[e][0].split(",")[0],  // Item Code
                    countedArrayList[e][0].split(",")[1],  // Item Name
                    countedArrayList[e][0].split(",")[2],  // Lot Number
                    countedArrayList[e][0].split(",")[3],  // Exp Date
                    resOfDivision,                         // Count in kits
                    countedArrayList[e][0].split(",")[5],  // UOM
                    countedArrayList[e][0].split(",")[6]   // Range for image
                    ]);
  }
  //console.log(countedFilteredStorePRItemsList)

  // To handle how many items there are in the array
  let countEachItemInArray = countedFilteredStorePRItemsList.length;
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
  ss.getSheetByName("Store-PR").getRange(5,11,1,1).setValue(updateInfo);

  // Print the last date logs
  const fromDate = startFilterDate.toLocaleDateString("en-UK");
  const endDate = endFilterDate.toLocaleDateString("en-UK");

  ss.getSheetByName("Store-PR").getRange(11,11,1,1).setValue(`${fromDate} to ${endDate}`+ "\r\n" +`${countEachItemInArray} items printed`);

}

function pasteStorePRArrayBasedOnPage(){

  getStorePROutgoing();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Clean up sheet for new arrays
  ss.getSheetByName("Store-PR").getRange(8,2,20,7).clearContent();

  // Look for choosen pages
  const choosenPageNumber = ss.getSheetByName("Store-PR").getRange(3,12,1,1).getValue();

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
    item = countedFilteredStorePRItemsList[index];
    if (item != undefined){
    pasteArrayToFormList.push([
              item[1],  // Item Name
              item[2],  // Lot Number
              item[3],  // Exp Date
              item[4],  // Count as One
              item[5],  // UOM
              item[6]   // Range for image
              ]);
    }
  }
  //console.log(pasteArrayToFormList)

  // Go through each pasteArrayToFormList to each cells for reporting
  for (b = 0; b < pasteArrayToFormList.length; b++){
    jumpCelltwice = b * 2;

    let getItemNameCell = ss.getSheetByName("Store-PR").getRange(8+jumpCelltwice,2,1,1);
    let getBarcodeImage = ss.getSheetByName("Store-PR").getRange(9+jumpCelltwice,2,1,1);
    let getLotNumber = ss.getSheetByName("Store-PR").getRange(8+jumpCelltwice,4,1,1);
    let getExpDate = ss.getSheetByName("Store-PR").getRange(8+jumpCelltwice,6,1,1);
    let getUOM = ss.getSheetByName("Store-PR").getRange(8+jumpCelltwice,7,1,1);
    let getQuantityReq = ss.getSheetByName("Store-PR").getRange(9+jumpCelltwice,7,1,1);
    //let getQuantityIssued = ss.getSheetByName("Store-PR").getRange(8+jumpCelltwice,8,1,1);
    
    // Add one row at a time for max 10 rows/items
    getItemNameCell.setValue(pasteArrayToFormList[b][0]);
    getBarcodeImage.setFormula(`=MasterL!R`+pasteArrayToFormList[b][5])
    getLotNumber.setValue(pasteArrayToFormList[b][1]);
    getExpDate.setValue(pasteArrayToFormList[b][2]);
    getUOM.setValue(pasteArrayToFormList[b][4]);
    getQuantityReq.setValue(pasteArrayToFormList[b][3]);
    //getQuantityIssued.setValue(pasteArrayToFormList[b][3]);

  }

    // Set fixed values to form
    ss.getSheetByName("Store-PR").getRange("C30").setValue("Biochem");
    ss.getSheetByName("Store-PR").getRange("C33").setValue("Biochem");
    ss.getSheetByName("Store-PR").getRange("F30").setFormula(`=today()`);
    ss.getSheetByName("Store-PR").getRange("F32").setFormula(`=today()`);
    ss.getSheetByName("Store-PR").getRange("F33").setFormula(`=today()`);
    ss.getSheetByName("Store-PR").getRange("A4").setValue(true);
    ss.getSheetByName("Store-PR").getRange("A5").setValue(false);
    ss.getSheetByName("Store-PR").getRange("C4").setValue(false);
    ss.getSheetByName("Store-PR").getRange("C5").setValue(false);

    // Select whole page for printing
    styleStoreSheet();
    ss.getSheetByName("Store-PR").getRange('A1:H36').activate();


}
