function updateStoreAlinityArrayBasedOnPage(){

  updateOUTGOING();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const getTblAlinityStoreLastRow = ss.getSheetByName("Store_Alinity").getLastRow(); 


  // Clean up sheet for new arrays
  ss.getSheetByName("Store_Alinity").getRange(8,2,20,7).clearContent();
  ss.getSheetByName("Store_Alinity").getRange(2,11,getTblAlinityStoreLastRow-1,7).clearContent().removeCheckboxes();
  styleStoreSheetReadableHeight(); // To set the row heights smaller

  // Paste the available list of items
  ss.getSheetByName("Store_Alinity").getRange(2,11,finalStoreTransfer.length,finalStoreTransfer[0].length).setValues(finalStoreTransfer).sort([{column: 13, ascending: true}, {column: 15, ascending: true}]);
  let getList = ss.getSheetByName("Store_Alinity").getRange(2,11,finalStoreTransfer.length,1);
  getList.insertCheckboxes();


}

function pasteStoreAlinityArrayBasedOnPage(){

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const getTblAlinityStoreLastRow = ss.getSheetByName("Store_Alinity").getLastRow(); 
  let arrayListAvailableOfItems = ss.getSheetByName("Store_Alinity").getRange(2,11,getTblAlinityStoreLastRow-1,7).getValues();
  let getMasterList = ss.getSheetByName("MasterL").getRange(2,1,getMasterListLastRow-1,14).getValues();
  //console.log(arrayListAvailableOfItems)

  // Find any choosen items, max ten only
  let choosenItemsArray = [];
  for (a = 0; a < arrayListAvailableOfItems.length; a++){
    for (b = 0; b < getMasterList.length; b++){
    if (arrayListAvailableOfItems[a][0] === true){
      

      if (arrayListAvailableOfItems[a][1] === getMasterList[b][0]){

      checkdate = new Date(arrayListAvailableOfItems[a][4]).toLocaleDateString("en-UK");
      if (checkdate === "Invalid Date"){
        checkdate = arrayListAvailableOfItems[a][4];
      } else {
        //checkdate = new Date(arrayListAvailableOfItems[a][5]).toLocaleDateString("en-UK");
      }

      itemName = getMasterList[b][10];
      itemUOM = getMasterList[b][13];
      getNumRowForImage = b+2; // Added 2 to include in the header and index starts from 0

      choosenItemsArray.push([
                itemName,                             // Item Name
                //arrayListAvailableOfItems[a][2],      // Item Type
                //arrayListAvailableOfItems[a][3],      // Location
                arrayListAvailableOfItems[a][3],      // Lot Number
                checkdate,                            // Expiry date
                itemUOM,                              // UOM
                getNumRowForImage                     // Number of rows
      ]);
      break; // New
      }
    }
  }
  }
  // console.log('Length of array', choosenItemsArray.length); // 3

  // Check for length, max only 10 items
  if (choosenItemsArray.length > 10){
  updateInfo = `There are ${choosenItemsArray.length} items chosen.`+ "\r\n" +`Please reduce to 10 items or less.`;

  // To prompt user of more than 10 items in array is selected
  const promptForExceededArray = SpreadsheetApp.getUi().alert(updateInfo, SpreadsheetApp.getUi().ButtonSet.OK);
  SpreadsheetApp.getActive().toast(promptForExceededArray);
  } else {}



  // Go through each pasteArrayToFormList to each cells for reporting
  for (b = 0; b < choosenItemsArray.length; b++){
    jumpCelltwice = b * 2;

    let getItemNameCell = ss.getSheetByName("Store_Alinity").getRange(8+jumpCelltwice,2,1,1);
    let getBarcodeImage = ss.getSheetByName("Store_Alinity").getRange(9+jumpCelltwice,2,1,1);
    let getLotNumber = ss.getSheetByName("Store_Alinity").getRange(8+jumpCelltwice,4,1,1);
    let getExpDate = ss.getSheetByName("Store_Alinity").getRange(8+jumpCelltwice,6,1,1);
    let getUOM = ss.getSheetByName("Store_Alinity").getRange(8+jumpCelltwice,7,1,1);

    
    // Add one row at a time for max 10 rows/items
    getItemNameCell.setValue(choosenItemsArray[b][0]);
    getBarcodeImage.setFormula(`=MasterL!R`+choosenItemsArray[b][4])
    getLotNumber.setValue(choosenItemsArray[b][1]);
    getExpDate.setValue(choosenItemsArray[b][2]);
    getUOM.setValue(choosenItemsArray[b][3]);

  }

    // Set fixed values to form
    ss.getSheetByName("Store_Alinity").getRange("C30").setValue("Biochem");
    ss.getSheetByName("Store_Alinity").getRange("C33").setValue("Biochem");
    ss.getSheetByName("Store_Alinity").getRange("F30").setFormula(`=today()`);
    ss.getSheetByName("Store_Alinity").getRange("F32").setFormula(`=today()`);
    ss.getSheetByName("Store_Alinity").getRange("F33").setFormula(`=today()`);
    ss.getSheetByName("Store_Alinity").getRange("A4").setValue(true);
    ss.getSheetByName("Store_Alinity").getRange("A5").setValue(false);
    ss.getSheetByName("Store_Alinity").getRange("C4").setValue(false);
    ss.getSheetByName("Store_Alinity").getRange("C5").setValue(false);

    // Select whole page for printing
    styleStoreSheet();
    ss.getSheetByName("Store_Alinity").getRange('A1:H36').activate();

}
