function getNewPR() {

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const getOrderPRLastRow = ss.getSheetByName("Order PR").getLastRow();
    const getTblPRLastRow = ss.getSheetByName("tblPR").getLastRow();

    // Get all of MasterL - longer columns than the others 9 to 14
    let getMasterList = ss.getSheetByName("MasterL").getRange(2,1,getMasterListLastRow-1,14).getValues();
    let getNewPRListFull = ss.getSheetByName("Order PR").getRange(2,1,getOrderPRLastRow-1,5).getValues();

    //console.log(getNewPRListFull);
    //console.log(getMasterList);

    // Only with PR number
    let getNewPRList = [];
    
    for (a = 0; a < getNewPRListFull.length; a ++){
      if (getNewPRListFull[a][3] != ''){
        getNewPRList.push([
                           getNewPRListFull[a][0],   // Item Code
                           getNewPRListFull[a][3],   // PR no.
                           getNewPRListFull[a][4],   // Quantity
        ]);

      }
    }

//console.log(getNewPRList);

    // Time stamp all inputs
    let stampUpAllItemsList = [];

    for (var n = 0; n < getNewPRList.length; n++){
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

    // Find Item Name, Type and Vesalius
    let lookFromMasterList = [];
    let getItemNameValue = '';
    let getItemTypeValue = '';
    let getVesaliusValue = '';

    for (let b = 0; b < getNewPRList.length; b++){
          for (let c = 0; c < getMasterList.length; c++){
            
            if (getNewPRList[b][0] === getMasterList[c][0]){
  
             getItemNameValue = getMasterList[c][1];
             getItemTypeValue = getMasterList[c][2];
             getVesaliusValue = getMasterList[c][9];

            }
          }
          lookFromMasterList.push([
                                  getItemNameValue,
                                  getItemTypeValue,
                                  getVesaliusValue
            ]);

    }

    //console.log(lookFromMasterList);

    // Collect all arrays and put into one
    let resArrayListForNewPR = [];

    for (d = 0; d < lookFromMasterList.length; d++) {
    resArrayListForNewPR.push(appendArrays(
                stampUpAllItemsList[d],       // Timestamp
                'PR'+getNewPRList[d][1],      // PR no.
                getNewPRList[d][0],           // Item Code
                lookFromMasterList[d][0],     // Item Name
                lookFromMasterList[d][1],     // Item Type
                lookFromMasterList[d][2],     // Vesalius Code
                getNewPRList[d][2],           // Quantity Ordered
                '',                           // Empty Value
                'Pending PO',                 // Starting PR status
                ''                            // Blank for remarks
    ));

    }
    //console.log(resArrayListForNewPR);

    // Paste New PR list in tblPR

    ss.getSheetByName("tblPR").getRange(getTblPRLastRow+1,1,resArrayListForNewPR.length,resArrayListForNewPR[0].length).setValues(resArrayListForNewPR);

    updatePRList();

    // Clear off OrderPR table after entry - For now Num of Rows are manually added
    ss.getSheetByName("Order PR").getRange(2,4,116,2).clearContent();

}


function getNewPRfromTOPR() {

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const getOrderPRLastRow = ss.getSheetByName("To PR").getLastRow();
    const getTblPRLastRow = ss.getSheetByName("tblPR").getLastRow();

    // Get all of MasterL - longer columns than the others 9 to 14
    let getMasterList = ss.getSheetByName("MasterL").getRange(2,1,getMasterListLastRow-1,14).getValues();
    let getNewPRListFull = ss.getSheetByName("To PR").getRange(3,1,getOrderPRLastRow-1,14).getValues();

    //console.log(getNewPRListFull);
    //console.log(getMasterList);

    // Only with PR number
    let getNewPRList = [];
    
    for (a = 0; a < getNewPRListFull.length; a ++){
      if (getNewPRListFull[a][9] != '' &&
          getNewPRListFull[a][12] != ''){
          getNewPRList.push([
                           getNewPRListFull[a][0],   // Item Code
                           getNewPRListFull[a][9],   // PR no.
                           getNewPRListFull[a][12],  // New Qty
        ]);

      }
    }
    //console.log(getNewPRList);

    // Time stamp all inputs
      const today = new Date();

      const dateEachItemDay = today.getDate();
      const dateEachItemMonth = today.getMonth()+1;
      const dateEachItemYear = today.getFullYear();
      const dateEachItemHours = today.getHours();
      const dateEachItemMinutes = today.getMinutes();
      const dateEachItemSeconds = today.getSeconds();

      let time = `${dateEachItemMonth}/${dateEachItemDay}/${dateEachItemYear} ${dateEachItemHours}:${dateEachItemMinutes}:${dateEachItemSeconds}`;

    // Find Item Name, Type and Vesalius
    let lookFromMasterList = [];
    let getItemNameValue = '';
    let getItemTypeValue = '';
    let getVesaliusValue = '';

    for (let b = 0; b < getNewPRList.length; b++){
          for (let c = 0; c < getMasterList.length; c++){
            
            if (getNewPRList[b][0] === getMasterList[c][0]){
  
             getItemNameValue = getMasterList[c][1];
             getItemTypeValue = getMasterList[c][2];
             getVesaliusValue = getMasterList[c][9];

            }
          }
          lookFromMasterList.push([
                                  getItemNameValue,
                                  getItemTypeValue,
                                  getVesaliusValue
            ]);

    }

    //console.log(lookFromMasterList);

    // Collect all arrays and put into one
    let resArrayListForNewPR = [];

    for (d = 0; d < lookFromMasterList.length; d++) {
    resArrayListForNewPR.push(appendArrays(
                time,                         // Timestamp
                'PR'+getNewPRList[d][1],      // PR no.
                getNewPRList[d][0],           // Item Code
                lookFromMasterList[d][0],     // Item Name
                lookFromMasterList[d][1],     // Item Type
                lookFromMasterList[d][2],     // Vesalius Code
                getNewPRList[d][2],           // Quantity Ordered
                getNewPRList[d][2],           // Same as ordered
                'Pending PO',                 // Starting PR status
                ''                            // Blank for remarks
    ));

    }
    //console.log(resArrayListForNewPR);

    // Paste New PR list in tblPR
    if (resArrayListForNewPR.length === 0){
    } else {
    ss.getSheetByName("tblPR").getRange(getTblPRLastRow+1,1,resArrayListForNewPR.length,resArrayListForNewPR[0].length).setValues(resArrayListForNewPR);
    }

    //updatePRList();

    // Clear off OrderPR table after entry - For now Num of Rows are manually added
    ss.getSheetByName("To PR").getRange(2,10,getOrderPRLastRow-1,1).clearContent();
    ss.getSheetByName("To PR").getRange(2,13,getOrderPRLastRow-1,1).clearContent();

}


function getNewOrderFoc() {

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const getOrderFOCLastRow = ss.getSheetByName("Order FOC").getLastRow();
    const getTblPRLastRow = ss.getSheetByName("tblPR").getLastRow();
    const getTblPOLastRow = ss.getSheetByName("tblPO").getLastRow();

    // Get all of MasterL - longer columns than the others 9 to 14
    let getMasterList = ss.getSheetByName("MasterL").getRange(2,1,getMasterListLastRow-1,14).getValues();
    let getNewFOCListFull = ss.getSheetByName("Order FOC").getRange(2,1,getOrderFOCLastRow-1,6).getValues();

    //console.log(getNewFOCListFull);
    //console.log(getMasterList);

    // Only with PR number
    let getNewFOCList = [];
    
    for (a = 0; a < getNewFOCListFull.length; a ++){
      if (getNewFOCListFull[a][3] != ''){
        getNewFOCList.push([
                           getNewFOCListFull[a][0],  // MasterCode
                           getNewFOCListFull[a][3],  // Order no.
                           getNewFOCListFull[a][4],  // Qty
                           getNewFOCListFull[a][5]   // Remarks
        ]);

      }
    }

//console.log(getNewFOCList);

    // Time stamp all inputs
    let stampUpAllItemsList = [];

    for (var n = 0; n < getNewFOCList.length; n++){
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

    // Find Item Name, Type and Vesalius
    let lookFromMasterList = [];
    let getItemNameValue = '';
    let getItemTypeValue = '';
    let getVesaliusValue = '';

    for (let b = 0; b < getNewFOCList.length; b++){
          for (let c = 0; c < getMasterList.length; c++){
            
            if (getNewFOCList[b][0] === getMasterList[c][0]){
  
             getItemNameValue = getMasterList[c][1];
             getItemTypeValue = getMasterList[c][2];
             getVesaliusValue = getMasterList[c][9];

            }
          }
          lookFromMasterList.push([
                                  getItemNameValue,
                                  getItemTypeValue,
                                  getVesaliusValue
            ]);

    }

    //console.log(lookFromMasterList);

    // Collect all arrays and put into one
    let resArrayListForNewFOC = [];
    let resArrayListForNewFOCPRPO = [];

    for (d = 0; d < lookFromMasterList.length; d++) {
    resArrayListForNewFOC.push(appendArrays(
                stampUpAllItemsList[d],       // Timestamp
                'FOC'+getNewFOCList[d][1],           // Order no.
                getNewFOCList[d][0],           // Item Code
                lookFromMasterList[d][0],     // Item Name
                lookFromMasterList[d][1],     // Item Type
                lookFromMasterList[d][2],     // Vesalius Code
                getNewFOCList[d][2],           // Quantity Ordered
                '',                            // Empty Value
                '',                            // Status
                getNewFOCList[d][3]            // Remarks
                ));

    resArrayListForNewFOCPRPO.push(appendArrays(
                stampUpAllItemsList[d],       // Timestamp
                "FOC Order",                   // FOC Order
                'FOC'+getNewFOCList[d][1],           // Order no.
                getNewFOCList[d][0],           // Item Code
                lookFromMasterList[d][0],     // Item Name
                lookFromMasterList[d][1],     // Item Type
                getNewFOCList[d][2],           // Quantity Ordered
                getNewFOCList[d][3]            // Remarks
    ));

    }
    //console.log(resArrayListForNewFOC);
    //console.log(resArrayListForNewFOCPRPO);

    // Paste New FOC list in tblPR

    ss.getSheetByName("tblPR").getRange(getTblPRLastRow+1,1,resArrayListForNewFOC.length,resArrayListForNewFOC[0].length).setValues(resArrayListForNewFOC);


    // Paste the Choosen FOC list to tblPO
    ss.getSheetByName("tblPO").getRange(getTblPOLastRow+1,1,resArrayListForNewFOCPRPO.length,resArrayListForNewFOCPRPO[0].length).setValues(resArrayListForNewFOCPRPO);

    // Clear off OrderFOC table after entry - For now Num of Rows are manually added
    ss.getSheetByName("Order FOC").getRange(2,4,9,3).clearContent();

    updatePRList();

}

function getNewPO() {

      // Match New PO with Active PR list
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      //const getPORemarksValue = ss.getSheetByName("PO Entry").getRange(2,2,1,1).getValue();
      const getPONumberEnteredValue = ss.getSheetByName("PO Entry").getRange(1,2,1,1).getValue();
      const getUrgentPOCheckBox = ss.getSheetByName("PO Entry").getRange(1,5,1,1).getValue();
      const getTblPOLastRow = ss.getSheetByName("tblPO").getLastRow();
      const getPOEntryLastRow = ss.getSheetByName("PO Entry").getLastRow();
      
      const today = new Date();
        const dateEachItemDay = today.getDate();
        const dateEachItemMonth = today.getMonth()+1;
        const dateEachItemYear = today.getFullYear();
        const dateEachItemHours = today.getHours();
        const dateEachItemMinutes = today.getMinutes();
        const dateEachItemSeconds = today.getSeconds();

      let time = `${dateEachItemMonth}/${dateEachItemDay}/${dateEachItemYear} ${dateEachItemHours}:${dateEachItemMinutes}:${dateEachItemSeconds}`;
      let searchThroughChoosenPRList = ss.getSheetByName("PO Entry").getRange(5,1,getPOEntryLastRow-1,6).getValues();
      let choosenPRListArray = [];
      let modifyPONumber = '';

      if (getUrgentPOCheckBox === true){
        modifyPONumber = 'Pending PO';
      } else {
        modifyPONumber = getPONumberEnteredValue;
      }

      for (b = 0; b < searchThroughChoosenPRList.length; b++){
          if (searchThroughChoosenPRList[b][0] === true){
            choosenPRListArray.push([
                                    time,                                   // Timestamp
                                    modifyPONumber,                         // PO Number OR Pending PO
                                    searchThroughChoosenPRList[b][1],       // PR Number
                                    searchThroughChoosenPRList[b][2],       // Item Code
                                    searchThroughChoosenPRList[b][3],       // Item Name
                                    searchThroughChoosenPRList[b][4],       // Item Type
                                    searchThroughChoosenPRList[b][5],       // Quantity Ordered
                                    "", //getPORemarksValue                       // PO Remarks
            ]);
        }
      }
      //console.log(choosenPRListArray);

      // Paste the Choosen PR list to tblPO
      if (choosenPRListArray.length === 0){
      
        const promptForEmptyArray = SpreadsheetApp.getUi().alert("Please select from PR list.", SpreadsheetApp.getUi().ButtonSet.OK);
        SpreadsheetApp.getActive().toast(promptForEmptyArray);

      } else {
      ss.getSheetByName("tblPO").getRange(getTblPOLastRow+1,1,choosenPRListArray.length,choosenPRListArray[0].length).setValues(choosenPRListArray);
      }

      //Clear off the list or update to exclude the chosen ones
      ss.getSheetByName("PO Entry").getRange(5,1,getPOEntryLastRow-1,6).clearContent().removeCheckboxes();
      ss.getSheetByName("PO Entry").getRange(1,2,1,1).clearContent();
      ss.getSheetByName("PO Entry").getRange(1,5,1,1).setValue(false);
      updatePOEntry();

}

function getNewDO() {

      // Match DO with Active PR PO list
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const getDONumberEnteredValue = ss.getSheetByName("DO Entry").getRange(1,2,1,1).getValue();
      const getDORemarksValue = ss.getSheetByName("DO Entry").getRange(2,2,1,1).getValue();
      const getTblDOLastRow = ss.getSheetByName("tblDO").getLastRow();
      let   getDOEntryLastRow = ss.getSheetByName("DO Entry").getLastRow();
      
      const today = new Date();
      const dateEachItemDay = today.getDate();
      const dateEachItemMonth = today.getMonth()+1;
      const dateEachItemYear = today.getFullYear();
      const dateEachItemHours = today.getHours();
      const dateEachItemMinutes = today.getMinutes();
      const dateEachItemSeconds = today.getSeconds();

      let time = `${dateEachItemMonth}/${dateEachItemDay}/${dateEachItemYear} ${dateEachItemHours}:${dateEachItemMinutes}:${dateEachItemSeconds}`;


      let choosenPRPOListArray = [];
      let searchThroughChoosenPRPOList = ss.getSheetByName("DO Entry").getRange(5,1,getDOEntryLastRow-4,8).getValues();

      for (b = 0; b < searchThroughChoosenPRPOList.length; b++){
          if (searchThroughChoosenPRPOList[b][0] === true){
            choosenPRPOListArray.push([
                                    time,                                   // Timestamp
                                    getDONumberEnteredValue,                // DO number
                                    searchThroughChoosenPRPOList[b][2],     // PO Number
                                    searchThroughChoosenPRPOList[b][1],     // PR Number
                                    searchThroughChoosenPRPOList[b][3],     // Item Code
                                    searchThroughChoosenPRPOList[b][4],     // Item Name
                                    searchThroughChoosenPRPOList[b][5],     // Item Type
                                    searchThroughChoosenPRPOList[b][7],     // Calculated quantity received
                                    getDORemarksValue                       // Remarks
            ]);
        }
      }
      //console.log(choosenPRPOListArray);

      // Paste the Choosen PR list to tblPO
      // If choosen nothing, need to break
      if (choosenPRPOListArray.length === 0){
      
        const promptForEmptyArray = SpreadsheetApp.getUi().alert("Please select from DO list.", SpreadsheetApp.getUi().ButtonSet.OK);
        SpreadsheetApp.getActive().toast(promptForEmptyArray);

      } else {
      ss.getSheetByName("tblDO").getRange(getTblDOLastRow+1,1,choosenPRPOListArray.length,choosenPRPOListArray[0].length).setValues(choosenPRPOListArray);
      }

      // TO clear off or update the DO list
      if (getDOEntryLastRow-4 === 0){
      } else {
      ss.getSheetByName("DO Entry").getRange(5,1,getDOEntryLastRow-4,8).clearContent().removeCheckboxes();
      }
      ss.getSheetByName("DO Entry").getRange(1,2,1,1).clearContent();

      
      //updateDOEntry();
      updatePRList();

      
}

function getBestExp() {

const ss = SpreadsheetApp.getActiveSpreadsheet(); 
const getBestExpLastRow = ss.getSheetByName("BestExp").getLastRow();

const getBatchNoValue = ss.getSheetByName("BestExp").getRange(1,2,1,1).getValue();
const getTblBestExpLastRow = ss.getSheetByName("tblBestExp").getLastRow();
const today = new Date();
      const dateEachItemDay = today.getDate();
      const dateEachItemMonth = today.getMonth()+1;
      const dateEachItemYear = today.getFullYear();
      const dateEachItemHours = today.getHours();
      const dateEachItemMinutes = today.getMinutes();
      const dateEachItemSeconds = today.getSeconds();
let time = `${dateEachItemMonth}/${dateEachItemDay}/${dateEachItemYear} ${dateEachItemHours}:${dateEachItemMinutes}:${dateEachItemSeconds}`;

let getNewBestExpArray = ss.getSheetByName("BestExp").getRange(4,1,getBestExpLastRow-3,8).getValues();

//console.log(getBatchNoValue);
//console.log(getNewBestExpArray);

// Append all rows with expiry dates
let getMatchedItemWithNewExp = [];
for (k = 0; k < getNewBestExpArray.length; k++){
    if (getNewBestExpArray[k][6]){
        getMatchedItemWithNewExp.push([
          time,
          getBatchNoValue,            // Batch number
          getNewBestExpArray[k][0],   // PR Number
          getNewBestExpArray[k][1],   // PO Number
          getNewBestExpArray[k][2],   // Item Code
          getNewBestExpArray[k][3],   // Item Name
          getNewBestExpArray[k][4],   // Item Type
          getNewBestExpArray[k][5],   // Quantity Ordered
          getNewBestExpArray[k][6],   // Quantity Remaining
          getNewBestExpArray[k][7],   // Exp Date Offered
          '' // Remarks
        ])
    }
}
//console.log(getMatchedItemWithNewExp);

// Paste the Best Exp List
//(TO UPDATE ON LAST ROW PASTE)
ss.getSheetByName("tblBestExp").getRange(getTblBestExpLastRow+1,1,getMatchedItemWithNewExp.length,getMatchedItemWithNewExp[0].length).setValues(getMatchedItemWithNewExp);

// To clear off the bestexp table
ss.getSheetByName("BestExp").getRange(4,8,getBestExpLastRow-3,1).clearContent();

updateBestExp();
// Successful message box after running the script

}

// New cycle here to cycle through choosen batch list and get the acceptable quantity
function getFinalisedBatchList() {

    const ss = SpreadsheetApp.getActiveSpreadsheet(); 
    const getBATCHLISTLastRow = ss.getSheetByName("Batch List").getLastRow();
    const getBatchNumberValue = ss.getSheetByName("Batch List").getRange(1,2,1,1).getValue();
    const getTblBatchLastRow = ss.getSheetByName("tblBatch").getLastRow();

    const today = new Date();
      const dateEachItemDay = today.getDate();
      const dateEachItemMonth = today.getMonth()+1;
      const dateEachItemYear = today.getFullYear();
      const dateEachItemHours = today.getHours();
      const dateEachItemMinutes = today.getMinutes();
      const dateEachItemSeconds = today.getSeconds();
    let time = `${dateEachItemMonth}/${dateEachItemDay}/${dateEachItemYear} ${dateEachItemHours}:${dateEachItemMinutes}:${dateEachItemSeconds}`;

    let lookForActiveBatchList = ss.getSheetByName("Batch List").getRange(4,1,getBATCHLISTLastRow-3,15).getValues();

    //console.log(getBatchNumberValue);
    //console.log(lookForActiveBatchList);

    // Go through each list to see if there are any choosen items under YES column
  let choosenBatchItemsArray = [];
  let getCorrectQty = 0;

  for (i = 0; i < lookForActiveBatchList.length; i++) {

  if (lookForActiveBatchList[i][11] === true){
    
  if (lookForActiveBatchList[i][12] != ''){
      
      getCorrectQty = lookForActiveBatchList[i][12];

  } else {
      
      getCorrectQty = lookForActiveBatchList[i][8];

  }

      choosenBatchItemsArray.push([
        false,
        time,                           // Date of confirmation
        getBatchNumberValue,            // Batch Number
        lookForActiveBatchList[i][0],   // PR Number
        lookForActiveBatchList[i][1],   // PO Number
        lookForActiveBatchList[i][2],   // Item Code
        lookForActiveBatchList[i][3],   // Item Name
        lookForActiveBatchList[i][4],   // Item Type
        lookForActiveBatchList[i][5],   // Quantity Ordered
        lookForActiveBatchList[i][6],   // Quantity Remaining
        lookForActiveBatchList[i][7],   // Exp Date Offered
        getCorrectQty,                  // Acceptable Quantity
        lookForActiveBatchList[i][9],   // QOH estimated date
        lookForActiveBatchList[i][10],  // New estimated date
        lookForActiveBatchList[i][13],  // Auto comment
        lookForActiveBatchList[i][14]   // Manual Remarks

      ])
    }
  }
  //console.log(choosenBatchItemsArray);

  // Paste the choosen Items into tblBatch sheet
  ss.getSheetByName("tblBatch").getRange(getTblBatchLastRow+1,1,choosenBatchItemsArray.length,choosenBatchItemsArray[0].length).setValues(choosenBatchItemsArray);
  ss.getSheetByName("tblBatch").getRange(getTblBatchLastRow+1,1,choosenBatchItemsArray.length,1).insertCheckboxes();

  // Clean up Batch List sheet
  ss.getSheetByName("Batch List").getRange(4,1,getBATCHLISTLastRow-3,15).clearContent().removeCheckboxes();


}