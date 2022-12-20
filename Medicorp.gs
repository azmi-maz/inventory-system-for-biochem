const mediSheet = SpreadsheetApp.open(DriveApp.getFileById('112cK9eQ46rDTaCUN1-XG0J1UbEctLL5OyU2uwpctmBY'));
const mainSheet = SpreadsheetApp.open(DriveApp.getFileById('1DA_8fUuL4t9OM61inJ-E4ElOMapTTt7QdZHvUHDeBkk'));
const mediQOHPR = mediSheet.getSheetByName("QOH PR");
const mediQOHFOC = mediSheet.getSheetByName("QOH FOC");


function toDisplayBatch() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const getBatchLastRow = ss.getSheetByName("tblBatch").getLastRow();
  const getMediBatchLastRow = mediSheet.getSheetByName("Batch List").getLastRow();


  let getMasterList = ss.getSheetByName("MasterL").getRange(2,1,getMasterListLastRow-1,17).getValues();
  let getMasterListJustCodes = ss.getSheetByName("MasterL").getRange(2,1,getMasterListLastRow-1,1).getValues();

  let cleanedgetMasterListJustCodes = [];
  for (a= 0; a < getMasterListJustCodes.length; a++){
      cleanedgetMasterListJustCodes.push(getMasterListJustCodes[a].toString());
  }

  let getBatchArray = ss.getSheetByName("tblBatch").getRange(2,1,getBatchLastRow-1,16).getValues();

  // Look for displays only
  let toDisplayList = [];
  for (a = 0; a < getBatchArray.length; a++){
    if (getBatchArray[a][0] === true){
      getListCode = getMasterList[cleanedgetMasterListJustCodes.indexOf(getBatchArray[a][5])][11];
      toDisplayList.push([
        getBatchArray[a][2],  // Batch no
        getBatchArray[a][3],  // PR
        getBatchArray[a][4],  // PO
        getListCode,          // List Code
        getBatchArray[a][6],  // Item Name
        getBatchArray[a][7],  // Item Type
        getBatchArray[a][8],  // Quantity Ordered
        getBatchArray[a][9],  // Quantity Remaining
        getBatchArray[a][11], // Acceptable Quantity
        getBatchArray[a][10], // Exp Date Offered
        getBatchArray[a][15]  // Remarks
      ])
    }
  }

  if (getMediBatchLastRow-1 === 0){
  } else {
  mediSheet.getSheetByName("Batch List").getRange(2,1,getMediBatchLastRow-1,11).clearContent();
  }

  mediSheet.getSheetByName("Batch List").getRange(2,1,toDisplayList.length,toDisplayList[0].length).setValues(toDisplayList).sort([{column: 1,ascending: false},{column: 3,ascending: true},{column: 5,ascending: true}]);

  // Successful message box after running the script
  const promptForSuccess = SpreadsheetApp.getUi().alert("Updated displayed Medicorp batch list", SpreadsheetApp.getUi().ButtonSet.OK);
  SpreadsheetApp.getActive().toast(promptForSuccess);


}