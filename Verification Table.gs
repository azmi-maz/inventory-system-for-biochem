function updateVerification() {

  // var t1 = new Date().getTime();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const getTblVerificationLastRow = ss.getSheetByName("Verification").getLastRow();

  let rawListFromTbl = ss.getSheetByName("Verification").getRange(2, 1, getTblVerificationLastRow - 1, 12).getValues();

  // Clear off filter
  ss.getSheetByName("Verification").getRange(1, 1, getTblVerificationLastRow - 1, 14).getFilter().remove();

  // Set borders for new items
  ss.getSheetByName("Verification").getRange(2, 1, getTblVerificationLastRow - 1, 14).setBorder(true, true, true, true, true, true);

  // Set datavalidation on STATUS column
  statusBuild = SpreadsheetApp.newDataValidation().requireValueInList(['Balum', 'Sudah']).build();
  ss.getSheetByName("Verification").getRange(2, 9, getTblVerificationLastRow - 1, 1).setDataValidation(statusBuild);

  // Set columns date format
  ss.getSheetByName("Verification").getRange(2, 1, getTblVerificationLastRow - 1, 1).setNumberFormat("DD/MM/YYYY");
  ss.getSheetByName("Verification").getRange(2, 4, getTblVerificationLastRow - 1, 1).setNumberFormat("DD/MM/YYYY");
  ss.getSheetByName("Verification").getRange(2, 6, getTblVerificationLastRow - 1, 1).setNumberFormat("DD/MM/YYYY");
  ss.getSheetByName("Verification").getRange(2, 7, getTblVerificationLastRow - 1, 1).setNumberFormat("DD/MM/YYYY");

  // Clear off Available Kits Column
  ss.getSheetByName("Verification").getRange(2, 12, getTblVerificationLastRow - 1, 3).clearContent();

  // Taken from BO-QOH and modified  __________________________________________________________________________
  const getTblUniqueINIDLastRow = ss.getSheetByName("tblUniqueINID").getLastRow();
  let getListOfUniqueID = ss.getSheetByName("tblUniqueINID").getRange(2, 1, getTblUniqueINIDLastRow - 1, 5).getValues();
  //let getMasterList = ss.getSheetByName("MasterL").getRange(2,1,getMasterListLastRow-1,17).getValues();

  updateOUTGOING();
  let listOfQOHSingleCount = remainingIDLeftForQOHList;
  let getItemCodeValueForLookUp = '';
  let getLotNumber = '';
  let extractListOfQOH = [];

  // Count current QOH and look for APM and Test Per Kit
  for (t = 0; t < listOfQOHSingleCount.length; t++) {
    for (u = 0; u < getListOfUniqueID.length; u++) {

      if (listOfQOHSingleCount[t] === getListOfUniqueID[u][0]) {

        getItemCodeValueForLookUp = getListOfUniqueID[u][2];
        getLotNumber = getListOfUniqueID[u][3];

        extractListOfQOH.push([
          getItemCodeValueForLookUp,
          getLotNumber
        ]);
        break; // New
      }
    }
  }

  //console.log(extractListOfQOH.sort())

  // Take extractListOfQOH and count each item
  let countExtractListOfQOHElem = countArrayElem(extractListOfQOH);
  let countExtractListOfQOH = countExtractListOfQOHElem[2];
  // Taken from BO-QOH and modified  __________________________________________________________________________

  // Match Verification table with QOH
  let matchVerificationArray = [];
  let lookForCritical = [];
  let statusUpdate = [];
  let matchedQOHValue = 0;
  let statusUpdateVal = '';
  let getQOHValue = 0;
  for (a = 0; a < rawListFromTbl.length; a++) {
    for (b = 0; b < countExtractListOfQOH.length; b++) {

      tblVerItemCode = rawListFromTbl[a][10];
      tblVerNewLotNumber = rawListFromTbl[a][2];
      tblVerCurrLotNumber = rawListFromTbl[a][4];
      tblStatusNow = rawListFromTbl[a][8];
      qohItemCode = countExtractListOfQOH[b][0].split(",")[0];
      qohLotNumber = countExtractListOfQOH[b][0].split(",")[1];

      if (tblVerItemCode === qohItemCode && tblVerNewLotNumber === qohLotNumber) {
        matchedQOHValue = countExtractListOfQOH[b][1];
        matchVerificationArray.push([a, matchedQOHValue]);
      }
      if (tblVerItemCode === qohItemCode && tblVerCurrLotNumber === qohLotNumber) {
        getQOHValue = countExtractListOfQOH[b][1];
        lookForCritical.push([a, getQOHValue]);

        if (tblVerItemCode === qohItemCode && getQOHValue <= 5 && tblStatusNow === 'Balum') {
          statusUpdateVal = "Critical";
          statusUpdate.push([a, statusUpdateVal]);
        }
      }
    }
  }
  // console.log(matchVerificationArray);
  //console.log(lookForCritical)
  //console.log(statusUpdate)

  // Paste the new array to the table
  for (a = 0; a < matchVerificationArray.length; a++) {
    ss.getSheetByName("Verification").getRange(matchVerificationArray[a][0] + 2, 12, 1, 1).setValue(matchVerificationArray[a][1]);
  }

  for (a = 0; a < lookForCritical.length; a++) {
    ss.getSheetByName("Verification").getRange(lookForCritical[a][0] + 2, 13, 1, 1).setValue(lookForCritical[a][1]);
  }

  for (a = 0; a < statusUpdate.length; a++) {
    ss.getSheetByName("Verification").getRange(statusUpdate[a][0] + 2, 14, 1, 1).setValue(statusUpdate[a][1]);
  }

  // Hide Column K
  const range = ss.getSheetByName("Verification").getRange("K2");
  ss.getSheetByName("Verification").hideColumn(range);

  // Create filter
  ss.getSheetByName("Verification").getRange(1, 1, getTblVerificationLastRow, 14).createFilter().sort(2, true);

  // var t2 = new Date().getTime();
  // var timeDiff = t2 - t1;
  // console.log(timeDiff); // 9984 ms before update. Reduced to 7361 ms after adding breaks.


}
