function updateFOCDate() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const getFOCCountLastRow = ss.getSheetByName("FOC-NonAbbott").getLastRow();
  const getFOCCompLastRow = ss.getSheetByName("FOCNonAbbottComp").getLastRow();

  let fullTableCountComp = ss.getSheetByName("FOC-NonAbbott").getRange(2, 1, getFOCCountLastRow - 1, 9).getValues();
  let compCountList = ss.getSheetByName("FOCNonAbbottComp").getRange(2, 1, getFOCCompLastRow - 1, 6).getValues();

  // Fill up CountComp with datevalues
  let getDateValues = [];
  for (b = 0; b < compCountList.length; b++) {
    for (a = 0; a < fullTableCountComp.length; a++) {
      if (compCountList[b][0] === fullTableCountComp[a][0]) {
        getDateValues.push([
          fullTableCountComp[a][3],     // Start Date
          fullTableCountComp[a][4]      // End Date
        ]);
      }
    }
  }
  ss.getSheetByName("FOCNonAbbottComp").getRange(2, 3, getDateValues.length, getDateValues[0].length).setValues(getDateValues);

  updateFOCAPM();

}

function updateFOCAPM() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const getFOCCountLastRow = ss.getSheetByName("FOC-NonAbbott").getLastRow();

  let fullTableCountComp = ss.getSheetByName("FOC-NonAbbott").getRange(2, 1, getFOCCountLastRow - 1, 9).getValues();
  let getMasterListArrayUpToAPM = ss.getSheetByName("MasterL").getRange(2, 1, getMasterListLastRow - 1, 9).getValues();

  // Update APM values
  let newAPMValues = [];
  for (a = 0; a < getMasterListArrayUpToAPM.length; a++) {
    for (b = 0; b < fullTableCountComp.length; b++) {
      if (getMasterListArrayUpToAPM[a][0] === fullTableCountComp[b][0]) {
        newAPMValues.push([a, fullTableCountComp[b][8]]);
      }
    }
  }
  //console.log(newAPMValues.sort());

  // Paste to the MasterL
  for (c = 0; c < newAPMValues.length; c++) {
    ss.getSheetByName("MasterL").getRange(newAPMValues[c][0] + 2, 9, 1, 1).setValue(newAPMValues[c][1]);
  }

}