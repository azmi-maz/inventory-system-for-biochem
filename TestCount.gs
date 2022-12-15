function getTestCount() {
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const getTestCountEntryLastRow = ss.getSheetByName("TestCountEntry").getLastRow();
  const getTestCountDataLastRow = ss.getSheetByName("Test_Count Data").getLastRow();
  const getTestCountDataLastCol = ss.getSheetByName("Test_Count Data").getLastColumn();

  let lastColTestCount = 0;

  // First run
  for (b = 1; b < getTestCountDataLastCol; b++){
    cellValue = ss.getSheetByName("Test_Count Data").getRange(2,b,1,1).getValue();
    if (cellValue !== ""){
      lastColTestCount += 1;
    } else {
      break;
    }
  }
  //console.log(lastColTestCount)


  let getTestCountNewEntry = ss.getSheetByName("TestCountEntry").getRange(2,1,getTestCountEntryLastRow-1,13).getValues();
  let getTestCountDataArray = ss.getSheetByName("Test_Count Data").getRange(2,1,getTestCountDataLastRow-1,lastColTestCount).getValues();

  //console.log(getTestCountNewEntry[0]);
  //console.log(getTestCountDataArray[0])

  // Make new array to match up with the current table
  let newArrayWithKeyValue = [];
  for (a = 0; a < getTestCountNewEntry.length; a++){

    primaryKey = getTestCountNewEntry[a][1]+getTestCountNewEntry[a][2]+getTestCountNewEntry[a][3];

    newArrayWithKeyValue.push([
                    primaryKey,                       // Primary key
                    getTestCountNewEntry[a][12]       // Total attempted tests
                    ]);
  }
  //console.log(newArrayWithKeyValue)

  // Match two arrays together
  let matchedKeys = [];
  for (c = 0; c < getTestCountDataArray.length; c++){
    for (d = 0; d < newArrayWithKeyValue.length; d++){
      
      value = [c,newArrayWithKeyValue[d][0],newArrayWithKeyValue[d][1]];

      if (getTestCountDataArray[c][3] === newArrayWithKeyValue[d][0]){

        matchedKeys.push(value);

      }
    }
  }
  //console.log(matchedKeys)
  
  // Paste based on the index match of each array element
  for (e = 0; e < matchedKeys.length; e++){
  
  // To paste only matched values
  ss.getSheetByName("Test_Count Data").getRange(matchedKeys[e][0]+2,lastColTestCount+1,1,1).setValue(matchedKeys[e][2]);

  }

  fillInTheBlanks();

}


function fillInTheBlanks() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const getTestCountDataLastRow = ss.getSheetByName("Test_Count Data").getLastRow();
  const getTestCountDataLastCol = ss.getSheetByName("Test_Count Data").getLastColumn();

  let lastColTestCount = 0;

  // First run
  for (b = 1; b < getTestCountDataLastCol; b++){
    cellValue = ss.getSheetByName("Test_Count Data").getRange(2,b,1,1).getValue();
    if (cellValue !== ""){
      lastColTestCount += 1;
    } else {
      break;
    }
  }
  //console.log(lastColTestCount)

  let getTestCountDataArray = ss.getSheetByName("Test_Count Data").getRange(2,1,getTestCountDataLastRow-1,lastColTestCount).getValues();

  // To fill the gap and put the blanks with zero value
  for (f = 0; f < getTestCountDataArray.length; f++){
    if (getTestCountDataArray[f][lastColTestCount-1] === ""){
     ss.getSheetByName("Test_Count Data").getRange(f+2,lastColTestCount,1,1).setValue(0);
    }
  }
}


function countComp() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const getAlinityCountLastRow = ss.getSheetByName("Alinity_Count").getLastRow();
  const getCountCompLastRow = ss.getSheetByName("CountComp").getLastRow();
  const getTestCountHeaderCol = ss.getSheetByName("Test_Count Data").getLastColumn();
  const getTestCountHeaderRow = ss.getSheetByName("Test_Count Data").getLastRow();

  let fullTableCountComp = ss.getSheetByName("Alinity_Count").getRange(2,1,getAlinityCountLastRow-1,10).getValues();
  let compCountList = ss.getSheetByName("CountComp").getRange(2,1,getCountCompLastRow-1,13).getValues();
  //console.log(getTestCountDataHeaderVal)

//___________________________________________________________________
  // Fill up CountComp with datevalues
  let getDateValues = [];
 for (b = 0; b < compCountList.length; b++){ 
  for (a = 0; a < fullTableCountComp.length; a++){
    if (compCountList[b][1] === fullTableCountComp[a][1]){
      getDateValues.push([
                  fullTableCountComp[a][4],     // Start Date
                  fullTableCountComp[a][5]      // End Date
      ]);
    }
  }
 }
 ss.getSheetByName("CountComp").getRange(2,3,getDateValues.length,getDateValues[0].length).setValues(getDateValues);

 updateAPM();
 //___________________________________________________________________

} 


function updateAPM(){

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const getAlinityCountLastRow = ss.getSheetByName("Alinity_Count").getLastRow();

  let fullTableCountComp = ss.getSheetByName("Alinity_Count").getRange(2,1,getAlinityCountLastRow-1,10).getValues();
  let getMasterListArrayUpToAPM = ss.getSheetByName("MasterL").getRange(2,1,getMasterListLastRow-1,9).getValues();

  // Update APM values
  let newAPMValues = [];
  for (a = 0; a < getMasterListArrayUpToAPM.length; a++){
    for (b = 0; b < fullTableCountComp.length; b++){
      if (getMasterListArrayUpToAPM[a][0] === fullTableCountComp[b][0]){
        newAPMValues.push([a,fullTableCountComp[b][9]]);
      }
    }
  }
//console.log(newAPMValues);

// Paste to the MasterL
for (c = 0; c < newAPMValues.length; c++){
  ss.getSheetByName("MasterL").getRange(newAPMValues[c][0]+2,9,1,1).setValue(newAPMValues[c][1]);
}

}
