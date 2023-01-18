const getListCodeLastRow = globalSheet.getSheetByName("ItemCodeL").getLastRow(); //Global Variable

function getUpdatedListCode() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  //const getListCodeLastRow = ss.getSheetByName("ItemCodeL").getLastRow();


  let itemCodeFullListArray = ss.getSheetByName("ItemCodeL").getRange(2, 1, getListCodeLastRow - 1, 2).getValues();
  let itemCodeList = ss.getSheetByName("ItemCodeL").getRange(2, 2, getListCodeLastRow - 1, 1).getValues();
  let masterListArray = ss.getSheetByName("MasterL").getRange(2, 1, getMasterListLastRow - 1, 17).getValues();
  //console.log(itemCodeList[0][0]);

  let cleanItemCodeList = [];
  for (a = 0; a < itemCodeList.length; a++) {
    cleanItemCodeList.push(itemCodeList[a][0]);
  }

  let cleanMasterList = [];
  for (a = 0; a < masterListArray.length; a++) {
    cleanMasterList.push(masterListArray[a][0]);
  }
  //console.log(cleanItemCodeList)

  // Get all the Item Code from MasterL
  let getItemCodeList = [];
  for (a = 0; a < masterListArray.length; a++) {
    for (b = 0; b < itemCodeList.length; b++) {
      if (masterListArray[a][0] === itemCodeList[b][0]) {
        getItemCodeList.push(masterListArray[a][0]);
        break;
      }
    }
  }
  //console.log(cleanItemCodeList.length)
  //console.log(getItemCodeList.length)
  //console.log(itemCodeFullListArray[26][0])

  // Go through each item code and find their index of latest item code
  let findTheLastestListCode = [];
  let findIndex = 0;
  for (a = 0; a < getItemCodeList.length; a++) {

    findIndex = cleanItemCodeList.lastIndexOf(getItemCodeList[a]);
    //console.log(itemCodeFullListArray[findIndex][0])
    getMasterIndex = cleanMasterList.indexOf(getItemCodeList[a]);
    getCompany = masterListArray[getMasterIndex][16];
    getLength = itemCodeFullListArray[findIndex][0].length;

    if (getLength === 6 && getCompany === "Medicorp") {
      listCodeFormat = itemCodeFullListArray[findIndex][0].toString().substring(0, 4) + "." + itemCodeFullListArray[findIndex][0].toString().substring(4, 6);
      findTheLastestListCode.push([getMasterIndex, getItemCodeList[a], listCodeFormat]);
    } else {
    }
  }
  // console.log(findTheLastestListCode);

  //Paste to MasterL
  for (a = 0; a < findTheLastestListCode.length; a++) {
    ss.getSheetByName("MasterL").getRange(findTheLastestListCode[a][0] + 2, 12, 1, 1).setValue(findTheLastestListCode[a][2]);
  }
}
