function statisticsOutput1() {

  // var t1 = new Date().getTime();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const getTblINLastRow = ss.getSheetByName("tblStockIN").getLastRow();
  const getTblUniqueLastRow = ss.getSheetByName("tblUniqueINID").getLastRow();
  const getTblOUTLastRow = ss.getSheetByName("tblStockOUT").getLastRow();

  const dateFrom = ss.getSheetByName("Statistics").getRange(3,2,1,1).getValue();
  const dateTo = new Date(ss.getSheetByName("Statistics").getRange(4,2,1,1).getValue().getTime()+23.9999*60*60*1000);

  let tblINList = ss.getSheetByName("tblStockIN").getRange(2,1,getTblINLastRow-1,3).getValues();
  let tblUniqueList = ss.getSheetByName("tblUniqueINID").getRange(2,1,getTblUniqueLastRow-1,5).getValues();
  let tblOUTList = ss.getSheetByName("tblStockOUT").getRange(2,1,getTblOUTLastRow-1,3).getValues();
  let masterList = ss.getSheetByName("MasterL").getRange(2,1,getMasterListLastRow-1,17).getValues();

  // For incoming count
  let newFilteredIncoming = [];
  for (a = 0; a < tblINList.length; a++){
    for (b = 0; b < tblUniqueList.length; b++){

        if (tblINList[a][0] > dateFrom && tblINList[a][0] < dateTo && tblINList[a][1] === tblUniqueList[b][0]){

            newFilteredIncoming.push(
                  tblUniqueList[b][2]
            );
            break; // New
            

          }
        }
      }
      // console.log('Length of array =', newFilteredIncoming.length); // 544

  // Expand Incoming
  let expandNewFilteredIncoming = [];
  let expandNewFilteredIncomingValue = [];
  let getMultiCounts = 0;
  let getPrice = 0;
  let getItemValue = 0;
  for (a = 0; a < newFilteredIncoming.length; a++){
    for (b = 0; b < masterList.length; b++){
      if (newFilteredIncoming[a] === masterList[b][0] && masterList[b][5] === "PR"){
        getMultiCounts = masterList[b][7];
        getPrice = masterList[b][15];
        getItemValue = getPrice/getMultiCounts;

        expandNewFilteredIncoming.push(
          newFilteredIncoming[a]
        );

        expandNewFilteredIncomingValue.push(
          getItemValue
        );
        break; // New

      }
    }
  }
  // console.log('Length of array =', expandNewFilteredIncoming.length); //342

  let sumOfIncomingItems = 0;
  if (expandNewFilteredIncomingValue.length === 0){
  sumOfIncomingItems = 0;
  } else {
  sumOfIncomingItems = expandNewFilteredIncomingValue.reduce(function(previousValue, currentValue) {
    return previousValue + currentValue;
  });
  }

  //console.log(sumOfIncomingItems)
  //console.log(expandNewFilteredIncoming.length)

  // For Outgoing count
  let newFilteredOutgoing = [];
  for (a = 0; a < tblOUTList.length; a++){
    for (b = 0; b < tblUniqueList.length; b++){

        if (tblOUTList[a][0] > dateFrom && tblOUTList[a][0] < dateTo && tblOUTList[a][1] === tblUniqueList[b][0]){

            newFilteredOutgoing.push(
                  tblUniqueList[b][2]
            );
            break; // New

          }
        }
      }
    // console.log('Length of array =', newFilteredOutgoing.length); // 486

  // Expand Outgoing
  let expandNewFilteredOutgoing = [];
  let expandNewFilteredOutgoingValue = [];
  let getOutMultiCounts = 0;
  let getOutPrice = 0;
  let getOutItemValue = 0;
  for (a = 0; a < newFilteredOutgoing.length; a++){
    for (b = 0; b < masterList.length; b++){
      if (newFilteredOutgoing[a] === masterList[b][0] && masterList[b][5] === "PR"){
        getOutMultiCounts = masterList[b][7];
        getOutPrice = masterList[b][15];
        getOutItemValue = getOutPrice/getOutMultiCounts;

        expandNewFilteredOutgoing.push(
          newFilteredOutgoing[a]
        );

        expandNewFilteredOutgoingValue.push(
          getOutItemValue
        );
        break; // New


      }
    }
  }
    // console.log('Length of array =', expandNewFilteredOutgoing.length); // 323

  let sumOfOutgoingItems = expandNewFilteredOutgoingValue.reduce(function(previousValue, currentValue) {
    return previousValue + currentValue;
  });

  //console.log(sumOfOutgoingItems)
  //console.log(expandNewFilteredOutgoing.length)

  // Paste

  // Incoming - Number of PR items
  ss.getSheetByName("Statistics").getRange(7,2,1,1).setValue(expandNewFilteredIncoming.length);
  // Incoming - Total Value
  ss.getSheetByName("Statistics").getRange(8,2,1,1).setValue(Math.floor(sumOfIncomingItems));
  // Outgoing - Number of PR items
  ss.getSheetByName("Statistics").getRange(7,3,1,1).setValue(expandNewFilteredOutgoing.length);
  // Outgoing - Total Value
  ss.getSheetByName("Statistics").getRange(8,3,1,1).setValue(Math.floor(sumOfOutgoingItems));

  // var t2 = new Date().getTime();
  // var timeDiff = t2 - t1;
  // console.log(timeDiff); // 25069 ms before update. Nope breaks didn't reduce the time.


}
