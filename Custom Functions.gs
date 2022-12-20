function occurrences(string, subString, allowOverlapping) {

    string += "";
    subString += "";
    if (subString.length <= 0) return (string.length + 1);

    var n = 0,
        pos = 0,
        step = allowOverlapping ? 1 : subString.length;

    while (true) {
        pos = string.indexOf(subString, pos);
        if (pos >= 0) {
            ++n;
            pos += step;
        } else break;
    }
    return n;
}

function appendArrays() {
  var temp = []
  for (var q = 0; q < arguments.length; q++) {
    temp.push(arguments[q]);
  }
  return temp;
}

function ExcelDateToJSDate(date) {
  return new Date(Math.round((date - 25569)*86400*1000));

}

function jsDateToExcelDate(serial) {
  let date = new Date(serial);
  return converted = 25569.0 + ((date.getTime() - (date.getTimezoneOffset() * 60 * 1000)) / (1000 * 60 * 60 * 24));
}

function createUniqueID(enterArrayHere, fromWhichSheet){

  const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
  let arrayLongList = [];
  let frontSheetID = '';

  if (fromWhichSheet === 'INCOMING'){
    frontSheetID = 'INid';
    // Probably not for outgoing because id must be the same with incoming
  } else if (fromWhichSheet === 'MANUAL'){
    frontSheetID = 'MANid';
  } else if (fromWhichSheet === 'OUTGOING'){
    frontSheetID = 'OUTid';
  }

  for (i = 0; i < enterArrayHere.length; i++){
    let randomNumber = Math.random() * 99999;
    let cleanUpRandomNumber = randomNumber.toString().replace(".","").substring(1,13);
    let frontLetters = alphabet[Math.floor(Math.random() * alphabet.length)] +
                          alphabet[Math.floor(Math.random() * alphabet.length)] +
                          alphabet[Math.floor(Math.random() * alphabet.length)] +
                          alphabet[Math.floor(Math.random() * alphabet.length)] +
                          alphabet[Math.floor(Math.random() * alphabet.length)] +
                          alphabet[Math.floor(Math.random() * alphabet.length)] +
                          alphabet[Math.floor(Math.random() * alphabet.length)];

    

    
    
    let a = frontSheetID + frontLetters + Math.floor(Math.random() * 10)+1 + cleanUpRandomNumber;
    
    arrayLongList.push([a]);
    
    return a;

  }
}

function findRemainingUniqueID(arrayIN, arrayOUT) {

  for (let i = 0; i < arrayIN.length; i++){
    for (let j = 0; j < arrayOUT.length; j++){
      if (arrayIN[i] === arrayOUT[j]){

        removeItemOnce(arrayIN, arrayOUT[j]);

      }
    }
  }
  return arrayIN;
  //console.log(arrayIN);
}

function removeItemOnce(arr, value) {
  var index = arr.indexOf(value);
  if (index > -1) {
    arr.splice(index, 1);
  }
  return arr;
}

function toCreateCheckBox(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const getOutgoingLastRow = ss.getSheetByName("OUTGOING").getLastRow();

  let getList = ss.getSheetByName("OUTGOING").getRange(2,1,getOutgoingLastRow-1,1);

  getList.insertCheckboxes();

}

function countArrayElem(arr){

  let countedItems = arr.reduce(function (allNames, name) {
  if (name in allNames) {
    allNames[name]++
  }
  else {
    allNames[name] = 1
  }
  return allNames
  }, {})

  // Return number 1
  let listOfItems = [];
  for (const prop in countedItems) {

  listOfItems.push(prop);  
  }

  // Return number 2
  let listCountedItems = [];
  for (const prop in countedItems) {
  listCountedItems.push(countedItems[prop]);
  }

  // Return number 3
  let listOfNewCount = [];
  for (const prop in countedItems) {

  listOfNewCount.push([prop, countedItems[prop]]);  
  }

  return [listOfItems, listCountedItems, listOfNewCount];

}

function sleep(milliseconds) {
  const date = Date.now();
  let currentDate = null;
  do {
    currentDate = Date.now();
  } while (currentDate - date < milliseconds);
}

function countRemainderFromPR(prList, doList){

  // Create object array for PR items
  listOfPR = prList.map(function(x) {
    return {    
        "PR": x[0],
        "Item": x[1],
        "Qty": x[2]
    }
  });
  //console.log(listOfPR);
  //console.log(listOfPR[0].Qty)

  // Create object array for DO items, with the Quantity in negative value
  listOfDO = doList.map(function(x) {
    return {    
        "PR": x[0],
        "Item": x[1],
        "Qty": -x[2]
    }
  });
  //console.log(listOfDO);

  // Combine objects from PR list and DO list
  Array.prototype.push.apply(listOfPR,listOfDO);

  //console.log(listOfPR);


  // Sum up the quantity in the object array
  let result = [];
  listOfPR.reduce(function(res, value) {
    if (!res[value.PR]) {
      res[value.PR] = { PR: value.PR, Item: value.Item, Qty: 0 };
      result.push(res[value.PR])
  }
  res[value.PR].Qty += value.Qty;
  return res;
  }, {});

  //console.log(result);

  // Convert object to array for further calculation or paste       
   let newArr = result.map(function(val, index){ 
            return [val.PR,val.Item,val.Qty]
   }) 

  return newArr;      
  //console.log(newArr); 

}

function countUpItemFromPR(prList){

  // Create object array for PR items
  listOfPR = prList.map(function(x) {
    return {    
        "Vesalius": x[0],
        "ItemCode": x[1],
        "ItemName": x[2],
        "ItemType": x[3],
        "Qty": x[4]
    }
  });
  //console.log(listOfPR);
  //console.log(listOfPR[0].Qty);

  // Combine objects from PR list and DO list
  //Array.prototype.push.apply(listOfPR);

  //console.log(listOfPR);

  // Sum up the quantity in the object array
  let result = [];
  listOfPR.reduce(function(res, value) {
    if (!res[value.Vesalius] &&
      !res[value.ItemCode] &&
      !res[value.ItemName]
      ) {

    res[value.Vesalius,
        value.ItemCode,
        value.ItemName
        ] = 
        
        { Vesalius: value.Vesalius, 
          ItemCode: value.ItemCode,
          ItemName: value.ItemName,
          Qty: 0 };
    
    result.push(res[
                    value.Vesalius,
                    value.ItemCode,
                    value.ItemName
                    ])
  }
    res[value.Vesalius,
        value.ItemCode,
        value.ItemName
        ]
        .Qty += value.Qty;
  
    return res;
  }, {});

  //console.log(result);

  // Convert object to array for further calculation or paste       
   let newArr = result.map(function(val, index){ 
            return [val.Vesalius,
                    val.ItemCode,
                    val.ItemName,
                    val.Qty]
   }) 

  return newArr;      
  //console.log(newArr); 

}

function makeArrayListOfNewOrderBasedOnTest (testperkit, apm, qohlast, bestexp, qtyremaining) {

  //testperkit = 200;
  //apm = 1000;
  //qohlast = 2000/1000;
  //const ss = SpreadsheetApp.getActiveSpreadsheet();
  //bestexp = ss.getSheetByName("Sheet54").getRange(23,3,1,1).getValue();
  //qtyremaining = 15;

  //testperkit = 200; // in tests
  //apm = 5; // in kits
  //qohlast = 2;
  //const ss = SpreadsheetApp.getActiveSpreadsheet();
  //bestexp = ss.getSheetByName("Sheet54").getRange(23,3,1,1).getValue();
  //qtyremaining = 15;


  const today = new Date();
  const convertToSeconds = (4*7*24*60*60*1000); //-28800000 to simulate using today()
  let arrayListOfNumItems = [];
  let numberOfKits = 0;
  let newPRoverAPMplusQOHlast = 0;
  let newExpDate = '';
  let result = 0;
  let resultRounded = 0;
  
  // calculate the 9999 length based on the maximum tests and more?
  for (i = 0; i < 99999; i++){
    numberOfKits = i / testperkit;
    newPRoverAPMplusQOHlast = i / (apm * testperkit) + qohlast;
    newExpDate = new Date(today.getTime()+(newPRoverAPMplusQOHlast*convertToSeconds));
    arrayListOfNumItems.push([i, numberOfKits, newPRoverAPMplusQOHlast, newExpDate]);
  }

  for (j = 0; j < arrayListOfNumItems.length; j++){
    if (bestexp > arrayListOfNumItems[j][3]){
      result = arrayListOfNumItems[j][1];
        if (result >= qtyremaining){
          result = qtyremaining;
         } else if (result < qtyremaining){
          result = arrayListOfNumItems[j][1];
          resultRounded = Math.round(arrayListOfNumItems[j][1]);

          
      }
    }
  }
  //console.log(arrayListOfNumItems);
  //return arrayListOfNumItems;
  //console.log(result);
  //console.log(resultRounded);
  return [result, resultRounded];

}

function makeArrayListOfNewOrderBasedOnKit (apm, qohlast) {

  apm = 5;
  qohlast = 14/5;
  const today = new Date();
  const convertToSeconds = (4*7*24*60*60*1000); //-28800000 to simulate using today()
  let arrayListOfNumKits = [];
  let newPRoverAPMplusQOHlast = 0;
  let newExpDate = '';
  
  for (i = 0; i < 100; i++){
    newPRoverAPMplusQOHlast = i / apm + qohlast;
    newExpDate = new Date(today.getTime()+(newPRoverAPMplusQOHlast*convertToSeconds));
    arrayListOfNumKits.push([i, newPRoverAPMplusQOHlast, newExpDate]);
  }
  //console.log(arrayListOfNumKits);
  return arrayListOfNumKits;

}

function sleep(milliseconds) {
  const date = Date.now();
  let currentDate = null;
  do {
    currentDate = Date.now();
  } while (currentDate - date < milliseconds);
}

function findLastPOforQOH(){

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const getTblPOLastRow = ss.getSheetByName("tblPO").getLastRow();

  let getTblPOArray = ss.getSheetByName("tblPO").getRange(2,1,getTblPOLastRow-1,4).getValues();
  let getTblPOArrayCol4 = ss.getSheetByName("tblPO").getRange(2,4,getTblPOLastRow-1,1).getValues();

  let newArrayToSearch = [];

  for (a = 0; a < getTblPOArrayCol4.length; a++) {
    newArrayToSearch.push(getTblPOArrayCol4[a][0]);
  }
  
  let resultFind = "";
  let resultArray = [];

  for (b = 0; b < matchItemCodeForQOH.length; b++){
  findIndex = newArrayToSearch.lastIndexOf(matchItemCodeForQOH[b]);
  if (findIndex  === -1) {
    resultFind = "No PO found";
    resultArray.push(resultFind);
  } else {
    resultFind = getTblPOArray[findIndex][1];
    resultArray.push(resultFind);
  }
  }

  //console.log(resultFind)
  return resultArray;

}

function findLastPOforQOHBO(){

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const getTblPOLastRow = ss.getSheetByName("tblPO").getLastRow();

  let getTblPOArray = ss.getSheetByName("tblPO").getRange(2,1,getTblPOLastRow-1,4).getValues();
  let getTblPOArrayCol4 = ss.getSheetByName("tblPO").getRange(2,4,getTblPOLastRow-1,1).getValues();

  let newArrayToSearch = [];

  for (a = 0; a < getTblPOArrayCol4.length; a++) {
    newArrayToSearch.push(getTblPOArrayCol4[a][0]);
  }
  
  let resultFind = "";
  let resultArray = [];

  for (b = 0; b < matchItemCodeForQOHBO.length; b++){
  findIndex = newArrayToSearch.lastIndexOf(matchItemCodeForQOHBO[b]);
  if (findIndex  === -1) {
    resultFind = "No PO found";
    resultArray.push(resultFind);
  } else {
    resultFind = getTblPOArray[findIndex][1];
    resultArray.push(resultFind);
  }
  }

  //console.log(resultFind)
  return resultArray;

}

function findLastPOforZeroQOH(){

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const getTblPOLastRow = ss.getSheetByName("tblPO").getLastRow();

  let getTblPOArray = ss.getSheetByName("tblPO").getRange(2,1,getTblPOLastRow-1,4).getValues();
  let getTblPOArrayCol4 = ss.getSheetByName("tblPO").getRange(2,4,getTblPOLastRow-1,1).getValues();

  let newArrayToSearch = [];

  for (a = 0; a < getTblPOArrayCol4.length; a++) {
    newArrayToSearch.push(getTblPOArrayCol4[a][0]);
  }
  
  let resultFind = "";
  let resultArray = [];

  for (b = 0; b < matchItemCodeForZeroQOH.length; b++){
  findIndex = newArrayToSearch.lastIndexOf(matchItemCodeForZeroQOH[b]);
  if (findIndex  === -1) {
    resultFind = "No PO found";
    resultArray.push(resultFind);
  } else {
    resultFind = getTblPOArray[findIndex][1];
    resultArray.push(resultFind);
  }
  }

  //console.log(resultFind)
  return resultArray;

}
