function testArrayQoh() {

  //let arrayIN = ['A001','A002','A003','A004','A005','A006','A007'];
  //let arrayOUT = ['A007','A002','A003','A004'];

  let arrayIN = ['INidbkkgogv81462290466156',
                 'INidkuarlsk01285951354703',
                 'INidizwxsha31863223854642',
                 'INidxgxmfbt01524976123247',
                 'INidpkcuaey41106112295354',
                 'INidfwziqvb91825117218259',
                 'INidcjafiqz31109966439289',
                 'INidxausrwm61683899338920',
                 'INidbvkgxug31438772801136',
                 'INidemqseio11794283227955'];


  let arrayOUT = ['INidbvkgxug31438772801136',
                  'INidxgxmfbt01524976123247',
                  'INidbkkgogv81462290466156'];


  let arrayOUTSort = arrayOUT.sort();
  let arrayMatched = [];

  for (let i = 0; i < arrayIN.length; i++){
    for (let j = 0; j < arrayOUTSort.length; j++){
      if (arrayIN[i] === arrayOUTSort[j]){

        removeItemOnce(arrayIN, arrayOUTSort[j]);

      } else {
        //console.log(arrayIN[i]);
      }
    }
  }


  // Use arrayIN - the remainder gives the QOH list. The matched array elements are removed.
  //console.log(arrayIN);
  //console.log(arrayOUTSort);
  
}

function removeItemAll(arr, value) {
  var i = 0;
  while (i < arr.length) {
    if (arr[i] === value) {
      arr.splice(i, 1);
    } else {
      ++i;
    }
  }
  return arr;
}

function tryItOut(){

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  //const getLastRow = ss.getSheetByName("Sheet32").getLastRow();
  //const getLastCol = ss.getSheetByName("Sheet32").getLastColumn();
  const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

  //let a = "id" + Math.floor(Math.random() * 100)
  let b = "id" + Math.random().toString(36).substring(2,9);
  let arrayLongList = [];

  for (i = 0; i < 10; i++){
    let randomNumber = Math.random() * 99999;
    let cleanUpRandomNumber = randomNumber.toString().replace(".","").substring(1,13);
    let frontLetters = alphabet[Math.floor(Math.random() * alphabet.length)] +
                          alphabet[Math.floor(Math.random() * alphabet.length)] +
                          alphabet[Math.floor(Math.random() * alphabet.length)] +
                          alphabet[Math.floor(Math.random() * alphabet.length)] +
                          alphabet[Math.floor(Math.random() * alphabet.length)] +
                          alphabet[Math.floor(Math.random() * alphabet.length)] +
                          alphabet[Math.floor(Math.random() * alphabet.length)];

    let a = "INid" + frontLetters + Math.floor(Math.random() * 10)+1 + cleanUpRandomNumber;
    
    arrayLongList.push([a]);
      console.log(a);

  }
  //console.log(random);

  let array1 = [
    ['Jen',5],
    ['Ben',6],
    ['Cen',7]
  ];

  let array2 = [
    ['Den',8],
    ['Len',9],
    ['Ren',10]
  ];

  let array3 = array1.concat(array2);

  let array4 = [
    'Den',
    'Len',
    'Ren'
  ];

  let array5 = [
    'Ten',
    'Ben',
    'Hen'
  ];

  let array6 = [array4,array5];


  let merged = []

  for (i = 0; i < array4.length; i++) {
  merged.push(appendArrays(array4[i], array5[i]));
  }

  // To get the last 240
  let f = '00024000024000024000';
  let g = f.lastIndexOf('240');
  //console.log(g);
  //console.log(merged);

}

function testingCounting(){
  const names = ['Alice', 'Bob', 'Tiff', 'Bruce', 'Alice']

  let countedNames = names.reduce(function (allNames, name) {
  if (name in allNames) {
    allNames[name]++
  }
  else {
    allNames[name] = 1
  }
  return allNames
  }, {})
  let personName = 'Alice';
  //console.log(countedNames);
  //console.log(countedNames.personName);
  //console.log(countedNames.Bob);
  let listOfItems = [];
  for (const prop in countedNames) {
  //console.log(`${prop} = ${countedNames[prop]}`);
  //console.log([prop, countedNames[prop]]);
  listOfItems.push([prop, countedNames[prop]]);
  
  }
  //console.log(listOfItems);
  //console.log(listOfItems[0][1]);
  //console.log(listOfItems[1][1]);


}

function testingCountinginObjects(){
  const names = {'Alice, Parrish':8,
               'Bob'  :7,
               'Tiff' :6,
               'Bruce':5,}

  console.log(Object.values(names));


  /*let listOfItems = [];
  for (const prop in countedNames) {
  //console.log(`${prop} = ${countedNames[prop]}`);
  //console.log([prop, countedNames[prop]]);
  listOfItems.push([prop, countedNames[prop]]);
  
  }
  console.log(listOfItems);
  console.log(listOfItems[0][1]);
  console.log(listOfItems[1][1]);

  */
}

function testOUTRanges() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  //const getLastRowL = ss.getSheetByName("Sheet31").getLastRow();
  //console.log(getLastRow);

  //ss.getSheetByName("Sheet31").getRangeList(['A2:A'+getLastRowL,'C3']).insertCheckboxes();

}

// To calculate PR quantities and minus the DO items
function countingQOHTest(){

  /*let prItems = [['PR01', 'Orange', 10],
               ['PR02', 'Apple', 20],
               ['PR03', 'Banana', 8],
               ['PR04', 'Grape', 5]];
  */

 let prItems = [ [ 'TEST', 'A003', 2 ],
  [ 'TEST', 'A005', 3 ],
  [ 'Order no 1', 'A001', 3 ],
  [ 'Order no 2', 'A002', 4 ],
  [ 'Order no 3', 'A004', 5 ],
  [ 'Order no 4', 'A006', 3 ],
  [ 'Order no 5', 'A011', 2 ],
  [ 'Order no 6', 'A016', 4 ],
  [ 'Order no 7', 'A018', 3 ],
  [ 'Order no 8', 'A019', 2 ],
  [ 'Order no 9', 'A020', 6 ],
  [ 'Order no 10', 'A026', 3 ] ];

  /*let doItems = [['PR01','Orange', 1],
               ['PR02','Apple', 1],
               ['PR02','Apple', 1],
               ['PR03','Banana', 1],
               ['PR04','Grape', 1]];
  */

  let doItems = [ [ 'TEST', 'A005', 2 ],
  [ 'TEST', 'A003', 1 ],
  [ 'TEST', 'A003', 1 ] ]

  // Create object array for PR items
  listOfPR = prItems.map(function(x) {
    return {    
        "PR": x[0],
        "Item": x[1],
        "Qty": x[2]
    }
  });
  //console.log(listOfPR);
  //console.log(listOfPR[0].Qty)

  // Create object array for DO items, with the Quantity in negative value
  listOfDO = doItems.map(function(x) {
    return {    
        "PR": x[0],
        "Item": x[1],
        "Qty": -x[2]
    }
  });
  //console.log(listOfDO);

  // Combine objects from PR list and DO list
  Array.prototype.push.apply(listOfPR,listOfDO);

  // console.log(listOfPR);


  // Sum up the quantity in the object array
  var result = [];
  listOfPR.reduce(function(res, value) {
  if (!res[value.PR] && !res[value.Item]) {
    res[value.PR, value.Item] = { PR: value.PR, Item: value.Item, Qty: 0 };
    result.push(res[value.PR, value.Item])
  }
  res[value.PR, value.Item].Qty += value.Qty;
  return res;
  }, {});

  //console.log(result);

  // Convert object to array for further calculation or paste       
   let newArr = result.map(function(val, index){ 
            return [val.PR,val.Item,val.Qty]
   }) 
          
  // console.log(newArr); 

}

function testOUTCount() {
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  //const getbestexp = ss.getSheetByName("Sheet54").getRange(23,3,1,1).getValue();

  const qoh = 2800;
  const testperkit = 200;
  const apm = 1000;
  const qohlast = qoh/apm;
  const bestexp = new Date(getbestexp.getTime());//(23.98333333333*60*60*1000));
  console.log(bestexp);

  let arrayList = makeArrayListOfNewOrderBasedOnTest (testperkit, apm, qohlast);

  for (i = 0; i < arrayList.length; i++){
    if (bestexp > arrayList[i][3]){
      result0 = arrayList[i][0];
      result1 = arrayList[i][1];
      result2 = arrayList[i][2];
      result3 = arrayList[i][3];
    }
  }
  //console.log(result0);
  //console.log(result1); // Use this one
  //console.log(result2);
  //console.log(result3);

    

  //console.log(arrayList[0][3]);
}

function getImageTest() {
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let getMasterList = ss.getSheetByName("MasterL").getRange(2,1,getMasterListLastRow-1,18).getValues();

  const findVal = 'A003';

  let rowList = [4,6,8];

  /*for (i = 0; i < getMasterList.length; i++) {
    if (getMasterList[i][0] === findVal){
      //getImage = getMasterList[i][17];
      getImageLocation = i;
      rowList.push(getImageLocation);
    }
  }*/
  //console.log(rowList);

  // This method is faster
  for (p = 0; p < 3; p++){
  //ss.getSheetByName("TEST").getRange(5+p,1,1,1).setFormula(`=MasterL!R`+rowList[p]);
 }

 let arrayList = [];

 let linkArray = ["https://drive.google.com/uc?id=1H_sriueE-usZoMgPtMVNIK9ANnMaznx4",
                 "https://drive.google.com/uc?id=1BQr1ZHeEAn8Ycawo1EiaDiJ74tLVEjMl",
                 "https://drive.google.com/uc?id=1lRsVGxLBtpXPcXR_4sXVElOh0wPlInGF"];

  //for (j = 0; j < 3; j++){
  for (k = 0; k < 3; k++){

  row = 5;
  //pasteCell = 1;
  //link = ss.getSheetByName("TEST").getRange(14+k,1,1,1).getValue();    // work by taking links without quotation from cells

  image = SpreadsheetApp
                 .newCellImage()
                 .setSourceUrl(linkArray[k])
                 .toBuilder();


  //ss.getSheetByName("TEST").getRange(row+j,1,1,1).setValue(image);
  //ss.getSheetByName("TEST").getRange(row+k,1,1,1).setValue(image); // WORK

  //arrayList.push(image);  
  }
  //}
  //console.log(arrayList);

  //ss.getSheetByName("TEST").getRange('A5').setValue(image);
  //ss.getSheetByName("TEST").getRange(5,1,3,1).setValue(arrayList);

  //ss.getSheetByName("TEST").getRange(5,1,1,1).setValue(arrayList);



}

function tryappendingtwoarrays(){

  let arr1 = [["A1"],
            ["C2"],
            ["B3"]];

  let arr2 = [["D",1],
            ["E",2]];

  let arr3 = arr1.concat(arr2);

  //let sortedArr = arr1.sort();

  //console.log(sortedArr);

  let firstArr = [["A",1],["B",2],["C",3]];
  let secondArr = [["D",1],["E",2],["F",3]];
  let thirdArr = [["G",1],["H",2],["I",3]];

  let combArr = firstArr.concat(secondArr).concat(thirdArr);

  // console.log(combArr);





  /*for (a = 0; a < arr1.length; a++){
  for (b = 0; b < arr2.length; b++){
    arr3.push(appendArrays(arr1[a]))


  }
  }*/
  let index = arr1.indexOf("A");
  //console.log(index)

  //console.log(arr1)
  //console.log(arr2)
  //console.log(arr3)

}


class Car {

  constructor(name, year){
    this.name = name;
    this.year = year;
  }

  age(num){
    return num + this.year
  }
}

let carOwner1 = new Car('Ali', 2001);

  const testSheet = SpreadsheetApp.getActiveSpreadsheet();
  const masterLastRow = testSheet.getSheetByName("MasterL").getLastRow();
  let openMasterList = testSheet.getSheetByName("MasterL").getRange(2,1,masterLastRow-1,9).getValues();

  let findThis = 'A001';
  let createObj = {
  id: 233,
  people:
  [{name: 'Ahmad', age: 25},
  {name: 'Ali', age: 45}]};

  // let resultFound = openMasterList.findIndex(findThis);
  var myArray = [
    {"id": 1, "name": "Alice"},
    {"id": 2, "name": "Peter"},
    {"id": 3, "name": "Harry"}
  ];

  // lookForThis = num => num.id === 2;
  let val = 3;

function lookForThis (num) {

  return num.id === val;
}

  var result = myArray.find(item => item.id === 2);

  // var res = myArray.find(lookForThis);
  // console.log(res);
  // console.log(res.name);


function myFunction() {

  // var t1 = new Date().getTime();
  // const sheetName = "MasterL";
  // const [headers, ...rows] = SpreadsheetApp.getActiveSpreadsheet()
  // .getSheetByName(sheetName)
  // .getDataRange()
  // .getValues();
  // const res = rows.map((r) =>
  // headers.reduce((o, h, j) => Object.assign(o, { [h]: r[j] }), {})
  // );

  // var result = res.find(item => item.Mastercode === 'A999');
  // console.log(result['Location']);
  // var t2 = new Date().getTime();
  // var timeDiff = t2-t1;

  // SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PASTE").getRange().setValues(res);
  // SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PASTE").getRange().setValues("Hello");
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PASTE").getRange(1,1,1,1).setValues();

  // console.log(timeDiff); // 575 ms

  // console.log(res[0]);
}

  for (i = 0; i < openMasterList.length; i++) {
    if (openMasterList[i][0] === 'A999') {
      // console.log(openMasterList[i]);
    }
  }


function testForClassConstruction(){

  var t1 = new Date().getTime();

  for (i = 0; i < openMasterList.length; i++) {
    if (openMasterList[i][0] === 'A200') {
      console.log(openMasterList[i]);
    }
  }
  var t2 = new Date().getTime();
  var timeDiff = t2-t1;
  // console.log(timeDiff); // 6 ms

}

function pracBreakLoop() {

  var t1 = new Date().getTime();

  let iterations = 10000;
  let someArray = [];

  for (i = 0; i < iterations; i++) {
      someArray.push(i);
  }

  for (i = 0; i < someArray.length; i++){
    if (someArray[i] === 4000){
      // console.log(someArray[i]);
    }

  }


  // someArray.forEach((num) => {
  //   if(num === 5999){
  //     console.log(num);
  //   }
  // });

  var t2 = new Date().getTime();
  var timeDiff = t2 - t1;
  console.log(timeDiff);

  // 4 ms - for loop

}












