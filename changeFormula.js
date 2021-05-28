let spreadSheet =  SpreadsheetApp.getActiveSpreadsheet();
let getLatestSheet = () => {//最新シート（先月シート）を取得します
  let numSheets = spreadSheet.getNumSheets();
  let rightSheet = spreadSheet.getSheets()[numSheets-1];
  return rightSheet;
}
let sheetName = getLatestSheet().getName();//シート名取得
console.log(sheetName);
let copiedSheet = getLatestSheet().copyTo(spreadSheet);


let regexpAsSheetName = /\d{4}\/\d{2}/;
let regexpAsKanji =  /(\d+)年(\d+)月売上/;
let regexpWitoutZero = /\d{4}\/\d/
let regYYYY = /[0-9]{4}/;
let regMM = /[0-9]{1,2}/;
let yyyy;
let mm;
[yyyy, mm] = sheetName.split('/');//yyyy = 2021, mm = 03

function addNumber(){
  let numYYYY = Number(yyyy);
  let numMM = Number(mm);
  if (numMM < 12){
    numMM++;
    numYYYY;
  }else{
    numMM = 1;
    numYYYY+=1;
  }
  numMM > 10 ? numMM:numMM= `0${numMM}`;//1桁の場合0を頭につける
  let strYYYY = String(numYYYY);
  let strMM = String(numMM);
  return [strYYYY, strMM];
}

yyyy = addNumber()[0];
mm  = addNumber()[1];//06
sheetName = `${yyyy}/${mm}`;
console.log("newSheetName = "+sheetName);
copiedSheet.setName(sheetName);
var valueA1 = copiedSheet.getRange("A1");
var replacedA1 = valueA1.setValue(`${yyyy}/${mm}`);

let mmWithout0 = Number(mm);
console.log(mmWithout0);
const templateForKanji = `${yyyy}年${mmWithout0}月売上`;
let editedList4E = [];
let editedList4H = [];
let editedList4I = [];

let getEColum = () => {
  let allEColumFormula = getLatestSheet().getRange(3,5, getLatestSheet().getLastRow(),1).getFormulas();
  return allEColumFormula;
}
let getHColum = () => {//全てのH1カラムの式を取得
  let allHColumFormula = getLatestSheet().getRange(3,8, getLatestSheet().getLastRow(),1).getFormulas();//各セルの式が二次元配列形式で取得されます
  return allHColumFormula;
}
let getIColum = () => {
  let allIColumFormula = getLatestSheet().getRange(3,9, getLatestSheet().getLastRow(),1).getFormulas();
  return allIColumFormula;
}

// console.log(getIColum());
console.log(getHColum().length + '　元々のH配列の長さ');
console.log(getIColum().length + '　元々のIの配列の長さ');

let createHColumesEditedFormulasList = () => {//"月"または”/”を含むカラムを取得し、setValues用に二次元配列に直して返します
  for (let count = 0;count < getHColum().length;count++){
   if(getHColum()[count][0].match(regexpAsKanji)){
  //返り値はtrueの時の処理
    let matched = getHColum()[count][0].replace(regexpAsKanji,templateForKanji);
    editedList4H.push(matched);//変換した後の式push
    } 
    else if(getHColum()[count][0].match(regexpAsSheetName)){
      let matched = getHColum()[count][0].replace(regexpAsSheetName,sheetName);
      editedList4H.push(matched);
    }
    // else if(getHColum()[count][0].match(regexpWitoutZero)){
    //   mm.slice(1);
    //   let sheetNameWithoutZero = `${yyyy}/${mm}`
    //   let matched = getHColum()[count][0].replace(regexpWitoutZero,sheetNameWithoutZero);
    //   editedList4H.push(matched);
    // }
  else {
    // console.log('false');
    editedList4H.push(getHColum()[count][0]);//オリジナルの式を何もせずにpush
    }
  }
  console.log('編集後のH配列が下に来ます');
  let twoDArray = [];
  for(var i = 0; i < editedList4H.length; i++) {
    twoDArray.push(editedList4H.slice(i, i+1));
  }
  console.log(twoDArray);//setValues用に二次元配列を作成
  return twoDArray;
}

let createIColumesEditedFormulasList = () => {//"月"または”/”を含むカラムを取得し、setValues用に二次元配列に直して返します
  for (let count = 0;count < getIColum().length;count++){
   if(getIColum()[count][0].match(regexpAsKanji)){
  //返り値はtrueの時の処理
    let matched = getIColum()[count][0].replace(regexpAsKanji,templateForKanji);
    editedList4I.push(matched);//変換した後の式push
    } 
    else if(getIColum()[count][0].match(regexpAsSheetName)){
      let matched = getIColum()[count][0].replace(regexpAsSheetName,sheetName);
      editedList4I.push(matched);
    }
    // else if(getIColum()[count][0].match(regexpWitoutZero)){
    //   mm.slice(1);
    //   let sheetNameWithoutZero = `${yyyy}/${mm}`
    //   let matched = getIColum()[count][0].replace(regexpWitoutZero,sheetNameWithoutZero);
    //   editedList4I.push(matched);
    // }
  else {
    editedList4I.push(getIColum()[count][0]);//オリジナルの式を何もせずにpush
    }
  }
  console.log('編集後のI配列が下に来ます');
  let twoDArray2 = [];
  for(var i = 0; i < editedList4I.length; i++) {
    twoDArray2.push(editedList4I.slice(i, i+1));
  }
  console.log(twoDArray2);//setValues用に二次元配列を作成
  return twoDArray2;
}

// let createEColumesEditedFormulasList = () => {//"月"または”/”を含むカラムを取得し、setValues用に二次元配列に直して返します
//   for (let count = 0;count < getEColum().length;count++){
//     let addedYYYY = addNumber()[0]
//     let addedMM = addNumber()[1];
//     // let tempAsKanji = `${addedYYYY}年${addedMM}月`;
//     let tempAsSheetName = `${addedYYYY}/${addedMM}`;
//    if(getEColum()[count][0].match(regexpAsKanji)){
//   //返り値はtrueの時の処理
//     let matched = getEColum()[count][0].replace(regexpAsKanji,templateForKanji);
//     editedList4E.push(matched);//変換した後の式push
//     } 
//     else if(getEColum()[count][0].match(regexpAsSheetName)){
//       let matched = getEColum()[count][0].replace(regexpAsSheetName,tempAsSheetName);
//       editedList4E.push(matched);
//     }
//     // else if(getEColum()[count][0].match(regexpWitoutZero)){
//     //   addedMM.slice(1);
//     //   let sheetNameWithoutZero = `${addedYYYY}/${addedMM}`;
//     //   let matched = getEColum()[count][0].replace(regexpWitoutZero,sheetNameWithoutZero);
//     //   editedList4E.push(matched);
//     // }
//   else {
//     editedList4E.push(getEColum()[count][0]);//オリジナルの式を何もせずにpush
//     }
//   }
//   console.log('編集後のE配列が下に来ます');
//   let twoDArray2 = [];
//   for(var i = 0; i < editedList4E.length; i++) {
//     twoDArray2.push(editedList4E.slice(i, i+1));
//   }
//   console.log(twoDArray2);//setValues用に二次元配列を作成
//   return twoDArray2;
// }

let changeFormula = () => {//各カラムの式を変更します
  let changeHColumsFormula = () => {
    let allHColum = getLatestSheet().getRange(3,8, getLatestSheet().getLastRow());
    allHColum.setValues(createHColumesEditedFormulasList());
  }
  let changeIColumsFormula = () => {
    let allIColum = getLatestSheet().getRange(3,9, getLatestSheet().getLastRow());
    allIColum.setValues(createIColumesEditedFormulasList());
  }
  // let changeEColumsFormula = () => {
  //   let allEColum = getLatestSheet().getRange(3,5, getLatestSheet().getLastRow());
  //   allEColum.setValues(createEColumesEditedFormulasList());
  // }

  changeHColumsFormula();
  changeIColumsFormula();
  // changeEColumsFormula();
}
