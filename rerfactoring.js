let spreadSheet =  SpreadsheetApp.getActiveSpreadsheet();
let getLatestSheet = () => {//最新シート（先月シート）を取得します
    let numSheets = spreadSheet.getNumSheets();
    let rightSheet = spreadSheet.getSheets()[numSheets-1];
    return rightSheet;
}

const COLUMN_DATA = [8,9];//取得・変更したいカラムを数値で指定します。カラムA ＝1、カラムB＝2、カラムC=3・・・です

let sheetName = getLatestSheet().getName();//シート名取得
console.log(sheetName);
let copiedSheet = getLatestSheet().copyTo(spreadSheet);
let regexpAsSheetName = /\d{4}\/\d{2}/;
let regexpAsKanji =  /(\d+)年(\d+)月売上/;
let regexpWitoutZero = /\d{4}\/\d/
let regYYYY = /[0-9]{4}/;
let regMM = /[0-9]{1,2}/;


const addNumber = (date) => {
    const [year, month] = date.split('/')
    let numYear = Number(year);
    let numMonth = Number(month);
    if (numMonth === 12){
        numMonth = 1;
        numYear+=1;
    }else{
        numMonth++;
    }
    return [String(numYear), numMonth >= 10 ? `${numMonth}`:`0${numMonth}`]; //1桁の場合0を頭につける
}
const [yyyy, mm] = addNumber(sheetName);
newSheetName = `${yyyy}/${mm}`;
console.log("newSheetName = "+newSheetName);
copiedSheet.setName(newSheetName);
let valueA1 = copiedSheet.getRange("A1");
let replacedA1 = valueA1.setValue(newSheetName);
const templateForKanji = `${yyyy}年${Number(mm)}月売上`; // ex) 08の0を外す

const createEditedFormulasList = (data) => {//"月"または"/"を含むカラムを取得し、setValues用に二次元配列に直して返します
    const column = getColumFormula(data);
    console.log(`${data}の式一覧`);
    console.log(column);
    const array = []
    for (let count = 0;count < column.length;count++){
        let matched = column[count][0];
        if(column[count][0].match(regexpAsKanji)){
            matched = matched.replace(regexpAsKanji,templateForKanji);
        } else if (column[count][0].match(regexpAsSheetName)){
            matched = matched.replace(regexpAsSheetName,newSheetName);
        }
        array.push([matched])
    }
    console.log(array)
    return array // [[],[],[]]
}

const changeFormula = () => {//各カラムの式を変更します
    COLUMN_DATA.map((data) => {
        const allColumn = getLatestSheet().getRange(3,data, getLatestSheet().getLastRow());
        const list = createEditedFormulasList(data); // [[],[],[]]
        allColumn.setValues(list);
    })
}

const getColumFormula = (data) => getLatestSheet().getRange(3,data, getLatestSheet().getLastRow(),1).getFormulas();
changeFormula();