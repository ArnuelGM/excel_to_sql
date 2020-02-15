const XLSX = require('xlsx');

function readExcel(path) {
    console.log('Reading file in ', path);
    let workbook = XLSX.readFile(path);
    return workbook;
}

async function wbEachSheet(workbook, callback) {
    let index = 0;
    for (sheetName of workbook.SheetNames) {
        let sheet = workbook.Sheets[sheetName];
        await callback(sheet, sheetName, index, workbook);
        index++;
    }
}

function wbSheetsToArray(wb) {
    const sheets = [];
    let i = 0;
    for (let sheetName of wb.SheetNames) {
        let sheet = wb.Sheets[sheetName];
        sheets.push({ name: sheetName, sheet: sheet, index: i });
        i++;
    }
    return sheets;
}

module.exports = {
    readExcel: readExcel,
    wbEachSheet: wbEachSheet,
    wbSheetsToArray: wbSheetsToArray
};