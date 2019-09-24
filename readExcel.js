const XLSX = require('xlsx');

async function readExcel(path) {
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

module.exports = {
    readExcel: readExcel,
    wbEachSheet: wbEachSheet
};