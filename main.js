const { app, BrowserWindow, ipcMain } = require('electron');
const { wbEachSheet, readExcel, wbSheetsToArray } = require('./readExcel');
const { sqlGenerator } = require('./sqlGenerator');

const sqlConfig = require('./knex.config');
const knex = require('knex')( sqlConfig.knex );
const mssql = require('mssql');

let win;
const createWindow = () => {
    win = new BrowserWindow({
        width: 800,
        height: 600,
        webPreferences: {
            nodeIntegration: true
        }
    });

    win.loadFile('index.html');
    win.on('close', () => {
        win = null;
    });
}

app.on('ready', createWindow);
app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});
app.on('active', () => { 
    if (win === null) {
        createWindow();
    }
});

ipcMain.on('process-file', async (ev, args) => {
    let wb = await readExcel(args);
    
    console.log('File readed... generating sql statements...');

    let sheets = wbSheetsToArray(wb);
    for (sheetOb of sheets) {
        // if (sheetOb.sheetName !== 'sis_paci') continue;
        let statements = await proccessSheet(sheetOb.sheet, sheetOb.name);
        break;
        // await executeStatements(statements);
    }

});

async function proccessSheet(sheet, sheetName) {
    console.log('Processing Sheet ' + sheetName);

    win.webContents.send('sql-query-to-execute', `--[${sheetName}]--\r\n`);

    // Statements es igual a un array en caso de exito
    // False en caso de error
    let statements = (await sqlGenerator(sheet, sheetName));
    let allOk = false;
    let response = null;

    if (statements) {
        allOk = !statements.length;
        response = allOk ? 'All OK\r\n\r\n' : (statements.join('\r\nGO\r\n')) + '\r\nGO\r\n\r\n';
    }
    else {
        response = 'No fue posible generar el script :(\r\n\r\n';
    }

    win.webContents.send('sql-query-to-execute', response);
    return statements;
}

async function executeStatements(statements) {
    console.log('Executing statements.');
    let c = 0;
    return await knex.transaction(async function (tr) {
        for (let st of statements) {
            await knex.raw(st).transacting(tr);
            c++;
            if (c % 50 === 0) console.log(c, ' statement executed.');
        }
        console.log(c, ' statement executed.');
    });
}