const { app, BrowserWindow, ipcMain } = require('electron');
const { wbEachSheet, readExcel } = require('./readExcel');
const { sqlGenerator } = require('./sqlGenerator');

const sqlConfig = require('./knex.config');
// const knex = require('knex')( sqlConfig );
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
    
    wbEachSheet(wb, async (sheet, sheetName, index, wb) => {
        win.webContents.send('sql-query-to-execute', `---[${sheetName}] not inserted---\r\n`);

        let statements = (await sqlGenerator(sheet, sheetName, index, wb));
        let allOk = statements.length == 0 ? true : false;

        win.webContents.send('sql-query-to-execute', allOk ? 'All OK\r\n\r\n' : (statements.join('\r\n')) + '\r\n\r\n');
    });
});