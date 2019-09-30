const { app, BrowserWindow, ipcMain } = require('electron');
const { wbEachSheet, readExcel } = require('./readExcel');
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
    
    wbEachSheet(wb, async (sheet, sheetName, index, wb) => {
        win.webContents.send('sql-query-to-execute', `--[${sheetName}]--\r\n`);

        // Statements es igual a un array en caso de exito
        // False en caso de error
        let statements = (await sqlGenerator(sheet, sheetName, index, wb));
        let allOk = false;
        let response = null;

        if (statements) { 
            allOk = !statements.length;
            response = allOk ? 'All OK\r\n\r\n' : (statements.join('\r\n')) + '\r\n\r\n';
        }
        else {
            response = 'No fue posible generar el script :(\r\n\r\n';
        }

        win.webContents.send('sql-query-to-execute', response);
        
    });
});