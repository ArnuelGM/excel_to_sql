const XLSX = require('xlsx');
const _ = require('lodash');
const moment = require('moment');
const knex = require('knex')( require('./knex.config').knex );
const sql = require('mssql');
const sqlConfig = require('./knex.config').mssql;

async function sqlGenerator(sheet, sheetName, index, wb) {
    
    // convertimos la hoja de excel en un objeto json
    let sheetJson = XLSX.utils.sheet_to_json(sheet);

    // obtenemos los tipos de datos de las columnas de la tabla 
    // (el nombre de la hoja debe ser igual al nombre de la table donde se insertaran los datos)
    let cdt = await getColumDataTypes(sheetName);

    // aplicamos la conversion de tipos de datos segun el tipo de dato de cada columna
    let sheetJsonTrueDataTypes = parseColumnDataTypes(sheetJson, cdt);
    console.log('rows to insert: ', sheetJsonTrueDataTypes.length);

    if (sheetJsonTrueDataTypes.length == 0) {
        console.log('SIN DATOS QUE INSERTAR.');
        return false;
    }

    // se etsablece un tamaño de insercion por lote
    let chunkSize = 200;
    console.log('chunk size: ', chunkSize);

    // generamos paquetes de insercion de tamaño [chunkSize]
    let chunks = _.chunk(sheetJsonTrueDataTypes, chunkSize);
    console.log('chunks count: ', chunks.length);

    let slqChunks = [];
    for (chunk of chunks) {
        let sql = knex(sheetName).insert(chunk).toString() + '\r\nGO';
        slqChunks.push(sql);
    }
    console.log('chunks generated: ', slqChunks.length);

    return slqChunks;
}

/* async function sqlGenerator(sheet, sheetName, index, wb) {
    // convertimos la hoja de excel en un objeto json
    let sheetJson = XLSX.utils.sheet_to_json(sheet);

    // obtenemos los tipos de datos de las columnas de la tabla 
    // (el nombre de la hoja debe ser igual al nombre de la table donde se insertaran los datos)
    let cdt = await getColumDataTypes(sheetName);
    
    // aplicamos la conversion de tipos de datos segun el tipo de dato de cada columna
    let sheetJsonTrueDataTypes = parseColumnDataTypes(sheetJson, cdt);

    let pool = await sql.connect(sqlConfig);
    let conteo = 0;
    let notInserted = [];
    for (let row of sheetJsonTrueDataTypes) {
        
        await insertOrUpdate(pool, row, 'num_id', row.num_id, sheetName, notInserted);
        // console.log('inserted or udated: ', );

        conteo++;

        if (conteo % 100 == 0) {
            console.log(notInserted.length, ' inserted');
            // break;
        }
    }

    return [];
} */

async function insertOrUpdate(poolConection, row, primary_key_name, primary_key_value, table_name, notInserted = []) {
    
    let { recordset } = await poolConection.request().query(`SELECT COUNT(*) AS total_reg FROM ${table_name} WHERE ${primary_key_name} = '${primary_key_value}'`);
    let total_reg = recordset[0].total_reg;
    // console.log(total_reg);
    // return false;
    
    let exists = total_reg ? true : false;
    
    let sql = '';
    if (!exists) {
        sql = knex(table_name).insert(row).toString();
        console.log(sql);
    }
    else {
        sql = knex(table_name).where(primary_key_name, '=', primary_key_value).update(row).toString();
        console.log(sql);
    }

    try {
        await poolConection.request().query(sql);
        notInserted.push(row);
        return true;
    }
    catch (error) {
        console.log(error);
        return false;
    }
}

async function getColumDataTypes(table) {
    let columnDataTypes = await knex.raw(`
        SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, IS_NULLABLE, COLUMN_DEFAULT
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = ?
    `, table);

    let cdt = {};
    columnDataTypes.forEach(e => {
        cdt[e.COLUMN_NAME.toLowerCase()] = {
            type: (e.DATA_TYPE).toLowerCase(),
            constructor: getDataTypeConstructor((e.DATA_TYPE).toLowerCase())
        };
    });
    console.log('column data types: \r\n', cdt);
    return cdt;
}

function getDataTypeConstructor(dataType) {
    switch (dataType) {
        case 'bit':
        case 'int':
        case 'money':
        case 'float':
        case 'double':
        case 'bigint':
        case 'integer':
        case 'numeric':
        case 'decimal':
        case 'tinyint':
        case 'smallint':
        case 'smallmoney':
            return Number;
            break;
        
        case 'date':
        case 'time':
        case 'datetime':
        case 'datetime2':
        case 'smalldatetime':
            // return function (date) { moment(date).format('YYYY-MM-DD HH:mm:ss.SSS'); }
            // break;
        
        case 'char':
        case 'text':
        case 'uuid':
        case 'nchar':
        case 'ntext':
        case 'varchar':
        case 'nvarchar':
        case 'uniqueidentifier':
            return String;
            break;
        
        default:
            return String;
    }
}

function parseColumnDataTypes(sheetJson, cdt) {
    let sheetJsonTrueDataTypes = _.map(sheetJson, (row) => {
        let rowTrueDataType = {};
        for (key in row) {
            try {
                let prop = key.toLowerCase();
                rowTrueDataType[prop] = cdt[prop].constructor(row[key]);
            }
            catch (error) {
                console.log(`Parece que la columna '${key}' no existe en la tabla destino.\nPor favor verifique e inténtelo nuevamente.`);
                return [];
            }
        }
        return rowTrueDataType;
    });
    return sheetJsonTrueDataTypes;
}

module.exports = {
    sqlGenerator: sqlGenerator,
    getColumDataTypes: getColumDataTypes,
    getDataTypeConstructor: getDataTypeConstructor,
    parseColumnDataTypes: parseColumnDataTypes
};