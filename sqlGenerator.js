const XLSX = require('xlsx');
const _ = require('lodash');
const knex = require('knex')( require('./knex.config').knex );

async function sqlGenerator(sheet, sheetName, index, wb) {
    
    // convertimos la hoja de excel en un objeto json
    let sheetJson = XLSX.utils.sheet_to_json(sheet);

    // obtenemos los tipos de datos de las columnas de la tabla 
    // (el nombre de la hoja debe ser igual al nombre de la table donde se insertaran los datos)
    let cdt = await getColumDataTypes(sheetName);

    // aplicamos la conversion de tipos de datos segun el tipo de dato de cada columna
    let sheetJsonTrueDataTypes = parseColumnDataTypes(sheetJson, cdt);

    // se etsablece un tamaño de insercion por lote
    let chunkSize = 100;

    // generamos paquetes de insercion (200 registros para inseratr a la vez en este caso)
    let chunks = _.chunk(sheetJsonTrueDataTypes, chunkSize);

    let sqlToClient = [];
    for (chunk of chunks) {

        let inserted = true;
        try {
            await knex.batchInsert(sheetName, chunk, chunk.length);
        } catch (error) {
            inserted = false;
        }

        if (!inserted) { 
            let sql = knex(sheetName).insert(chunk).toString() + '\r\nGO';
            sqlToClient.push(sql);
        }
    }

    return sqlToClient;
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
        
        // horas y fechas se guardan como string
        case 'date':
        case 'time':
        case 'datetime':
        case 'datetime2':
        case 'smalldatetime':
        
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
            let prop = key.toLowerCase();
            rowTrueDataType[prop] = cdt[prop].constructor(row[key]);
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