'use strict';

const XLSX = require('xlsx');
const extend = require('node.extend');

const excelToJson = (function() {

    let config = {};
    const requiredConfig = ['sourceFile'];

    const getCellRow = cell => Number(cell.replace(/[A-z]/gi, ''));
    const getCellColumn = cell => cell.replace(/[0-9]/g, '').toUpperCase();
    const getRangeBegin = cell => cell.match(/^[^:]*/)[0];
    const getRangeEnd = cell => cell.match(/[^:]*$/)[0];

    const parseSheet = (sheetData, workbook) => {
        const sheetName = (sheetData.constructor == String) ? sheetData : sheetData.name;
        const sheet = workbook.Sheets[sheetName];
        const columnToKey = sheetData.columnToKey || config.columnToKey;
        const range = sheetData.range || config.range;
        const headerRows = (sheetData.header && sheetData.header.rows) || (config.header && config.header.rows);
        const headerRowToKeys = (sheetData.header && sheetData.header.rowToKeys) || (config.header && config.header.rowToKeys);

        let strictRangeColumns;
        let strictRangeRows;
        if (range) {
            strictRangeColumns = {
                from: getCellColumn(getRangeBegin(range)),
                to: getCellColumn(getRangeEnd(range))
            };

            strictRangeRows = {
                from: getCellRow(getRangeBegin(range)),
                to: getCellRow(getRangeEnd(range))
            };
        }

        let rows = [];
        for (let cell in sheet) {

            // !ref is not a data to be retrieved || this cell doesn't have a value
            if (cell == '!ref' || sheet[cell].v === undefined) {
                continue;
            }

            const row = getCellRow(cell);
            const column = getCellColumn(cell);

            // Is a Header row
            if (headerRows && row <= headerRows) {
                continue;
            }

            // This column is not configured to be retrieved
            if (columnToKey && !(columnToKey[column] || columnToKey['*'])) {
                continue;
            }

            // This cell is out of the configured range
            if ((strictRangeColumns && strictRangeRows) && (column < strictRangeColumns.from || column > strictRangeColumns.to || row < strictRangeRows.from || row > strictRangeRows.to)) {
                continue;
            }


            const rowData = rows[row] = rows[row] || {};
            let columnData = (columnToKey && (columnToKey[column] || columnToKey['*'])) ?
                columnToKey[column] || columnToKey['*'] :
                (headerRowToKeys) ?
                `{{${column}${headerRowToKeys}}}` :
                column;

            let dataVariables = columnData.match(/{{([^}}]+)}}/g);
            if (dataVariables) {
                dataVariables.forEach(dataVariable => {
                    let dataVariableRef = dataVariable.replace(/[\{\}]*/gi, '');
                    let variableValue;
                    switch (dataVariableRef) {
                        case 'columnHeader':
                            dataVariableRef = (headerRows) ? `${column}${headerRows}` : `${column + 1}`;
                        default:
                            variableValue = (sheet[dataVariableRef] && sheet[dataVariableRef].w) ? sheet[dataVariableRef].w.trim() : null;
                    }

                    columnData = columnData.replace(dataVariable, variableValue);
                });
            }

            if (columnData === '') {
                continue;
            }

            rowData[columnData] = (sheet[cell].t == 'n') ? sheet[cell].w.trim() : sheet[cell].v.trim();
            if (sheetData.appendData) {
                extend(true, rowData, sheetData.appendData);
            }
        }

        // Cleaning empty
        rows = rows.filter(v => v !== null && v !== undefined);
        return rows;
    };

    const convertExcelToJson = function(configSet) {
        config = configSet.constructor === String ? JSON.parse(configSet) : configSet;

        for (let i in requiredConfig) {
            if (Object.keys(config).indexOf(requiredConfig[i]) < 0)
                throw new Error(':: Missing required config :: ' + requiredConfig[i]);
        }

        let workbook = XLSX.readFile(config.sourceFile, {
            sheetStubs: true
        });

        let sheetsToGet = (config.sheets && config.sheets.constructor === Array) ?
            config.sheets :
            Object.keys(workbook.Sheets).slice(0, (config && config.sheets && config.sheets.numberOfSheetsToGet) || undefined);

        let parsedData = {};
        sheetsToGet.forEach(sheet => {
            sheet = (sheet.constructor == String) ? {
                name: sheet
            } : sheet;

            parsedData[sheet.name] = parseSheet(sheet, workbook);
        });

        return parsedData;
    };

    return convertExcelToJson;
}());

module.exports = excelToJson;