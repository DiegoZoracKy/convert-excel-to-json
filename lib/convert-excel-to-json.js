'use strict';

const XLSX = require('xlsx');
const extend = require('node.extend');

const excelToJson = (function() {

    let _config = {};

    const getCellRow = cell => Number(cell.replace(/[A-z]/gi, ''));
    const getCellColumn = cell => cell.replace(/[0-9]/g, '').toUpperCase();
    const getRangeBegin = cell => cell.match(/^[^:]*/)[0];
    const getRangeEnd = cell => cell.match(/[^:]*$/)[0];

    function getSheetCellValue(sheetCell, sheetRow, functionOrExplicitType) {
        
        if (!sheetCell) {
            return undefined;
        }
        if (sheetCell.t === 'z' && _config.sheetStubs) {
            return null;
        }

        var rawVal = (sheetCell.t === 'n' || sheetCell.t === 'd') ? sheetCell.v : (sheetCell.w && sheetCell.w.trim && sheetCell.w.trim()) || sheetCell.w;
        var returnVal;
        
        if(typeof functionOrExplicitType === 'undefined' || functionOrExplicitType === null){
            return rawVal;
        }
        else if(typeof functionOrExplicitType == 'string'){
            // step into type conversion logic
            switch(functionOrExplicitType.toLowerCase()){
                case 'string':
                    returnvalue = String(rawVal);
                break;
                case 'number':
                    returnvalue = Number(rawVal);
                break;
                case 'boolean':
                    returnvalue = Boolean(rawVal);
                break;
                case 'date':
                    if(sheetCell.t === 'd'){
                        returnVal = rawVal
                    }
                    else{
                        returnvalue = new Date(rawVal);
                    }
                break;
            }
        }
        else if(typeof functionOrExplicitType === 'function'){

            // Try to run the function with returned value
            functionOrExplicitType.call(rawVal, sheetRow)
        }
       
        
    };

    

    const parseSheet = (sheetData, workbook) => {
        const sheetName = (sheetData.constructor == String) ? sheetData : sheetData.name;
        const sheet = workbook.Sheets[sheetName];
        const columnToKey = sheetData.columnToKey || _config.columnToKey;
        const range = sheetData.range || _config.range;
        const headerRows = (sheetData.header && sheetData.header.rows) || (_config.header && _config.header.rows);
        const headerRowToKeys = (sheetData.header && sheetData.header.rowToKeys) || (_config.header && _config.header.rowToKeys);

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
            if (cell == '!ref' || (sheet[cell].v === undefined && !(_config.sheetStubs && sheet[cell].t === 'z'))) {
                continue;
            }

            const row = getCellRow(cell);
            const column = getCellColumn(cell);


            // Is a Header row
            if (headerRows && row <= headerRows) {
                continue;
            }

            // This column is not _configured to be retrieved
            if (columnToKey && !(columnToKey[column] || columnToKey['*'])) {
                continue;
            }

            // This cell is out of the _configured range
            if ((strictRangeColumns && strictRangeRows) && (column < strictRangeColumns.from || column > strictRangeColumns.to || row < strictRangeRows.from || row > strictRangeRows.to)) {
                continue;
            }


            // Need this declared above this block now
            const rowData = rows[row] = rows[row] || {};
            var columnData;
            var transformValue = null;

            // Work out what config is in the columToKey mapping
            if(typeof columnToKey[column] === 'object' && columnToKey[column] !== null){
                // Then expect either an explicit type conversion or a transform function
                // Either way we must always expect a property called 'property' in this instance
                if(columnToKey[column].hasOwnProperty('property') && typeof columnToKey[column]['property'] === 'string'){
                    columnData = columnToKey[column]['property'];
                }   
                else{
                    //Throw invalid config error
                }
                if(columnToKey[column].hasOwnProperty('transform') && typeof columnToKey[column]['transform'] === 'string'){
                    transformValue = columnToKey[column]['transform'];
                }   
            }
            else if (typeof columnToKey[column] == 'string'){

                // Then this must just be a columnToKey mapping with no explicit type conversion specific or no transform function specific
                let columnData = (columnToKey && (columnToKey[column] || columnToKey['*'])) ?
                columnToKey[column] || columnToKey['*'] :
                (headerRowToKeys) ?
                `{{${column}${headerRowToKeys}}}` :
                column;

            }

            let dataVariables = columnData.match(/{{([^}}]+)}}/g);
            if (dataVariables) {
                dataVariables.forEach(dataVariable => {
                    let dataVariableRef = dataVariable.replace(/[\{\}]*/gi, '');
                    let variableValue;
                    switch (dataVariableRef) {
                        case 'columnHeader':
                            dataVariableRef = (headerRows) ? `${column}${headerRows}` : `${column + 1}`;
                        default:
                            variableValue = getSheetCellValue(sheet[dataVariableRef], row, transformValue);
                    }
                    columnData = columnData.replace(dataVariable, variableValue);
                });
            }

            if (columnData === '') {
                continue;
            }

            rowData[columnData] = getSheetCellValue(sheet[cell], row, transformValue);
            
            if (sheetData.appendData) {
                extend(true, rowData, sheetData.appendData);
            }
        }

        // removing first row i.e. 0th rows because first cell itself starts from A1
        rows.shift();

        // Cleaning empty if required
        if (!_config.includeEmptyLines) {
            rows = rows.filter(v => v !== null && v !== undefined);
        }

        return rows;
    };

    const convertExcelToJson = function(config = {}, sourceFile) {
        _config = config.constructor === String ? JSON.parse(config) : config;
        _config.sourceFile = _config.sourceFile || sourceFile;
        
        // ignoring empty lines by default
        _config.includeEmptyLines = _config.includeEmptyLines || false;

        // at least sourceFile or source has to be defined and have a value
        if (!(_config.sourceFile || _config.source)) {
            throw new Error(':: \'sourceFile\' or \'source\' required for _config :: ');
        }

        let workbook = {};

        if (_config.source) {
            workbook = XLSX.read(_config.source, {
                sheetStubs: true,
                cellDates: true
            });
        } else {
            workbook = XLSX.readFile(_config.sourceFile, {
                sheetStubs: true,
                cellDates: true
            });
        }

        let sheetsToGet = (_config.sheets && _config.sheets.constructor === Array) ?
            _config.sheets :
            Object.keys(workbook.Sheets).slice(0, (_config && _config.sheets && _config.sheets.numberOfSheetsToGet) || undefined);

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
