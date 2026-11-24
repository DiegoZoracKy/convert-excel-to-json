'use strict';

const XLSX = require('xlsx');
const extend = require('node.extend');

const excelToJson = (function() {

    let _config = {};

    const getCellRow = cell => Number(cell.replace(/[A-z]/gi, ''));
    const getCellColumn = cell => cell.replace(/[0-9]/g, '').toUpperCase();
    const getRangeBegin = cell => cell.match(/^[^:]*/)[0];
    const getRangeEnd = cell => cell.match(/[^:]*$/)[0];
    function getSheetCellValue(sheetCell) {
        if (!sheetCell) {
            return undefined;
        }
        if (sheetCell.t === 'z' && _config.sheetStubs) {
            return null;
        }
        return (sheetCell.t === 'n' || sheetCell.t === 'd') ? sheetCell.v : (sheetCell.w && sheetCell.w.trim && sheetCell.w.trim()) || sheetCell.w;
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
                            variableValue = getSheetCellValue(sheet[dataVariableRef]);
                    }
                    columnData = columnData.replace(dataVariable, variableValue);
                });
            }

            if (columnData === '') {
                continue;
            }

            rowData[columnData] = getSheetCellValue(sheet[cell]);
            
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

    // Encodings accepted for a String `source`. SheetJS also accepts 'file' as a type,
    // which would read from the file system, so only inline encodings are allowed.
    const STRING_SOURCE_TYPES = ['base64', 'binary', 'string'];

    // Realm-safe checks (an ArrayBuffer created in another realm, e.g. a same-origin
    // iframe or a vm context, is not an `instanceof` this realm's ArrayBuffer).
    const getTag = value => Object.prototype.toString.call(value);
    const isArrayBuffer = value => getTag(value) === '[object ArrayBuffer]' || getTag(value) === '[object SharedArrayBuffer]';

    // Picks the right xlsx `type` for the given source (#44).
    // Node.js Buffer -> 'buffer'; ArrayBuffer / Uint8Array (e.g. from a browser FileReader or
    // file.arrayBuffer()) -> 'array'; Strings are base64 by default (or `sourceType`).
    const readSource = function(source, readOptions) {
        if (typeof Buffer !== 'undefined' && Buffer.isBuffer(source)) {
            return XLSX.read(source, Object.assign({ type: 'buffer' }, readOptions));
        }
        if (isArrayBuffer(source)) {
            return XLSX.read(new Uint8Array(source), Object.assign({ type: 'array' }, readOptions));
        }
        if (typeof ArrayBuffer !== 'undefined' && ArrayBuffer.isView(source)) {
            return XLSX.read(new Uint8Array(source.buffer, source.byteOffset, source.byteLength), Object.assign({ type: 'array' }, readOptions));
        }
        if (typeof source === 'string') {
            const sourceType = _config.sourceType || 'base64';
            if (STRING_SOURCE_TYPES.indexOf(sourceType) === -1) {
                throw new Error(`:: 'sourceType' must be one of ${STRING_SOURCE_TYPES.map(type => `'${type}'`).join(', ')}. Use 'sourceFile' to read a file :: `);
            }
            return XLSX.read(source, Object.assign({ type: sourceType }, readOptions));
        }
        return XLSX.read(source, readOptions);
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
            workbook = readSource(_config.source, {
                sheetStubs: true,
                cellDates: true
            });
        } else {
            try {
                workbook = XLSX.readFile(_config.sourceFile, {
                    sheetStubs: true,
                    cellDates: true
                });
            } catch (err) {
                // There is no file system on the browser (#44)
                if (/readFileSync is not a function/.test(err.message)) {
                    throw new Error(':: \'sourceFile\' requires a file system (Node.js). On the browser, read the file as an ArrayBuffer and pass it in as \'source\' :: ');
                }
                throw err;
            }
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
