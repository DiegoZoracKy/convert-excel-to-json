'use strict';

const XLSX = require('xlsx');
const extend = require('node.extend');

const excelToJson = (function() {

    let _config = {};
    let _date1904 = false;

    const getCellRow = cell => Number(cell.replace(/[A-z]/gi, ''));
    const getCellColumn = cell => cell.replace(/[0-9]/g, '').toUpperCase();
    const getRangeBegin = cell => cell.match(/^[^:]*/)[0];
    const getRangeEnd = cell => cell.match(/[^:]*$/)[0];
    const MS_PER_DAY = 24 * 60 * 60 * 1000;

    // SheetJS >= 0.20 represents dates as UTC wall-clock times; older versions use local time.
    const sheetJsUsesUTCDates = (function() {
        const [major, minor] = String(XLSX.version).split('.').map(Number);
        return major > 0 || minor >= 20;
    }());

    // Builds the Date returned to the user from wall-clock parts (#33).
    // By default the Date represents the same wall-clock time in the local timezone
    // (as before). With `utcDates: true` it represents it in UTC.
    function wallClockToDate(y, m, d, H, M, S, ms) {
        return _config.utcDates ?
            new Date(Date.UTC(y, m, d, H, M, S, ms)) :
            new Date(y, m, d, H, M, S, ms);
    }

    // Builds a Date from an Excel serial date number (#33).
    // The xlsx `cellDates` conversion drifts by some seconds (and sometimes hours, on
    // historical DST/LMT offsets) depending on the timezone, e.g. 1/1/2019 became
    // 2018-12-31T15:59:14.000Z. Using the date parts avoids any timezone math, and the
    // time of the day is rounded to the millisecond (12:00:00.500 stays .500).
    function excelSerialToDate(serial) {
        let days = Math.floor(serial);
        let ms = Math.round((serial - days) * MS_PER_DAY);
        if (ms >= MS_PER_DAY) {
            days += 1;
            ms -= MS_PER_DAY;
        }
        const parts = XLSX.SSF.parse_date_code(days, { date1904: _date1904 });
        if (!parts) {
            return serial;
        }
        return wallClockToDate(parts.y, parts.m - 1, parts.d,
            Math.floor(ms / 3600000), Math.floor(ms / 60000) % 60, Math.floor(ms / 1000) % 60, ms % 1000);
    }

    // Dates parsed by SheetJS (cells explicitly typed as dates, t="d", without a date
    // number format) -> the wall-clock time they represent.
    function sheetJsDateToDate(date) {
        return sheetJsUsesUTCDates ?
            wallClockToDate(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate(), date.getUTCHours(), date.getUTCMinutes(), date.getUTCSeconds(), date.getUTCMilliseconds()) :
            wallClockToDate(date.getFullYear(), date.getMonth(), date.getDate(), date.getHours(), date.getMinutes(), date.getSeconds(), date.getMilliseconds());
    }

    // The workbook is read with cellDates, so every date cell (numbers with a date
    // format, and cells explicitly typed as dates) comes with t === 'd'.
    // Numbers with a date format are rebuilt from their exact serial number, read again
    // without cellDates (only when the selected sheets have dates).
    function normalizeDates(workbook, sheetNames, readWorkbook) {
        let serialsWorkbook;
        sheetNames.forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName];
            if (!sheet) {
                return;
            }
            Object.keys(sheet).forEach(ref => {
                const cell = sheet[ref];
                if (ref[0] === '!' || !cell || cell.t !== 'd' || !(cell.v instanceof Date)) {
                    return;
                }
                if (cell.z && XLSX.SSF.is_date(cell.z)) {
                    serialsWorkbook = serialsWorkbook || readWorkbook({ cellDates: false, cellNF: false, cellText: false, sheets: sheetNames });
                    const serialSheet = serialsWorkbook.Sheets[sheetName];
                    const serialCell = serialSheet && serialSheet[ref];
                    if (serialCell && serialCell.t === 'n') {
                        cell.v = excelSerialToDate(serialCell.v);
                        return;
                    }
                }
                cell.v = sheetJsDateToDate(cell.v);
            });
        });
    }

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

    const convertExcelToJson = function(config = {}, sourceFile) {
        _config = config.constructor === String ? JSON.parse(config) : config;
        _config.sourceFile = _config.sourceFile || sourceFile;
        
        // ignoring empty lines by default
        _config.includeEmptyLines = _config.includeEmptyLines || false;

        // at least sourceFile or source has to be defined and have a value
        if (!(_config.sourceFile || _config.source)) {
            throw new Error(':: \'sourceFile\' or \'source\' required for _config :: ');
        }

        const readWorkbook = options => _config.source ?
            XLSX.read(_config.source, options) :
            XLSX.readFile(_config.sourceFile, options);

        const workbook = readWorkbook({
            sheetStubs: true,
            cellDates: true,
            cellNF: true
        });

        _date1904 = !!(workbook.Workbook && workbook.Workbook.WBProps && workbook.Workbook.WBProps.date1904);

        let sheetsToGet = (_config.sheets && _config.sheets.constructor === Array) ?
            _config.sheets :
            Object.keys(workbook.Sheets).slice(0, (_config && _config.sheets && _config.sheets.numberOfSheetsToGet) || undefined);

        normalizeDates(workbook, sheetsToGet.map(sheet => (sheet.constructor == String) ? sheet : sheet.name), readWorkbook);

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
