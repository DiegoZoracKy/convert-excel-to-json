'use strict';

const XLSX = require('xlsx');
const extend = require('node.extend');

const excelToJson = (function() {

    let _config = {};
    let _date1904 = false;

    const getCellRow = cell => Number(cell.replace(/[A-z]/gi, ''));
    const getCellColumn = cell => cell.replace(/[0-9]/g, '').toUpperCase();
    // 'A' -> 1, 'Z' -> 26, 'AA' -> 27. Column letters can't be compared as strings ('AA' < 'Z').
    const getColumnNumber = column => column.split('').reduce((number, letter) => number * 26 + letter.charCodeAt(0) - 64, 0);
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
        // formattedValues: return the text as displayed on Excel (e.g. '22.5%', '1/1/19') (#62, #65).
        // The display text of numbers/dates/booleans is kept as it is: padding added by the
        // number format (e.g. accounting formats) is significant. Text cells follow the
        // default text handling below.
        if (_config.formattedValues && sheetCell.t !== 's' && sheetCell.w !== undefined) {
            return sheetCell.w;
        }
        if (sheetCell.t === 'n' || sheetCell.t === 'd') {
            return sheetCell.v;
        }
        // `trim: false` keeps leading/trailing whitespace of the cell values (#72)
        if (_config.trim === false) {
            return sheetCell.w;
        }
        return (sheetCell.w && sheetCell.w.trim && sheetCell.w.trim()) || sheetCell.w;
    };

    // parseBooleans: boolean cells and 'true'/'false' texts become Booleans (#55).
    // Applied only to values, never to header cells used as keys.
    function getSheetCellDataValue(sheetCell) {
        if (_config.parseBooleans && sheetCell) {
            if (sheetCell.t === 'b') {
                return sheetCell.v;
            }
            if (sheetCell.t === 's' && typeof sheetCell.v === 'string' && /^\s*(true|false)\s*$/i.test(sheetCell.v)) {
                return sheetCell.v.trim().toLowerCase() === 'true';
            }
        }
        return getSheetCellValue(sheetCell);
    }

    const parseSheet = (sheetData, workbook) => {
        const sheetName = (sheetData.constructor == String) ? sheetData : sheetData.name;
        const sheet = workbook.Sheets[sheetName];
        const columnToKey = sheetData.columnToKey || _config.columnToKey;
        const range = sheetData.range || _config.range;
        const headerRows = (sheetData.header && sheetData.header.rows) || (_config.header && _config.header.rows);
        const headerRowToKeys = (sheetData.header && sheetData.header.rowToKeys) || (_config.header && _config.header.rowToKeys);
        // header.include: the header rows are used for {{columnHeader}} but are also kept on the result (#36)
        const headerInclude = (sheetData.header && sheetData.header.include !== undefined) ? sheetData.header.include : (_config.header && _config.header.include);
        const suffixDuplicateKeys = sheetData.suffixDuplicateKeys !== undefined ? sheetData.suffixDuplicateKeys : _config.suffixDuplicateKeys;

        // The key (and transform) of a column depends only on the column (mapping and header cells).
        // Returns { key, transform }; key is undefined when a columnToKey function returned null/undefined.
        const getColumnSpec = column => {
            let columnData = (columnToKey && (columnToKey[column] || columnToKey['*'])) ?
                columnToKey[column] || columnToKey['*'] :
                (headerRowToKeys) ?
                `{{${column}${headerRowToKeys}}}` :
                column;

            // { property: 'keyName', transform: (value, row) => newValue } (#4)
            let transform;
            if (columnData && typeof columnData === 'object') {
                transform = columnData.transform;
                columnData = columnData.property || column;
                if (transform !== undefined && typeof transform !== 'function') {
                    throw new TypeError(`:: columnToKey['${column}'].transform must be a Function :: `);
                }
            }

            // A function can be used to build the key (e.g. { '*': _.camelCase }).
            // It receives the column header value and the column letter.
            // Keys returned by a function are used as they are: no {{variables}} substitution,
            // so a header like 'Cost {{USD}}' is kept (review of #78).
            let isFunctionKey = false;
            if (typeof columnData === 'function') {
                isFunctionKey = true;
                const headerCell = headerRows ? `${column}${headerRows}` : `${column}1`;
                columnData = columnData(getSheetCellValue(sheet[headerCell]), column);
                if (columnData === undefined || columnData === null) {
                    return { key: undefined, transform };
                }
                columnData = String(columnData);
            } else if (typeof columnData !== 'string') {
                throw new TypeError(`:: columnToKey['${column}'] must be a String or a Function, got ${typeof columnData} :: `);
            }

            let dataVariables = !isFunctionKey && columnData.match(/{{([^}}]+)}}/g);
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
            return { key: columnData, transform };
        };
        const getColumnKey = column => getColumnSpec(column).key;

        // Two columns resolving to the same key would overwrite each other.
        // With suffixDuplicateKeys they become 'key', 'key_2', 'key_3'... assigned in
        // column order (A, B, C...), regardless of which cells have values (#21).
        const keyByColumn = {};
        if (suffixDuplicateKeys && sheet) {
            const usedKeys = {};
            const columns = {};
            Object.keys(sheet).forEach(cell => {
                if (cell[0] !== '!') {
                    columns[getCellColumn(cell)] = true;
                }
            });
            Object.keys(columns)
                .filter(column => !columnToKey || columnToKey[column] || columnToKey['*'])
                .sort((a, b) => XLSX.utils.decode_col(a) - XLSX.utils.decode_col(b))
                .forEach(column => {
                    const key = getColumnKey(column);
                    if (key === '' || key === undefined) {
                        return;
                    }
                    let uniqueKey = key;
                    let counter = 1;
                    while (usedKeys[uniqueKey]) {
                        uniqueKey = `${key}_${++counter}`;
                    }
                    usedKeys[uniqueKey] = true;
                    keyByColumn[column] = uniqueKey;
                });
        }

        let strictRangeColumns;
        let strictRangeRows;
        if (range) {
            strictRangeColumns = {
                from: getColumnNumber(getCellColumn(getRangeBegin(range))),
                to: getColumnNumber(getCellColumn(getRangeEnd(range)))
            };

            strictRangeRows = {
                from: getCellRow(getRangeBegin(range)),
                to: getCellRow(getRangeEnd(range))
            };
        }

        const hiddenColumns = (sheet && sheet['!cols'] || []).map(col => !!(col && col.hidden));
        const hiddenRows = (sheet && sheet['!rows'] || []).map(row => !!(row && row.hidden));

        let rows = [];
        const rowsTransforms = [];
        for (let cell in sheet) {

            // Keys starting with '!' (!ref, !merges, !margins, !cols...) are sheet metadata, not cells || this cell doesn't have a value
            if (cell[0] === '!' || (sheet[cell].v === undefined && !(_config.sheetStubs && sheet[cell].t === 'z'))) {
                continue;
            }

            const row = getCellRow(cell);
            const column = getCellColumn(cell);

            // skipHiddenColumns / skipHiddenRows (#46)
            if (_config.skipHiddenColumns && hiddenColumns[XLSX.utils.decode_col(column)]) {
                continue;
            }
            if (_config.skipHiddenRows && hiddenRows[row - 1]) {
                continue;
            }

            // Is a Header row
            if (headerRows && row <= headerRows && !headerInclude) {
                continue;
            }

            // This column is not _configured to be retrieved
            if (columnToKey && !(columnToKey[column] || columnToKey['*'])) {
                continue;
            }

            // This cell is out of the _configured range
            if ((strictRangeColumns && strictRangeRows) && (getColumnNumber(column) < strictRangeColumns.from || getColumnNumber(column) > strictRangeColumns.to || row < strictRangeRows.from || row > strictRangeRows.to)) {
                continue;
            }


            const rowData = rows[row] = rows[row] || {};
            const columnSpec = getColumnSpec(column);
            const transform = columnSpec.transform;
            let columnData = suffixDuplicateKeys ? keyByColumn[column] : columnSpec.key;

            if (columnData === '' || columnData === undefined) {
                continue;
            }

            let value = getSheetCellDataValue(sheet[cell]);

            // whitespaceToNull: cells with only blank spaces become null (#9)
            if (_config.whitespaceToNull && typeof value === 'string' && value.trim() === '') {
                value = null;
            }

            rowData[columnData] = value;

            // One pending transform per key: the column that wrote the value last owns the key,
            // so a transform never runs on the value of another column mapped to the same key.
            const rowTransforms = rowsTransforms[row] = rowsTransforms[row] || {};
            if (transform) {
                rowTransforms[columnData] = { transform, value: rowData[columnData] };
            } else {
                delete rowTransforms[columnData];
            }
            
            if (sheetData.appendData) {
                extend(true, rowData, sheetData.appendData);
            }
        }

        // Transforms run once the row is complete, so they can read any column of it.
        // They receive the original (not transformed) values.
        rowsTransforms.forEach((rowTransforms, row) => {
            const rowData = rows[row];
            const originalRow = extend({}, rowData);
            Object.keys(rowTransforms).forEach(key => {
                const { transform, value } = rowTransforms[key];
                rowData[key] = transform(value, originalRow);
            });
        });

        // whitespaceToNull + sheetStubs: every column of the range (or of the sheet) is present
        // on each row, with null for empty, blank and absent cells (#9)
        if (_config.whitespaceToNull && _config.sheetStubs && sheet && (range || sheet['!ref'])) {
            const bounds = XLSX.utils.decode_range(range || sheet['!ref']);
            const keys = [];
            for (let c = bounds.s.c; c <= bounds.e.c; c++) {
                const column = XLSX.utils.encode_col(c);
                if (columnToKey && !(columnToKey[column] || columnToKey['*'])) {
                    continue;
                }
                if (_config.skipHiddenColumns && hiddenColumns[c]) {
                    continue;
                }
                const key = suffixDuplicateKeys ? keyByColumn[column] : getColumnKey(column);
                if (key !== '' && key !== undefined && keys.indexOf(key) === -1) {
                    keys.push(key);
                }
            }
            rows.forEach((rowData, row) => {
                const filled = {};
                keys.forEach(key => {
                    filled[key] = rowData[key] === undefined ? null : rowData[key];
                });
                rows[row] = extend(filled, rowData);
            });
        }

        // removing first row i.e. 0th rows because first cell itself starts from A1
        rows.shift();

        // Cleaning empty if required
        if (!_config.includeEmptyLines) {
            rows = rows.filter(v => v !== null && v !== undefined);
        }

        return rows;
    };

    // The config is kept on module-level state while a conversion runs. A callback (e.g. a
    // columnToKey function or a transform) may start another conversion, so the outer config
    // is restored when the inner one ends: conversions are reentrant (review of #78 / #82).
    const convertExcelToJson = function(config, sourceFile) {
        const outerConfig = _config;
        try {
            return convert(config, sourceFile);
        } finally {
            _config = outerConfig;
        }
    };

    const convert = function(config = {}, sourceFile) {
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
            cellNF: true,
            // needed to know which columns/rows are hidden
            cellStyles: !!(_config.skipHiddenColumns || _config.skipHiddenRows)
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
