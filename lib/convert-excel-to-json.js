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

    // Copies the value of each merged range (stored only on its top-left cell)
    // to all the cells of the range (#64). Only the cells inside `range` (when given)
    // are filled, so big merged areas outside of it cost nothing.
    const fillMergedCells = (sheet, range) => {
        if (!sheet || !sheet['!merges'] || !sheet['!merges'].length) {
            return sheet;
        }
        const limits = range ? XLSX.utils.decode_range(range) : null;
        const filled = extend({}, sheet);
        // A merge may extend beyond the recorded bounds (!ref) of the worksheet, e.g. when
        // only its anchor has a value: the filled cells are added to !ref
        let bounds = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : null;
        sheet['!merges'].forEach(merge => {
            const topLeftCell = sheet[XLSX.utils.encode_cell(merge.s)];
            if (!topLeftCell) {
                return;
            }
            const fromR = limits ? Math.max(merge.s.r, limits.s.r) : merge.s.r;
            const toR = limits ? Math.min(merge.e.r, limits.e.r) : merge.e.r;
            const fromC = limits ? Math.max(merge.s.c, limits.s.c) : merge.s.c;
            const toC = limits ? Math.min(merge.e.c, limits.e.c) : merge.e.c;
            if (fromR > toR || fromC > toC) {
                return;
            }
            for (let r = fromR; r <= toR; r++) {
                for (let c = fromC; c <= toC; c++) {
                    filled[XLSX.utils.encode_cell({ r, c })] = extend({}, topLeftCell);
                }
            }
            bounds = bounds ? {
                s: { r: Math.min(bounds.s.r, fromR), c: Math.min(bounds.s.c, fromC) },
                e: { r: Math.max(bounds.e.r, toR), c: Math.max(bounds.e.c, toC) }
            } : { s: { r: fromR, c: fromC }, e: { r: toR, c: toC } };
        });
        if (bounds) {
            filled['!ref'] = XLSX.utils.encode_range(bounds);
        }
        return filled;
    };

    // The top-left (anchor) cell of the merged range that contains `ref`, if any
    const getMergeAnchorCell = (sheet, ref) => {
        const merges = sheet && sheet['!merges'];
        if (!merges || !merges.length) {
            return undefined;
        }
        const { r, c } = XLSX.utils.decode_cell(ref);
        const merge = merges.find(m => r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c);
        return merge ? sheet[XLSX.utils.encode_cell(merge.s)] : undefined;
    };

    const parseSheet = (sheetData, workbook) => {
        const sheetName = (sheetData.constructor == String) ? sheetData : sheetData.name;
        const shouldFillMergedCells = sheetData.fillMergedCells !== undefined ? sheetData.fillMergedCells : _config.fillMergedCells;
        const sheet = shouldFillMergedCells ? fillMergedCells(workbook.Sheets[sheetName], sheetData.range || _config.range) : workbook.Sheets[sheetName];
        const columnToKey = sheetData.columnToKey || _config.columnToKey;
        const range = sheetData.range || _config.range;
        const headerRows = (sheetData.header && sheetData.header.rows) || (_config.header && _config.header.rows);
        const headerRowToKeys = (sheetData.header && sheetData.header.rowToKeys) || (_config.header && _config.header.rowToKeys);
        // Cells referenced by keys (e.g. header cells) may be outside the range, where merged
        // cells are not filled: they are resolved through the anchor of their merged range.
        const anchorCells = new Map();
        // A cell covered by a merge always resolves to the merge's top-left anchor, whatever the
        // covered cell holds: nothing (a styled/bordered stub), an empty string or a stale value
        // kept by the file (Excel only shows the anchor).
        const getReferencedCell = ref => {
            const cell = sheet[ref];
            if (!shouldFillMergedCells) {
                return cell;
            }
            if (!anchorCells.has(ref)) {
                anchorCells.set(ref, getMergeAnchorCell(sheet, ref));
            }
            return anchorCells.get(ref) || cell;
        };

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
                            variableValue = getSheetCellValue(getReferencedCell(dataVariableRef));
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
