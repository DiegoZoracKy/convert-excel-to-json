'use strict';

const assert = require('assert');
const excelToJson = require('../');
const buildWorkbook = require('./helpers/workbook');

// #25: define sheets by index
describe('"sheets" by index (#25)', function() {
	const source = buildWorkbook({
		Zeta: [['h'], ['zeta']],
		Alpha: [['h'], ['alpha']],
		Middle: [['h'], ['middle']]
	});

	it('should get sheets by their position on the workbook', function() {
		const result = excelToJson({ source, sheets: [0, 2] });
		assert.deepStrictEqual(result, { Zeta: [{ A: 'h' }, { A: 'zeta' }], Middle: [{ A: 'h' }, { A: 'middle' }] });
	});

	it('should accept { index } with a per-sheet config, mixed with names', function() {
		const result = excelToJson({ source, sheets: [{ index: 1, header: { rows: 1 }, columnToKey: { A: 'value' } }, 'Middle'] });
		assert.deepStrictEqual(result, { Alpha: [{ value: 'alpha' }], Middle: [{ A: 'h' }, { A: 'middle' }] });
	});

	it('should throw a clear error for an index that does not exist', function() {
		assert.throws(() => excelToJson({ source, sheets: [3] }), /no sheet at index 3/);
	});

	it('should keep "numberOfSheetsToGet" behaving as before', function() {
		const result = excelToJson({ source, sheets: { numberOfSheetsToGet: 2 } });
		assert.deepStrictEqual(Object.keys(result), ['Zeta', 'Alpha']);
	});
});

describe('"numberOfSheetsToGet" backward compatibility (review of #87)', function() {
	// Object.keys lists integer-like sheet names first: callers relying on the previous
	// behaviour must keep getting '1' for a workbook ordered ['Sales', '1']
	const XLSX = buildWorkbook.XLSX;
	const workbook = XLSX.utils.book_new();
	XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['sales']]), 'Sales');
	XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['one']]), '1');
	const source = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });

	it('should be a workbook ordered [Sales, 1]', function() {
		assert.deepStrictEqual(XLSX.read(source, { type: 'buffer' }).SheetNames, ['Sales', '1']);
	});

	it('should keep returning the same sheet as before for numberOfSheetsToGet: 1', function() {
		assert.deepStrictEqual(Object.keys(excelToJson({ source, sheets: { numberOfSheetsToGet: 1 } })), ['1']);
	});

	it('should select by workbook position only when an index is used', function() {
		assert.deepStrictEqual(Object.keys(excelToJson({ source, sheets: [0] })), ['Sales']);
		assert.deepStrictEqual(excelToJson({ source, sheets: [1] }), { 1: [{ A: 'one' }] });
	});
});
