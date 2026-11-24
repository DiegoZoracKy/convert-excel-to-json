'use strict';

const assert = require('assert');
const excelToJson = require('../');
const buildWorkbook = require('./helpers/workbook');

// #9: empty space cells as null values
describe('"whitespaceToNull" config (#9)', function() {
	const source = buildWorkbook({
		sheet1: [
			['id', 'col1', 'col2', 'col3'],
			[11111, 'x', 'x', '   ']
		]
	});

	it('should keep blank spaces by default (backward compatible)', function() {
		const result = excelToJson({ source, header: { rows: 1 } });
		assert.strictEqual(result.sheet1[0].D, '   ');
	});

	it('should return null for blank cells with "whitespaceToNull"', function() {
		const result = excelToJson({ source, header: { rows: 1 }, whitespaceToNull: true });
		assert.deepStrictEqual(result.sheet1[0], { A: 11111, B: 'x', C: 'x', D: null });
	});

	it('should give null for blank and empty cells on a real file when used with "sheetStubs"', function() {
		const sourceFile = `${__dirname}/test-data.xlsx`;
		const result = excelToJson({ sourceFile, sheets: ['sheet3'], sheetStubs: true, whitespaceToNull: true });
		// on sheet3 there is a row made only of blank spaces, and a row with an empty C cell
		const blankRow = result.sheet3.find(row => row.F === null);
		assert.deepStrictEqual(blankRow, { A: null, B: null, C: null, D: null, E: null, F: null });
		const rowWithEmptyCell = result.sheet3.find(row => row.A === 42);
		assert.strictEqual(rowWithEmptyCell.C, null);
	});
});
