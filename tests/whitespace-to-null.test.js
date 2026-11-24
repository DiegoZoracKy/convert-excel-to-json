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

describe('"whitespaceToNull" + "sheetStubs" fill every column (review of #86)', function() {
	const source = buildWorkbook({
		sheet1: [
			['id', 'col1', 'col2', 'col3'],
			[11111, 'x', 'x', '   '],
			[22222, null, 'x'] // B3 and D3 are absent (no cell at all)
		]
	});

	it('should give null for absent cells too, keeping the column order', function() {
		const result = excelToJson({ source, header: { rows: 1 }, sheetStubs: true, whitespaceToNull: true, columnToKey: { '*': '{{columnHeader}}' } });
		assert.deepStrictEqual(result.sheet1, [
			{ id: 11111, col1: 'x', col2: 'x', col3: null },
			{ id: 22222, col1: null, col2: 'x', col3: null }
		]);
		assert.deepStrictEqual(Object.keys(result.sheet1[1]), ['id', 'col1', 'col2', 'col3']);
	});

	it('should follow columnToKey and range', function() {
		const result = excelToJson({ source, sheetStubs: true, whitespaceToNull: true, range: 'A2:C3', columnToKey: { A: 'id', B: 'first' } });
		assert.deepStrictEqual(result.sheet1, [{ id: 11111, first: 'x' }, { id: 22222, first: null }]);
	});

	it('should not add properties without both options (backward compatible)', function() {
		const result = excelToJson({ source, header: { rows: 1 }, sheetStubs: true });
		assert.deepStrictEqual(result.sheet1[1], { A: 22222, C: 'x' });
	});

	it('should not create rows for empty lines', function() {
		const src = buildWorkbook({ sheet1: [['a', 'b'], [1, 2], [], [3]] });
		const result = excelToJson({ source: src, sheetStubs: true, whitespaceToNull: true });
		assert.deepStrictEqual(result.sheet1, [{ A: 'a', B: 'b' }, { A: 1, B: 2 }, { A: 3, B: null }]);
	});
});
