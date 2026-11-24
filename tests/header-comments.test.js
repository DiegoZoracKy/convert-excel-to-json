'use strict';

const assert = require('assert');
const excelToJson = require('../');

// #23: "When header is having comment it's not working"
// tests/test-data-comments.xlsx has a threaded comment (Excel 365) on A1, a legacy note on B1
// and a note on a data cell (C2).
describe('cells with comments (#23)', function() {
	const sourceFile = `${__dirname}/test-data-comments.xlsx`;
	const expected = [
		{ Name: 'Ann', Age: 30, City: 'Lisbon' },
		{ Name: 'Bob', Age: 25, City: 'Recife' }
	];

	it('should use header cells with comments on {{columnHeader}}', function() {
		const result = excelToJson({ sourceFile, header: { rows: 1 }, columnToKey: { '*': '{{columnHeader}}' } });
		assert.deepStrictEqual(result.sheet1, expected);
	});

	it('should use the first row as header (without header.rows) when it has comments', function() {
		const result = excelToJson({ sourceFile, columnToKey: { '*': '{{columnHeader}}' } });
		assert.deepStrictEqual(result.sheet1, [{ Name: 'Name', Age: 'Age', City: 'City' }].concat(expected));
	});

	it('should work with comments together with "sheetStubs"', function() {
		const result = excelToJson({ sourceFile, header: { rows: 1 }, sheetStubs: true, columnToKey: { '*': '{{columnHeader}}' } });
		assert.deepStrictEqual(result.sheet1, expected);
	});

	it('should not return the comments as data', function() {
		const result = excelToJson({ sourceFile });
		assert.deepStrictEqual(result.sheet1[1], { A: 'Ann', B: 30, C: 'Lisbon' });
	});
});
