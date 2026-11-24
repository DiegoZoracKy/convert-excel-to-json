'use strict';

const assert = require('assert');
const excelToJson = require('../');
const buildWorkbook = require('./helpers/workbook');

// #50: rows as headers
describe('"transpose" config (#50)', function() {
	const source = buildWorkbook({
		sheet1: [
			['Type', 'TypeData', 'OtherType'],
			['Date', 'DateData', 'OtherDate']
		],
		sheet2: [
			['ignored', 'ignored', 'ignored'],
			['ignored', 'Season', 'Autumn'],
			['ignored', 'Year', 2021],
			['ignored', 'Round', 1]
		]
	});

	it('should not transpose by default', function() {
		const result = excelToJson({ source, sheets: ['sheet1'] });
		assert.deepStrictEqual(result.sheet1[0], { A: 'Type', B: 'TypeData', C: 'OtherType' });
	});

	it('should swap rows and columns', function() {
		const result = excelToJson({ source, sheets: ['sheet1'], transpose: true });
		assert.deepStrictEqual(result.sheet1, [
			{ A: 'Type', B: 'Date' },
			{ A: 'TypeData', B: 'DateData' },
			{ A: 'OtherType', B: 'OtherDate' }
		]);
	});

	it('should use the first column as header with header.rows and {{columnHeader}}', function() {
		const result = excelToJson({ source, sheets: ['sheet1'], transpose: true, header: { rows: 1 }, columnToKey: { '*': '{{columnHeader}}' } });
		assert.deepStrictEqual(result.sheet1, [
			{ Type: 'TypeData', Date: 'DateData' },
			{ Type: 'OtherType', Date: 'OtherDate' }
		]);
	});

	it('should apply the range on the original sheet and start the transposed data at A1', function() {
		const result = excelToJson({
			source,
			sheets: [{ name: 'sheet2', transpose: true, range: 'B2:C4', header: { rows: 1 }, columnToKey: { '*': '{{columnHeader}}' } }]
		});
		assert.deepStrictEqual(result.sheet2, [{ Season: 'Autumn', Year: 2021, Round: 1 }]);
	});
});
