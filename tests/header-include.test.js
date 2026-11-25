'use strict';

const assert = require('assert');
const excelToJson = require('../');
const buildWorkbook = require('./helpers/workbook');

// #36: include the header row on the result even when it's specified
describe('"header.include" config (#36)', function() {
	const source = buildWorkbook({
		sheet1: [
			['Report title'],
			['id', 'name'],
			[1, 'Judith'],
			[2, 'Denise']
		]
	});

	it('should skip the header rows by default (backward compatible)', function() {
		const result = excelToJson({ source, header: { rows: 2 }, columnToKey: { '*': '{{columnHeader}}' } });
		assert.deepStrictEqual(result.sheet1, [{ id: 1, name: 'Judith' }, { id: 2, name: 'Denise' }]);
	});

	it('should keep the header rows with "header.include"', function() {
		const result = excelToJson({ source, header: { rows: 2, include: true }, columnToKey: { '*': '{{columnHeader}}' } });
		assert.deepStrictEqual(result.sheet1, [
			{ id: 'Report title' },
			{ id: 'id', name: 'name' },
			{ id: 1, name: 'Judith' },
			{ id: 2, name: 'Denise' }
		]);
	});

	it('should be configurable per sheet', function() {
		const result = excelToJson({ source, sheets: [{ name: 'sheet1', header: { rows: 2, include: true } }], columnToKey: { A: 'a' } });
		assert.equal(result.sheet1.length, 4);
	});
});
