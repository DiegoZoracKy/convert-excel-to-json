'use strict';

const assert = require('assert');
const excelToJson = require('../');
const buildWorkbook = require('./helpers/workbook');

// #46: skip hidden columns (and rows)
describe('"skipHiddenColumns" and "skipHiddenRows" config (#46)', function() {
	const XLSX = buildWorkbook.XLSX;
	const ws = XLSX.utils.aoa_to_sheet([
		['id', 'secret', 'name'],
		[1, 's1', 'Judith'],
		[2, 's2', 'Denise'],
		[3, 's3', 'Antonio']
	]);
	ws['!cols'] = [{}, { hidden: true }, {}];
	ws['!rows'] = [{}, {}, { hidden: true }];
	const source = buildWorkbook({ sheet1: ws });

	it('should return hidden columns and rows by default (backward compatible)', function() {
		const result = excelToJson({ source, header: { rows: 1 } });
		assert.deepStrictEqual(result.sheet1, [
			{ A: 1, B: 's1', C: 'Judith' },
			{ A: 2, B: 's2', C: 'Denise' },
			{ A: 3, B: 's3', C: 'Antonio' }
		]);
	});

	it('should skip hidden columns with "skipHiddenColumns"', function() {
		const result = excelToJson({ source, header: { rows: 1 }, columnToKey: { '*': '{{columnHeader}}' }, skipHiddenColumns: true });
		assert.deepStrictEqual(result.sheet1, [
			{ id: 1, name: 'Judith' },
			{ id: 2, name: 'Denise' },
			{ id: 3, name: 'Antonio' }
		]);
	});

	it('should skip hidden rows with "skipHiddenRows"', function() {
		const result = excelToJson({ source, header: { rows: 1 }, skipHiddenRows: true });
		assert.deepStrictEqual(result.sheet1.map(row => row.A), [1, 3]);
	});

	it('should skip both', function() {
		const result = excelToJson({ source, header: { rows: 1 }, skipHiddenColumns: true, skipHiddenRows: true });
		assert.deepStrictEqual(result.sheet1, [{ A: 1, C: 'Judith' }, { A: 3, C: 'Antonio' }]);
	});
});
