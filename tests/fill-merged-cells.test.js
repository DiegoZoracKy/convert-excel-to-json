'use strict';

const assert = require('assert');
const excelToJson = require('../');
const buildWorkbook = require('./helpers/workbook');

// #64: rows with merged cells inside
describe('"fillMergedCells" config (#64)', function() {
	const XLSX = buildWorkbook.XLSX;
	const ws = XLSX.utils.aoa_to_sheet([
		['Category', 'Place', 'Term'],
		['FirstCat', 'Inside', 12],
		[null, null, 360],
		['SecondCat', 'Outside', 24]
	]);
	ws['!merges'] = [
		{ s: { r: 1, c: 0 }, e: { r: 2, c: 0 } }, // A2:A3
		{ s: { r: 1, c: 1 }, e: { r: 2, c: 1 } } // B2:B3
	];
	const source = buildWorkbook({ sheet1: ws });
	const config = { source, header: { rows: 1 }, columnToKey: { '*': '{{columnHeader}}' } };

	it('should keep the value only on the first cell by default (backward compatible)', function() {
		const result = excelToJson(config);
		assert.deepStrictEqual(result.sheet1[1], { Term: 360 });
	});

	it('should repeat the merged value on every cell of the range with "fillMergedCells"', function() {
		const result = excelToJson(Object.assign({ fillMergedCells: true }, config));
		assert.deepStrictEqual(result.sheet1, [
			{ Category: 'FirstCat', Place: 'Inside', Term: 12 },
			{ Category: 'FirstCat', Place: 'Inside', Term: 360 },
			{ Category: 'SecondCat', Place: 'Outside', Term: 24 }
		]);
	});

	it('should fill horizontally merged header cells too', function() {
		const hws = XLSX.utils.aoa_to_sheet([['Group', null], ['a', 'b']]);
		hws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }];
		const result = excelToJson({ source: buildWorkbook({ sheet1: hws }), fillMergedCells: true });
		assert.deepStrictEqual(result.sheet1[0], { A: 'Group', B: 'Group' });
	});

	it('should be configurable per sheet', function() {
		const result = excelToJson(Object.assign({ sheets: [{ name: 'sheet1', fillMergedCells: true }] }, config));
		assert.deepStrictEqual(result.sheet1[1], { Category: 'FirstCat', Place: 'Inside', Term: 360 });
	});
});
