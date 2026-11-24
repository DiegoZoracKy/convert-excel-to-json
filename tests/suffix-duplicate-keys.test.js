'use strict';

const assert = require('assert');
const excelToJson = require('../');
const buildWorkbook = require('./helpers/workbook');

// #21: header rows with duplicated names lose data
describe('"suffixDuplicateKeys" config (#21)', function() {
	const source = buildWorkbook({
		sheet1: [
			['test', 'test', 1, 'test'],
			[1, 2, 2, 3],
			[2, 3, 2, 4]
		]
	});
	const config = { source, header: { rows: 1 }, columnToKey: { '*': '{{columnHeader}}' } };

	it('should keep the current behaviour by default (last column wins)', function() {
		const result = excelToJson(config);
		assert.deepStrictEqual(result.sheet1, [{ test: 3, 1: 2 }, { test: 4, 1: 2 }]);
	});

	it('should suffix duplicated keys when "suffixDuplicateKeys" is true', function() {
		const result = excelToJson(Object.assign({ suffixDuplicateKeys: true }, config));
		assert.deepStrictEqual(result.sheet1, [
			{ test: 1, test_2: 2, 1: 2, test_3: 3 },
			{ test: 2, test_2: 3, 1: 2, test_3: 4 }
		]);
	});

	it('should not collide with a column that is already named like a suffixed key', function() {
		const src = buildWorkbook({ sheet1: [['a', 'a_2', 'a'], [1, 2, 3]] });
		const result = excelToJson({ source: src, header: { rows: 1 }, columnToKey: { '*': '{{columnHeader}}' }, suffixDuplicateKeys: true });
		assert.deepStrictEqual(result.sheet1, [{ a: 1, a_2: 2, a_3: 3 }]);
	});

	it('should be configurable per sheet', function() {
		const result = excelToJson({
			source,
			header: { rows: 1 },
			columnToKey: { '*': '{{columnHeader}}' },
			sheets: [{ name: 'sheet1', suffixDuplicateKeys: true }]
		});
		assert.deepStrictEqual(Object.keys(result.sheet1[0]), ['1', 'test', 'test_2', 'test_3']);
	});
});
