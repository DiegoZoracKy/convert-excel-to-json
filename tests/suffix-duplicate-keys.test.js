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

describe('"suffixDuplicateKeys" order (review of #83)', function() {
	it('should assign suffixes by column order, even when the first data row has blank cells', function() {
		const source = buildWorkbook({
			sheet1: [
				['fee', 'fee', 'name'],
				[null, 20, 'Ann'], // A2 blank: B is found first
				[10, 30, 'Bob']
			]
		});
		const result = excelToJson({ source, header: { rows: 1 }, columnToKey: { '*': '{{columnHeader}}' }, suffixDuplicateKeys: true });
		assert.deepStrictEqual(result.sheet1, [
			{ fee_2: 20, name: 'Ann' },
			{ fee: 10, fee_2: 30, name: 'Bob' }
		]);
	});

	it('should follow the column order after column Z', function() {
		const header = [];
		for (let i = 0; i < 28; i++) {
			header.push(i === 1 || i === 27 ? 'dup' : `h${i}`);
		}
		const values = header.map((h, i) => (i === 1 ? null : i));
		const source = buildWorkbook({ sheet1: [header, values, header.map((h, i) => i)] });
		const result = excelToJson({ source, header: { rows: 1 }, columnToKey: { '*': '{{columnHeader}}' }, suffixDuplicateKeys: true });
		assert.strictEqual(result.sheet1[0].dup, undefined); // B2 is blank
		assert.strictEqual(result.sheet1[0].dup_2, 27); // AB2
		assert.strictEqual(result.sheet1[1].dup, 1);
	});
});
