'use strict';

const assert = require('assert');
const excelToJson = require('../');
const buildWorkbook = require('./helpers/workbook');

// #4 (transform data during conversion) and #31 (apply toFixed to a column)
describe('columnToKey with { property, transform } (#4, #31)', function() {
	const source = buildWorkbook({
		sheet1: [
			['id', 'first', 'last', 'phone', 'price'],
			[1, 'Judith', 'Bennett', 5511999990000, 1.4400000000000002],
			[2, 'Denise', 'Ramirez', 5511888880000, 1.2960000000000003]
		]
	});

	it('should transform values and rename the key', function() {
		const result = excelToJson({
			source,
			header: { rows: 1 },
			columnToKey: {
				A: 'id',
				D: { property: 'phoneNumber', transform: value => value.toString() },
				E: { property: 'price', transform: value => Number(value.toFixed(2)) }
			}
		});
		assert.deepStrictEqual(result.sheet1, [
			{ id: 1, phoneNumber: '5511999990000', price: 1.44 },
			{ id: 2, phoneNumber: '5511888880000', price: 1.3 }
		]);
	});

	it('should give the whole row (original values) to the transform, regardless of column order', function() {
		const result = excelToJson({
			source,
			header: { rows: 1 },
			columnToKey: {
				A: { property: 'label', transform: (value, row) => `${value}-${row.last}` },
				C: { property: 'last', transform: value => value.toUpperCase() }
			}
		});
		assert.deepStrictEqual(result.sheet1, [
			{ label: '1-Bennett', last: 'BENNETT' },
			{ label: '2-Ramirez', last: 'RAMIREZ' }
		]);
	});

	it('should support special variables on "property" and "*"', function() {
		const result = excelToJson({
			source,
			header: { rows: 1 },
			columnToKey: {
				'*': { property: '{{columnHeader}}', transform: value => typeof value === 'number' ? Number(value.toFixed(1)) : value }
			}
		});
		assert.deepStrictEqual(result.sheet1[0], { id: 1, first: 'Judith', last: 'Bennett', phone: 5511999990000, price: 1.4 });
	});

	it('should default "property" to the column letter', function() {
		const result = excelToJson({ source, header: { rows: 1 }, columnToKey: { B: { transform: v => v.length } } });
		assert.deepStrictEqual(result.sheet1, [{ B: 6 }, { B: 6 }]);
	});

	it('should throw a clear error when transform is not a function', function() {
		assert.throws(() => excelToJson({ source, columnToKey: { A: { property: 'id', transform: 'x' } } }), /transform must be a Function/);
	});

	it('should keep string mappings working as before', function() {
		const result = excelToJson({ source, header: { rows: 1 }, columnToKey: { A: 'id', B: 'name' } });
		assert.deepStrictEqual(result.sheet1, [{ id: 1, name: 'Judith' }, { id: 2, name: 'Denise' }]);
	});
});

describe('transform (review of #82)', function() {
	it('should run a transform only on the value of its own column when two columns map to the same key', function() {
		const source = buildWorkbook({ sheet1: [['a', 'b'], ['from A', 'from B']] });

		// B (no transform) is written last, so it owns the key: A's transform must not run on B's value
		const lastWithoutTransform = excelToJson({
			source, header: { rows: 1 },
			columnToKey: { A: { property: 'x', transform: v => `transformed ${v}` }, B: 'x' }
		});
		assert.deepStrictEqual(lastWithoutTransform.sheet1, [{ x: 'from B' }]);

		// B (with transform) is written last: its transform runs on its own value
		const lastWithTransform = excelToJson({
			source, header: { rows: 1 },
			columnToKey: { A: 'x', B: { property: 'x', transform: v => `transformed ${v}` } }
		});
		assert.deepStrictEqual(lastWithTransform.sheet1, [{ x: 'transformed from B' }]);
	});

	it('should not let a conversion started inside a transform change the outer conversion config', function() {
		const source = buildWorkbook({ sheet1: [['code'], ['A1'], [null], ['B2']] });
		const lookupSource = buildWorkbook({ sheet1: [['A1', 'Apple'], ['B2', 'Banana']] });

		const result = excelToJson({
			source,
			header: { rows: 1 },
			includeEmptyLines: true,
			columnToKey: {
				A: {
					property: 'name',
					transform: value => excelToJson({ source: lookupSource, columnToKey: { A: 'code', B: 'name' } }).sheet1.find(row => row.code === value).name
				}
			}
		});

		// includeEmptyLines of the outer call still applies: slots for the skipped header row 1 and the empty row 3
		assert.equal(result.sheet1.length, 4);
		assert.deepStrictEqual(result.sheet1[1], { name: 'Apple' });
		assert.strictEqual(result.sheet1[2], undefined);
		assert.deepStrictEqual(result.sheet1[3], { name: 'Banana' });
	});
});

describe('transform callbacks (second review of #82)', function() {
	it('should receive every cell of the sheet row (mapped or not) as the 3rd argument', function() {
		const source = buildWorkbook({ s: [['qty', 'price', 'note'], [2, 10, 'x']] });
		const seen = [];
		const result = excelToJson({
			source,
			header: { rows: 1 },
			range: 'A1:A2',
			columnToKey: {
				A: { property: 'total', transform: (value, row, cells) => { seen.push(row, cells); return value * cells.B; } }
			}
		});
		assert.deepStrictEqual(result.s, [{ total: 20 }]);
		assert.deepStrictEqual(seen, [{ total: 2 }, { A: 2, B: 10, C: 'x' }]);
	});

	it('should not leave an empty row when the only cell of a row gets an empty key', function() {
		const source = buildWorkbook({ s: [['a', ''], [1, null], [null, 2]] });
		// the header of column B is empty, so its key is '' and its cells are skipped
		const result = excelToJson({ source, header: { rows: 1 }, columnToKey: { '*': { property: '{{columnHeader}}', transform: v => v } } });
		assert.deepStrictEqual(result.s, [{ a: 1 }]);
	});
});
