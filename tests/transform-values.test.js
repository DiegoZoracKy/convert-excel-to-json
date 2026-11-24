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
