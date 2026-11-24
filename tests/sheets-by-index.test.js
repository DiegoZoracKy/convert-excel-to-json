'use strict';

const assert = require('assert');
const excelToJson = require('../');
const buildWorkbook = require('./helpers/workbook');

// #25: define sheets by index
describe('"sheets" by index (#25)', function() {
	const source = buildWorkbook({
		Zeta: [['h'], ['zeta']],
		Alpha: [['h'], ['alpha']],
		Middle: [['h'], ['middle']]
	});

	it('should get sheets by their position on the workbook', function() {
		const result = excelToJson({ source, sheets: [0, 2] });
		assert.deepStrictEqual(result, { Zeta: [{ A: 'h' }, { A: 'zeta' }], Middle: [{ A: 'h' }, { A: 'middle' }] });
	});

	it('should accept { index } with a per-sheet config, mixed with names', function() {
		const result = excelToJson({ source, sheets: [{ index: 1, header: { rows: 1 }, columnToKey: { A: 'value' } }, 'Middle'] });
		assert.deepStrictEqual(result, { Alpha: [{ value: 'alpha' }], Middle: [{ A: 'h' }, { A: 'middle' }] });
	});

	it('should throw a clear error for an index that does not exist', function() {
		assert.throws(() => excelToJson({ source, sheets: [3] }), /no sheet at index 3/);
	});

	it('should follow the workbook order on "numberOfSheetsToGet"', function() {
		const result = excelToJson({ source, sheets: { numberOfSheetsToGet: 2 } });
		assert.deepStrictEqual(Object.keys(result), ['Zeta', 'Alpha']);
	});
});
