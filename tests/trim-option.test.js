'use strict';

const assert = require('assert');
const excelToJson = require('../');
const buildWorkbook = require('./helpers/workbook');

// #72: values with leading/trailing whitespace lose it on the result
describe('"trim" config (#72)', function() {
	const source = buildWorkbook({
		sheet1: [
			['  leading', 'trailing  ', '  both  ', 'none', 42]
		]
	});

	it('should trim text values by default (backward compatible)', function() {
		const result = excelToJson({ source });
		assert.deepStrictEqual(result.sheet1, [{ A: 'leading', B: 'trailing', C: 'both', D: 'none', E: 42 }]);
	});

	it('should trim text values when "trim" is true', function() {
		const result = excelToJson({ source, trim: true });
		assert.deepStrictEqual(result.sheet1, [{ A: 'leading', B: 'trailing', C: 'both', D: 'none', E: 42 }]);
	});

	it('should keep whitespace when "trim" is false', function() {
		const result = excelToJson({ source, trim: false });
		assert.deepStrictEqual(result.sheet1, [{ A: '  leading', B: 'trailing  ', C: '  both  ', D: 'none', E: 42 }]);
	});

	it('should keep whitespace on header values used as keys when "trim" is false', function() {
		const src = buildWorkbook({ sheet1: [[' key '], ['value']] });
		const result = excelToJson({ source: src, trim: false, header: { rows: 1 }, columnToKey: { '*': '{{columnHeader}}' } });
		assert.deepStrictEqual(result.sheet1, [{ ' key ': 'value' }]);
	});
});
