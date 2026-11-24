'use strict';

const assert = require('assert');
const excelToJson = require('../');
const buildWorkbook = require('./helpers/workbook');

// #65: percentage cells; #62: dates as strings
describe('"formattedValues" config (#62, #65)', function() {
	const source = buildWorkbook({
		sheet1: {
			'!ref': 'A1:E1',
			A1: { t: 'n', v: 0.225, z: '0.0%' },
			B1: { t: 'n', v: 43466, z: 'm/d/yy' },
			C1: { t: 'n', v: 1234.5, z: '#,##0.00' },
			D1: { t: 's', v: '  text  ' },
			E1: { t: 'n', v: 7 }
		}
	});

	it('should return raw values by default (backward compatible)', function() {
		const row = excelToJson({ source }).sheet1[0];
		assert.strictEqual(row.A, 0.225);
		assert.ok(row.B instanceof Date);
		assert.strictEqual(row.C, 1234.5);
		assert.strictEqual(row.D, 'text');
		assert.strictEqual(row.E, 7);
	});

	it('should return the values as displayed on Excel with "formattedValues"', function() {
		const row = excelToJson({ source, formattedValues: true }).sheet1[0];
		assert.deepStrictEqual(row, { A: '22.5%', B: '1/1/19', C: '1,234.50', D: 'text', E: '7' });
	});
});
