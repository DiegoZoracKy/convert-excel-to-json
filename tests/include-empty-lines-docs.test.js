'use strict';

const assert = require('assert');
const excelToJson = require('../');
const buildWorkbook = require('./helpers/workbook');

// #61: documents (and pins) the behaviour of "includeEmptyLines"
describe('"includeEmptyLines" as documented on README (#61)', function() {
	const source = buildWorkbook({
		sheet1: {
			'!ref': 'A1:B4',
			A2: { t: 's', v: 'data of cell A2' }, B2: { t: 's', v: 'data of cell B2' },
			A4: { t: 's', v: 'data of cell A4' }, B4: { t: 's', v: 'data of cell B4' }
		}
	});

	it('should remove empty rows by default', function() {
		assert.equal(excelToJson({ source }).sheet1.length, 2);
	});

	it('should keep a position for each row, with empty slots that become null on JSON', function() {
		const rows = excelToJson({ source, includeEmptyLines: true }).sheet1;
		assert.equal(rows.length, 4);
		assert.strictEqual(rows[0], undefined);
		assert.strictEqual(rows[2], undefined);
		assert.deepStrictEqual(rows[1], { A: 'data of cell A2', B: 'data of cell B2' });
		assert.equal(JSON.stringify(rows), '[null,{"A":"data of cell A2","B":"data of cell B2"},null,{"A":"data of cell A4","B":"data of cell B4"}]');
	});
});
