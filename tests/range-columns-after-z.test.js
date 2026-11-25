'use strict';

const assert = require('assert');
const excelToJson = require('../');
const buildWorkbook = require('./helpers/workbook');

// #42: ranges crossing the Z -> AA boundary returned undefined values
describe('range crossing column Z (#42)', function() {
	const header = [];
	const values = [];
	for (let i = 0; i < 30; i++) { // A..AD
		header.push(`h${i + 1}`);
		values.push(i + 1);
	}
	const source = buildWorkbook({ sheet1: [header, values, values] });

	it('should read columns after Z when the range starts before Z (A2:AA2)', function() {
		const result = excelToJson({
			source,
			header: { rows: 1 },
			sheets: [{ name: 'sheet1', range: 'A2:AA2' }],
			columnToKey: { '*': '{{columnHeader}}' }
		});
		assert.equal(result.sheet1.length, 1);
		assert.equal(Object.keys(result.sheet1[0]).length, 27);
		assert.equal(result.sheet1[0].h26, 26);
		assert.equal(result.sheet1[0].h27, 27);
		assert.equal(result.sheet1[0].h28, undefined);
	});

	it('should read a range that only spans columns around the boundary (Z2:AB3)', function() {
		const result = excelToJson({ source, range: 'Z2:AB3' });
		assert.deepStrictEqual(result.sheet1, [
			{ Z: 26, AA: 27, AB: 28 },
			{ Z: 26, AA: 27, AB: 28 }
		]);
	});

	it('should keep working for ranges that do not cross Z (B2:C2)', function() {
		const result = excelToJson({ source, range: 'B2:C2' });
		assert.deepStrictEqual(result.sheet1, [{ B: 2, C: 3 }]);
	});
});
