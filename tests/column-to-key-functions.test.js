'use strict';

const assert = require('assert');
const excelToJson = require('../');
const buildWorkbook = require('./helpers/workbook');

// #70: "columnData.match is not a function" when a function is used on columnToKey
describe('columnToKey with functions (#70)', function() {
	const source = buildWorkbook({
		sheet1: [
			['First Name', 'Last Name', 'E-mail Address'],
			['Judith', 'Bennett', 'jb@example.com'],
			['Denise', 'Ramirez', 'dr@example.com']
		]
	});
	const camelCase = str => str.toLowerCase().replace(/[^a-z0-9]+(.)/g, (m, chr) => chr.toUpperCase());

	it('should accept a function on "*" and call it with the column header value', function() {
		const result = excelToJson({ source, header: { rows: 1 }, columnToKey: { '*': camelCase } });
		assert.deepStrictEqual(result.sheet1, [
			{ firstName: 'Judith', lastName: 'Bennett', eMailAddress: 'jb@example.com' },
			{ firstName: 'Denise', lastName: 'Ramirez', eMailAddress: 'dr@example.com' }
		]);
	});

	it('should pass the column letter as the second argument', function() {
		const result = excelToJson({ source, header: { rows: 1 }, columnToKey: { '*': (header, column) => `${column}_${header}` } });
		assert.deepStrictEqual(Object.keys(result.sheet1[0]), ['A_First Name', 'B_Last Name', 'C_E-mail Address']);
	});

	it('should accept functions mixed with strings on specific columns', function() {
		const result = excelToJson({ source, header: { rows: 1 }, columnToKey: { A: 'id', C: header => header.toUpperCase() } });
		assert.deepStrictEqual(result.sheet1[0], { id: 'Judith', 'E-MAIL ADDRESS': 'jb@example.com' });
	});

	it('should skip the column when the function returns undefined', function() {
		const result = excelToJson({ source, header: { rows: 1 }, columnToKey: { '*': (header, column) => column === 'B' ? undefined : column } });
		assert.deepStrictEqual(result.sheet1[0], { A: 'Judith', C: 'jb@example.com' });
	});

	it('should use the first row as header when header.rows is not defined', function() {
		const result = excelToJson({ source, columnToKey: { A: header => `key:${header}` } });
		assert.deepStrictEqual(result.sheet1[1], { 'key:First Name': 'Judith' });
	});

	it('should throw a clear error for values that are not a String or a Function', function() {
		assert.throws(() => excelToJson({ source, columnToKey: { A: 42 } }), /must be a String or a Function/);
	});
});

describe('columnToKey functions (review of #78)', function() {
	it('should keep keys returned by a function as they are, without {{variable}} substitution', function() {
		const source = buildWorkbook({ sheet1: [['Cost {{USD}}', 'Cost {{EUR}}'], [10, 9]] });
		const result = excelToJson({ source, header: { rows: 1 }, columnToKey: { '*': header => header } });
		assert.deepStrictEqual(result.sheet1, [{ 'Cost {{USD}}': 10, 'Cost {{EUR}}': 9 }]);
	});

	it('should not let a conversion started inside a callback change the outer conversion config', function() {
		const source = buildWorkbook({ sheet1: [['First Name', 'Notes'], ['Ann', '   '], [null, null], ['Bob', 'x']] });
		const lookupSource = buildWorkbook({ sheet1: [['First Name', 'first_name']] });

		const result = excelToJson({
			source,
			header: { rows: 1 },
			includeEmptyLines: true,
			columnToKey: {
				'*': header => {
					// e.g. a lookup table read from another spreadsheet, with a different config
					const lookup = excelToJson({ source: lookupSource, columnToKey: { A: 'from', B: 'to' } }).sheet1;
					const match = lookup.find(row => row.from === header);
					return match ? match.to : header;
				}
			}
		});

		// includeEmptyLines of the outer call must still apply: a slot per row
		// (row 1 is the skipped header, row 3 is empty)
		assert.equal(result.sheet1.length, 4);
		assert.deepStrictEqual(result.sheet1[1], { first_name: 'Ann', Notes: '   ' });
		assert.strictEqual(result.sheet1[2], undefined);
		assert.deepStrictEqual(result.sheet1[3], { first_name: 'Bob', Notes: 'x' });
	});
});
