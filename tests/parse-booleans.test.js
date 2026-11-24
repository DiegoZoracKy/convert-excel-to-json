'use strict';

const assert = require('assert');
const excelToJson = require('../');
const buildWorkbook = require('./helpers/workbook');

// #55: TRUE/FALSE as Booleans
describe('"parseBooleans" config (#55)', function() {
	const source = buildWorkbook({
		sheet1: [
			['Name', 'Bool cell', 'Text true', 'Text false', 'Other text'],
			['Anthony', true, 'TRUE', ' False ', 'truely'],
			['Maria', false, 'true', 'FALSE', 'no']
		]
	});
	const config = { source, header: { rows: 1 }, columnToKey: { '*': '{{columnHeader}}' } };

	it('should keep returning the text by default (backward compatible)', function() {
		const result = excelToJson(config);
		assert.deepStrictEqual(result.sheet1[0], { Name: 'Anthony', 'Bool cell': 'TRUE', 'Text true': 'TRUE', 'Text false': 'False', 'Other text': 'truely' });
	});

	it('should return Booleans with "parseBooleans"', function() {
		const result = excelToJson(Object.assign({ parseBooleans: true }, config));
		assert.deepStrictEqual(result.sheet1, [
			{ Name: 'Anthony', 'Bool cell': true, 'Text true': true, 'Text false': false, 'Other text': 'truely' },
			{ Name: 'Maria', 'Bool cell': false, 'Text true': true, 'Text false': false, 'Other text': 'no' }
		]);
	});
});
