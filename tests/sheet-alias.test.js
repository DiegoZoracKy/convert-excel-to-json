'use strict';

const assert = require('assert');
const excelToJson = require('../');
const buildWorkbook = require('./helpers/workbook');

// #58: different ranges from the same sheet
describe('sheet "alias" (#58)', function() {
	const source = buildWorkbook({
		Ladder: [
			['Season', 'Autumn'],
			['Year', 2021],
			[],
			['Ranking', 'Team', 'Total'],
			[1, 'Team1', 0],
			[2, 'Team2', 0]
		]
	});

	it('should read different ranges of the same sheet into different keys', function() {
		const result = excelToJson({
			source,
			sheets: [{
				name: 'Ladder', alias: 'info', range: 'A1:B2', columnToKey: { A: 'key', B: 'value' }
			}, {
				name: 'Ladder', alias: 'standings', range: 'A4:C6', header: { rows: 4 }, columnToKey: { '*': '{{columnHeader}}' }
			}]
		});
		assert.deepStrictEqual(result, {
			info: [{ key: 'Season', value: 'Autumn' }, { key: 'Year', value: 2021 }],
			standings: [{ Ranking: 1, Team: 'Team1', Total: 0 }, { Ranking: 2, Team: 'Team2', Total: 0 }]
		});
	});

	it('should keep using the sheet name when there is no alias', function() {
		const result = excelToJson({ source, sheets: [{ name: 'Ladder', range: 'A1:B1' }] });
		assert.deepStrictEqual(result, { Ladder: [{ A: 'Season', B: 'Autumn' }] });
	});
});

describe('sheet "alias" collisions (review of #91)', function() {
	const source = buildWorkbook({ sheet1: [['one']], sheet2: [['two']] });

	it('should throw when an alias matches the name of another selected sheet (alias last)', function() {
		assert.throws(() => excelToJson({ source, sheets: ['sheet1', { name: 'sheet2', alias: 'sheet1' }] }), /result key 'sheet1' is used by more than one sheet config/);
	});

	it('should throw when an alias matches the name of another selected sheet (alias first)', function() {
		assert.throws(() => excelToJson({ source, sheets: [{ name: 'sheet2', alias: 'sheet1' }, 'sheet1'] }), /result key 'sheet1'/);
	});

	it('should throw on two equal aliases', function() {
		assert.throws(() => excelToJson({ source, sheets: [{ name: 'sheet1', alias: 'x' }, { name: 'sheet2', alias: 'x' }] }), /result key 'x'/);
	});

	it('should allow an alias equal to its own sheet name', function() {
		assert.deepStrictEqual(excelToJson({ source, sheets: [{ name: 'sheet1', alias: 'sheet1' }] }), { sheet1: [{ A: 'one' }] });
	});

	it('should keep the previous behaviour for the same sheet listed twice without aliases', function() {
		assert.deepStrictEqual(excelToJson({ source, sheets: ['sheet1', 'sheet1'] }), { sheet1: [{ A: 'one' }] });
	});
});

describe('sheet names that are Object.prototype properties (second review of #91)', function() {
	it('should not report a collision for a sheet named "toString" selected twice without aliases', function() {
		const source = buildWorkbook({ toString: [['x']] });
		assert.deepStrictEqual(excelToJson({ source, sheets: ['toString', 'toString'] }), { toString: [{ A: 'x' }] });
	});

	it('should still report a real collision with an alias named "toString"', function() {
		const source = buildWorkbook({ toString: [['x']], other: [['y']] });
		assert.throws(() => excelToJson({ source, sheets: ['toString', { name: 'other', alias: 'toString' }] }), /result key 'toString'/);
	});
});

describe('a sheet named "__proto__" (second review of #91)', function() {
	it('should be a regular key of the result', function() {
		const source = buildWorkbook({ ['__proto__']: [['x']] });
		const result = excelToJson({ source, sheets: ['__proto__'] });
		assert.ok(Object.prototype.hasOwnProperty.call(result, '__proto__'));
		assert.equal(JSON.stringify(result), '{"__proto__":[{"A":"x"}]}');
	});
});
