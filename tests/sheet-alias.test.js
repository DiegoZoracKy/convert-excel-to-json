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
