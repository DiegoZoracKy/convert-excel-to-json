'use strict';

const assert = require('assert');
const excelToJson = require('../');
const buildWorkbook = require('./helpers/workbook');

// #50: rows as headers
describe('"transpose" config (#50)', function() {
	const source = buildWorkbook({
		sheet1: [
			['Type', 'TypeData', 'OtherType'],
			['Date', 'DateData', 'OtherDate']
		],
		sheet2: [
			['ignored', 'ignored', 'ignored'],
			['ignored', 'Season', 'Autumn'],
			['ignored', 'Year', 2021],
			['ignored', 'Round', 1]
		]
	});

	it('should not transpose by default', function() {
		const result = excelToJson({ source, sheets: ['sheet1'] });
		assert.deepStrictEqual(result.sheet1[0], { A: 'Type', B: 'TypeData', C: 'OtherType' });
	});

	it('should swap rows and columns', function() {
		const result = excelToJson({ source, sheets: ['sheet1'], transpose: true });
		assert.deepStrictEqual(result.sheet1, [
			{ A: 'Type', B: 'Date' },
			{ A: 'TypeData', B: 'DateData' },
			{ A: 'OtherType', B: 'OtherDate' }
		]);
	});

	it('should use the first column as header with header.rows and {{columnHeader}}', function() {
		const result = excelToJson({ source, sheets: ['sheet1'], transpose: true, header: { rows: 1 }, columnToKey: { '*': '{{columnHeader}}' } });
		assert.deepStrictEqual(result.sheet1, [
			{ Type: 'TypeData', Date: 'DateData' },
			{ Type: 'OtherType', Date: 'OtherDate' }
		]);
	});

	it('should apply the range on the original sheet and start the transposed data at A1', function() {
		const result = excelToJson({
			source,
			sheets: [{ name: 'sheet2', transpose: true, range: 'B2:C4', header: { rows: 1 }, columnToKey: { '*': '{{columnHeader}}' } }]
		});
		assert.deepStrictEqual(result.sheet2, [{ Season: 'Autumn', Year: 2021, Round: 1 }]);
	});
});

// Review of #89 + #90: transpose discarded the visibility metadata, so with
// transpose + skipHiddenColumns a hidden column became a visible output row.
describe('"transpose" with "skipHiddenColumns" / "skipHiddenRows" (#50, #46)', function() {
	const XLSX = buildWorkbook.XLSX;
	const ws = XLSX.utils.aoa_to_sheet([
		['Type', 'TypeData', 'SecretData', 'OtherType'],
		['Date', 'DateData', 'SecretDate', 'OtherDate'],
		['Internal', 'i1', 'i2', 'i3']
	]);
	ws['!cols'] = [{}, {}, { hidden: true }, {}]; // column C hidden
	ws['!rows'] = [{}, {}, { hidden: true }]; // row 3 hidden
	const source = buildWorkbook({ sheet1: ws });
	const config = { source, transpose: true, header: { rows: 1 }, columnToKey: { '*': '{{columnHeader}}' } };

	it('should skip the hidden column of the file (an output row once transposed)', function() {
		const result = excelToJson(Object.assign({ skipHiddenColumns: true }, config));
		assert.deepStrictEqual(result.sheet1, [
			{ Type: 'TypeData', Date: 'DateData', Internal: 'i1' },
			{ Type: 'OtherType', Date: 'OtherDate', Internal: 'i3' }
		]);
	});

	it('should skip the hidden row of the file (an output key once transposed)', function() {
		const result = excelToJson(Object.assign({ skipHiddenRows: true }, config));
		assert.deepStrictEqual(result.sheet1, [
			{ Type: 'TypeData', Date: 'DateData' },
			{ Type: 'SecretData', Date: 'SecretDate' },
			{ Type: 'OtherType', Date: 'OtherDate' }
		]);
	});

	it('should skip both, also with a range', function() {
		const result = excelToJson(Object.assign({ skipHiddenColumns: true, skipHiddenRows: true, range: 'A1:D3' }, config));
		assert.deepStrictEqual(result.sheet1, [
			{ Type: 'TypeData', Date: 'DateData' },
			{ Type: 'OtherType', Date: 'OtherDate' }
		]);
	});

	it('should keep everything without the skip options', function() {
		assert.equal(excelToJson(config).sheet1.length, 3);
	});
});

describe('"transpose" coordinates without a range (second review of #90)', function() {
	it('should keep leading empty rows and columns (B3 -> C2)', function() {
		const ws = { '!ref': 'B3:C4', B3: { t: 's', v: 'b3' }, C3: { t: 's', v: 'c3' }, B4: { t: 's', v: 'b4' } };
		const source = buildWorkbook({ s: ws });
		const result = excelToJson({ source, transpose: true, includeEmptyLines: true });
		assert.equal(JSON.stringify(result.s), '[null,{"C":"b3","D":"b4"},{"C":"c3"}]');
	});
});
