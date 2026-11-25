'use strict';

const assert = require('assert');
const path = require('path');
const excelToJson = require('../');
const buildWorkbook = require('./helpers/workbook');

// #64: rows with merged cells inside
describe('"fillMergedCells" config (#64)', function() {
	const XLSX = buildWorkbook.XLSX;
	const ws = XLSX.utils.aoa_to_sheet([
		['Category', 'Place', 'Term'],
		['FirstCat', 'Inside', 12],
		[null, null, 360],
		['SecondCat', 'Outside', 24]
	]);
	ws['!merges'] = [
		{ s: { r: 1, c: 0 }, e: { r: 2, c: 0 } }, // A2:A3
		{ s: { r: 1, c: 1 }, e: { r: 2, c: 1 } } // B2:B3
	];
	const source = buildWorkbook({ sheet1: ws });
	const config = { source, header: { rows: 1 }, columnToKey: { '*': '{{columnHeader}}' } };

	it('should keep the value only on the first cell by default (backward compatible)', function() {
		const result = excelToJson(config);
		assert.deepStrictEqual(result.sheet1[1], { Term: 360 });
	});

	it('should repeat the merged value on every cell of the range with "fillMergedCells"', function() {
		const result = excelToJson(Object.assign({ fillMergedCells: true }, config));
		assert.deepStrictEqual(result.sheet1, [
			{ Category: 'FirstCat', Place: 'Inside', Term: 12 },
			{ Category: 'FirstCat', Place: 'Inside', Term: 360 },
			{ Category: 'SecondCat', Place: 'Outside', Term: 24 }
		]);
	});

	it('should fill horizontally merged header cells too', function() {
		const hws = XLSX.utils.aoa_to_sheet([['Group', null], ['a', 'b']]);
		hws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }];
		const result = excelToJson({ source: buildWorkbook({ sheet1: hws }), fillMergedCells: true });
		assert.deepStrictEqual(result.sheet1[0], { A: 'Group', B: 'Group' });
	});

	it('should be configurable per sheet', function() {
		const result = excelToJson(Object.assign({ sheets: [{ name: 'sheet1', fillMergedCells: true }] }, config));
		assert.deepStrictEqual(result.sheet1[1], { Category: 'FirstCat', Place: 'Inside', Term: 360 });
	});
});

describe('"fillMergedCells" limited to the range (review of #92)', function() {
	const XLSX = buildWorkbook.XLSX;

	it('should only fill merged cells inside the range, including merges starting before it', function() {
		const ws = XLSX.utils.aoa_to_sheet([
			['Group', 'v'],
			[null, 1],
			[null, 2],
			['Other', 3]
		]);
		ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 2, c: 0 } }]; // A1:A3, value on A1 (outside the range)
		const result = excelToJson({ source: buildWorkbook({ sheet1: ws }), range: 'A2:B3', fillMergedCells: true });
		assert.deepStrictEqual(result.sheet1, [{ A: 'Group', B: 1 }, { A: 'Group', B: 2 }]);
	});

	it('should not expand merged areas that are outside of the range', function() {
		const ws = XLSX.utils.aoa_to_sheet([['a', 'b'], [1, 2]]);
		ws.Z1 = { t: 's', v: 'huge' };
		ws['!ref'] = 'A1:Z2';
		ws['!merges'] = [{ s: { r: 0, c: 25 }, e: { r: 1048575, c: 25 } }]; // Z1:Z1048576
		const source = buildWorkbook({ sheet1: ws });
		// count the cells created by the expansion
		const encodeCell = XLSX.utils.encode_cell;
		let encodedCells = 0;
		XLSX.utils.encode_cell = function() {
			encodedCells++;
			return encodeCell.apply(this, arguments);
		};
		let result;
		try {
			result = excelToJson({ source, range: 'A1:B2', fillMergedCells: true });
		} finally {
			XLSX.utils.encode_cell = encodeCell;
		}
		assert.deepStrictEqual(result.sheet1, [{ A: 'a', B: 'b' }, { A: 1, B: 2 }]);
		assert.ok(encodedCells < 100, `the merged column outside of the range must not be expanded (${encodedCells} cells)`);
	});
});

describe('"fillMergedCells" with merged header cells outside the range (second review of #92)', function() {
	const XLSX = buildWorkbook.XLSX;

	it('should resolve a referenced header cell through the anchor of its merged range', function() {
		const ws = XLSX.utils.aoa_to_sheet([['Group', null], [4, 5]]);
		ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }]; // A1:B1 = 'Group'
		const source = buildWorkbook({ sheet1: ws });
		const result = excelToJson({ source, header: { rows: 1 }, range: 'B2:B2', fillMergedCells: true, columnToKey: { '*': '{{columnHeader}}' } });
		assert.deepStrictEqual(result.sheet1, [{ Group: 5 }]);
	});

	it('should resolve cell variables through the anchor too', function() {
		const ws = XLSX.utils.aoa_to_sheet([['Group', null], [4, 5]]);
		ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }];
		const source = buildWorkbook({ sheet1: ws });
		const result = excelToJson({ source, range: 'B2:B2', fillMergedCells: true, columnToKey: { B: '{{B1}}' } });
		assert.deepStrictEqual(result.sheet1, [{ Group: 5 }]);
	});
});

// Owner review checks (round 3), assertions kept as delivered.
// tests/bordered-header.xlsx, sheet "s": A1:B1 merged with 'Group' and a border, so B1 is a
// styled cell without a value (a stub, as the workbook is read with sheetStubs); A2 = 10, B2 = 20.
describe('"fillMergedCells": bordered merged headers and merges beyond !ref (owner review checks)', function() {
	const XLSX = buildWorkbook.XLSX;

	it('uses the merged anchor for a bordered header outside the selected data range', function() {
		const sourceFile = path.join(__dirname, 'bordered-header.xlsx');
		const options = { sourceFile, header: { rows: 1 }, fillMergedCells: true, columnToKey: { '*': '{{columnHeader}}' } };
		// (without suffixDuplicateKeys, #83, both columns are named 'Group' and B wins)
		const expected = [{ Group: 20 }];
		assert.deepStrictEqual(excelToJson(Object.assign({}, options, { range: 'A1:B2' })).s, expected);
		assert.deepStrictEqual(excelToJson(Object.assign({}, options, { range: 'A2:B2' })).s, expected);
	});

	it('fills the full extent of a merge beyond the last populated cell', function() {
		const sheet = XLSX.utils.aoa_to_sheet([['Group']]);
		sheet['!merges'] = [XLSX.utils.decode_range('A1:C2')];
		const source = buildWorkbook({ s: sheet });
		const expected = Array.from({ length: 2 }, () => ({ A: 'Group', B: 'Group', C: 'Group' }));
		assert.deepStrictEqual(excelToJson({ source, fillMergedCells: true, range: 'A1:C2' }).s, expected);
		assert.deepStrictEqual(excelToJson({ source, fillMergedCells: true }).s, expected);
		// (on its own this already passed: the transposed variant, #90 + #92, is on #96)
		assert.deepStrictEqual(excelToJson({ source, fillMergedCells: true, sheetStubs: true }).s, expected);
	});
});

// Owner review (round 4): an xlsx file can keep a value on a cell covered by a merge. Excel only
// shows the anchor, so the covered header cell must always resolve through the anchor.
describe('"fillMergedCells": covered header cells holding a stale value (owner review round 4)', function() {
	const XLSX = buildWorkbook.XLSX;
	const buildSource = b1 => {
		const ws = XLSX.utils.aoa_to_sheet([['Group', b1], [null, 20]]);
		ws['!merges'] = [XLSX.utils.decode_range('A1:B1')];
		return buildWorkbook({ s: ws });
	};
	const options = { header: { rows: 1 }, fillMergedCells: true, columnToKey: { B: '{{columnHeader}}' } };

	it('should use the anchor, not a stale "Old" kept on B1, for ranges A1:B2 and A2:B2', function() {
		const source = buildSource('Old');
		assert.deepStrictEqual(excelToJson(Object.assign({ source, range: 'A1:B2' }, options)).s, [{ Group: 20 }]);
		assert.deepStrictEqual(excelToJson(Object.assign({ source, range: 'A2:B2' }, options)).s, [{ Group: 20 }]);
	});

	it('should use the anchor when B1 keeps an empty string (instead of dropping the row)', function() {
		const source = buildSource('');
		assert.equal(XLSX.read(source).Sheets.s.B1.v, '');
		assert.deepStrictEqual(excelToJson(Object.assign({ source, range: 'A1:B2' }, options)).s, [{ Group: 20 }]);
		assert.deepStrictEqual(excelToJson(Object.assign({ source, range: 'A2:B2' }, options)).s, [{ Group: 20 }]);
	});
});
