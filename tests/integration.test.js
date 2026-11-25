'use strict';

// Options from different PRs used together (release-candidate verification).
const assert = require('assert');
const excelToJson = require('../');
const buildWorkbook = require('./helpers/workbook');
const XLSX = buildWorkbook.XLSX;

describe('options combined (integration)', function() {
	it('sheets by index + dates: dates are normalized on sheets selected by index', function() {
		const ws = { A1: { t: 'n', v: 43466, z: 'yyyy-mm-dd' }, '!ref': 'A1' }; // 2019-01-01
		const source = buildWorkbook({ S: ws });
		const byIndex = excelToJson({ source, sheets: [0] });
		const byName = excelToJson({ source, sheets: ['S'] });
		assert.deepStrictEqual(byIndex, byName);
		assert.ok(byIndex.S[0].A instanceof Date);
		assert.strictEqual(byIndex.S[0].A.getFullYear(), 2019);
		assert.strictEqual(byIndex.S[0].A.getHours(), 0);
	});

	it('sheets by index + alias', function() {
		const source = buildWorkbook({ first: [['a']], second: [['b']] });
		assert.deepStrictEqual(excelToJson({ source, sheets: [{ index: 1, alias: 'x' }, 'first'] }), { x: [{ A: 'b' }], first: [{ A: 'a' }] });
		assert.throws(() => excelToJson({ source, sheets: ['first', { index: 1, alias: 'first' }] }), /result key 'first'/);
	});

	it('transform + suffixDuplicateKeys: each transform runs on its own column', function() {
		const source = buildWorkbook({ s: [['n', 'n'], [1, 2]] });
		const result = excelToJson({
			source,
			header: { rows: 1 },
			suffixDuplicateKeys: true,
			columnToKey: { A: { property: '{{A1}}', transform: v => v * 10 }, B: '{{B1}}' }
		});
		assert.deepStrictEqual(result.s, [{ n: 10, n_2: 2 }]);
	});

	it('columnToKey function + suffixDuplicateKeys', function() {
		const source = buildWorkbook({ s: [['Name', 'name'], ['a', 'b']] });
		const result = excelToJson({ source, header: { rows: 1 }, suffixDuplicateKeys: true, columnToKey: { '*': header => header.toLowerCase() } });
		assert.deepStrictEqual(result.s, [{ name: 'a', name_2: 'b' }]);
	});

	it('whitespaceToNull + sheetStubs + skipHiddenColumns: hidden columns are not filled back', function() {
		const ws = XLSX.utils.aoa_to_sheet([['a', 'hidden', ' '], ['b', 'x', null]]);
		ws['!cols'] = [{}, { hidden: true }, {}];
		const source = buildWorkbook({ s: ws });
		const result = excelToJson({ source, sheetStubs: true, whitespaceToNull: true, skipHiddenColumns: true });
		assert.deepStrictEqual(result.s, [{ A: 'a', C: null }, { A: 'b', C: null }]);
	});

	it('parseBooleans + transform: the transform receives the Boolean', function() {
		const source = buildWorkbook({ s: [['TRUE']] });
		const result = excelToJson({ source, parseBooleans: true, columnToKey: { A: { property: 'ok', transform: v => typeof v } } });
		assert.deepStrictEqual(result.s, [{ ok: 'boolean' }]);
	});

	it('fillMergedCells + transpose: merges are filled on the original layout', function() {
		const ws = XLSX.utils.aoa_to_sheet([['key', 'v1', 'v2'], ['group', null, null]]);
		ws['!merges'] = [{ s: { r: 1, c: 1 }, e: { r: 1, c: 2 } }]; // B2:C2, value on B2
		ws.B2 = { t: 's', v: 'merged' };
		const source = buildWorkbook({ s: ws });
		const result = excelToJson({ source, fillMergedCells: true, transpose: true, header: { rows: 1 }, columnToKey: { '*': '{{columnHeader}}' } });
		assert.deepStrictEqual(result.s, [{ key: 'v1', group: 'merged' }, { key: 'v2', group: 'merged' }]);
	});

	it('ArrayBuffer source + skipHiddenRows (the workbook is read with cellStyles)', function() {
		const ws = XLSX.utils.aoa_to_sheet([['a'], ['hidden'], ['c']]);
		ws['!rows'] = [{}, { hidden: true }, {}];
		const buffer = buildWorkbook({ s: ws });
		const arrayBuffer = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
		assert.deepStrictEqual(excelToJson({ source: arrayBuffer, skipHiddenRows: true }).s, [{ A: 'a' }, { A: 'c' }]);
	});

	it('formattedValues + trim: false', function() {
		const ws = XLSX.utils.aoa_to_sheet([[0.5, '  text  ']]);
		ws.A1.z = '0%';
		const source = buildWorkbook({ s: ws });
		assert.deepStrictEqual(excelToJson({ source, formattedValues: true, trim: false }).s, [{ A: '50%', B: '  text  ' }]);
	});
});
