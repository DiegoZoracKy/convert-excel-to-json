'use strict';

// Options from different PRs used together (release-candidate verification).
const assert = require('assert');
const excelToJson = require('../');
const buildWorkbook = require('./helpers/workbook');
const XLSX = buildWorkbook.XLSX;
const child_process = require('child_process');
const path = require('path');

// Runs `body` (which writes JSON to stdout) in a child process with the given TZ
function runOnTimezone(timezone, body) {
	const script = `const excelToJson = require(${JSON.stringify(path.resolve(__dirname, '..'))});\n${body}`;
	const result = child_process.spawnSync(process.execPath, ['-e', script], {
		env: Object.assign({}, process.env, { TZ: timezone }),
		encoding: 'utf8'
	});
	const out = JSON.parse(result.stdout || 'null');
	assert.ok(out, result.stderr);
	return out;
}

describe('options combined (integration)', function() {
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

	it('formattedValues + trim: false', function() {
		const ws = XLSX.utils.aoa_to_sheet([[0.5, '  text  ']]);
		ws.A1.z = '0%';
		const source = buildWorkbook({ s: ws });
		assert.deepStrictEqual(excelToJson({ source, formattedValues: true, trim: false }).s, [{ A: '50%', B: '  text  ' }]);
	});
});

describe('options combined: second review of #96', function() {
	it('suffixDuplicateKeys + key callback + range: a column outside the range is never keyed', function() {
		const source = buildWorkbook({ s: [[5, 'Name'], [1, 'x']] }); // A1 is a number: toLowerCase would throw
		const result = excelToJson({ source, header: { rows: 1 }, range: 'B1:B2', suffixDuplicateKeys: true, columnToKey: { '*': header => header.toLowerCase() } });
		assert.deepStrictEqual(result.s, [{ name: 'x' }]);
	});

	it('suffixDuplicateKeys + whitespaceToNull + sheetStubs: a wholly absent column is still emitted as null', function() {
		const ws = { '!ref': 'A1:C1', A1: { t: 's', v: 'x' }, C1: { t: 's', v: 'y' } };
		const source = buildWorkbook({ s: ws });
		const result = excelToJson({ source, suffixDuplicateKeys: true, sheetStubs: true, whitespaceToNull: true });
		assert.equal(JSON.stringify(result.s), '[{"A":"x","B":null,"C":"y"}]');
	});

	it('suffixDuplicateKeys + whitespaceToNull + sheetStubs + "toString" header', function() {
		const source = buildWorkbook({ s: [['toString', 'toString'], [null, 1]] });
		const result = excelToJson({ source, header: { rows: 1 }, suffixDuplicateKeys: true, sheetStubs: true, whitespaceToNull: true, columnToKey: { '*': '{{columnHeader}}' } });
		assert.equal(JSON.stringify(result.s), '[{"toString":null,"toString_2":1}]');
	});

	it('key callback skipping a cell + transforms: no empty row is left', function() {
		const source = buildWorkbook({ s: [['keep', 'skip'], ['a', null], [null, 'b']] });
		const result = excelToJson({ source, header: { rows: 1 }, columnToKey: { A: { property: 'keep', transform: v => v.toUpperCase() }, B: () => null } });
		assert.deepStrictEqual(result.s, [{ keep: 'A' }]);
	});

	it('transform cells + parseBooleans: cells are the converted values of the whole row', function() {
		const source = buildWorkbook({ s: [[1, 'TRUE', 'note']] });
		const result = excelToJson({ source, parseBooleans: true, columnToKey: { A: { property: 'n', transform: (v, row, cells) => [v, cells.B, cells.C] } } });
		assert.deepStrictEqual(result.s, [{ n: [1, true, 'note'] }]);
	});
});
