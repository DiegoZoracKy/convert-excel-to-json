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
	// Pinned to America/Sao_Paulo (child process): with the original bug (dates not normalized on
	// sheets selected by index) this passes on UTC and fails there.
	it('sheets by index + dates: dates are normalized on sheets selected by index (America/Sao_Paulo)', function() {
		const ws = { A1: { t: 'n', v: 43466, z: 'yyyy-mm-dd' }, '!ref': 'A1' }; // 2019-01-01
		const source = buildWorkbook({ S: ws });
		const out = runOnTimezone('America/Sao_Paulo', `
			const source = Buffer.from(${JSON.stringify(source.toString('base64'))}, 'base64');
			const p = d => [d.getFullYear(), d.getMonth() + 1, d.getDate(), d.getHours(), d.getMinutes()].join('-');
			process.stdout.write(JSON.stringify([p(excelToJson({ source, sheets: [0] }).S[0].A), p(excelToJson({ source, sheets: [{ index: 0 }] }).S[0].A), p(excelToJson({ source, sheets: ['S'] }).S[0].A)]));
		`);
		assert.deepStrictEqual(out, ['2019-1-1-0-0', '2019-1-1-0-0', '2019-1-1-0-0']);
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

describe('options combined: second review of #96', function() {
	it('aliases + dates: a worksheet selected under two aliases is normalized once (America/Sao_Paulo)', function() {
		const sourceFile = path.resolve(__dirname, 'test-data-typed-dates.xlsx');
		const out = runOnTimezone('America/Sao_Paulo', `
			const once = excelToJson({ sourceFile: ${JSON.stringify(sourceFile)}, sheets: ['s'] });
			const twice = excelToJson({ sourceFile: ${JSON.stringify(sourceFile)}, sheets: [{ name: 's', alias: 'a' }, { name: 's', alias: 'b' }] });
			process.stdout.write(JSON.stringify([once.s[0].A, twice.a[0].A, twice.b[0].A]));
		`);
		assert.deepStrictEqual(out, ['2019-01-01T12:20:30.500Z', '2019-01-01T12:20:30.500Z', '2019-01-01T12:20:30.500Z']);
	});

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

	it('fillMergedCells + key callback: a merged header outside the range is read from its anchor', function() {
		const ws = XLSX.utils.aoa_to_sheet([['Group', null], [4, 5]]);
		ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }];
		const source = buildWorkbook({ s: ws });
		const result = excelToJson({ source, header: { rows: 1 }, range: 'B2:B2', fillMergedCells: true, columnToKey: { '*': header => header } });
		assert.deepStrictEqual(result.s, [{ Group: 5 }]);
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

	it('transpose without a range + sheets by index: leading empty coordinates are kept', function() {
		const ws = { '!ref': 'B2:B2', B2: { t: 's', v: 'v' } };
		const source = buildWorkbook({ s: ws });
		assert.equal(JSON.stringify(excelToJson({ source, sheets: [0], transpose: true, includeEmptyLines: true }).s), '[null,{"B":"v"}]');
	});
});
