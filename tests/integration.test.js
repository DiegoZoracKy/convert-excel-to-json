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
});

describe('options combined: second review of #96', function() {
	it('suffixDuplicateKeys + key callback + range: a column outside the range is never keyed', function() {
		const source = buildWorkbook({ s: [[5, 'Name'], [1, 'x']] }); // A1 is a number: toLowerCase would throw
		const result = excelToJson({ source, header: { rows: 1 }, range: 'B1:B2', suffixDuplicateKeys: true, columnToKey: { '*': header => header.toLowerCase() } });
		assert.deepStrictEqual(result.s, [{ name: 'x' }]);
	});

	it('key callback skipping a cell + transforms: no empty row is left', function() {
		const source = buildWorkbook({ s: [['keep', 'skip'], ['a', null], [null, 'b']] });
		const result = excelToJson({ source, header: { rows: 1 }, columnToKey: { A: { property: 'keep', transform: v => v.toUpperCase() }, B: () => null } });
		assert.deepStrictEqual(result.s, [{ keep: 'A' }]);
	});
});
