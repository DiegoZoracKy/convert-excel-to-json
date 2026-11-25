'use strict';

const assert = require('assert');
const child_process = require('child_process');
const path = require('path');
const excelToJson = require('../');
const buildWorkbook = require('./helpers/workbook');

// #33: "false date conversion" (1/1/2019 -> 2018-12-31T15:59:14.000Z on UTC+8)
// #52: date and time cells
function buildDateWorkbook() {
	const ws = {
		'!ref': 'A1:D1',
		A1: { t: 'n', v: 43466, z: 'm/d/yy' }, // 1/1/2019
		B1: { t: 'n', v: 43923.791666666664, z: 'yyyy-mm-dd hh:mm:ss' }, // 2020-04-02 19:00:00
		C1: { t: 'n', v: 43923.791666666664 }, // a plain number, not a date
		D1: { t: 'n', v: 44927.5, z: 'dd/mm/yyyy hh:mm' } // 01/01/2023 12:00
	};
	return buildWorkbook({ sheet1: ws });
}

describe('dates (#33, #52)', function() {
	const source = buildDateWorkbook();

	it('should return date cells as Dates with the exact local date/time shown on Excel', function() {
		const row = excelToJson({ source }).sheet1[0];
		assert.ok(row.A instanceof Date);
		assert.equal(row.A.getTime(), new Date(2019, 0, 1).getTime());
		assert.equal(row.B.getTime(), new Date(2020, 3, 2, 19, 0, 0).getTime());
		assert.equal(row.D.getTime(), new Date(2023, 0, 1, 12, 0, 0).getTime());
	});

	it('should keep plain numbers as numbers', function() {
		assert.strictEqual(excelToJson({ source }).sheet1[0].C, 43923.791666666664);
	});

	it('should return the wall-clock time in UTC with "utcDates"', function() {
		const row = excelToJson({ source, utcDates: true }).sheet1[0];
		assert.equal(row.A.toISOString(), '2019-01-01T00:00:00.000Z');
		assert.equal(row.B.toISOString(), '2020-04-02T19:00:00.000Z');
	});

	it('should handle workbooks using the 1904 date system', function() {
		const ws = { '!ref': 'A1', A1: { t: 'n', v: 40908, z: 'm/d/yy' } }; // 1/1/2016 in the 1904 system (42370 in the 1900 system)
		const src = buildWorkbook({ sheet1: ws }, {});
		const XLSX = buildWorkbook.XLSX;
		const wb = XLSX.read(src, { type: 'buffer', cellNF: true });
		wb.Workbook = wb.Workbook || {};
		wb.Workbook.WBProps = Object.assign({}, wb.Workbook.WBProps, { date1904: true });
		const src1904 = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
		const row = excelToJson({ source: src1904, utcDates: true }).sheet1[0];
		assert.equal(row.A.toISOString(), '2016-01-01T00:00:00.000Z');
	});

	['Asia/Kuala_Lumpur', 'America/Sao_Paulo', 'America/Los_Angeles', 'Europe/Berlin', 'Pacific/Apia', 'UTC'].forEach(timezone => {
		it(`should not drift on timezone ${timezone}`, function() {
			const script = `
				const excelToJson = require(${JSON.stringify(path.resolve(__dirname, '..'))});
				const source = Buffer.from(${JSON.stringify(source.toString('base64'))}, 'base64');
				const row = excelToJson({ source }).sheet1[0];
				const p = d => [d.getFullYear(), d.getMonth() + 1, d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds()].join('-');
				process.stdout.write(JSON.stringify([p(row.A), p(row.B)]));
			`;
			const result = child_process.spawnSync(process.execPath, ['-e', script], {
				env: Object.assign({}, process.env, { TZ: timezone }),
				encoding: 'utf8'
			});
			assert.deepEqual(JSON.parse(result.stdout), ['2019-1-1-0-0-0', '2020-4-2-19-0-0']);
		});
	});
});

// Review follow-up: explicitly typed date cells (t="d") without a date number format were
// returned as numbers, and milliseconds were truncated (12:00:00.500 -> 12:00:00.000).
// tests/test-data-typed-dates.xlsx, sheet "s":
//   A1: t="d" 2019-01-01T10:20:30.500Z, no number format
//   B1: number 43466.500005787 formatted as yyyy-mm-dd hh:mm:ss.000 (2019-01-01 12:00:00.500)
//   C1: t="d" 2020-02-29T23:59:59.000Z with a date number format
//   D1: plain number 43466.5
describe('explicitly typed dates and milliseconds (review of #79)', function() {
	const sourceFile = path.resolve(__dirname, 'test-data-typed-dates.xlsx');

	['Asia/Kuala_Lumpur', 'America/Sao_Paulo', 'America/Los_Angeles', 'Pacific/Apia', 'UTC'].forEach(timezone => {
		it(`should return typed dates as Dates and keep milliseconds on ${timezone}`, function() {
			const script = `
				const excelToJson = require(${JSON.stringify(path.resolve(__dirname, '..'))});
				const local = excelToJson({ sourceFile: ${JSON.stringify(sourceFile)} }).s[0];
				const utc = excelToJson({ sourceFile: ${JSON.stringify(sourceFile)}, utcDates: true }).s[0];
				const p = d => d instanceof Date ? [d.getFullYear(), d.getMonth() + 1, d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds(), d.getMilliseconds()].join('-') : d;
				const u = d => d instanceof Date ? d.toISOString() : d;
				process.stdout.write(JSON.stringify({ local: [p(local.A), p(local.B), p(local.C), local.D], utc: [u(utc.A), u(utc.B), u(utc.C), utc.D] }));
			`;
			const result = child_process.spawnSync(process.execPath, ['-e', script], {
				env: Object.assign({}, process.env, { TZ: timezone }),
				encoding: 'utf8'
			});
			const out = JSON.parse(result.stdout || 'null');
			assert.ok(out, result.stderr);
			assert.deepStrictEqual(out.local.slice(0, 2), ['2019-1-1-10-20-30-500', '2019-1-1-12-0-0-500']);
			assert.deepStrictEqual(out.utc.slice(0, 2), ['2019-01-01T10:20:30.500Z', '2019-01-01T12:00:00.500Z']);
			assert.strictEqual(out.local[3], 43466.5);
			// A typed date that also has a date format is exact too (SheetJS 0.20, #76)
			assert.equal(out.local[2], '2020-2-29-23-59-59-0');
			assert.equal(out.utc[2], '2020-02-29T23:59:59.000Z');
		});
	});

	it('should keep milliseconds of serial numbers (in memory workbook)', function() {
		const source = buildWorkbook({ sheet1: { '!ref': 'A1', A1: { t: 'n', v: 43466 + (12 * 3600 + 0.5) / 86400, z: 'yyyy-mm-dd hh:mm:ss.000' } } });
		assert.equal(excelToJson({ source, utcDates: true }).sheet1[0].A.toISOString(), '2019-01-01T12:00:00.500Z');
	});

	it('should round a time of 23:59:59.9999 to the next day instead of overflowing', function() {
		const source = buildWorkbook({ sheet1: { '!ref': 'A1', A1: { t: 'n', v: 43466.99999999999, z: 'yyyy-mm-dd hh:mm:ss' } } });
		assert.equal(excelToJson({ source, utcDates: true }).sheet1[0].A.toISOString(), '2019-01-02T00:00:00.000Z');
	});
});

// Runs `body` (which must write JSON to stdout) in a child process with the given TZ, so the
// host timezone is part of the test.
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

describe('dates: second review of #79', function() {
	const typedDates = path.resolve(__dirname, 'test-data-typed-dates.xlsx');

	it('should normalize a worksheet once, even when it is selected more than once (America/Sao_Paulo)', function() {
		const out = runOnTimezone('America/Sao_Paulo', `
			const once = excelToJson({ sourceFile: ${JSON.stringify(typedDates)}, sheets: ['s'] }).s[0];
			const twice = excelToJson({ sourceFile: ${JSON.stringify(typedDates)}, sheets: ['s', 's'] }).s[0];
			process.stdout.write(JSON.stringify([once.A, once.B, twice.A, twice.B]));
		`);
		assert.deepStrictEqual(out, ['2019-01-01T12:20:30.500Z', '2019-01-01T14:00:00.500Z', '2019-01-01T12:20:30.500Z', '2019-01-01T14:00:00.500Z']);
	});

	// tests/test-data-typed-dates-2011.xlsx, sheet "s": A1 and B1 are typed dates (t="d",
	// 2011-12-31T00:00:00Z) with date formats, C1 is the serial number 40908 (2011-12-31) formatted as a date
	['UTC', 'America/Sao_Paulo', 'Asia/Tokyo', 'America/Los_Angeles', 'Pacific/Kiritimati', 'Pacific/Apia'].forEach(timezone => {
		it(`should not depend on the host timezone with utcDates (${timezone})`, function() {
			const out = runOnTimezone(timezone, `
				const row = excelToJson({ sourceFile: ${JSON.stringify(path.resolve(__dirname, 'test-data-typed-dates-2011.xlsx'))}, utcDates: true }).s[0];
				process.stdout.write(JSON.stringify([row.A, row.B, row.C]));
			`);
			assert.deepStrictEqual(out, ['2011-12-31T00:00:00.000Z', '2011-12-31T00:00:00.000Z', '2011-12-31T00:00:00.000Z']);
		});
	});
});

describe('dates on sheets named like Object.prototype properties', function() {
	it('should convert a sheet named "__proto__" with dates', function() {
		const source = buildWorkbook({ ['__proto__']: { '!ref': 'A1', A1: { t: 'n', v: 43466, z: 'yyyy-mm-dd' } } });
		const result = excelToJson({ source, sheets: ['__proto__'], utcDates: true });
		assert.equal(JSON.stringify(result['__proto__']), '[{"A":"2019-01-01T00:00:00.000Z"}]');
	});
});

// Owner review check (round 3), assertions kept as delivered: Pacific/Apia skipped 2011-12-30
// when it moved across the international date line, and SheetJS' parse_date_code (local Dates)
// turned serial 40907 into 2011-12-31 there.
describe('serial dates across the international date line (owner review check)', function() {
	function inTimeZone(timezone, read) {
		const original = process.env.TZ;
		process.env.TZ = timezone;
		try {
			return read();
		} finally {
			if (original === undefined) {
				delete process.env.TZ;
			} else {
				process.env.TZ = original;
			}
		}
	}

	it('returns the same UTC serial date on either side of the international date line', function() {
		const source = buildWorkbook({ s: { '!ref': 'A1', A1: { t: 'n', v: 40907, z: 'yyyy-mm-dd' } } });
		const read = () => excelToJson({ source, utcDates: true }).s[0].A.toISOString();
		assert.equal(inTimeZone('UTC', read), '2011-12-30T00:00:00.000Z');
		assert.equal(inTimeZone('Pacific/Apia', read), '2011-12-30T00:00:00.000Z');
	});

	it('returns the same UTC serial date on Pacific/Apia (child process)', function() {
		const source = buildWorkbook({ s: { '!ref': 'A1:C1', A1: { t: 'n', v: 40907, z: 'yyyy-mm-dd' }, B1: { t: 'n', v: 40907.75, z: 'yyyy-mm-dd hh:mm' }, C1: { t: 'n', v: 40908, z: 'yyyy-mm-dd' } } });
		const out = runOnTimezone('Pacific/Apia', `
			const source = Buffer.from(${JSON.stringify(source.toString('base64'))}, 'base64');
			const row = excelToJson({ source, utcDates: true }).s[0];
			process.stdout.write(JSON.stringify([row.A, row.B, row.C]));
		`);
		assert.deepStrictEqual(out, ['2011-12-30T00:00:00.000Z', '2011-12-30T18:00:00.000Z', '2011-12-31T00:00:00.000Z']);
	});

	it('matches Excel on the 1900 leap year bug and on the 1904 date system edges', function() {
		const ws = { '!ref': 'A1:D1', A1: { t: 'n', v: 59, z: 'yyyy-mm-dd' }, B1: { t: 'n', v: 61, z: 'yyyy-mm-dd' }, C1: { t: 'n', v: 1, z: 'yyyy-mm-dd' }, D1: { t: 'n', v: 2958465, z: 'yyyy-mm-dd' } };
		const row = excelToJson({ source: buildWorkbook({ sheet1: ws }), utcDates: true }).sheet1[0];
		assert.deepStrictEqual([row.A, row.B, row.C, row.D].map(d => d.toISOString().slice(0, 10)), ['1900-02-28', '1900-03-01', '1900-01-01', '9999-12-31']);
	});
});
