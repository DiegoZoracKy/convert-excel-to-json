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

	['Asia/Kuala_Lumpur', 'America/Sao_Paulo', 'America/Los_Angeles', 'Europe/Berlin', 'UTC'].forEach(timezone => {
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
	const XLSX = require('xlsx');
	const [major, minor] = XLSX.version.split('.').map(Number);
	const sheetJsUTC = major > 0 || minor >= 20;
	const sourceFile = path.resolve(__dirname, 'test-data-typed-dates.xlsx');

	['Asia/Kuala_Lumpur', 'America/Sao_Paulo', 'America/Los_Angeles', 'UTC'].forEach(timezone => {
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
			// A typed date that also has a date format is exact on SheetJS >= 0.20 (#76). On 0.12 its
			// serial number is computed by SheetJS with the historical (LMT) offset, so only check the day.
			if (sheetJsUTC) {
				assert.equal(out.local[2], '2020-2-29-23-59-59-0');
				assert.equal(out.utc[2], '2020-02-29T23:59:59.000Z');
			} else {
				assert.ok(/^2020-(2-29|3-1)-/.test(out.local[2]), out.local[2]);
			}
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
