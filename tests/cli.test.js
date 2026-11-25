'use strict';

const assert = require('assert');
const child_process = require('child_process');
const path = require('path');

const cliPath = path.resolve(`${__dirname}/../bin/cli.js`);
const sourceFile = path.resolve(`${__dirname}/test-data.xlsx`);
const pkg = require('../package.json');
const { parseArgs, NESTED_CONFIG } = require('../bin/cli');
const excelToJson = require('../');

const runCli = args => child_process.spawnSync(process.execPath, [cliPath].concat(args), { encoding: 'utf8' });
const parsed = args => {
	const result = parseArgs(args);
	const plain = Object.assign({}, result);
	if (result[NESTED_CONFIG]) {
		plain.config = result[NESTED_CONFIG];
		delete plain[NESTED_CONFIG];
	}
	return plain;
};

describe('CLI (dependency-free)', function() {
	it('should print the help with --help and exit 0', function() {
		const result = runCli(['--help']);
		assert.equal(result.status, 0);
		assert.ok(result.stdout.includes('--config'));
		assert.ok(result.stdout.includes('--sourceFile'));
	});

	it('should print the help and exit 1 when no option is given', function() {
		const result = runCli([]);
		assert.equal(result.status, 1);
		assert.ok(result.stdout.includes('Usage'));
	});

	it('should print the version with --version', function() {
		assert.equal(runCli(['--version']).stdout.trim(), pkg.version);
	});

	it('should fail with a clear message on invalid JSON passed to --config', function() {
		const result = runCli(['--config={not json']);
		assert.equal(result.status, 1);
		assert.ok(result.stderr.includes('Invalid JSON'));
	});

	describe('compatibility with the argument forms of the previous (magicli / yargs-parser) CLI', function() {
		const all = excelToJson({ sourceFile });
		const sheet2WithHeader = excelToJson({ sourceFile, header: { rows: 1 }, sheets: ['sheet2', 'sheet3'] });

		[
			['--sourceFile', sourceFile],
			[`--sourceFile=${sourceFile}`],
			['--source-file', sourceFile],
			[`--source-file=${sourceFile}`],
			[`--config.sourceFile=${sourceFile}`],
			['--config.sourceFile', sourceFile],
			['--config', JSON.stringify({ sourceFile })],
			[`--config=${JSON.stringify({ sourceFile })}`]
		].forEach(args => {
			it(`should convert with ${args.join(' ').replace(sourceFile, '<file>')}`, function() {
				const result = runCli(args);
				assert.equal(result.status, 0, result.stderr);
				assert.deepStrictEqual(JSON.parse(result.stdout), JSON.parse(JSON.stringify(all)));
			});
		});

		it('should build nested configs, numbers and arrays from dot notation (--config.header.rows=1 --config.sheets=a --config.sheets=b)', function() {
			const result = runCli([`--config.sourceFile=${sourceFile}`, '--config.header.rows=1', '--config.sheets=sheet2', '--config.sheets=sheet3']);
			assert.equal(result.status, 0, result.stderr);
			assert.deepStrictEqual(JSON.parse(result.stdout), JSON.parse(JSON.stringify(sheet2WithHeader)));
		});

		it('should combine --sourceFile with --config.* options', function() {
			const result = runCli([`--sourceFile=${sourceFile}`, '--config.header.rows=1', '--config.sheets=sheet2', '--config.sheets=sheet3']);
			assert.deepStrictEqual(JSON.parse(result.stdout), JSON.parse(JSON.stringify(sheet2WithHeader)));
		});

		it('should merge --config JSON with --config.* options', function() {
			const result = runCli([`--config=${JSON.stringify({ sourceFile, sheets: ['sheet2', 'sheet3'] })}`, '--config.header.rows=1']);
			assert.deepStrictEqual(JSON.parse(result.stdout), JSON.parse(JSON.stringify(sheet2WithHeader)));
		});

		// Expected values were generated with the previous parser:
		// yargs-parser@7 (via magicli > cliss) with { string: ['sourceFile'], json: ['config'] }
		[
			[['--sourceFile=a.xlsx'], { sourceFile: 'a.xlsx' }],
			[['--source-file', 'a.xlsx'], { sourceFile: 'a.xlsx' }],
			[['--sourceFile=a=b.xlsx'], { sourceFile: 'a=b.xlsx' }],
			[['--config.sourceFile=a.xlsx', '--config.header.rows=1', '--config.sheets=s2', '--config.sheets=s3'], { config: { sourceFile: 'a.xlsx', header: { rows: 1 }, sheets: ['s2', 's3'] } }],
			[['--config.sourceFile=a.xlsx', '--config.sheetStubs'], { config: { sourceFile: 'a.xlsx', sheetStubs: true } }],
			[['--config.sourceFile=a.xlsx', '--config.sheetStubs=true'], { config: { sourceFile: 'a.xlsx', sheetStubs: 'true' } }],
			[['--config.range=A1:B2', '--config.x=1.5', '--config.y=-2', '--config.z=007', '--config.w=1e3'], { config: { range: 'A1:B2', x: 1.5, y: -2, z: 7, w: 1000 } }],
			[['--config.columnToKey.A=id', '--config.columnToKey.*={{columnHeader}}'], { config: { columnToKey: { A: 'id', '*': '{{columnHeader}}' } } }],
			[['--config.sheets.numberOfSheetsToGet=2', '--source-file', 'x.xlsx'], { config: { sheets: { numberOfSheetsToGet: 2 } }, sourceFile: 'x.xlsx' }],
			[['--config={"sourceFile":"a.xlsx"}'], { config: '{"sourceFile":"a.xlsx"}' }]
		].forEach(([args, expected]) => {
			it(`should parse ${args.join(' ')} like the previous CLI`, function() {
				assert.deepStrictEqual(parsed(args), expected);
			});
		});

		it('should ignore __proto__ paths on dot notation', function() {
			parseArgs(['--config.__proto__.polluted=yes']);
			assert.strictEqual({}.polluted, undefined);
		});
	});
});
