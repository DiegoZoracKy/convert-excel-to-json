'use strict';

const assert = require('assert');
const child_process = require('child_process');
const path = require('path');

const cliPath = path.resolve(`${__dirname}/../bin/cli.js`);
const pkg = require('../package.json');

const runCli = args => child_process.spawnSync(process.execPath, [cliPath].concat(args), { encoding: 'utf8' });

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
		const result = runCli(['--version']);
		assert.equal(result.stdout.trim(), pkg.version);
	});

	it('should accept "--sourceFile <path>" (space separated)', function() {
		const result = runCli(['--sourceFile', path.resolve(`${__dirname}/test-data.xlsx`)]);
		assert.equal(result.status, 0);
		assert.deepEqual(Object.keys(JSON.parse(result.stdout)), ['sheet1', 'sheet2', 'sheet3']);
	});

	it('should fail with a clear message on invalid JSON passed to --config', function() {
		const result = runCli(['--config={not json']);
		assert.equal(result.status, 1);
		assert.ok(result.stderr.includes('Invalid JSON'));
	});
});
