#!/usr/bin/env node
'use strict';

// Minimal, dependency-free CLI. It replaces `magicli`, whose transitive
// dependency (`cliss` > `yargs-parser`) was flagged by `npm audit` (#45, #68).
// Supported options are the same as before: --config, --sourceFile and --help.

const excelToJson = require('../lib/convert-excel-to-json');
const pkg = require('../package.json');

const HELP = `
Description:

  ${pkg.description}

Usage:

  $ convert-excel-to-json [options]

Options:

  --config JSON         A full config in a valid JSON format
  --sourceFile String   The sourceFile path (to be used without the 'config'
                        parameter)
  --help                Show this help
  --version             Show the version
`;

function parseArgs(argv) {
	const args = {};
	for (let i = 0; i < argv.length; i++) {
		const arg = argv[i];
		if (!arg.startsWith('--')) {
			continue;
		}

		const eqIndex = arg.indexOf('=');
		if (eqIndex > -1) {
			args[arg.slice(2, eqIndex)] = arg.slice(eqIndex + 1);
		} else if (argv[i + 1] !== undefined && !argv[i + 1].startsWith('--')) {
			args[arg.slice(2)] = argv[++i];
		} else {
			args[arg.slice(2)] = true;
		}
	}
	return args;
}

function run(argv) {
	const args = parseArgs(argv);

	if (args.version) {
		process.stdout.write(`${pkg.version}\n`);
		return 0;
	}

	if (args.help || (!args.config && !args.sourceFile)) {
		process.stdout.write(HELP);
		return args.help ? 0 : 1;
	}

	let config = {};
	if (args.config) {
		try {
			config = JSON.parse(args.config);
		} catch (err) {
			process.stderr.write(`Invalid JSON passed to --config: ${err.message}\n`);
			return 1;
		}
	}

	const result = excelToJson(config, args.sourceFile);
	process.stdout.write(`${JSON.stringify(result)}\n`);
	return 0;
}

try {
	process.exitCode = run(process.argv.slice(2));
} catch (err) {
	process.stderr.write(`${err.message}\n`);
	process.exitCode = 1;
}
