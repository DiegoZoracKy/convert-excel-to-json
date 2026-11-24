#!/usr/bin/env node
'use strict';

// Minimal, dependency-free CLI. It replaces `magicli`, whose transitive
// dependency (`cliss` > `yargs-parser`) was flagged by `npm audit` (#45, #68).
//
// It accepts the same argument forms the magicli/yargs-parser based CLI accepted:
//   --config='<json>'  |  --config '<json>'
//   --config.some.nested.option=value   (dot notation, builds the config object;
//                                        repeated options become arrays, numeric values become numbers)
//   --sourceFile=<path> | --sourceFile <path> | --source-file=<path> | --source-file <path>
//   --help, --version

const excelToJson = require('../lib/convert-excel-to-json');
const pkg = require('../package.json');

const HELP = `
Description:

  ${pkg.description}

Usage:

  $ convert-excel-to-json [options]

Options:

  --config JSON         A full config in a valid JSON format. Options can also be
                        set one by one with dot notation, e.g.
                        --config.sourceFile=file.xlsx --config.header.rows=1
  --sourceFile String   The sourceFile path (to be used without the 'config'
                        parameter). Also accepted as --source-file
  --help                Show this help
  --version             Show the version
`;

// Same rules as yargs-parser (used by the previous CLI) to turn values into numbers
const isNumber = value => /^0x[0-9a-f]+$/i.test(value) || /^[-+]?(?:\d+(?:\.\d*)?|\.\d+)(e[-+]?\d+)?$/.test(value);
const coerceValue = value => (typeof value === 'string' && isNumber(value) && Number.isSafeInteger(Math.floor(Number(value)))) ? Number(value) : value;

const NESTED_CONFIG = Symbol('nestedConfig');

const kebabToCamel = str => str.replace(/-+([^-])/g, (match, chr) => chr.toUpperCase());

// Repeated options become arrays (yargs-parser's duplicate-arguments-array)
function assign(target, key, value) {
	if (Object.prototype.hasOwnProperty.call(target, key)) {
		target[key] = [].concat(target[key], value);
	} else {
		target[key] = value;
	}
}

function setPath(target, path, value) {
	let current = target;
	path.slice(0, -1).forEach(key => {
		if (current[key] === undefined || current[key] === null || typeof current[key] !== 'object' || Array.isArray(current[key])) {
			current[key] = {};
		}
		current = current[key];
	});
	assign(current, path[path.length - 1], value);
}

function parseArgs(argv) {
	const args = {};
	const nestedConfig = {};
	let hasNestedConfig = false;

	for (let i = 0; i < argv.length; i++) {
		const arg = argv[i];
		if (arg === '--') {
			break;
		}
		if (!arg.startsWith('--') || arg.length === 2) {
			continue;
		}

		let name;
		let value;
		const eqIndex = arg.indexOf('=');
		if (eqIndex > -1) {
			name = arg.slice(2, eqIndex);
			value = arg.slice(eqIndex + 1);
		} else {
			name = arg.slice(2);
			value = (argv[i + 1] !== undefined && !argv[i + 1].startsWith('--')) ? argv[++i] : true;
		}

		const path = name.split('.').map(kebabToCamel);
		if (path[0] === 'config' && path.length > 1) {
			hasNestedConfig = true;
			// prototype pollution guard
			if (path.some(key => key === '__proto__' || key === 'constructor' || key === 'prototype')) {
				continue;
			}
			setPath(nestedConfig, path.slice(1), coerceValue(value));
		} else {
			// --config and --sourceFile are strings, as on the previous CLI
			assign(args, path.join('.'), value);
		}
	}

	if (hasNestedConfig) {
		args[NESTED_CONFIG] = nestedConfig;
	}
	return args;
}

function deepMerge(target, source) {
	Object.keys(source).forEach(key => {
		if (source[key] && typeof source[key] === 'object' && !Array.isArray(source[key]) &&
			target[key] && typeof target[key] === 'object' && !Array.isArray(target[key])) {
			deepMerge(target[key], source[key]);
		} else {
			target[key] = source[key];
		}
	});
	return target;
}

function run(argv) {
	const args = parseArgs(argv);

	if (args.version) {
		process.stdout.write(`${pkg.version}\n`);
		return 0;
	}

	if (args.help || (!args.config && !args.sourceFile && !args[NESTED_CONFIG])) {
		process.stdout.write(HELP);
		return args.help ? 0 : 1;
	}

	let config = {};
	if (typeof args.config === 'string') {
		try {
			config = JSON.parse(args.config);
		} catch (err) {
			process.stderr.write(`Invalid JSON passed to --config: ${err.message}\n`);
			return 1;
		}
	}
	if (args[NESTED_CONFIG]) {
		config = deepMerge(config, args[NESTED_CONFIG]);
	}

	const sourceFile = typeof args.sourceFile === 'string' ? args.sourceFile : undefined;
	const result = excelToJson(config, sourceFile);
	process.stdout.write(`${JSON.stringify(result)}\n`);
	return 0;
}

if (require.main === module) {
	try {
		process.exitCode = run(process.argv.slice(2));
	} catch (err) {
		process.stderr.write(`${err.message}\n`);
		process.exitCode = 1;
	}
}

module.exports = { parseArgs, run, NESTED_CONFIG };
