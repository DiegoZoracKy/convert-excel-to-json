#!/usr/bin/env node

'use strict';

const excelToJson = require('../lib/convert-excel-to-json');

const HELP_TEXT = [
    'Usage: convert-excel-to-json [--config=<json>] [--sourceFile=<path>] [--help]',
    '',
    'Options:',
    '  --config      A full config in a valid JSON format',
    '  --sourceFile  The source file path (to be used without the config parameter)',
    '  --help        Show this message'
].join('\n');

function getArgValue(args, name) {
    const prefixed = `--${name}=`;
    const inline = args.find(arg => arg.startsWith(prefixed));
    if (inline) {
        return inline.slice(prefixed.length);
    }

    const index = args.indexOf(`--${name}`);
    if (index !== -1) {
        return args[index + 1];
    }

    return undefined;
}

function main() {
    const args = process.argv.slice(2);

    if (args.includes('--help') || args.includes('-h')) {
        console.log(HELP_TEXT);
        return;
    }

    const configArg = getArgValue(args, 'config');
    const sourceFileArg = getArgValue(args, 'sourceFile');

    let config;
    if (configArg) {
        try {
            config = JSON.parse(configArg);
        } catch (error) {
            console.error('Invalid JSON passed to --config.');
            process.exitCode = 1;
            return;
        }
    } else {
        config = { sourceFile: sourceFileArg };
    }

    try {
        const result = excelToJson(config);
        process.stdout.write(`${JSON.stringify(result)}\n`);
    } catch (error) {
        console.error(error.message || error);
        process.exitCode = 1;
    }
}

main();
