//
// FIXME :: THIS IS NOT FINISHED YET!
//

#!/usr/bin/env node

'use strict';

let excelToJSON = require('./lib/convert-excel-to-json');
let ArgumentParser = require('argparse').ArgumentParser;
let parser = new ArgumentParser({
	version: '1.0.0', // TODO :: GET CURRENT MODULE VERSION
	addHelp: true,
	description: 'Converts Excel to JSON'
});

parser.addArgument(
	['-s', '--src'], {
		help: 'Source Excel File',
		required: true
	}
);

parser.addArgument(
	['-d', '--dest'], {
		help: 'JSON file. It can be a full path (e.g. full/path/data.json)'
	}
);

let args = parser.parseArgs();
let outputJSONData = excelToJSON({
	sourceFile: args.src,
	outputFile: args.dest
});

if(!args.dest){
	console.log(outputJSONData);
}