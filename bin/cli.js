#!/usr/bin/env node

require('magicli')({
	commands: {
		'convert-excel-to-json': {
			options: [{
				name: 'config',
				description: 'A full config in a valid JSON format',
				type: 'JSON'

			},{
				name: 'sourceFile',
				description: `The sourceFile path (to be used without the 'config' parameter`,
				type: 'String'

			}]
		}
	},
	pipe: {
		after: JSON.stringify
		
		// Sample used for testing transform functions
		// before: (args, positionalArgs, argsAfterEndOfOptions) => {

		// 	if (args.config) {
		// 		var parsedConfigArg = JSON.parse(args.config);
		// 		parsedConfigArg.columnToKey.A = {
		// 			property: 'account',
		// 			transform: (rawCellValue, rowCells) => {
		// 				return `Account: ${String(rowCells['A'].v)} - BSB: ${String(rowCells['C'].v)}`;
		// 			}
		// 		}
		// 		args.config = parsedConfigArg;
		// 	}
		// 	return args;
		// }
	}
});