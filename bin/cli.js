#!/usr/bin/env node


const convertXlsxToJsonTransformNumberTypes = (rawCellValue, rowCells) => {
    return GetValidTransformedNumberVal(rawCellValue);
}

function GetValidTransformedNumberVal(value){

    console.log(`GetValidTransformedNumberVal = ${value}`);

    if(typeof value === 'undefined' || value === null || value === ''){
        return Number(0);
    }
    else{
        return Number(value);
    }
}


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
		after: JSON.stringify,
		
		// // Sample used for testing transform functions
		// before: (args, positionalArgs, argsAfterEndOfOptions) => {

		// 	if (args.config) {
		// 		var parsedConfigArg = JSON.parse(args.config);
		// 		parsedConfigArg.columnToKey.G = {
		// 			property: 'debit',
		// 			transform: convertXlsxToJsonTransformNumberTypes
		// 		}
		// 		parsedConfigArg.columnToKey.H = {
		// 			property: 'credit',
		// 			transform: convertXlsxToJsonTransformNumberTypes
		// 		}
		// 		args.config = parsedConfigArg;
		// 	}
		// 	return args;
		// }
	}
});