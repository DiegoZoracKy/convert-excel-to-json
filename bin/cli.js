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
	}
});