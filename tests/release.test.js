'use strict';

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const pkg = require('../package.json');

// Release prep (2.0.0): the Node.js 22 requirement is a breaking change
describe('release metadata', function() {
	it('should be a major version that requires Node.js >= 22', function() {
		assert.equal(pkg.engines.node, '>=22');
		assert.equal(pkg.version, '2.0.0');
		assert.equal(require('../package-lock.json').version, pkg.version);
	});

	it('should document the Node.js 22 requirement on the changelog', function() {
		const changelog = fs.readFileSync(path.resolve(__dirname, '..', 'CHANGELOG.md'), 'utf8');
		const unreleased = changelog.split(/^## \[/m).find(section => section.startsWith('Unreleased'));
		assert.ok(/2\.0\.0/.test(unreleased) && /Node\.js 22 or later is required/.test(unreleased));
	});
});
