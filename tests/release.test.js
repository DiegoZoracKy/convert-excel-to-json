'use strict';

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const pkg = require('../package.json');

// The Node.js 22 note must be on the changelog section of this release: either still under
// "Unreleased" (announcing 2.0.0) or under the dated "2.0.0" heading once it is released.
function nodeRequirementSection(changelog, version) {
	const sections = changelog.split(/^## \[/m).slice(1);
	return sections.find(section => {
		const heading = section.split('\n')[0];
		const isRelease = heading.startsWith(`${version}]`);
		const isUnreleasedForVersion = heading.startsWith('Unreleased]') && section.includes(version);
		return (isRelease || isUnreleasedForVersion) && /Node\.js 22 or later is required/.test(section);
	});
}

// Release prep (2.0.0): the Node.js 22 requirement is a breaking change
describe('release metadata', function() {
	it('should be a major version that requires Node.js >= 22', function() {
		assert.equal(pkg.engines.node, '>=22');
		assert.equal(pkg.version, '2.0.0');
		assert.equal(require('../package-lock.json').version, pkg.version);
	});

	it('should document the Node.js 22 requirement on the changelog', function() {
		const changelog = fs.readFileSync(path.resolve(__dirname, '..', 'CHANGELOG.md'), 'utf8');
		assert.ok(nodeRequirementSection(changelog, pkg.version));
	});

	describe('changelog layouts', function() {
		const note = '- **Breaking:** Node.js 22 or later is required.\n';
		const older = '## [1.7.0] - 2020-06-08\n### Added\n- `sheetStubs` option.\n';

		it('should find the note under "Unreleased" (to be released as 2.0.0)', function() {
			const changelog = `# Changelog\n\n## [Unreleased]\nTo be released as **2.0.0**.\n\n### Changed\n${note}\n${older}`;
			assert.ok(nodeRequirementSection(changelog, '2.0.0'));
		});

		it('should find the note under the dated 2.0.0 release', function() {
			const changelog = `# Changelog\n\n## [Unreleased]\n\n## [2.0.0] - 2026-12-01\n### Changed\n${note}\n${older}`;
			assert.ok(nodeRequirementSection(changelog, '2.0.0'));
		});

		it('should not accept the note on an older release or on an Unreleased section of another version', function() {
			assert.ok(!nodeRequirementSection(`# Changelog\n\n## [Unreleased]\n\n## [1.7.0] - 2020-06-08\n${note}`, '2.0.0'));
			assert.ok(!nodeRequirementSection(`# Changelog\n\n## [Unreleased]\nTo be released as **3.0.0**.\n${note}`, '2.0.0'));
		});
	});
});
