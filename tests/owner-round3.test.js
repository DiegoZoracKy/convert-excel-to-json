'use strict';

// Owner review checks (round 3), adapted from the delivered round3-workbooks.test.js: only
// the repository path changed (repo = this repository, not process.cwd()). The fixture
// tests/bordered-header.xlsx is kept beside this file.
const assert = require('node:assert/strict');
const path = require('node:path');
const repo = path.resolve(__dirname, '..');
const convert = require(repo);
const build = require(path.join(repo, 'tests/helpers/workbook'));
const XLSX = build.XLSX;

describe('Checks for the updated release candidate', () => {
    it('reads the unmapped sibling cell through the documented third argument', () => {
        const source = build({ s: [['Ada', 'Lovelace']] });
        const columnToKey = {
            A: { property: 'name', transform: (value, row, cells) => {
                assert.equal(row.B, undefined);
                return value + ' ' + cells.B;
            } }
        };
        assert.deepEqual(convert({ source, columnToKey }).s, [{ name: 'Ada Lovelace' }]);
    });

    it('uses the merged anchor for a bordered header outside the selected data range', () => {
        const sourceFile = path.join(__dirname, 'bordered-header.xlsx');
        const options = {
            sourceFile, header: { rows: 1 }, fillMergedCells: true,
            suffixDuplicateKeys: true, columnToKey: { '*': '{{columnHeader}}' }
        };
        const expected = [{ Group: 10, Group_2: 20 }];
        assert.deepEqual(convert({ ...options, range: 'A1:B2' }).s, expected);
        assert.deepEqual(convert({ ...options, range: 'A2:B2' }).s, expected);
    });

    it('transposes the full extent of a merge beyond the last populated cell', () => {
        const sheet = XLSX.utils.aoa_to_sheet([['Group']]);
        sheet['!merges'] = [XLSX.utils.decode_range('A1:C2')];
        const source = build({ s: sheet });
        const opts = { source, fillMergedCells: true, transpose: true };
        const expected = Array.from({ length: 3 }, () => ({ A: 'Group', B: 'Group' }));
        assert.deepEqual(convert({ ...opts, range: 'A1:C2' }).s, expected);
        assert.deepEqual(convert(opts).s, expected);
    });

    it('excludes a hidden column before assigning duplicate header suffixes', () => {
        const sheet = XLSX.utils.aoa_to_sheet([['Name', 'Name'], ['hidden', 'visible']]);
        sheet['!cols'] = [{ hidden: true }, {}];
        const source = build({ s: sheet });
        const result = convert({
            source, header: { rows: 1 }, columnToKey: { '*': '{{columnHeader}}' },
            skipHiddenColumns: true, suffixDuplicateKeys: true
        });
        assert.deepEqual(result.s, [{ Name: 'visible' }]);
    });
});
