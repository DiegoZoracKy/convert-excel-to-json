'use strict';

// Owner review checks (second review, 11 checks), adapted from the delivered
// convert-excel-to-json-rereview.test.js: only the paths changed (repo = this repository, not
// process.cwd()), and the transform check uses the documented 3-arg contract,
// transform(value, row, cells), where unmapped columns are on `cells` (README).
const assert = require('node:assert/strict');
const path = require('node:path');
const repo = path.resolve(__dirname, '..');
const convert = require(repo);
const workbook = require(path.join(repo, 'tests/helpers/workbook'));
const XLSX = workbook.XLSX;

function inTimeZone(timezone, read) {
    const original = process.env.TZ;
    process.env.TZ = timezone;
    try {
        return read();
    } finally {
        if (original === undefined) {
            delete process.env.TZ;
        } else {
            process.env.TZ = original;
        }
    }
}

describe('Date conversion', function () {
    it('returns the same typed date when extracting one sheet under two aliases', function () {
        inTimeZone('America/Sao_Paulo', () => {
            const sourceFile = path.join(repo, 'tests/test-data-typed-dates.xlsx');
            const once = convert({ sourceFile, sheets: ['s'] }).s[0].A;
            const both = convert({
                sourceFile,
                sheets: [{ name: 's', alias: 'first' }, { index: 0, alias: 'second' }]
            });
            assert.equal(both.first[0].A.toISOString(), once.toISOString());
            assert.equal(both.second[0].A.toISOString(), once.toISOString());
        });
    });

    it('returns the same UTC serial date on either side of the international date line', function () {
        const source = workbook({ s: { '!ref': 'A1', A1: { t: 'n', v: 40907, z: 'yyyy-mm-dd' } } });
        const read = () => convert({ source, utcDates: true }).s[0].A.toISOString();
        assert.equal(inTimeZone('UTC', read), '2011-12-30T00:00:00.000Z');
        assert.equal(inTimeZone('Pacific/Apia', read), '2011-12-30T00:00:00.000Z');
    });
});

describe('Keys, selected columns and absent cells', function () {
    it('keeps a unique toString header when duplicate suffixes are enabled', function () {
        const source = workbook({ s: [['toString'], ['present']] });
        const result = convert({
            source, header: { rows: 1 }, columnToKey: { '*': '{{columnHeader}}' }, suffixDuplicateKeys: true
        });
        assert.deepEqual(result.s, [{ toString: 'present' }]);
    });

    it('only invokes header callbacks for columns selected by range', function () {
        const source = workbook({ s: [[12, 'Keep'], [1, 2]] });
        const result = convert({
            source, range: 'B1:B2', header: { rows: 1 }, suffixDuplicateKeys: true,
            columnToKey: { '*': header => header.toLowerCase() }
        });
        assert.deepEqual(result.s, [{ keep: 2 }]);
    });

    it('keeps an absent interior column as null with duplicate suffixes enabled', function () {
        const source = workbook({ s: [[1, null, 3], [2, null, 4]] });
        const result = convert({ source, sheetStubs: true, whitespaceToNull: true, suffixDuplicateKeys: true });
        assert.deepEqual(result.s, [{ A: 1, B: null, C: 3 }, { A: 2, B: null, C: 4 }]);
    });

    it('emits null for an absent value whose header is toString', function () {
        const source = workbook({ s: [['id', 'toString', 'price'], [1, null, 2]] });
        const result = convert({
            source, header: { rows: 1 }, columnToKey: { '*': '{{columnHeader}}' },
            sheetStubs: true, whitespaceToNull: true
        });
        assert.deepEqual(JSON.parse(JSON.stringify(result.s)), [{ id: 1, toString: null, price: 2 }]);
    });

    it('drops a row when its only populated column was skipped by a callback', function () {
        const source = workbook({ s: [['ignore', 'keep'], [9], [null, 3]] });
        const result = convert({
            source, header: { rows: 1 }, columnToKey: { '*': header => header === 'ignore' ? null : header }
        });
        assert.deepEqual(result.s, [{ keep: 3 }]);
    });

    it('passes an unmapped sibling column to a transform as promised in the README', function () {
        const source = workbook({ s: [['Ada', 'Lovelace']] });
        const result = convert({
            source, columnToKey: { A: { property: 'name', transform: (value, row, cells) => value + ' ' + cells.B } }
        });
        assert.deepEqual(result.s, [{ name: 'Ada Lovelace' }]);
    });
});

describe('Sheet layout and selection', function () {
    it('preserves the sheet origin when transposing without a crop', function () {
        const ws = { '!ref': 'A2:B3', A2: { t: 's', v: 'Type' }, A3: { t: 's', v: 'Date' },
            B2: { t: 's', v: 'Book' }, B3: { t: 'n', v: 12 } };
        const result = convert({ source: workbook({ s: ws }), transpose: true });
        assert.deepEqual(result.s, [{ B: 'Type', C: 'Date' }, { B: 'Book', C: 12 }]);
    });

    it('resolves a merged header when the selected data starts below it', function () {
        const ws = XLSX.utils.aoa_to_sheet([['Group', null], [null, 5]]);
        ws['!merges'] = [XLSX.utils.decode_range('A1:B1')];
        const result = convert({
            source: workbook({ s: ws }), range: 'A2:B2', header: { rows: 1 },
            columnToKey: { B: '{{columnHeader}}' }, fillMergedCells: true
        });
        assert.deepEqual(result.s, [{ Group: 5 }]);
    });

    it('allows duplicate unaliased selections for a worksheet named toString', function () {
        const source = workbook({ toString: [['visible']] });
        const one = convert({ source, sheets: ['toString'] });
        assert.deepEqual(convert({ source, sheets: ['toString', 'toString'] }), one);
    });
});
