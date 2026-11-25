# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html). Dates are the npm release dates.

## [Unreleased]

## [2.0.0] - 2026-11-25
Major version, because of the new Node.js requirement.

### Added
- `trim: false` option, to keep leading/trailing whitespace ([#72](https://github.com/DiegoZoracKy/convert-excel-to-json/issues/72), [#81](https://github.com/DiegoZoracKy/convert-excel-to-json/pull/81)).
- `header.include` option, to keep the header rows on the result ([#36](https://github.com/DiegoZoracKy/convert-excel-to-json/issues/36), [#88](https://github.com/DiegoZoracKy/convert-excel-to-json/pull/88)).
- `skipHiddenColumns` and `skipHiddenRows` options ([#46](https://github.com/DiegoZoracKy/convert-excel-to-json/issues/46), [#89](https://github.com/DiegoZoracKy/convert-excel-to-json/pull/89)).
- Functions on `columnToKey`, to build the keys from the header (e.g. `{ '*': _.camelCase }`) ([#70](https://github.com/DiegoZoracKy/convert-excel-to-json/issues/70), [#78](https://github.com/DiegoZoracKy/convert-excel-to-json/pull/78)).
- `{ property, transform }` on `columnToKey`, to transform values during the conversion; `transform(value, row, cells)` ([#4](https://github.com/DiegoZoracKy/convert-excel-to-json/issues/4), [#31](https://github.com/DiegoZoracKy/convert-excel-to-json/issues/31), [#82](https://github.com/DiegoZoracKy/convert-excel-to-json/pull/82)).
- `suffixDuplicateKeys` option, to keep columns with repeated keys (`key`, `key_2`...) ([#21](https://github.com/DiegoZoracKy/convert-excel-to-json/issues/21), [#83](https://github.com/DiegoZoracKy/convert-excel-to-json/pull/83)).
- `formattedValues` option, to get values as displayed on Excel ([#62](https://github.com/DiegoZoracKy/convert-excel-to-json/issues/62), [#65](https://github.com/DiegoZoracKy/convert-excel-to-json/issues/65), [#84](https://github.com/DiegoZoracKy/convert-excel-to-json/pull/84)).
- `parseBooleans` option ([#55](https://github.com/DiegoZoracKy/convert-excel-to-json/issues/55), [#85](https://github.com/DiegoZoracKy/convert-excel-to-json/pull/85)).
- `whitespaceToNull` option ([#9](https://github.com/DiegoZoracKy/convert-excel-to-json/issues/9), [#86](https://github.com/DiegoZoracKy/convert-excel-to-json/pull/86)).
- `ArrayBuffer`/`Uint8Array` as `source`, and browser usage ([#44](https://github.com/DiegoZoracKy/convert-excel-to-json/issues/44), [#35](https://github.com/DiegoZoracKy/convert-excel-to-json/issues/35), [#80](https://github.com/DiegoZoracKy/convert-excel-to-json/pull/80)).
- Sheets by index (`sheets: [0, { index: 1 }]`) ([#25](https://github.com/DiegoZoracKy/convert-excel-to-json/issues/25), [#87](https://github.com/DiegoZoracKy/convert-excel-to-json/pull/87)).
- `transpose` option (rows as headers) ([#50](https://github.com/DiegoZoracKy/convert-excel-to-json/issues/50), [#90](https://github.com/DiegoZoracKy/convert-excel-to-json/pull/90)).
- `alias` on a sheet config, to read the same sheet more than once ([#58](https://github.com/DiegoZoracKy/convert-excel-to-json/issues/58), [#91](https://github.com/DiegoZoracKy/convert-excel-to-json/pull/91)).
- `fillMergedCells` option ([#64](https://github.com/DiegoZoracKy/convert-excel-to-json/issues/64), [#92](https://github.com/DiegoZoracKy/convert-excel-to-json/pull/92)).

### Changed
- **Breaking:** Node.js 22 or later is required (`"engines": { "node": ">=22" }`). CI runs on Node.js 22, 24 and 26 ([#76](https://github.com/DiegoZoracKy/convert-excel-to-json/pull/76)).
- The CLI no longer depends on `magicli`; it accepts the same arguments ([#76](https://github.com/DiegoZoracKy/convert-excel-to-json/pull/76)).
- README: documented the `sheetStubs` option ([#9](https://github.com/DiegoZoracKy/convert-excel-to-json/issues/9), [#41](https://github.com/DiegoZoracKy/convert-excel-to-json/pull/41)).

### Fixed
- Ranges crossing column `Z` (e.g. `A1:AB10`) ([#42](https://github.com/DiegoZoracKy/convert-excel-to-json/issues/42), [#77](https://github.com/DiegoZoracKy/convert-excel-to-json/pull/77)).
- Dates are returned with the exact date and time shown on Excel, without drifting by seconds/hours depending on the host timezone; explicitly typed date cells are returned as Dates and milliseconds are kept. New `utcDates` option to get that wall-clock time in UTC ([#33](https://github.com/DiegoZoracKy/convert-excel-to-json/issues/33), [#52](https://github.com/DiegoZoracKy/convert-excel-to-json/issues/52), [#79](https://github.com/DiegoZoracKy/convert-excel-to-json/pull/79)).

### Security
- `xlsx` (SheetJS) updated to 0.20.3, installed from the official SheetJS CDN, as the versions on the npm registry have known vulnerabilities ([#76](https://github.com/DiegoZoracKy/convert-excel-to-json/pull/76)).

## [1.7.0] - 2020-06-08
### Added
- `sheetStubs` option, to include empty cells on the result with `null` values ([#9](https://github.com/DiegoZoracKy/convert-excel-to-json/issues/9), [#39](https://github.com/DiegoZoracKy/convert-excel-to-json/pull/39)).

## [1.6.1] - 2019-02-23
### Security
- Dependencies updated with `npm audit fix`.

## [1.6.0] - 2018-12-04
### Added
- `source` option, to convert an xlsx that you have as a Buffer instead of a file path ([#22](https://github.com/DiegoZoracKy/convert-excel-to-json/pull/22)).

## [1.5.2] - 2018-11-07
### Security
- Dependencies updated with `npm audit fix`.

## [1.5.1] - 2018-11-06
### Fixed
- Error when a cell referenced by a key variable (e.g. `{{columnHeader}}`) doesn't exist ([#19](https://github.com/DiegoZoracKy/convert-excel-to-json/pull/19)).

## [1.5.0] - 2018-11-06
### Added
- `includeEmptyLines` option, to keep empty rows on the result ([#20](https://github.com/DiegoZoracKy/convert-excel-to-json/pull/20)).

## [1.4.1] - 2018-09-21
### Fixed
- Values are trimmed only when the cell has a formatted text (`w`).

## [1.4.0] - 2018-07-10
### Changed
- Improvements to the command-line usage (`--config` and `--sourceFile`).
- Dropped support for Node.js 4.

## [1.3.0] - 2018-02-19
### Changed
- `xlsx` updated to 0.12.1.
- Date cells are returned as `Date` objects (`cellDates: true`).

## [1.2.1] - 2017-11-09
### Changed
- Command-line interface built with `magicli`, with a `--help` section.

## [1.2.0] - 2017-08-01
### Changed
- Numeric cells are returned as Numbers instead of their formatted text.

## [1.1.1] - 2017-05-19
### Fixed
- `.trim()` was being called on values that aren't strings.

## 1.1.0 - 2016-12-31 (not published to npm)
### Added
- Command-line usage, installing the module globally.
- The config can be passed in as a JSON string.

## [1.0.0] - 2016-08-17
### Added
- First release: convert Excel to JSON, with `sheets`, `header.rows`, `columnToKey` (including cell variables and `{{columnHeader}}`) and `range` options.

[Unreleased]: https://github.com/DiegoZoracKy/convert-excel-to-json/compare/2.0.0...HEAD
[2.0.0]: https://github.com/DiegoZoracKy/convert-excel-to-json/compare/3719b02...2.0.0
[1.7.0]: https://www.npmjs.com/package/convert-excel-to-json/v/1.7.0
[1.6.1]: https://www.npmjs.com/package/convert-excel-to-json/v/1.6.1
[1.6.0]: https://www.npmjs.com/package/convert-excel-to-json/v/1.6.0
[1.5.2]: https://www.npmjs.com/package/convert-excel-to-json/v/1.5.2
[1.5.1]: https://www.npmjs.com/package/convert-excel-to-json/v/1.5.1
[1.5.0]: https://www.npmjs.com/package/convert-excel-to-json/v/1.5.0
[1.4.1]: https://www.npmjs.com/package/convert-excel-to-json/v/1.4.1
[1.4.0]: https://www.npmjs.com/package/convert-excel-to-json/v/1.4.0
[1.3.0]: https://www.npmjs.com/package/convert-excel-to-json/v/1.3.0
[1.2.1]: https://www.npmjs.com/package/convert-excel-to-json/v/1.2.1
[1.2.0]: https://www.npmjs.com/package/convert-excel-to-json/v/1.2.0
[1.1.1]: https://www.npmjs.com/package/convert-excel-to-json/v/1.1.1
[1.0.0]: https://www.npmjs.com/package/convert-excel-to-json/v/1.0.0
