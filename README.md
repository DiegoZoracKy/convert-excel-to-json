# convert-excel-to-json

[![test](https://github.com/DiegoZoracKy/convert-excel-to-json/actions/workflows/test.yml/badge.svg)](https://github.com/DiegoZoracKy/convert-excel-to-json/actions/workflows/test.yml)

Convert Excel to JSON, mapping sheet columns to object keys.

Key features:
- Define a specific Range (e.g. `'A1:E6'`)
- Keep empty rows on the result with `includeEmptyLines`
- Specify a column to key mapping (e.g. `{porperty1: 'CELLVALUE A1', property2: 'CELLVALUE B1'}`)
- Get just specific sheets (e.g. `{sheets: ['sheet1', 'sheet2']}`)

See the [CHANGELOG](CHANGELOG.md) for the changes on each release.

## Install

### NPM / Node

```javascript
npm install convert-excel-to-json
```

or to use it via command-line

```javascript
npm install -g convert-excel-to-json
```

**OBS:** The spreadsheet parsing is done by [SheetJS](https://sheetjs.com). SheetJS no longer publishes new versions of the `xlsx` package to the npm registry (the last one there, `0.18.5`, has known vulnerabilities), so this module depends on the official tarball served by the [SheetJS CDN](https://cdn.sheetjs.com) (`https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz`). npm installs it transparently, no extra step is needed.

## Usage / Examples

For all the examples, lets suppose that our excel file has two sheets, named as 'sheet1' and 'sheet2'.

### CLI

OBS: All the following examples can be used via command-line, in this case, the `--config` parameter expects a valid JSON string.

```javascript
$ convert-excel-to-json --config='{"sourceFile": "tests/test-data.xlsx"}'
```

In order to use it passing in only the **sourceFile** without extra configuration:

```javascript
$ convert-excel-to-json --sourceFile="tests/test-data.xlsx"
```

To check the help section:

```javascript
$ convert-excel-to-json --help
```

### Simple conversion

Just gets all the rows, for each sheet, where each row will be represented by an object with a structure like `{ COLUMN: 'CELLVALUE' }`, e.g. from a sheet with only one column ( the column A) and two rows `[{A: 'VALUE OF A1'}, {A: 'VALUE OF A2'}]`

```javascript
'use strict';
const excelToJson = require('convert-excel-to-json');

const result = excelToJson({
	sourceFile: 'SOME-EXCEL-FILE.xlsx'
});

// result will be an Object containing keys with the same name as the sheets found on the excel file. Each of the keys will have an array of objects where each of them represents a row of the container sheet. e.g. for a excel file that has two sheets ('sheet1', 'sheet2')
{
    sheet1: [{
        A: 'data of cell A1',
        B: 'data of cell B1',
        C: 'data of cell C1'
    }],
    sheet2: [{
        A: 'data of cell A1',
        B: 'data of cell B1',
        C: 'data of cell C1'
    }]
}


```

### Converting an xlsx that you have as a Buffer

```javascript
'use strict';
const excelToJson = require('convert-excel-to-json');
const fs = require('fs');

const result = excelToJson({
	source: fs.readFileSync('SOME-EXCEL-FILE.xlsx') // fs.readFileSync return a Buffer
});

// result will be an Object containing keys with the same name as the sheets found on the excel file. Each of the keys will have an array of objects where each of them represents a row of the container sheet. e.g. for a excel file that has two sheets ('sheet1', 'sheet2')
{
    sheet1: [{
        A: 'data of cell A1',
        B: 'data of cell B1',
        C: 'data of cell C1'
    }],
    sheet2: [{
        A: 'data of cell A1',
        B: 'data of cell B1',
        C: 'data of cell C1'
    }]
}
```

### Identifying header rows

You will notice that if your sheet has some top rows setup as a header (it is very common), the first position of our result will have this data, which in this case it should not be very useful. So we can tell the module how many of the rows are headers, so we can skip them and get only the data.

```javascript
'use strict';
const excelToJson = require('convert-excel-to-json');

const result = excelToJson({
	sourceFile: 'SOME-EXCEL-FILE.xlsx',
	header:{
	    // Is the number of rows that will be skipped and will not be present at our result object. Counting from top to bottom
	    rows: 1 // 2, 3, 4, etc.
	}
});

// result will be an Object like the previous example, but without the rows that was defined as headers
```

To also keep the header rows on the result (e.g. to check if the headers are the expected ones) while still using them for `{{columnHeader}}`, set `header.include` to `true`:

```javascript
'use strict';
const excelToJson = require('convert-excel-to-json');

const result = excelToJson({
	sourceFile: 'SOME-EXCEL-FILE.xlsx',
	header:{
	    rows: 1,
	    include: true
	},
	columnToKey: {
		'*': '{{columnHeader}}'
	}
});

// result will be an Object like:
{
    sheet1: [{
        id: 'id',           // the header row
        name: 'name'
    },{
        id: 'data of cell A2',
        name: 'data of cell B2'
    }]
}
```

### Only to specific sheets

Just gets all the rows for each sheet defined on the config object

```javascript
'use strict';
const excelToJson = require('convert-excel-to-json');

const result = excelToJson({
	sourceFile: 'SOME-EXCEL-FILE.xlsx',
	header:{
	    rows: 1
	},
	sheets: ['sheet2']
});

// result will be an Object like:
{
    sheet2: [{
        A: 'data of cell A1',
        B: 'data of cell B1',
        C: 'data of cell C1'
    }]
}
```

### Mapping columns to keys

#### One config to all sheets
Gets all the rows, for each sheet, but defining which columns should be returned and how they should be named on the result object.

```javascript
'use strict';
const excelToJson = require('convert-excel-to-json');

const result = excelToJson({
	sourceFile: 'SOME-EXCEL-FILE.xlsx',
	columnToKey: {
		A: 'id',
		B: 'firstName'
	}
});

// result will be an Object like:
{
    sheet1: [{
        id: 'data of cell A1',
        firstName: 'data of cell B1'
    }],
    sheet2: [{
        id: 'data of cell A1',
        firstName: 'data of cell B1'
    }]
}
```
#### Config per sheet
Gets all the rows, for each sheet, but defining which columns should be returned and how they should be named on the result object, **per sheet**.

```javascript
'use strict';
const excelToJson = require('convert-excel-to-json');

const result = excelToJson({
	sourceFile: 'SOME-EXCEL-FILE.xlsx',
	sheets:[{
	    name: 'sheet1',
	    columnToKey: {
	    	A: 'id',
    		B: 'ProductName'
	    }
	},{
	    name: 'sheet2',
	    columnToKey: {
	    	A: 'id',
    		B: 'ProductDescription'
	    }
	}]
});

// result will be an Object like:
{
    sheet1: [{
        id: 'data of cell A1',
        ProductName: 'data of cell B1'
    }],
    sheet2: [{
        id: 'data of cell A1',
        ProductDescription: 'data of cell B1'
    }]
}
```

**OBS:** The config *header.rows* can also be defined per sheet, like in the previous example of *columnToKey*. e.g.

```javascript
{
	sourceFile: 'SOME-EXCEL-FILE.xlsx',
	sheets:[{
	    name: 'sheet1',
	    header:{
	        rows: 1
	    },
	    columnToKey: {
	    	A: 'id',
    		B: 'ProductName'
	    }
	},{
	    name: 'sheet2',
	    header:{
	        rows: 3
	    },
	    columnToKey: {
	    	A: 'id',
    		B: 'ProductDescription'
	    }
	}]
}
```

### Mapping columns to keys :: Special Variables

#### Cell Variables

A value from a specific cell can be defined as a key name (e.g. `{ A: '{{A1}}' }`). e.g. if we have 3 rows allocated for a header, but the text value is specified at the first row:

```javascript
'use strict';
const excelToJson = require('convert-excel-to-json');

const result = excelToJson({
	sourceFile: 'SOME-EXCEL-FILE.xlsx',
	header:{
	    rows: 3
	}
	columnToKey: {
		'A': '{{A1}}',
		'B': '{{B1}}'
	}
});

// result will be an Object like:
{
    sheet1: [{
        THE-VALUE-OF-THE-CELL-A1: 'data of cell A1',
        THE-VALUE-OF-THE-CELL-B1: 'data of cell B1'
    }],
    sheet2: [{
        THE-VALUE-OF-THE-CELL-A1: 'data of cell A1',
        THE-VALUE-OF-THE-CELL-B1: 'data of cell B1'
    }]
}
```

**OBS:** {{columnHeader}} will follow the config *header.rows* or, in case it is not specified, it will always treat the first row as a header.

#### Automatic key/property naming following the column header {{columnHeader}}

To return all the data but having the object keys named as a row header found at the excel, instead of the column letters, is just use two special configs. Check the following *columnToKey*:

```javascript
'use strict';
const excelToJson = require('convert-excel-to-json');

const result = excelToJson({
	sourceFile: 'SOME-EXCEL-FILE.xlsx',
	columnToKey: {
		'*': '{{columnHeader}}'
	}
});

// result will be an Object like:
{
    sheet1: [{
        THE-VALUE-OF-THE-HEADER-CELL-A1: 'data of cell A1',
        THE-VALUE-OF-THE-HEADER-CELL-B1: 'data of cell B1',
        THE-VALUE-OF-THE-HEADER-CELL-C1: 'data of cell C1'
    }],
    sheet2: [{
        THE-VALUE-OF-THE-HEADER-CELL-A1: 'data of cell A1',
        THE-VALUE-OF-THE-HEADER-CELL-B1: 'data of cell B1',
        THE-VALUE-OF-THE-HEADER-CELL-C1: 'data of cell C1'
    }]
}
```

**OBS:** {{columnHeader}} will follow the config *header.rows* or, in case it is not specified, it will always treat the first row as a header.

#### Using a function to build the keys

Instead of a String, a column (or `'*'`) can be mapped to a Function. It receives the value of the column header cell (following *header.rows*, like `{{columnHeader}}`) and the column letter, and must return the key name. Returning `undefined`, `null` or `''` skips the column.

```javascript
'use strict';
const excelToJson = require('convert-excel-to-json');
const camelCase = require('lodash/camelCase');

const result = excelToJson({
	sourceFile: 'SOME-EXCEL-FILE.xlsx',
	header: {
		rows: 1
	},
	columnToKey: {
		'*': (columnHeader, column) => camelCase(columnHeader) // 'First Name' -> 'firstName'
	}
});
```

### Null values

If you would like to include null values in the result, you can specify the sheetStubs option.

```javascript
'use strict';
const excelToJson = require('convert-excel-to-json');

const result = excelToJson({
	sourceFile: sourceFile,
    sheetStubs: true
});

// result will include null values for empty cells C1, B2 in the example below:
{
    sheet1: [{
        A: 'data of cell A1',
        B: 'data of cell B1',
        C: null
    }],
    sheet2: [{
        A: 'data of cell A1',
        B: null,
        C: 'data of cell C1'
    }]
}
```

### Dates

Cells formatted as dates (or date and time) are returned as JavaScript `Date` objects representing the same date and time shown on Excel, in the local timezone of the machine running the conversion. This includes cells explicitly typed as dates in the file, even without a date format, and keeps milliseconds. If you prefer the Date to represent that wall-clock time in UTC (so it's the same regardless of the server timezone), set `utcDates` to `true`.

```javascript
'use strict';
const excelToJson = require('convert-excel-to-json');

// cell A1 shows 1/1/2019 18:30
const local = excelToJson({ sourceFile: 'SOME-EXCEL-FILE.xlsx' });
// local.sheet1[0].A -> new Date(2019, 0, 1, 18, 30) (e.g. 2019-01-01T10:30:00.000Z on a server at UTC+8)

const utc = excelToJson({ sourceFile: 'SOME-EXCEL-FILE.xlsx', utcDates: true });
// utc.sheet1[0].A -> 2019-01-01T18:30:00.000Z, on any server
```

### Empty lines (includeEmptyLines)

By default, empty rows are removed from the result, so the position of an object on the result array doesn't necessarily match the row number on the sheet. To keep a position for each row (up to the last row with data), set `includeEmptyLines` to `true`. Empty rows will be represented as empty slots (`undefined`) on the array, which become `null` when the result is serialized with `JSON.stringify`.

```javascript
'use strict';
const excelToJson = require('convert-excel-to-json');

// sheet1: row 1 is empty, row 2 has data, row 3 is empty, row 4 has data
const result = excelToJson({
	sourceFile: 'SOME-EXCEL-FILE.xlsx',
	includeEmptyLines: true
});

// result will be an Object like:
{
    sheet1: [
        undefined, // row 1
        { A: 'data of cell A2', B: 'data of cell B2' },
        undefined, // row 3
        { A: 'data of cell A4', B: 'data of cell B4' }
    ]
}

// JSON.stringify(result) -> {"sheet1":[null,{"A":"data of cell A2","B":"data of cell B2"},null,{"A":"data of cell A4","B":"data of cell B4"}]}
```

**OBS:** The empty rows are *holes* on the array, so `.forEach()`, `.map()` and `.filter()` skip them. Use a `for` loop or `Array.from(result.sheet1)` if you need to iterate over them too.

The position on the array follows the row number on the sheet (`result.sheet1[0]` is row 1), so the rows skipped by `header.rows` also take a position, as empty slots. e.g. with `header: { rows: 1 }` and data only on `A2`, the result is `[undefined, { A: 'data of cell A2' }]` (`[null,{"A":"data of cell A2"}]` as JSON).

### Skipping hidden columns and rows

Set `skipHiddenColumns` and/or `skipHiddenRows` to `true` to leave the columns and rows hidden on Excel out of the result.

```javascript
'use strict';
const excelToJson = require('convert-excel-to-json');

// column B and row 3 are hidden
const result = excelToJson({
	sourceFile: 'SOME-EXCEL-FILE.xlsx',
	skipHiddenColumns: true,
	skipHiddenRows: true
});

// result will be an Object like:
{
    sheet1: [{
        A: 'data of cell A1',
        C: 'data of cell C1'
    },{
        A: 'data of cell A2',
        C: 'data of cell C2'
    },{
        A: 'data of cell A4',
        C: 'data of cell C4'
    }]
}
```

**OBS:** Header cells of hidden columns can still be referenced by cell variables, e.g. `{{B1}}`.

### Keeping whitespace (trim)

By default, leading and trailing whitespace is removed from text values. To keep the values exactly as they are in the cells, set `trim` to `false`.

```javascript
'use strict';
const excelToJson = require('convert-excel-to-json');

const result = excelToJson({
	sourceFile: 'SOME-EXCEL-FILE.xlsx',
	trim: false
});

// a cell with the value '  some text  ' will be returned as '  some text  ' instead of 'some text'
```

### Range

A specific range can be defined. Also like the previous configs, for all the sheets or per sheet.

#### One Range for all sheets

```javascript
'use strict';
const excelToJson = require('convert-excel-to-json');

const result = excelToJson({
	sourceFile: 'SOME-EXCEL-FILE.xlsx',
	range: 'A2:B3',
	sheets: ['sheet1', 'sheet2']
});

// result will be an Object like:
{
    sheet1: [{
        A: 'data of cell A2',
        B: 'data of cell B2'
    },{
        A: 'data of cell A3',
        B: 'data of cell B3'
    }],
    sheet2: [{
        A: 'data of cell A2',
        B: 'data of cell B2'
    },{
        A: 'data of cell A3',
        B: 'data of cell B3'
    }]
}
```

#### A Range per sheet

```javascript
'use strict';
const excelToJson = require('convert-excel-to-json');

const result = excelToJson({
	sourceFile: 'SOME-EXCEL-FILE.xlsx',
	sheets: [{
	    name: 'sheet1',
	    range: 'A2:B2'
	},{
	    name: 'sheet2',
	    range: 'A3:B4'
	}]
});

// result will be an Object like this:
{
    sheet1: [{
        A: 'data of cell A2',
        B: 'data of cell B2'
    ],
    sheet2: [{
        A: 'data of cell A3',
        B: 'data of cell B3'
    },{
        A: 'data of cell A4',
        B: 'data of cell B4'
    }]
}
```
