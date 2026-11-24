# convert-excel-to-json

[![test](https://github.com/DiegoZoracKy/convert-excel-to-json/actions/workflows/test.yml/badge.svg)](https://github.com/DiegoZoracKy/convert-excel-to-json/actions/workflows/test.yml)

Convert Excel to JSON, mapping sheet columns to object keys.

Key features:
- Define a specific Range (e.g. `'A1:E6'`)
- Specify a column to key mapping (e.g. `{porperty1: 'CELLVALUE A1', property2: 'CELLVALUE B1'}`)
- Get just specific sheets (e.g. `{sheets: ['sheet1', 'sheet2']}`)

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

### On the browser

The module can be used on a web page through a bundler (e.g. [Vite](https://vitejs.dev), [webpack](https://webpack.js.org), [esbuild](https://esbuild.github.io)). As there's no file system on the browser, `sourceFile` can't be used: read the file chosen by the user as an `ArrayBuffer` and pass it in as `source` (an `ArrayBuffer`, including one coming from an iframe, a `Uint8Array`, a Node.js `Buffer` or a base64 String are accepted; for other String encodings set `sourceType` to `'binary'` or `'string'`).

```html
<input type="file" id="excel-file" accept=".xlsx,.xls">
<script type="module" src="./main.js"></script>
```

```javascript
// main.js (bundled with Vite, webpack, esbuild...)
import excelToJson from 'convert-excel-to-json';

document.getElementById('excel-file').addEventListener('change', async event => {
	const file = event.target.files[0];

	const result = excelToJson({
		source: await file.arrayBuffer(), // or the result of FileReader.readAsArrayBuffer()
		header: {
			rows: 1
		},
		columnToKey: {
			'*': '{{columnHeader}}'
		}
	});

	console.log(result); // { sheet1: [...], sheet2: [...] }
});
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

### Sheets by index

Sheets can also be referenced by their position on the workbook (zero-based), whatever their names are. Use a Number, or an object with `index` (instead of `name`) to also define a config for that sheet. The result keys are still the sheet names.

```javascript
'use strict';
const excelToJson = require('convert-excel-to-json');

const result = excelToJson({
	sourceFile: 'SOME-EXCEL-FILE.xlsx',
	sheets: [1, {
	    index: 2,
	    header: {
	        rows: 1
	    }
	}]
});

// result will be an Object like (for a workbook with the sheets 'sheet1', 'sheet2' and 'sheet3'):
{
    sheet2: [ ... ],
    sheet3: [ ... ]
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

### Transforming values during the conversion

A column (or `'*'`) can also be mapped to an object with the key name on `property` and a `transform` function. The function receives the cell value and the entire row (with the original values of all the other columns) and returns the value to be used on the result.

```javascript
'use strict';
const excelToJson = require('convert-excel-to-json');

const result = excelToJson({
	sourceFile: 'SOME-EXCEL-FILE.xlsx',
	header: {
		rows: 1
	},
	columnToKey: {
		A: 'id',
		B: 'firstName',
		C: {
			property: 'fullName',
			transform: (value, row) => `${row.firstName} ${value}` // concatenating with the value of another column
		},
		D: {
			property: 'phoneNumber',
			transform: value => value.toString()
		},
		E: {
			property: 'price',
			transform: value => Number(value.toFixed(2)) // 1.4400000000000002 -> 1.44
		}
	}
});
```

`property` accepts the same special variables as a String (e.g. `{{columnHeader}}`), so a transform can be applied to all the columns at once:

```javascript
columnToKey: {
	'*': {
		property: '{{columnHeader}}',
		transform: value => typeof value === 'number' ? Number(value.toFixed(2)) : value
	}
}
```

**OBS:** `transform` is called only for cells that have a value (empty cells are not present on the result unless `sheetStubs` is set).

### Duplicated keys

When two or more columns end up with the same key (e.g. the header has repeated names and `{{columnHeader}}` is used), the value of the last column overwrites the others. Set `suffixDuplicateKeys` to `true` to keep all of them, by adding a numeric suffix to the repeated keys (in the order the columns appear). It can also be set per sheet.

```javascript
'use strict';
const excelToJson = require('convert-excel-to-json');

// header row: | test | test | other | test |
const result = excelToJson({
	sourceFile: 'SOME-EXCEL-FILE.xlsx',
	header: {
		rows: 1
	},
	columnToKey: {
		'*': '{{columnHeader}}'
	},
	suffixDuplicateKeys: true
});

// result will be an Object like:
{
    sheet1: [{
        test: 'data of cell A2',
        test_2: 'data of cell B2',
        other: 'data of cell C2',
        test_3: 'data of cell D2'
    }]
}
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

### Rows as headers (transpose)

When the headers are on the first column and each record is a column (instead of a row), set `transpose` to `true`. Rows and columns are swapped before the conversion, so the first column becomes the first row, and all the other configs (`header`, `columnToKey`, `{{columnHeader}}`...) work as usual. It can also be set per sheet.

If a `range` is defined, it refers to the original sheet, and the transposed range is treated as if it started at `A1`.

```javascript
'use strict';
const excelToJson = require('convert-excel-to-json');

// sheet1:
// |   | A    | B          | C           |
// | 1 | Type | TypeData   | OtherType   |
// | 2 | Date | DateData   | OtherDate   |
const result = excelToJson({
	sourceFile: 'SOME-EXCEL-FILE.xlsx',
	transpose: true,
	header: {
		rows: 1 // the first column
	},
	columnToKey: {
		'*': '{{columnHeader}}'
	}
});

// result will be an Object like:
{
    sheet1: [{
        Type: 'TypeData',
        Date: 'DateData'
    },{
        Type: 'OtherType',
        Date: 'OtherDate'
    }]
}
```

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

### Formatted values (as displayed on Excel)

By default, numbers and dates are returned with their raw values (e.g. a cell showing `22.5%` is returned as `0.225`, and a date as a `Date` object). To get the text exactly as it is displayed on Excel, following the cell format, set `formattedValues` to `true`.

```javascript
'use strict';
const excelToJson = require('convert-excel-to-json');

const result = excelToJson({
	sourceFile: 'SOME-EXCEL-FILE.xlsx',
	formattedValues: true
});

// result will be an Object like:
{
    sheet1: [{
        A: '22.5%',       // instead of 0.225
        B: '1/1/19',      // instead of a Date object
        C: '$1,234.50'    // instead of 1234.5
    }]
}
```

The display text of numbers and dates is returned exactly as Excel shows it, including any padding added by the number format (e.g. accounting formats such as `_(* #,##0.00_)` add spaces). Text cells keep the default whitespace handling.

If you only need the dates as strings, you can also keep the default and convert them with `.map()` on the result (e.g. `row.date.toISOString()`).

### Booleans

By default, boolean cells are returned as the text displayed on Excel (`'TRUE'` / `'FALSE'`). Set `parseBooleans` to `true` to get them as JavaScript Booleans. Text cells containing `true` or `false` (case insensitive) are also converted.

```javascript
'use strict';
const excelToJson = require('convert-excel-to-json');

const result = excelToJson({
	sourceFile: 'SOME-EXCEL-FILE.xlsx',
	header: {
		rows: 1
	},
	columnToKey: {
		'*': '{{columnHeader}}'
	},
	parseBooleans: true
});

// result will be an Object like:
{
    sheet1: [{
        Name: 'Anthony',
        Membership: true // instead of 'TRUE'
    }]
}
```

### Blank spaces as null values

Cells containing only blank spaces are returned as they are (e.g. `'   '`). Set `whitespaceToNull` to `true` to get `null` for them instead. Together with `sheetStubs`, every column of the sheet (or of the `range`, following `columnToKey`) will be present on every row of the result, with `null` for the empty, blank and absent cells.

```javascript
'use strict';
const excelToJson = require('convert-excel-to-json');

const result = excelToJson({
	sourceFile: 'SOME-EXCEL-FILE.xlsx',
	sheetStubs: true,
	whitespaceToNull: true
});

// result will be an Object like (B2 is empty, C1 has only blank spaces):
{
    sheet1: [{
        A: 'data of cell A1',
        B: 'data of cell B1',
        C: null
    }, {
        A: 'data of cell A2',
        B: null,
        C: 'data of cell C2'
    }]
}
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
