'use strict';

// Test helper: builds an .xlsx file in memory (as a Buffer) so each test can
// describe exactly the spreadsheet it needs, without binary fixtures.
const XLSX = require('xlsx');

// sheets: { sheetName: arrayOfArrays | worksheetObject }
function buildWorkbook(sheets, writeOptions) {
	const workbook = XLSX.utils.book_new();
	Object.keys(sheets).forEach(name => {
		const data = sheets[name];
		const worksheet = Array.isArray(data) ? XLSX.utils.aoa_to_sheet(data) : data;
		XLSX.utils.book_append_sheet(workbook, worksheet, name);
	});
	return XLSX.write(workbook, Object.assign({ type: 'buffer', bookType: 'xlsx' }, writeOptions));
}

module.exports = buildWorkbook;
module.exports.XLSX = XLSX;
