'use strict';

const assert = require('assert');
const fs = require('fs');
const excelToJson = require('../');

// #44: "_fs.readFileSync is not a function" in the browser, where the file comes as an ArrayBuffer
// #35: using it on a web page
describe('"source" as ArrayBuffer / Uint8Array / base64 (#44, #35)', function() {
	const buffer = fs.readFileSync(`${__dirname}/test-data.xlsx`);
	const expected = excelToJson({ source: buffer });
	const arrayBuffer = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.length);

	it('should accept an ArrayBuffer (FileReader.readAsArrayBuffer / file.arrayBuffer())', function() {
		assert.deepStrictEqual(excelToJson({ source: arrayBuffer }), expected);
	});

	it('should accept a Uint8Array', function() {
		assert.deepStrictEqual(excelToJson({ source: new Uint8Array(arrayBuffer) }), expected);
	});

	it('should accept a Uint8Array that is a view on part of a bigger ArrayBuffer', function() {
		const bigger = new Uint8Array(buffer.length + 20);
		bigger.set(buffer, 10);
		assert.deepStrictEqual(excelToJson({ source: bigger.subarray(10, 10 + buffer.length) }), expected);
	});

	it('should keep accepting a base64 string', function() {
		assert.deepStrictEqual(excelToJson({ source: buffer.toString('base64') }), expected);
	});

	it('should accept a binary string with sourceType: "binary"', function() {
		assert.deepStrictEqual(excelToJson({ source: buffer.toString('binary'), sourceType: 'binary' }), expected);
	});
});

describe('"sourceFile" without a file system (#44)', function() {
	it('should explain that "source" must be used on the browser', function() {
		const XLSX = require('xlsx');
		const originalReadFile = XLSX.readFile;
		XLSX.readFile = () => { throw new TypeError('_fs.readFileSync is not a function'); };
		try {
			assert.throws(() => excelToJson({ sourceFile: 'some-file.xlsx' }), /On the browser, read the file as an ArrayBuffer/);
		} finally {
			XLSX.readFile = originalReadFile;
		}
	});
});
