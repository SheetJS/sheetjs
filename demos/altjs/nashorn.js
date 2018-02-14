#!/usr/bin/env jjs
/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */

/* load module */
var global = (function(){ return this; }).call(null);
load('xlsx.full.min.js');

/* helper to convert byte array to plain JS array */
function b2a(b) {
	var out = new Array(b.length);
	for(var i = 0; i < out.length; i++) out[i] = (b[i] < 0 ? b[i] + 256 : b[i]);
	return out;
}

function process_file(path) {
	java.lang.System.out.println(path);

	/* read file */
	var path = java.nio.file.Paths.get(path);
	var bytes = java.nio.file.Files.readAllBytes(path);
	var u8a = b2a(bytes);

	/* read data */
	var wb = XLSX.read(u8a, {type:"array"});

	/* get first worksheet as an array of arrays */
	var ws = wb.Sheets[wb.SheetNames[0]];
	var js = XLSX.utils.sheet_to_json(ws, {header:1});

	/* print out every line */
	js.forEach(function(l) { java.lang.System.out.println(JSON.stringify(l)); });
}

process_file('sheetjs.xlsx');
process_file('sheetjs.xlsb');
process_file('sheetjs.biff8.xls');
