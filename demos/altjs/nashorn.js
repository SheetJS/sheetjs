#!/usr/bin/env jjs
/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
/* read file */
var path = java.nio.file.Paths.get('sheetjs.xlsx');
var fileArray = java.nio.file.Files.readAllBytes(path);

/* convert to plain JS array */
function b2a(b) {
	var out = new Array(b.length);
	for(var i = 0; i < out.length; i++) out[i] = b[i];
	return out;
}
var u8a = b2a(fileArray);

/* load module */
load('./jvm-npm.js');
JSZip = require('../../jszip.js');
cptable = require('../../dist/cpexcel.js');
XLSX = require('../../xlsx.js');

/* read file */
var wb = XLSX.read(u8a, {type:"array"});

/* get first worksheet */
var ws = wb.Sheets[wb.SheetNames[0]];
var js = XLSX.utils.sheet_to_json(ws, {header:1});

/* print out every line */
js.forEach(function(l) { java.lang.System.out.println(JSON.stringify(l)); });
