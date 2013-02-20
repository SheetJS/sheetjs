#!/usr/bin/env node

var XLSX = require('../xlsx');
var utils = XLSX.utils;
var filename = process.argv[2];
if(!filename || filename == "-h" || filename === "--help") {
	console.log("usage:",process.argv[1],"<workbook> [sheet]");
	console.log("  when sheet = :list, print a list of sheets in the workbook");
	process.exit(0);
}
var fs = require('fs');
if(!fs.existsSync(filename)) {
	console.error("error:",filename,"does not exist!");
	process.exit(1);
}
var xlsx = XLSX.readFile(filename);
var sheetname = process.argv[3] || xlsx.SheetNames[0];
if(sheetname === ":list") {
	xlsx.SheetNames.forEach(function(x) { console.log(x); });
	process.exit(0);
}
if(xlsx.SheetNames.indexOf(sheetname)===-1) {
	console.error("Sheet", sheetname, "not found in", filename, ".  I see:");
	xlsx.SheetNames.forEach(function(x) { console.error(" - " + x); });
	process.exit(1);
}

var sheet = xlsx.Sheets[sheetname];
console.log(XLSX.utils.sheet_to_csv(sheet));
