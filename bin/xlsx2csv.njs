#!/usr/bin/env node
/* xlsx.js (C) 2013-2014 SheetJS -- http://sheetjs.com */
var n = "xlsx";
/* vim: set ts=2: */
var X = require('../');
var fs = require('fs'), program = require('commander');
program
	.version(X.version)
	.usage('[options] <file> [sheetname]')
	.option('-f, --file <file>', 'use specified workbook')
	.option('-s, --sheet <sheet>', 'print specified sheet (default first sheet)')
	.option('-l, --list-sheets', 'list sheet names and exit')
	.option('-S, --formulae', 'print formulae')
	.option('-j, --json', 'emit formatted JSON rather than CSV (all fields text)')
	.option('-J, --raw-js', 'emit raw JS object rather than CSV (raw numbers)')
	.option('-F, --field-sep <sep>', 'CSV field separator', ",")
	.option('-R, --row-sep <sep>', 'CSV row separator', "\n")
	.option('--dev', 'development mode')
	.option('--read', 'read but do not print out contents')
	.option('-q, --quiet', 'quiet mode');

program.on('--help', function() {
	console.log('  Support email: dev@sheetjs.com');
	console.log('  Web Demo: http://oss.sheetjs.com/js-'+n+'/');
});

program.parse(process.argv);

var filename, sheetname = '';
if(program.args[0]) {
	filename = program.args[0];
	if(program.args[1]) sheetname = program.args[1];
}
if(program.sheet) sheetname = program.sheet;
if(program.file) filename = program.file;

if(!filename) {
	console.error(n + "2csv: must specify a filename");
	process.exit(1);
}

if(!fs.existsSync(filename)) {
	console.error(n + "2csv: " + filename + ": No such file or directory");
	process.exit(2);
}

if(program.dev) X.verbose = 2;

var wb;
if(program.dev) wb = X.readFile(filename);
else try {
	wb = X.readFile(filename);
} catch(e) {
	var msg = (program.quiet) ? "" : n + "2csv: error parsing ";
	msg += filename + ": " + e;
	console.error(msg);
	process.exit(3);
}
if(program.read) process.exit(0);

if(program.listSheets) {
	console.log(wb.SheetNames.join("\n"));
	process.exit(0);
}

var target_sheet = sheetname || '';
if(target_sheet === '') target_sheet = wb.SheetNames[0];

var ws;
try {
	ws = wb.Sheets[target_sheet];
	if(!ws) throw "Sheet " + target_sheet + " cannot be found";
} catch(e) {
	console.error(n + "2csv: error parsing "+filename+" "+target_sheet+": " + e);
	process.exit(4);
}

if(!program.quiet) console.error(target_sheet);
if(program.formulae) console.log(X.utils.get_formulae(ws).join("\n"));
else if(program.json) console.log(JSON.stringify(X.utils.sheet_to_row_object_array(ws)));
else if(program.rawJs) console.log(JSON.stringify(X.utils.sheet_to_row_object_array(ws,{raw:true})));
else console.log(X.utils.make_csv(ws, {FS:program.fieldSep, RS:program.rowSep}));
