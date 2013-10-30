#!/usr/bin/env node

/* vim: set ts=2: */
var XLSX = require('../');
var fs = require('fs'), program = require('commander');
program
	.version('0.3.1')
	.usage('[options] <file> [sheetname]')
	.option('-f, --file <file>', 'use specified workbook')
	.option('-s, --sheet <sheet>', 'print specified sheet (default first sheet)')
	.option('-l, --list-sheets', 'list sheet names and exit')
	.option('-F, --formulae', 'print formulae')
	.option('--dev', 'development mode')
	.option('--read', 'read but do not print out contents')
	.option('-q, --quiet', 'quiet mode')
	.parse(process.argv);

var filename, sheetname = '';
if(program.args[0]) {
	filename = program.args[0];
	if(program.args[1]) sheetname = program.args[1];
}
if(program.sheet) sheetname = program.sheet;
if(program.file) filename = program.file;

if(!filename) {
	console.error("xlsx2csv: must specify a filename");
	process.exit(1);
}

if(!fs.existsSync(filename)) {
	console.error("xlsx2csv: " + filename + ": No such file or directory");
	process.exit(2);
}

if(program.dev) XLSX.verbose = 2;

var wb;
if(program.dev) wb = XLSX.readFile(filename);
else try {
	wb = XLSX.readFile(filename);
} catch(e) {
	var msg = (program.quiet) ? "" : "xlsx2csv: error parsing ";
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
	console.error("xlsx2csv: error parsing "+filename+" "+target_sheet+": " + e);
	process.exit(4);
}

if(!program.quiet) console.error(target_sheet);
if(program.formulae) console.log(XLSX.utils.get_formulae(ws).join("\n"));
else console.log(XLSX.utils.make_csv(ws));
