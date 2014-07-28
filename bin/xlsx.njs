#!/usr/bin/env node
/* xlsx.js (C) 2013-2014 SheetJS -- http://sheetjs.com */
var n = "xlsx";
/* vim: set ts=2 ft=javascript: */
var X = require('../');
var fs = require('fs'), program = require('commander');
program
	.version(X.version)
	.usage('[options] <file> [sheetname]')
	.option('-f, --file <file>', 'use specified workbook')
	.option('-s, --sheet <sheet>', 'print specified sheet (default first sheet)')
	.option('-l, --list-sheets', 'list sheet names and exit')
	.option('-o, --output <file>', 'output to specified file')
	.option('-B, --xlsb', 'emit XLSB to <sheetname> or <file>.xlsb')
	.option('-M, --xlsm', 'emit XLSM to <sheetname> or <file>.xlsm')
	.option('-X, --xlsx', 'emit XLSX to <sheetname> or <file>.xlsx')
	.option('-S, --formulae', 'print formulae')
	.option('-j, --json', 'emit formatted JSON (all fields text)')
	.option('-J, --raw-js', 'emit raw JS object (raw numbers)')
	.option('-F, --field-sep <sep>', 'CSV field separator', ",")
	.option('-R, --row-sep <sep>', 'CSV row separator', "\n")
	.option('-n, --sheet-rows <num>', 'Number of rows to process (0=all rows)')
	.option('--sst', 'generate sst')
	.option('--perf', 'do not generate output')
	.option('--all', 'parse everything; XLS[XMB] write as much as possible')
	.option('--dev', 'development mode')
	.option('--read', 'read but do not print out contents')
	.option('-q, --quiet', 'quiet mode');

program.on('--help', function() {
	console.log('  Default output format is CSV');
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
	console.error(n + ": must specify a filename");
	process.exit(1);
}

if(!fs.existsSync(filename)) {
	console.error(n + ": " + filename + ": No such file or directory");
	process.exit(2);
}

var opts = {}, wb;
if(program.listSheets) opts.bookSheets = true;
if(program.sheetRows) opts.sheetRows = program.sheetRows;
if(program.xlsx || program.xlsm || program.xlsb) {
	opts.cellNF = true;
	if(program.output) sheetname = program.output;
}
else if(program.formulae);
else opts.cellFormula = false;

if(program.all) {
	opts.cellFormula = true;
	opts.cellNF = true;
	opts.cellStyles = true;
}

if(program.dev) {
	X.verbose = 2;
	opts.WTF = true;
	wb = X.readFile(filename, opts);
}
else try {
	wb = X.readFile(filename, opts);
} catch(e) {
	var msg = (program.quiet) ? "" : n + ": error parsing ";
	msg += filename + ": " + e;
	console.error(msg);
	process.exit(3);
}
if(program.read) process.exit(0);

if(program.listSheets) {
	console.log(wb.SheetNames.join("\n"));
	process.exit(0);
}

var wopts = {WTF:opts.WTF, bookSST:program.sst};

if(program.xlsx) return X.writeFile(wb, sheetname || (filename + ".xlsx"), wopts);
if(program.xlsm) return X.writeFile(wb, sheetname || (filename + ".xlsm"), wopts);
if(program.xlsb) return X.writeFile(wb, sheetname || (filename + ".xlsb"), wopts);

var target_sheet = sheetname || '';
if(target_sheet === '') target_sheet = wb.SheetNames[0];

var ws;
try {
	ws = wb.Sheets[target_sheet];
	if(!ws) throw "Sheet " + target_sheet + " cannot be found";
} catch(e) {
	console.error(n + ": error parsing "+filename+" "+target_sheet+": " + e);
	process.exit(4);
}

if(program.perf) return;

var oo = "";
if(!program.quiet) console.error(target_sheet);
if(program.formulae) oo = X.utils.get_formulae(ws).join("\n");
else if(program.json) oo = JSON.stringify(X.utils.sheet_to_row_object_array(ws));
else if(program.rawJs) oo = JSON.stringify(X.utils.sheet_to_row_object_array(ws,{raw:true}));
else oo = X.utils.make_csv(ws, {FS:program.fieldSep, RS:program.rowSep});

if(program.output) fs.writeFileSync(program.output, oo);
else console.log(oo);
