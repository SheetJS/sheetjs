#!/usr/bin/env node
/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
var n = "xlsx";
/* vim: set ts=2 ft=javascript: */
var X = require('../');
require('exit-on-epipe');
var fs = require('fs'), program = require('commander');
program
	.version(X.version)
	.usage('[options] <file> [sheetname]')
	.option('-f, --file <file>', 'use specified workbook')
	.option('-s, --sheet <sheet>', 'print specified sheet (default first sheet)')
	.option('-N, --sheet-index <idx>', 'use specified sheet index (0-based)')
	.option('-p, --password <pw>', 'if file is encrypted, try with specified pw')
	.option('-l, --list-sheets', 'list sheet names and exit')
	.option('-o, --output <file>', 'output to specified file')

	.option('-B, --xlsb', 'emit XLSB to <sheetname> or <file>.xlsb')
	.option('-M, --xlsm', 'emit XLSM to <sheetname> or <file>.xlsm')
	.option('-X, --xlsx', 'emit XLSX to <sheetname> or <file>.xlsx')
	.option('-Y, --ods',  'emit ODS  to <sheetname> or <file>.ods')
	.option('-2, --biff2','emit XLS  to <sheetname> or <file>.xls (BIFF2)')
	.option('-6, --xlml', 'emit SSML to <sheetname> or <file>.xls (2003 XML)')
	.option('-T, --fods', 'emit FODS to <sheetname> or <file>.fods (Flat ODS)')

	.option('-S, --formulae', 'print formulae')
	.option('-j, --json', 'emit formatted JSON (all fields text)')
	.option('-J, --raw-js', 'emit raw JS object (raw numbers)')
	.option('-A, --arrays', 'emit rows as JS objects (raw numbers)')
	.option('-D, --dif', 'emit data interchange format (dif)')
	.option('-K, --sylk', 'emit symbolic link (sylk)')
	.option('-P, --prn', 'emit formatted text (prn)')
	.option('-t, --txt', 'emit delimited text (txt)')

	.option('-F, --field-sep <sep>', 'CSV field separator', ",")
	.option('-R, --row-sep <sep>', 'CSV row separator', "\n")
	.option('-n, --sheet-rows <num>', 'Number of rows to process (0=all rows)')
	.option('--sst', 'generate shared string table for XLS* formats')
	.option('--compress', 'use compression when writing XLSX/M/B and ODS')
	.option('--perf', 'do not generate output')
	.option('--all', 'parse everything; write as much as possible')
	.option('--dev', 'development mode')
	.option('--read', 'read but do not print out contents')
	.option('-q, --quiet', 'quiet mode');

program.on('--help', function() {
	console.log('  Default output format is CSV');
	console.log('  Support email: dev@sheetjs.com');
	console.log('  Web Demo: http://oss.sheetjs.com/js-'+n+'/');
});

/* output formats, update list with full option name */
var workbook_formats = ['xlsx', 'xlsm', 'xlsb', 'ods', 'fods'];
/* flag, bookType, default ext */
var wb_formats_2 = [
	['xlml', 'xlml', 'xls']
];
program.parse(process.argv);

/* see https://github.com/SheetJS/j/issues/4 */
if(process.version === 'v0.10.31') {
	var msgs = [
		"node v0.10.31 is known to crash on OSX and Linux, refusing to proceed.",
		"see https://github.com/SheetJS/j/issues/4 for the relevant discussion.",
		"see https://github.com/joyent/node/issues/8208 for the relevant node issue"
	];
	msgs.forEach(function(m) { console.error(m); });
	process.exit(1);
}

var filename/*:?string*/, sheetname = '';
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
/*:: if(filename) { */
if(!fs.existsSync(filename)) {
	console.error(n + ": " + filename + ": No such file or directory");
	process.exit(2);
}

var opts = {}, wb/*:?Workbook*/;
if(program.listSheets) opts.bookSheets = true;
if(program.sheetRows) opts.sheetRows = program.sheetRows;
if(program.password) opts.password = program.password;
var seen = false;
function wb_fmt() {
	seen = true;
	opts.cellFormula = true;
	opts.cellNF = true;
	if(program.output) sheetname = program.output;
}
workbook_formats.forEach(function(m) { if(program[m]) { wb_fmt(); } });
wb_formats_2.forEach(function(m) { if(program[m[0]]) { wb_fmt(); } });
if(seen);
else if(program.formulae) opts.cellFormula = true;
else opts.cellFormula = false;

if(program.all) {
	opts.cellFormula = true;
	opts.bookVBA = true;
	opts.cellNF = true;
	opts.cellStyles = true;
	opts.sheetStubs = true;
	opts.cellDates = true;
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

/*::   if(wb) { */
if(program.listSheets) {
	console.log((wb.SheetNames||[]).join("\n"));
	process.exit(0);
}

var wopts = ({WTF:opts.WTF, bookSST:program.sst}/*:any*/);
if(program.compress) wopts.compression = true;

/* full workbook formats */
workbook_formats.forEach(function(m) { if(program[m]) {
		X.writeFile(wb, sheetname || ((filename || "") + "." + m), wopts);
		process.exit(0);
} });

wb_formats_2.forEach(function(m) { if(program[m[0]]) {
		wopts.bookType = m[1];
		X.writeFile(wb, sheetname || ((filename || "") + "." + m[2]), wopts);
		process.exit(0);
} });

var target_sheet = sheetname || '';
if(target_sheet === '') {
	if(program.sheetIndex < (wb.SheetNames||[]).length) target_sheet = wb.SheetNames[program.sheetIndex];
	else target_sheet = (wb.SheetNames||[""])[0];
}

var ws;
try {
	ws = wb.Sheets[target_sheet];
	if(!ws) throw "Sheet " + target_sheet + " cannot be found";
} catch(e) {
	console.error(n + ": error parsing "+filename+" "+target_sheet+": " + e);
	process.exit(4);
}

if(program.perf) process.exit(0);

/* single worksheet formats */
[
	['biff2', '.xls'],
	['sylk', '.slk'],
	['prn', '.prn'],
	['txt', '.txt'],
	['dif', '.dif']
].forEach(function(m) { if(program[m[0]]) {
		wopts.bookType = m[0];
		X.writeFile(wb, sheetname || ((filename || "") + m[1]), wopts);
		process.exit(0);
} });

var oo = "";
if(!program.quiet) console.error(target_sheet);
if(program.formulae) oo = X.utils.get_formulae(ws).join("\n");
else if(program.json) oo = JSON.stringify(X.utils.sheet_to_row_object_array(ws));
else if(program.rawJs) oo = JSON.stringify(X.utils.sheet_to_row_object_array(ws,{raw:true}));
else if(program.arrays) oo = JSON.stringify(X.utils.sheet_to_row_object_array(ws,{raw:true, header:1}));
else oo = X.utils.make_csv(ws, {FS:program.fieldSep, RS:program.rowSep});

if(program.output) fs.writeFileSync(program.output, oo);
else console.log(oo);
/*::   } */
/*:: } */
