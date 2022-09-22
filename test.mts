/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
/*jshint mocha:true */
/* eslint-env mocha */
/*global process, document, require */
/*global ArrayBuffer, Uint8Array */
/*::
declare type EmptyFunc = (() => void) | null;
declare var afterEach:(test:EmptyFunc)=>void;
declare var cptable: any;
*/
import * as assert_ from 'https://deno.land/std/testing/asserts.ts';
import * as base64_ from 'https://deno.land/std/encoding/base64.ts';
const assert: any = {...assert_};
assert.throws = function(f: () => void) { assert.assertThrows(function() { try { f(); } catch(e) { throw e instanceof Error ? e : new Error(e); }})};
assert.doesNotThrow = function(f: ()=>void) { f(); };
assert.equal = assert.assertEquals;
assert.notEqual = assert.assertNotEquals;
assert.deepEqual = (x: any,y: any, z?: string) => assert.equal(JSON.stringify(x), JSON.stringify(y), z);

// @deno-types="./types/index.d.ts"
import * as X from './xlsx.mjs';
import * as cptable from './dist/cpexcel.full.mjs';
X.set_cptable(cptable);
import XLSX_ZAHL from './dist/xlsx.zahl.mjs';
var DIF_XL = true;

type BSEncoding = 'utf-8' | 'binary' | 'base64';
type BFEncoding = 'buffer';
type ShEncoding = BSEncoding | BFEncoding;
function readFileSync2(x: string): Uint8Array;
function readFileSync2(x: string, e: BSEncoding): string;
function readFileSync2(x: string, e: BFEncoding): Uint8Array;
function readFileSync2(x: string, e?: ShEncoding): Uint8Array | string {
	const u8 = Deno.readFileSync(x);
	if(!e) return u8;
	switch(e) {
		case 'utf-8': return new TextDecoder().decode(u8);
		case 'base64': return base64_.encode(u8);
		case 'buffer': return u8;
		case 'binary': return Array.from({length: u8.length}, (_,i) => String.fromCharCode(u8[i])).join("");
	}
	throw new Error(`unsupported encoding ${e}`)
}
var fs = {
	readFileSync: readFileSync2,
	existsSync: (x: string): boolean => { try { Deno.readFileSync(x); } catch(e) { return false; } return true; },
	readdirSync: (x: string) => ([...Deno.readDirSync(x)].filter(x => x.isFile).map(x => x.name))
};

var browser = false;

var Buffer_from = function(data: any, enc?: string): Uint8Array {
	if(data instanceof Uint8Array) return data;
	if((!enc || enc == "utf8") && typeof data == "string") return new TextEncoder().encode(data);
	if(enc == "binary" && typeof data == "string") return Uint8Array.from({length: data.length}, (_,i) => data[i].charCodeAt(0))
	throw `unsupported encoding ${enc}`;
}

var opts: any = {cellNF: true};
var TYPE: 'buffer' = "buffer";
opts.type = TYPE;
var fullex = [".xlsb", /*".xlsm",*/ ".xlsx"/*, ".xlml", ".xls"*/];
var ofmt: X.BookType[] = ["xlsb", "xlsm", "xlsx", "ods", "biff2", "biff5", "biff8", "xlml", "sylk", "dif", "dbf", "eth", "fods", "csv", "txt", "html"];
var ex = fullex.slice(); ex = ex.concat([".ods", ".xls", ".xml", ".fods"]);
	opts.WTF = true;
	opts.cellStyles = true;
	if(Deno.env.get("FMTS") === "full") Deno.env.set("FMTS", ex.join(":"));
	if(Deno.env.get("FMTS")) ex=(Deno.env.get("FMTS")||"").split(":").map(function(x){return x[0]==="."?x:"."+x;});
var exp = ex.map(function(x: string){ return x + ".pending"; });
function test_file(x: string){ return ex.indexOf(x.slice(-5))>=0||exp.indexOf(x.slice(-13))>=0 || ex.indexOf(x.slice(-4))>=0||exp.indexOf(x.slice(-12))>=0; }

var files: string[]  = [], fileA: string[] = [];
if(!browser) {
	var _files = fs.existsSync('tests.lst') ? fs.readFileSync('tests.lst', 'utf-8').split("\n").map(function(x) { return x.trim(); }) : fs.readdirSync('test_files');
	for(var _filesi = 0; _filesi < _files.length; ++_filesi) if(test_file(_files[_filesi])) files.push(_files[_filesi]);
	var _fileA = fs.existsSync('tests/testA.lst') ? fs.readFileSync('tests/testA.lst', 'utf-8').split("\n").map(function(x) { return x.trim(); }) : [];
	for(var _fileAi = 0; _fileAi < _fileA.length; ++_fileAi) if(test_file(_fileA[_fileAi])) fileA.push(_fileA[_fileAi]);
}

var can_write_numbers = typeof Set !== "undefined" && typeof Array.prototype.findIndex == "function" && typeof Uint8Array !== "undefined" && typeof Uint8Array.prototype.indexOf == "function";

/* Excel enforces 31 character sheet limit, although technical file limit is 255 */
function fixsheetname(x: string): string { return x.substr(0,31); }

function stripbom(x: string): string { return x.replace(/^\ufeff/,""); }
function fixcsv(x: string): string { return stripbom(x).replace(/\t/g,",").replace(/#{255}/g,"").replace(/"/g,"").replace(/[\n\r]+/g,"\n").replace(/\n*$/,""); }
function fixjson(x: string): string { return x.replace(/[\r\n]+$/,""); }

var dir = "./test_files/";

var dirwp = dir + "artifacts/wps/", dirqp = dir + "artifacts/quattro/";
var paths: any = {
	aadbf:  dirwp + 'write.dbf',
	aadif:  dirwp + 'write.dif',
	aaxls:  dirwp + 'write.xls',
	aaxlsx: dirwp + 'write.xlsx',
	aaxml:  dirwp + 'write.xml',

	abcsv:  dirqp + 'write_.csv',
	abdif:  dirqp + 'write_.dif',
	abslk:  dirqp + 'write_.slk',
	abx57:  dirqp + 'write_57.xls',
	abwb2:  dirqp + 'write_6.wb2',
	abwb2b: dirqp + 'write_6b.wb2',
	abwb3:  dirqp + 'write_8.wb3',
	abqpw:  dirqp + 'write_9.qpw',
	abx97:  dirqp + 'write_97.xls',
	abwks:  dirqp + 'write_L1.wks',
	abwk1:  dirqp + 'write_L2.wk1',
	abwk3:  dirqp + 'write_L3.wk3',
	abwk4:  dirqp + 'write_L45.wk4',
	ab123:  dirqp + 'write_L9.123',
	ab124:  dirqp + 'write_L97.123',
	abwke2: dirqp + 'write_Led.wke',
	abwq1:  dirqp + 'write_qpdos.wq1',
	abwb1:  dirqp + 'write_qpw.wb1',

	afxls:   dir + 'AutoFilter.xls',
	afxml:   dir + 'AutoFilter.xml',
	afods:   dir + 'AutoFilter.ods',
	afxlsx:  dir + 'AutoFilter.xlsx',
	afxlsb:  dir + 'AutoFilter.xlsb',

	asxls:   dir + 'author_snowman.xls',
	asxls5:  dir + 'author_snowman.xls5',
	asxml:   dir + 'author_snowman.xml',
	asods:   dir + 'author_snowman.ods',
	asxlsx:  dir + 'author_snowman.xlsx',
	asxlsb:  dir + 'author_snowman.xlsb',

	cpxls:   dir + 'custom_properties.xls',
	cpxml:   dir + 'custom_properties.xls.xml',
	cpxlsx:  dir + 'custom_properties.xlsx',
	cpxlsb:  dir + 'custom_properties.xlsb',

	cssxls: dir + 'cell_style_simple.xls',
	cssxml: dir + 'cell_style_simple.xml',
	cssxlsx: dir + 'cell_style_simple.xlsx',
	cssxlsb: dir + 'cell_style_simple.xlsb',

	cstxls: dir + 'comments_stress_test.xls',
	cstxml: dir + 'comments_stress_test.xls.xml',
	cstxlsx: dir + 'comments_stress_test.xlsx',
	cstxlsb: dir + 'comments_stress_test.xlsb',
	cstods: dir + 'comments_stress_test.ods',

	cwxls:  dir + 'column_width.xls',
	cwxls5:  dir + 'column_width.biff5',
	cwxml:  dir + 'column_width.xml',
	cwxlsx:  dir + 'column_width.xlsx',
	cwxlsb:  dir + 'column_width.xlsb',
	cwslk:  dir + 'column_width.slk',

	dnsxls: dir + 'defined_names_simple.xls',
	dnsxml: dir + 'defined_names_simple.xml',
	dnsxlsx: dir + 'defined_names_simple.xlsx',
	dnsxlsb: dir + 'defined_names_simple.xlsb',
	dnsslk: dir + 'defined_names_simple.slk',

	dnuxls: dir + 'defined_names_unicode.xls',
	dnuxml: dir + 'defined_names_unicode.xml',
	dnuods: dir + 'defined_names_unicode.ods',
	dnuxlsx: dir + 'defined_names_unicode.xlsx',
	dnuxlsb: dir + 'defined_names_unicode.xlsb',

	dtxls:  dir + 'xlsx-stream-d-date-cell.xls',
	dtxml:  dir + 'xlsx-stream-d-date-cell.xls.xml',
	dtxlsx:  dir + 'xlsx-stream-d-date-cell.xlsx',
	dtxlsb:  dir + 'xlsx-stream-d-date-cell.xlsb',

	dtfxlsx: dir + 'DataTypesFormats.xlsx',

	fstxls: dir + 'formula_stress_test.xls',
	fstxml: dir + 'formula_stress_test.xls.xml',
	fstxlsx: dir + 'formula_stress_test.xlsx',
	fstxlsb: dir + 'formula_stress_test.xlsb',
	fstods: dir + 'formula_stress_test.ods',

	hlxls:  dir + 'hyperlink_stress_test_2011.xls',
	hlxml:  dir + 'hyperlink_stress_test_2011.xml',
	hlxlsx:  dir + 'hyperlink_stress_test_2011.xlsx',
	hlxlsb:  dir + 'hyperlink_stress_test_2011.xlsb',

	ilxls:   dir + 'internal_link.xls',
	ilxls5:  dir + 'internal_link.biff5',
	ilxml:   dir + 'internal_link.xml',
	ilxlsx:  dir + 'internal_link.xlsx',
	ilxlsb:  dir + 'internal_link.xlsb',
	ilods:   dir + 'internal_link.ods',

	lonxls: dir + 'LONumbers.xls',
	lonxlsx: dir + 'LONumbers.xlsx',

	m19xlsx:  dir + 'metadata_2019.xlsx',
	m19xlsb:  dir + 'metadata_2019.xlsb',

	mcxls:  dir + 'merge_cells.xls',
	mcxml:  dir + 'merge_cells.xls.xml',
	mcxlsx:  dir + 'merge_cells.xlsx',
	mcxlsb:  dir + 'merge_cells.xlsb',
	mcods:  dir + 'merge_cells.ods',

	nfxls:  dir + 'number_format.xls',
	nfxml:  dir + 'number_format.xls.xml',
	nfxlsx:  dir + 'number_format.xlsm',
	nfxlsb:  dir + 'number_format.xlsb',

	olxls:  dir + 'outline.xls',
	olxls5:  dir + 'outline.biff5',
	olxlsx:  dir + 'outline.xlsx',
	olxlsb:  dir + 'outline.xlsb',
	olods:  dir + 'outline.ods',

	pmxls:  dir + 'page_margins_2016.xls',
	pmxls5: dir + 'page_margins_2016_5.xls',
	pmxml:  dir + 'page_margins_2016.xml',
	pmxlsx: dir + 'page_margins_2016.xlsx',
	pmxlsb: dir + 'page_margins_2016.xlsb',

	rhxls:  dir + 'row_height.xls',
	rhxls5:  dir + 'row_height.biff5',
	rhxml:  dir + 'row_height.xml',
	rhxlsx:  dir + 'row_height.xlsx',
	rhxlsb:  dir + 'row_height.xlsb',
	rhslk:  dir + 'row_height.slk',

	svxls:  dir + 'sheet_visibility.xls',
	svxls5: dir + 'sheet_visibility5.xls',
	svxml:  dir + 'sheet_visibility.xml',
	svxlsx: dir + 'sheet_visibility.xlsx',
	svxlsb: dir + 'sheet_visibility.xlsb',
	svods:  dir + 'sheet_visibility.ods',

	swcxls: dir + 'apachepoi_SimpleWithComments.xls',
	swcxml: dir + '2011/apachepoi_SimpleWithComments.xls.xml',
	swcxlsx: dir + 'apachepoi_SimpleWithComments.xlsx',
	swcxlsb: dir + '2013/apachepoi_SimpleWithComments.xlsx.xlsb'
};

function pathit(p: string, ext: string[]) { return ext.map(function(n) { return paths[p + n]; }); }
var FSTPaths = pathit("fst", ["xlsx", "xlsb", "xls", "xml", "ods"]);
var CSTPaths = pathit("cst", ["xlsx", "xlsb", "xls", "xml", "ods"]);
var MCPaths =  pathit("mc",  ["xlsx", "xlsb", "xls", "xml", "ods"]);
var CSSPaths = pathit("css", ["xlsx", "xlsb", "xls", "xml"]);
var NFPaths =  pathit("nf",  ["xlsx", "xlsb", "xls", "xml"]);
var DTPaths =  pathit("dt",  ["xlsx", "xlsb", "xls", "xml"]);
var HLPaths =  pathit("hl",  ["xlsx", "xlsb", "xls", "xml"]);
var ILPaths =  pathit("il",  ["xlsx", "xlsb", "xls", "xml", "ods", "xls5"]);
var OLPaths =  pathit("ol",  ["xlsx", "xlsb", "xls", "ods", "xls5"]);
var PMPaths =  pathit("pm",  ["xlsx", "xlsb", "xls", "xml", "xls5"]);
var SVPaths =  pathit("sv",  ["xlsx", "xlsb", "xls", "xml", "xls5"]);
var CWPaths =  pathit("cw",  ["xlsx", "xlsb", "xls", "xml", "xls5", "slk"]);
var RHPaths =  pathit("rh",  ["xlsx", "xlsb", "xls", "xml", "xls5", "slk"]);

var artifax = [
	"cstods", "cstxls", "cstxlsb", "cstxlsb", "cstxml", "aadbf", "aadif",
	"aaxls", "aaxml", "aaxlsx", "ab123", "ab124", "abcsv", "abdif", "abqpw",
	"abslk", "abwb1", "abwb2", "abwb3", "abwk1", "abwk3", "abwk4", "abwks",
	"abwq1", "abx57", "abx97", "abwke2", "abwb2b"
].map(function(x) { return paths[x]; });

async function parsetest(x: string, wb: X.WorkBook, full: boolean, ext: string, t: Deno.TestContext) {
	ext = (ext ? " [" + ext + "]": "");
	if(!full && ext) return;
	describe(x + ext + ' should have all bits', function() {
		var sname = dir + '2016/' + x.substr(x.lastIndexOf('/')+1) + '.sheetnames';
		if(!fs.existsSync(sname)) sname = dir + '2011/' + x.substr(x.lastIndexOf('/')+1) + '.sheetnames';
		if(!fs.existsSync(sname)) sname = dir + '2013/' + x.substr(x.lastIndexOf('/')+1) + '.sheetnames';
		it('should have all sheets', function() {
			wb.SheetNames.forEach(function(y) { assert.ok(wb.Sheets[y], 'bad sheet ' + y); });
		});
		if(fs.existsSync(sname)) it('should have the right sheet names', function() {
			var file = fs.readFileSync(sname, 'utf-8').replace(/\r/g,"");
			var names = wb.SheetNames.map(fixsheetname).join("\n") + "\n";
			if(file.length && !x.match(/artifacts/)) assert.equal(names, file);
		});
	});
	describe(x + ext + ' should generate CSV', function() {
		for(var i = 0; i < wb.SheetNames.length; ++i) { var ws = wb.SheetNames[i];
			it('#' + i + ' (' + ws + ')', function() {
				X.utils.sheet_to_csv(wb.Sheets[ws]);
			});
		}
	});
	describe(x + ext + ' should generate JSON', function() {
		for(var i = 0; i < wb.SheetNames.length; ++i) { var ws = wb.SheetNames[i];
			it('#' + i + ' (' + ws + ')', function() {
				X.utils.sheet_to_json(wb.Sheets[ws]);
			});
		}
	});
	describe(x + ext + ' should generate formulae', function() {
		for(var i = 0; i < wb.SheetNames.length; ++i) { var ws = wb.SheetNames[i];
			it('#' + i + ' (' + ws + ')', function() {
				X.utils.sheet_to_formulae(wb.Sheets[ws]);
			});
		}
	});
	if(!full) return;
	var getfile = function(dir: string, x: string, i: number, type: string) {
		var name = (dir + x + '.' + i + type);
		var root = "";
		if(x.substr(-5) === ".xlsb") {
			root = x.slice(0,-5);
			if(!fs.existsSync(name)) name=(dir + root + '.xlsx.' + i + type);
			if(!fs.existsSync(name)) name=(dir + root + '.xlsm.' + i + type);
			if(!fs.existsSync(name)) name=(dir + root + '.xls.'  + i + type);
		}
		if(x.substr(-4) === ".xls") {
			root = x.slice(0,-4);
			if(!fs.existsSync(name)) name=(dir + root + '.xlsx.' + i + type);
			if(!fs.existsSync(name)) name=(dir + root + '.xlsm.' + i + type);
			if(!fs.existsSync(name)) name=(dir + root + '.xlsb.' + i + type);
		}
		return name;
	};
	describe(x + ext + ' should generate correct CSV output', function() {
		for(var i = 0; i < wb.SheetNames.length; ++i) { var ws = wb.SheetNames[i];
			var name = getfile(dir, x, i, ".csv");
			if(fs.existsSync(name)) it('#' + i + ' (' + ws + ')', function() {
				var file = fs.readFileSync(name, 'utf-8');
				var csv = X.utils.sheet_to_csv(wb.Sheets[ws]);
				assert.equal(fixcsv(csv), fixcsv(file), "CSV badness");
			});
		}
	});
	if(typeof JSON !== 'undefined') describe(x + ext + ' should generate correct JSON output', function() {
		for(var i = 0; i < wb.SheetNames.length; ++i) { var ws = wb.SheetNames[i];
			var rawjson = getfile(dir, x, i, ".rawjson");
			if(fs.existsSync(rawjson)) it('#' + i + ' (' + ws + ')', function() {
				var file = fs.readFileSync(rawjson, 'utf-8');
				var json: Array<any> = X.utils.sheet_to_json(wb.Sheets[ws],{raw:true});
				assert.equal(JSON.stringify(json), fixjson(file), "JSON badness");
			});

			var jsonf = getfile(dir, x, i, ".json");
			if(fs.existsSync(jsonf)) it('#' + i + ' (' + ws + ')', function() {
				var file = fs.readFileSync(jsonf, 'utf-8');
				var json: Array<any> = X.utils.sheet_to_json(wb.Sheets[ws], {raw:false});
				assert.equal(JSON.stringify(json), fixjson(file), "JSON badness");
			});
		}
	});
	if(fs.existsSync(dir + '2011/' + x + '.xml'))
	describe(x + ext + '.xml from 2011', function() {
		it('should parse', function() {
			/*var wb = */X.readFile(dir + '2011/' + x + '.xml', opts);
		});
	});
	if(fs.existsSync(dir + '2013/' + x + '.xlsb'))
	describe(x + ext + '.xlsb from 2013', function() {
		it('should parse', function() {
			/*var wb = */X.readFile(dir + '2013/' + x + '.xlsb', opts);
		});
	});
	if(fs.existsSync(dir + x + '.xml' + ext))
	describe(x + '.xml', function() {
		it('should parse', function() {
			/*var wb = */X.readFile(dir + x + '.xml', opts);
		});
	});
}

describe('should parse test files', function() {
	var idx = 0, x = "";
	for(idx = 0; idx < files.length; ++idx) { var x = files[idx];
		if(x.slice(-8) == ".pending" || !fs.existsSync(dir + x)) continue;
		it(x, function() {
			var _wb = X.readFile(dir + x, opts);
			await parsetest(x, _wb, true, "", t);
			for(var fidx = 0; fidx < fullex.length; ++fidx) { var ext = fullex[fidx];
				it(x + ' [' + ext + ']', function(){
					var wb = _wb;
					wb = X.read(X.write(wb, {type:"buffer", bookType:(ext.replace(/\./,"") as X.BookType)}), {WTF:opts.WTF, cellNF: true});
					await parsetest(x, wb, ext.replace(/\./,"") !== "xlsb", ext, t);
				});
			}
		});
	}
	for(idx = 0; idx < fileA.length; ++idx) { var x = fileA[idx];
		if(x.slice(-8) == ".pending" || !fs.existsSync(dir + x)) continue;
		it(x, function() {
			var wb = X.readFile(dir + x, {WTF:opts.WTF, sheetRows:10});
			await parsetest(x, wb, false, "", t);
		});
	}
});

function get_cell(ws: X.WorkSheet, addr: string) {
	if(!Array.isArray(ws)) return ws[addr];
	var a = X.utils.decode_cell(addr);
	return (ws[a.r]||[])[a.c];
}

function each_cell(ws: X.WorkSheet, f: (c: X.CellObject) => any) {
	if(Array.isArray(ws)) ws.forEach(function(row) { if(row) row.forEach(f); });
	else Object.keys(ws).forEach(function(addr) { if(addr[0] === "!" || !ws.hasOwnProperty(addr)) return; f(ws[addr]); });
}

function each_sheet(wb: X.WorkBook, f: (ws: X.WorkSheet, i: number)=>any) { wb.SheetNames.forEach(function(n, i) { f(wb.Sheets[n], i); }); }

/* comments_stress_test family */
function check_comments(wb: X.WorkBook) {
	var ws0 = wb.Sheets["Sheet2"];
	assert.equal(get_cell(ws0,"A1").c[0].a, 'Author');
	assert.equal(get_cell(ws0,"A1").c[0].t, 'Author:\nGod thinks this is good');
	assert.equal(get_cell(ws0,"C1").c[0].a, 'Author');
	assert.equal(get_cell(ws0,"C1").c[0].t, 'I really hope that xlsx decides not to use magic like rPr');

	var ws3 = wb.Sheets["Sheet4"];
	assert.equal(get_cell(ws3,"B1").c[0].a, 'Author');
	assert.equal(get_cell(ws3,"B1").c[0].t, 'The next comment is empty');
	assert.equal(get_cell(ws3,"B2").c[0].a, 'Author');
	assert.equal(get_cell(ws3,"B2").c[0].t, '');
}

describe('parse options', function() {
	var html_cell_types = ['s'];
	describe('cell', function() {
		it('XLSX should generate HTML by default', function() {
			var wb = X.read(fs.readFileSync(paths.cstxlsx), {type:TYPE});
			var ws = wb.Sheets["Sheet1"];
			each_cell(ws, function(cell) {
				assert.ok(html_cell_types.indexOf(cell.t) === -1 || cell.h);
			});
		});
		it('XLSX should not generate HTML when requested', function() {
			var wb = X.read(fs.readFileSync(paths.cstxlsx), {type:TYPE, cellHTML:false});
			var ws = wb.Sheets["Sheet1"];
			each_cell(ws, function(cell) {
				assert.ok(typeof cell.h === 'undefined');
			});
		});
		it('should generate formulae by default', function() {
			FSTPaths.forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:TYPE});
				var found = false;
				wb.SheetNames.forEach(function(s) {
					each_cell(wb.Sheets[s], function(cell) {
						if(typeof cell.f !== 'undefined') return (found = true);
					});
				});
				assert.ok(found);
			});
		});
		it('should not generate formulae when requested', function() {
			FSTPaths.forEach(function(p) {
				var wb =X.read(fs.readFileSync(p),{type:TYPE,cellFormula:false});
				wb.SheetNames.forEach(function(s) {
					each_cell(wb.Sheets[s], function(cell) {
						assert.ok(typeof cell.f === 'undefined');
					});
				});
			});
		});
		it('should generate formatted text by default', function() {
			FSTPaths.forEach(function(p) {
				var wb = X.read(fs.readFileSync(p),{type:TYPE});
				var found = false;
				wb.SheetNames.forEach(function(s) {
					var ws = wb.Sheets[s];
					each_cell(ws, function(cell) {
						if(typeof cell.w !== 'undefined') return (found = true);
					});
				});
				assert.ok(found);
			});
		});
		it('should not generate formatted text when requested', function() {
			FSTPaths.forEach(function(p) {
				var wb =X.read(fs.readFileSync(p),{type:TYPE, cellText:false});
				wb.SheetNames.forEach(function(s) {
					var ws = wb.Sheets[s];
					each_cell(ws, function(cell) {
						assert.ok(typeof cell.w === 'undefined');
					});
				});
			});
		});
		it('should not generate number formats by default', function() {
			NFPaths.forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:TYPE});
				wb.SheetNames.forEach(function(s) {
					var ws = wb.Sheets[s];
					each_cell(ws, function(cell) {
						assert.ok(typeof cell.z === 'undefined');
					});
				});
			});
		});
		it('should generate number formats when requested', function() {
			NFPaths.forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:TYPE, cellNF: true});
				wb.SheetNames.forEach(function(s) {
					var ws = wb.Sheets[s];
					each_cell(ws, function(cell) {
						assert.ok(cell.t!== 'n' || typeof cell.z !== 'undefined');
					});
				});
			});
		});
		it('should not generate cell styles by default', function() {
			CSSPaths.forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:TYPE, WTF:true});
				wb.SheetNames.forEach(function(s) {
					var ws = wb.Sheets[s];
					each_cell(ws, function(cell) {
						assert.ok(typeof cell.s === 'undefined');
					});
				});
			});
		});
		it('should generate cell styles when requested', function() {
			/* TODO: XLS / XLML */
			[paths.cssxlsx /*, paths.cssxlsb, paths.cssxls, paths.cssxml*/].forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:TYPE, cellStyles:true});
				var found = false;
				each_sheet(wb, function(ws/*::, i*/) { /*:: void i; */each_cell(ws, function(cell) {
					if(typeof cell.s !== 'undefined') return (found = true);
				}); });
				assert.ok(found);
			});
		});
		it('should not generate cell dates by default', function() {
			DTPaths.forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:TYPE});
				each_sheet(wb, function(ws/*::, i*/) { /*:: void i; */each_cell(ws, function(cell) {
					assert.ok(cell.t !== 'd');
				}); });
			});
		});
		it('should generate cell dates when requested', function() {
			DTPaths.forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:TYPE, cellDates: true, WTF:true});
				var found = false;
				each_sheet(wb, function(ws/*::, i*/) { /*:: void i; */each_cell(ws, function(cell) {
					if(cell.t === 'd') return (found = true);
				}); });
				assert.ok(found);
			});
		});
		it('should preserve `_xlfn.` only when requested', function() {
			var wb = { SheetNames: ["Sheet1"], Sheets: { Sheet1: {
				"!ref": "A1:A1",
				"A1": { t:"n", v:2, f:"_xlfn.IFS(2>3,1,3>2,2)"}
			} } };
			var str = X.write(wb, {bookType: "xlsx", type: "binary"});
			var wb2 = X.read(str, {type: "binary"});
			/*jshint -W069 */
			assert.equal(wb2.Sheets["Sheet1"]["A1"].f, "IFS(2>3,1,3>2,2)");
			var wb3 = X.read(str, {type: "binary", xlfn: true});
			assert.equal(wb3.Sheets["Sheet1"]["A1"].f, "_xlfn.IFS(2>3,1,3>2,2)");
			/*jshint +W069 */
		});
	});
	describe('sheet', function() {
		it('should not generate sheet stubs by default', function() {
			MCPaths.forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:TYPE});
				assert.throws(function() { return get_cell(wb.Sheets["Merge"], "A2").v; });
			});
		});
		it('should generate sheet stubs when requested', function() {
			MCPaths.forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:TYPE, sheetStubs:true});
				assert.ok(get_cell(wb.Sheets["Merge"], "A2").t == 'z');
			});
		});
		it('should handle stub cells', function() {
			MCPaths.forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:TYPE, sheetStubs:true});
				X.utils.sheet_to_csv(wb.Sheets["Merge"]);
				X.utils.sheet_to_json(wb.Sheets["Merge"]);
				X.utils.sheet_to_formulae(wb.Sheets["Merge"]);
				ofmt.forEach(function(f) { if(f != "dbf") X.write(wb, {type:TYPE, bookType:f}); });
			});
		});
		function checkcells(wb: X.WorkBook, A46: boolean, B26: boolean, C16: boolean, D2: boolean) {
			([ ["A46", A46], ["B26", B26], ["C16", C16], ["D2", D2] ] as Array<[string, boolean]>).forEach(function(r: [string, boolean]) {
				assert.ok((typeof get_cell(wb.Sheets["Text"], r[0]) !== 'undefined') == r[1]);
			});
		}
		it('should read all cells by default', function() { FSTPaths.forEach(function(p) {
			checkcells(X.read(fs.readFileSync(p), {type:TYPE}), true, true, true, true);
		}); });
		it('sheetRows n=30', function() { FSTPaths.forEach(function(p) {
			checkcells(X.read(fs.readFileSync(p), {type:TYPE, sheetRows:30}), false, true, true, true);
		}); });
		it('sheetRows n=20', function() { FSTPaths.forEach(function(p) {
			checkcells(X.read(fs.readFileSync(p), {type:TYPE, sheetRows:20}), false, false, true, true);
		}); });
		it('sheetRows n=10', function() { FSTPaths.forEach(function(p) {
			checkcells(X.read(fs.readFileSync(p), {type:TYPE, sheetRows:10}), false, false, false, true);
		}); });
		it('sheetRows n=1', function() { ofmt.forEach(function(fmt) {
		([TYPE, "base64", "binary", "array"] as Array<'binary' | 'base64' | 'binary' | 'array'>).forEach(function(ot) {
			var data = [[1,2],[3,4],[5,6]];
			var ws = X.utils.aoa_to_sheet(data);
			assert.ok(ws['!ref'] === "A1:B3");
			var wb = X.utils.book_new();
			X.utils.book_append_sheet(wb, ws, "Sheet1");
			var bs = X.write(wb, { bookType: fmt, type: ot, WTF:true });

			var wb0 = X.read(bs, { type: ot, WTF:true });
			var ws0 = wb0.Sheets["Sheet1"];
			assert.equal(ws0['!ref'], "A1:B3");
			assert.equal(get_cell(ws0, "A1").v, fmt == "dbf" ? "1" : 1);
			assert.equal(get_cell(ws0, "B2").v, 4);
			assert.equal(get_cell(ws0, "A3").v, 5);

			var wb1 = X.read(bs, { type: ot, sheetRows: 1 });
			var ws1 = wb1.Sheets["Sheet1"];
			assert.equal(ws1['!ref'], "A1:B1");
			assert.equal(get_cell(ws1, "A1").v, fmt == "dbf" ? "1" : 1);
			assert.ok(!get_cell(ws1, "B2"));
			assert.ok(!get_cell(ws1, "A3"));
			if(ws1['!fullref']) assert.equal(ws1['!fullref'], "A1:B3");

			var wb2 = X.read(bs, { type: ot, sheetRows: 2 });
			var ws2 = wb2.Sheets["Sheet1"];
			assert.equal(ws2['!ref'], "A1:B2");
			assert.equal(get_cell(ws2, "A1").v, fmt == "dbf" ? "1" : 1);
			assert.equal(get_cell(ws2, "B2").v, 4);
			assert.ok(!get_cell(ws2, "A3"));
			if(ws2['!fullref']) assert.equal(ws2['!fullref'], "A1:B3");

			var wb3 = X.read(bs, { type: ot, sheetRows: 3 });
			var ws3 = wb3.Sheets["Sheet1"];
			assert.equal(ws3['!ref'], "A1:B3");
			assert.equal(get_cell(ws3, "A1").v, fmt == "dbf" ? "1" : 1);
			assert.equal(get_cell(ws3, "B2").v, 4);
			assert.equal(get_cell(ws3, "A3").v, 5);
			if(ws3['!fullref']) assert.equal(ws3['!fullref'], "A1:B3");
		}); }); });
	});
	describe('book', function() {
		it('bookSheets should not generate sheets', function() {
			MCPaths.forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:TYPE, bookSheets:true});
				assert.ok(typeof wb.Sheets === 'undefined');
			});
		});
		it('bookProps should not generate sheets', function() {
			NFPaths.forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:TYPE, bookProps:true});
				assert.ok(typeof wb.Sheets === 'undefined');
			});
		});
		it('bookProps && bookSheets should not generate sheets', function() {
			PMPaths.forEach(function(p) {
				if(!fs.existsSync(p)) return;
				var wb = X.read(fs.readFileSync(p), {type:TYPE, bookProps:true, bookSheets:true});
				assert.ok(typeof wb.Sheets === 'undefined');
			});
		});

		var FSTXL = [
			[paths.fstxlsx, true],
			[paths.fstxlsb, true],
			[paths.fstxls, false]
		];
		it('should not generate deps by default', function() {
			FSTPaths.forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:TYPE});
				//assert.ok(typeof wb.Deps === 'undefined' || !(wb.Deps && wb.Deps.length>0));
			});
		});
		it('bookDeps should generate deps (XLSX/XLSB)', function() {
			FSTXL.forEach(function(p) {
				if(!p[1]) return;
				var wb = X.read(fs.readFileSync(p[0]), {type:TYPE, bookDeps:true});
				//assert.ok(typeof wb.Deps !== 'undefined' && wb.Deps.length > 0);
			});
		});

		var ckf = function(wb: X.WorkBook, fields: string[], exists: boolean) { fields.forEach(function(f) { assert.ok((typeof (wb as any)[f] !== 'undefined') == exists); }); };
		it('should not generate book files by default', function() {FSTXL.forEach(function(r) {
			var wb = X.read(fs.readFileSync(r[0]), {type:TYPE});
			ckf(wb, r[1] ? ['files', 'keys'] : ['cfb'], false);
		}); });
		it('bookFiles should generate book files', function() {FSTXL.forEach(function(r) {
			var wb = X.read(fs.readFileSync(r[0]), {type:TYPE, bookFiles:true});
			ckf(wb, r[1] ? ['files', 'keys'] : ['cfb'], true);
		}); });

		var NFVBA = ["nfxlsx", "nfxlsb", "nfxls"].map(function(n) { return paths[n]; });
		it('should not generate VBA by default', function() { NFPaths.forEach(function(p) {
			var wb = X.read(fs.readFileSync(p), {type:TYPE}); assert.ok(typeof wb.vbaraw === 'undefined');
		}); });
		it('bookVBA should generate vbaraw', function() { NFVBA.forEach(function(p) {
			var wb = X.read(fs.readFileSync(p),{type: TYPE, bookVBA: true});
			assert.ok(wb.vbaraw);
			var cfb = X.CFB.read(wb.vbaraw, {type: 'array'});
			assert.ok(X.CFB.find(cfb, '/VBA/ThisWorkbook'));
		}); });
	});
});

describe('input formats', function() {
	it('should read binary strings', function() { artifax.forEach(function(p) {
		X.read(fs.readFileSync(p, 'binary'), {type: 'binary'});
	}); });
	it('should read base64 strings', function() { artifax.forEach(function(p) {
		X.read(fs.readFileSync(p, 'base64'), {type: 'base64'});
	}); });
	it('handles base64 within data URI scheme (gh-2762)', function() {
		var data = 'TmFtZXMNCkhhZmV6DQpTYW0NCg==';

		var wb0 = X.read(data, { type: 'base64' }); // raw base64 string
		var wb1 = X.read('data:;base64,' + data, { type: 'base64' }); // data URI, no media type
		var wb2 = X.read('data:text/csv;base64,' + data, { type: 'base64' }); // data URI, CSV type
		var wb3 = X.read('data:application/vnd.ms-excel;base64,' + data, { type: 'base64' }); // data URI, Excel

		[wb0, wb1, wb2, wb3].forEach(function(wb) {
			var ws = wb.Sheets.Sheet1;
			assert.equal(ws["!ref"], "A1:A3");
			assert.equal(get_cell(ws, "A1").v, "Names");
			assert.equal(get_cell(ws, "A2").v, "Hafez");
			assert.equal(get_cell(ws, "A3").v, "Sam");
		});
	});
	if(typeof Uint8Array !== 'undefined') it('should read array', function() { artifax.forEach(function(p) {
		X.read(fs.readFileSync(p, 'binary').split("").map(function(x) { return x.charCodeAt(0); }), {type:'array'});
	}); });
	it('should read Buffers', function() { artifax.forEach(function(p) {
		X.read(fs.readFileSync(p), {type: 'buffer'});
	}); });
	if(typeof Uint8Array !== 'undefined') it('should read ArrayBuffer / Uint8Array', function() { artifax.forEach(function(p) {
		var payload: any = fs.readFileSync(p, "buffer");
		var ab = new ArrayBuffer(payload.length), vu = new Uint8Array(ab);
		for(var i = 0; i < payload.length; ++i) vu[i] = payload[i];
		X.read(ab, {type: 'array'});
		X.read(vu, {type: 'array'});
	}); });
	it('should throw if format is unknown', function() { artifax.forEach(function(p) {
		assert.throws(function() { X.read(fs.readFileSync(p), ({type: 'dafuq'} as any)); });
	}); });

	var T = browser ? 'base64' : 'buffer';
	it('should default to "' + T + '" type', function() { artifax.forEach(function(p) {
		X.read(fs.readFileSync(p));
	}); });
	if(!browser) it('should read files', function() { artifax.forEach(function(p) { X.readFile(p); }); });
});

describe('output formats', function() {
	var fmts = [
		/* fmt   unicode   str */
		["xlsx",   true,  false],
		["xlsb",   true,  false],
		["xls",    true,  false],
		["xlml",   true,   true],
		["ods",    true,  false],
		["fods",   true,   true],
		["csv",    true,   true],
		["txt",    true,   true],
		["rtf",   false,   true],
		["sylk",  false,   true],
		["eth",   false,   true],
		["html",   true,   true],
		["dif",   false,   true],
		["dbf",   false,  false],
		["prn",   false,   true]
	];
	function RT(T: 'string' | 'binary' | 'base64' | 'array' | 'buffer') {
		fmts.forEach(function(fmt) {
			var wb = X.utils.book_new();
			X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet([['R',"\u2603"],["\u0BEE",2]]), "Sheet1");
			if(T == 'string' && !fmt[2]) return assert.throws(function() {X.write(wb, {type: T, bookType:(fmt[0] as X.BookType), WTF:true});});
			var out = X.write(wb, {type: T, bookType:(fmt[0] as X.BookType), WTF:true});
			var nwb = X.read(out, {type: T, PRN: fmt[0] == 'prn', WTF:true});
			var nws = nwb.Sheets[nwb.SheetNames[0]];
			assert.equal(get_cell(nws, "B2").v, 2);
			assert.equal(get_cell(nws, "A1").v, "R");
			if(fmt[1]) assert.equal(get_cell(nws, "A2").v, "\u0BEE");
			if(fmt[1]) assert.equal(get_cell(nws, "B1").v, "\u2603");
		});
	}
	it('should write binary strings', function() { RT('binary'); });
	it('should write base64 strings', function() { RT('base64'); });
	it('should write JS strings', function() { RT('string'); });
	if(typeof ArrayBuffer !== 'undefined') it('should write array buffers', function() { RT('array'); });
	if(!browser) it('should write buffers', function() { RT('buffer'); });
	it('should throw if format is unknown', function() { assert.throws(function() { RT('dafuq' as any); }); });
});

function eqarr(a: any[],b: any[]) {
	assert.equal(a.length, b.length);
	a.forEach(function(x, i) { assert.equal(x, b[i]); });
}

describe('API', function() {
	it('book_append_sheet', function() {
		var wb = X.utils.book_new();
		X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet([[1,2,3],[4],[5]]), "A");
		X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet([[1,2,3],[4],[5]]));
		X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet([[1,2,3],[4],[5]]));
		X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet([[1,2,3],[4],[5]]), "B");
		X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet([[1,2,3],[4],[5]]));
		eqarr(wb.SheetNames, ["A","Sheet1","Sheet2","B","Sheet3"]);
	});
	it('sheet_add_json', function() {
		var ws = X.utils.json_to_sheet([{A:"S", B:"h", C:"e", D:"e", E:"t", F:"J", G:"S"}], {header:["A","B","C","D","E","F","G"], skipHeader:true});
		X.utils.sheet_add_json(ws, [{A:1, B:2}, {A:2, B:3}, {A:3, B:4}], {skipHeader:true, origin:"A2"});
		X.utils.sheet_add_json(ws, [{A:5, B:6, C:7}, {A:6, B:7, C:8}, {A:7, B:8, C:9}], {skipHeader:true, origin:{r:1, c:4}, header:["A","B","C"]});
		X.utils.sheet_add_json(ws, [{A:4, B:5, C:6, D:7, E:8, F:9, G:0}], {header:["A","B","C","D","E","F","G"], skipHeader:true, origin:-1});
		assert.equal(X.utils.sheet_to_csv(ws).trim(), "S,h,e,e,t,J,S\n1,2,,,5,6,7\n2,3,,,6,7,8\n3,4,,,7,8,9\n4,5,6,7,8,9,0");
	});
	it('sheet_add_aoa', function() {
		var ws = X.utils.aoa_to_sheet([ "SheetJS".split("") ]);
		X.utils.sheet_add_aoa(ws, [[1,2], [2,3], [3,4]], {origin: "A2"});
		X.utils.sheet_add_aoa(ws, [[5,6,7], [6,7,8], [7,8,9]], {origin:{r:1, c:4}});
		X.utils.sheet_add_aoa(ws, [[4,5,6,7,8,9,0]], {origin: -1});
		assert.equal(X.utils.sheet_to_csv(ws).trim(), "S,h,e,e,t,J,S\n1,2,,,5,6,7\n2,3,,,6,7,8\n3,4,,,7,8,9\n4,5,6,7,8,9,0");
	});
	it('sheet_add_aoa support object cell', function() {
		var data = X.utils.aoa_to_sheet([
			['url', 'name', 'id'],
			[ { l: { Target: 'https://123.com' }, v: 'url', t: 's' }, 'tom', 'xxx' ]
		]);
		if(assert.deepEqual) assert.deepEqual(data.A2, { l: { Target: 'https://123.com' }, v: 'url', t: 's' });
	});
	it('decode_range', function() {
		var _c = "ABC", _r = "123", _C = "DEF", _R = "456";

		var r = X.utils.decode_range(_c + _r + ":" + _C + _R);
		assert.ok(r.s != r.e);
		assert.equal(r.s.c, X.utils.decode_col(_c)); assert.equal(r.s.r, X.utils.decode_row(_r));
		assert.equal(r.e.c, X.utils.decode_col(_C)); assert.equal(r.e.r, X.utils.decode_row(_R));

		r = X.utils.decode_range(_c + _r);
		assert.ok(r.s != r.e);
		assert.equal(r.s.c, X.utils.decode_col(_c)); assert.equal(r.s.r, X.utils.decode_row(_r));
		assert.equal(r.e.c, X.utils.decode_col(_c)); assert.equal(r.e.r, X.utils.decode_row(_r));
	});
});

function coreprop(props: X.Properties) {
	assert.equal(props.Title, 'Example with properties');
	assert.equal(props.Subject, 'Test it before you code it');
	assert.equal(props.Author, 'Pony Foo');
	assert.equal(props.Manager, 'Despicable Drew');
	assert.equal(props.Company, 'Vector Inc');
	assert.equal(props.Category, 'Quirky');
	assert.equal(props.Keywords, 'example humor');
	assert.equal(props.Comments, 'some comments');
	assert.equal(props.LastAuthor, 'Hugues');
}
function custprop(props: any) {
	assert.equal(props['I am a boolean'], true);
	assert.equal(props['Date completed'].toISOString(), '1967-03-09T16:30:00.000Z');
	assert.equal(props.Status, 2);
	assert.equal(props.Counter, -3.14);
}

function cmparr(x: any[]){ for(var i=1;i<x.length;++i) assert.deepEqual(x[0], x[i]); }

function deepcmp(x: any, y: any, k: string, m?: string, c: number = 0): void {
	var s = k.indexOf(".");
	m = (m||"") + "|" + (s > -1 ? k.substr(0,s) : k);
	if(s < 0) return assert[c<0?'notEqual':'equal'](x[k], y[k], m);
	return deepcmp(x[k.substr(0,s)],y[k.substr(0,s)],k.substr(s+1),m,c);
}

var styexc = [
	'A2|H10|bgColor.rgb',
	'F6|H1|patternType'
];
var stykeys = [
	"patternType",
	"fgColor.rgb",
	"bgColor.rgb"
];
function diffsty(ws: X.WorkSheet, r1: string,r2: string) {
	var c1 = get_cell(ws,r1).s, c2 = get_cell(ws,r2).s;
	stykeys.forEach(function(m) {
		var c = -1;
		if(styexc.indexOf(r1+"|"+r2+"|"+m) > -1) c = 1;
		else if(styexc.indexOf(r2+"|"+r1+"|"+m) > -1) c = 1;
		deepcmp(c1,c2,m,r1+","+r2,c);
	});
}

function hlink1(ws: X.WorkSheet) {[
	["A1", "http://www.sheetjs.com"],
	["A2", "http://oss.sheetjs.com"],
	["A3", "http://oss.sheetjs.com#foo"],
	["A4", "mailto:dev@sheetjs.com"],
	["A5", "mailto:dev@sheetjs.com?subject=hyperlink"],
	["A6", "../../sheetjs/Documents/Test.xlsx"],
	["A7", "http://sheetjs.com", "foo bar baz"]
].forEach(function(r) {
	assert.equal(get_cell(ws, r[0]).l.Target, r[1]);
	if(r[2]) assert.equal(get_cell(ws, r[0]).l.Tooltip, r[2]);
}); }

function hlink2(ws: X.WorkSheet) { [
	["A1", "#Sheet2!A1"],
	["A2", "#WBScope"],
	["A3", "#Sheet1!WSScope1", "#Sheet1!C7:E8"],
	["A5", "#Sheet1!A5"]
].forEach(function(r) {
	if(r[2] && get_cell(ws, r[0]).l.Target == r[2]) return;
	assert.equal(get_cell(ws, r[0]).l.Target, r[1]);
}); }

function check_margin(margins: X.MarginInfo, exp: number[]) {
	(["left", "right", "top", "bottom", "header", "footer"] as Array<keyof X.MarginInfo>).forEach(function(m: keyof X.MarginInfo, i: number) {
		assert.equal(margins[m],   exp[i]);
	});
}

describe('parse features', function() {
	describe('sheet visibility', function() {
		var wbs = SVPaths.map(function(p) { return X.read(fs.readFileSync(p), {type:TYPE}); });
		it('should detect visible sheets', function() {
			wbs.forEach(function(wb) {
				assert.ok(!wb?.Workbook?.Sheets?.[0]?.Hidden);
			});
		});
		it('should detect all hidden sheets', function() {
			wbs.forEach(function(wb) {
				assert.ok(wb?.Workbook?.Sheets?.[1]?.Hidden);
				assert.ok(wb?.Workbook?.Sheets?.[2]?.Hidden);
			});
		});
		it('should distinguish very hidden sheets', function() {
			wbs.forEach(function(wb) {
				assert.equal(wb?.Workbook?.Sheets?.[1]?.Hidden,1);
				assert.equal(wb?.Workbook?.Sheets?.[2]?.Hidden,2);
			});
		});
	});

	describe('comments', function() {
		if(fs.existsSync(paths.swcxlsx)) it('should have comment as part of cell properties', function(){
			var sheet = 'Sheet1';
			var wb1=X.read(fs.readFileSync(paths.swcxlsx), {type:TYPE});
			var wb2=X.read(fs.readFileSync(paths.swcxlsb), {type:TYPE});
			var wb3=X.read(fs.readFileSync(paths.swcxls), {type:TYPE});
			var wb4=X.read(fs.readFileSync(paths.swcxml), {type:TYPE});

			[wb1,wb2,wb3,wb4].map(function(wb: X.WorkBook) { return wb.Sheets[sheet]; }).forEach(function(ws, i) {
				assert.equal(get_cell(ws, "B1").c.length, 1,"must have 1 comment");
				assert.equal(get_cell(ws, "B1").c[0].a, "Yegor Kozlov","must have the same author");
				assert.equal(get_cell(ws, "B1").c[0].t, "Yegor Kozlov:\nfirst cell", "must have the concatenated texts");
				if(i > 0) return;
				assert.equal(get_cell(ws, "B1").c[0].r, '<r><rPr><b/><sz val="8"/><color indexed="81"/><rFont val="Tahoma"/></rPr><t>Yegor Kozlov:</t></r><r><rPr><sz val="8"/><color indexed="81"/><rFont val="Tahoma"/></rPr><t xml:space="preserve">\r\nfirst cell</t></r>', "must have the rich text representation");
				assert.equal(get_cell(ws, "B1").c[0].h, '<span style="font-size:8pt;"><b>Yegor Kozlov:</b></span><span style="font-size:8pt;"><br/>first cell</span>', "must have the html representation");
			});
		});
		var stst = [
			['xlsx', paths.cstxlsx],
			['xlsb', paths.cstxlsb],
			['xls', paths.cstxls],
			['xlml', paths.cstxml],
			['ods', paths.cstods]
		]; for(var ststi = 0; ststi < stst.length; ++ststi) { let m = stst[ststi]; it(m[0] + ' stress test', function() {
			var wb = X.read(fs.readFileSync(m[1]), {type:TYPE});
			check_comments(wb);
			var ws0 = wb.Sheets["Sheet2"];
			assert.equal(get_cell(ws0,"A1").c[0].a, 'Author');
			assert.equal(get_cell(ws0,"A1").c[0].t, 'Author:\nGod thinks this is good');
			assert.equal(get_cell(ws0,"C1").c[0].a, 'Author');
			assert.equal(get_cell(ws0,"C1").c[0].t, 'I really hope that xlsx decides not to use magic like rPr');
		}); }
	});

	describe('should parse core properties and custom properties', function() {
			var wbs = [
				X.read(fs.readFileSync(paths.cpxlsx), {type:TYPE, WTF:true}),
				X.read(fs.readFileSync(paths.cpxlsb), {type:TYPE, WTF:true}),
				X.read(fs.readFileSync(paths.cpxls), {type:TYPE, WTF:true}),
				X.read(fs.readFileSync(paths.cpxml), {type:TYPE, WTF:true})
			];

		var s1 = ['XLSX', 'XLSB', 'XLS', 'XML']; for(var i = 0; i < s1.length; ++i) { let x = s1[i];
			it(x + ' should parse core properties', function() { var P = wbs?.[i]?.Props; if(typeof P == "undefined") throw "missing props"; coreprop(P); });
			it(x + ' should parse custom properties', function() { custprop(wbs?.[i]?.Custprops); });
		}
		var s2 = [
			["asxls",  "BIFF8", "\u2603"],
			["asxls5", "BIFF5", "_"],
			["asxml",   "XLML", "\u2603"],
			["asods",    "ODS", "God"],
			["asxlsx",  "XLSX", "\u2603"],
			["asxlsb",  "XLSB", "\u2603"]
		]; for(var ii = 0; ii < s2.length; ++ii) { let x = s2[ii];
		it(x[1] + ' should read ' + (x[2] == "\u2603" ? 'unicode ' : "") + 'author', function() {
			var wb = X.read(fs.readFileSync(paths[x[0]]), {type:TYPE});
			assert.equal(wb?.Props?.Author, x[2]);
		}); }
		var BASE = "இராமா";
		/* TODO: ODS, XLS */
		var s3:Array<X.BookType> = [ "xlsx", "xlsb", "xlml"/*, "ods", "xls" */]; for(var i = 0; i < s3.length; ++i) { let n = s3[i];
		it(n + ' should round-trip unicode category', function() {
			var wb = X.utils.book_new();
			X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet([["a"]]), "Sheet1");
			if(!wb.Props) wb.Props = {};
			(wb.Props || (wb.Props = {})).Category = BASE;
			var wb2 = X.read(X.write(wb, {bookType:n, type:TYPE}), {type:TYPE});
			assert.equal(wb2?.Props?.Category,BASE);
		}); }
	});

	describe('sheetRows', function() {
		it('should use original range if not set', function() {
			var opts = {type:TYPE};
			FSTPaths.map(function(p) { return X.read(fs.readFileSync(p), opts); }).forEach(function(wb) {
				assert.equal(wb.Sheets["Text"]["!ref"],"A1:F49");
			});
		});
		it('should adjust range if set', function() {
			var opts = {type:TYPE, sheetRows:10};
			var wbs = FSTPaths.map(function(p) { return X.read(fs.readFileSync(p), opts); });
			/* TODO: XLS, XML, ODS */
			wbs.slice(0,2).forEach(function(wb) {
				assert.equal(wb.Sheets["Text"]["!fullref"],"A1:F49");
				assert.equal(wb.Sheets["Text"]["!ref"],"A1:F10");
			});
		});
		it('should not generate comment cells', function() {
			var opts = {type:TYPE, sheetRows:10};
			var wbs = CSTPaths.map(function(p) { return X.read(fs.readFileSync(p), opts); });
			/* TODO: XLS, XML, ODS */
			wbs.slice(0,2).forEach(function(wb) {
				assert.equal(wb.Sheets["Sheet7"]["!fullref"],"A1:N34");
				assert.equal(wb.Sheets["Sheet7"]["!ref"],"A1");
			});
		});
	});

	describe('column properties', function() {
		var wbs: X.WorkBook[] = [], wbs_no_slk: X.WorkBook[] = [];
			wbs = CWPaths.map(function(n) { return X.read(fs.readFileSync(n), {type:TYPE, cellStyles:true}); });
			wbs_no_slk = wbs.slice(0, 5);
		it('should have "!cols"', function() {
			wbs.forEach(function(wb) { assert.ok(wb.Sheets["Sheet1"]['!cols']); });
		});
		it('should have correct widths', function() {
			/* SYLK rounds wch so skip non-integral */
			wbs_no_slk.map(function(x) { return x.Sheets["Sheet1"]['!cols']; }).forEach(function(x) {
				assert.equal(x?.[1]?.width, 0.1640625);
				assert.equal(x?.[2]?.width, 16.6640625);
				assert.equal(x?.[3]?.width, 1.6640625);
			});
			wbs.map(function(x) { return x.Sheets["Sheet1"]['!cols']; }).forEach(function(x) {
				assert.equal(x?.[4]?.width, 4.83203125);
				assert.equal(x?.[5]?.width, 8.83203125);
				assert.equal(x?.[6]?.width, 12.83203125);
				assert.equal(x?.[7]?.width, 16.83203125);
			});
		});
		it('should have correct pixels', function() {
			/* SYLK rounds wch so skip non-integral */
			wbs_no_slk.map(function(x) { return x.Sheets["Sheet1"]['!cols']; }).forEach(function(x) {
				assert.equal(x?.[1].wpx, 1);
				assert.equal(x?.[2].wpx, 100);
				assert.equal(x?.[3].wpx, 10);
			});
			wbs.map(function(x) { return x.Sheets["Sheet1"]['!cols']; }).forEach(function(x) {
				assert.equal(x?.[4].wpx, 29);
				assert.equal(x?.[5].wpx, 53);
				assert.equal(x?.[6].wpx, 77);
				assert.equal(x?.[7].wpx, 101);
			});
		});
	});

	describe('row properties', function() {
			var wbs = RHPaths.map(function(p) { return X.read(fs.readFileSync(p), {type:TYPE, cellStyles:true}); });
			/* */
			var ols = OLPaths.map(function(p) { return X.read(fs.readFileSync(p), {type:TYPE, cellStyles:true}); });
		it('should have "!rows"', function() {
			wbs.forEach(function(wb) { assert.ok(wb.Sheets["Sheet1"]['!rows']); });
		});
		it('should have correct points', function() {
			wbs.map(function(x) { return x.Sheets["Sheet1"]['!rows']; }).forEach(function(x) {
				assert.equal(x?.[1].hpt, 1);
				assert.equal(x?.[2].hpt, 10);
				assert.equal(x?.[3].hpt, 100);
			});
		});
		it('should have correct pixels', function() {
			wbs.map(function(x) { return x.Sheets["Sheet1"]['!rows']; }).forEach(function(x) {
				/* note: at 96 PPI hpt == hpx */
				assert.equal(x?.[1].hpx, 1);
				assert.equal(x?.[2].hpx, 10);
				assert.equal(x?.[3].hpx, 100);
			});
		});
		it('should have correct outline levels', function() {
			ols.map(function(x) { return x.Sheets["Sheet1"]; }).forEach(function(ws) {
				var rows = ws['!rows'];
				for(var i = 0; i < 29; ++i) {
					var cell = get_cell(ws, "A" + X.utils.encode_row(i));
					var lvl = (rows?.[i]||{}).level||0;
					if(!cell || cell.t == 's') assert.equal(lvl, 0);
					else if(cell.t == 'n') {
						if(cell.v === 0) assert.equal(lvl, 0);
						else assert.equal(lvl, cell.v);
					}
				}
				assert.equal(rows?.[29]?.level, 7);
			});
		});
	});

	describe('merge cells',function() {
			var wbs = MCPaths.map(function(p) { return X.read(fs.readFileSync(p), {type:TYPE}); });
		it('should have !merges', function() {
			wbs.forEach(function(wb) {
				assert.ok(wb.Sheets["Merge"]['!merges']);
			});
			var m = wbs.map(function(x) { return x.Sheets["Merge"]?.['!merges']?.map(function(y) { return X.utils.encode_range(y); });});
			m?.slice(1)?.forEach(function(x) {
				assert.deepEqual(m?.[0]?.sort(),x?.sort());
			});
		});
	});

	describe('should find hyperlinks', function() {
			var wb1 = HLPaths.map(function(p) { return X.read(fs.readFileSync(p), {type:TYPE, WTF:true}); });
			var wb2 = ILPaths.map(function(p) { return X.read(fs.readFileSync(p), {type:TYPE, WTF:true}); });

		var ext1 = ['xlsx', 'xlsb', 'xls', 'xml']; for(let i = 0; i < ext1.length; ++i) { let x = ext1[i];
			it(x + " external", function() { hlink1(wb1[i].Sheets["Sheet1"]); });
		}
		var ext2 = ['xlsx', 'xlsb', 'xls', 'xml', 'ods']; for(let i = 0; i < ext2.length; ++i) { let x = ext2[i];
			it(x + " internal", function() { hlink2(wb2[i].Sheets["Sheet1"]); });
		}
	});

	describe('should parse cells with date type (XLSX/XLSM)', function() {
		it('Must have read the date', function() {
			var wb: X.WorkBook, ws: X.WorkSheet;
			var sheetName = 'Sheet1';
			wb = X.read(fs.readFileSync(paths.dtxlsx), {type:TYPE});
			ws = wb.Sheets[sheetName];
			var sheet: Array<any> = X.utils.sheet_to_json(ws, {raw: false});
			assert.equal(sheet[3]['てすと'], '2/14/14');
		});
		it('cellDates should not affect formatted text', function() {
			var sheetName = 'Sheet1';
			var ws1 = X.read(fs.readFileSync(paths.dtxlsx), {type:TYPE}).Sheets[sheetName];
			var ws2 = X.read(fs.readFileSync(paths.dtxlsb), {type:TYPE}).Sheets[sheetName];
			assert.equal(X.utils.sheet_to_csv(ws1),X.utils.sheet_to_csv(ws2));
		});
	});

	describe('cellDates', function() {
		var fmts = [
			/* desc     path        sheet     cell   formatted */
			['XLSX', paths.dtxlsx, 'Sheet1',  'B5',  '2/14/14'],
			['XLSB', paths.dtxlsb, 'Sheet1',  'B5',  '2/14/14'],
			['XLS',  paths.dtxls,  'Sheet1',  'B5',  '2/14/14'],
			['XLML', paths.dtxml,  'Sheet1',  'B5',  '2/14/14'],
			['XLSM', paths.nfxlsx, 'Implied', 'B13', '18-Oct-33']
		];
		it('should not generate date cells by default', function() { fmts.forEach(function(f) {
			var wb, ws;
			wb = X.read(fs.readFileSync(f[1]), {type:TYPE});
			ws = wb.Sheets[f[2]];
			assert.equal(get_cell(ws, f[3]).w, f[4]);
			assert.equal(get_cell(ws, f[3]).t, 'n');
		}); });
		it('should generate date cells if cellDates is true', function() { fmts.forEach(function(f) {
			var wb, ws;
			wb = X.read(fs.readFileSync(f[1]), {type:TYPE, cellDates:true});
			ws = wb.Sheets[f[2]];
			assert.equal(get_cell(ws, f[3]).w, f[4]);
			assert.equal(get_cell(ws, f[3]).t, 'd');
		}); });
	});

	describe('defined names', function() {var dnp = [
		/* desc     path        cmnt */
		['xlsx', paths.dnsxlsx,  true],
		['xlsb', paths.dnsxlsb,  true],
		['xls',  paths.dnsxls,   true],
		['xlml', paths.dnsxml,  false],
		['slk',  paths.dnsslk,  false]
	] as Array<[string, string, boolean]>; for(var i = 0; i < dnp.length; ++i) { let m: [string, string, boolean] = dnp[i]; it(m[0], function() {
		var wb = X.read(fs.readFileSync(m[1]), {type:TYPE});
		var names = wb?.Workbook?.Names;

		if(names) {
		if(m[0] != 'slk') {
		for(var i = 0; i < names?.length; ++i) if(names[i].Name == "SheetJS") break;
		assert.ok(i < names?.length, "Missing name");
		assert.equal(names[i].Sheet, void 0);
		assert.equal(names[i].Ref, "Sheet1!$A$1");
		if(m[2]) assert.equal(names[i].Comment, "defined names just suck  excel formulae are bad  MS should feel bad");
		}

		for(i = 0; i < names.length; ++i) if(names[i].Name == "SHEETjs") break;
		assert.ok(i < names.length, "Missing name");
		assert.equal(names[i].Sheet, 0);
		assert.equal(names[i].Ref, "Sheet1!$A$2");
		}
	}); } });

	describe('defined names unicode', function() {var dnu=[
		/* desc     path          RT */
		['xlsx', paths.dnuxlsx,  true],
		['xlsb', paths.dnuxlsb,  true],
		['ods',  paths.dnuods,   true],
		['xls',  paths.dnuxls,  false],
		['xlml', paths.dnuxml,  false]
	]; for(var i = 0; i < dnu.length; ++i) { let m = dnu[i]; it(m[0], function() {
		var wb = X.read(fs.readFileSync(m[1]), {type:TYPE});
		var wb2 = X.read(X.write(wb, {type:TYPE, bookType: m[0]}), {type:TYPE});
		[
			"NoContainsJapanese",
			"\u65E5\u672C\u8a9e\u306e\u307f",
			"sheet\u65e5\u672c\u8a9e",
			"\u65e5\u672c\u8a9esheet",
			"sheet\u65e5\u672c\u8a9esheet"
		].forEach(function(n, i) { assert.equal(wb.SheetNames[i], n); });
		[
			["name\u65e5\u672c\u8a9e", "sheet\u65e5\u672c\u8a9e!$A$1"],
			["name\u65e5\u672c\u8a9ename", "sheet\u65e5\u672c\u8a9esheet!$B$2"],
			["NoContainsJapaneseName", "\u65e5\u672c\u8a9e\u306e\u307f!$A$1"],
			["sheet\u65e5\u672c\u8a9e", "sheet\u65e5\u672c\u8a9e!$A$1"],
			["\u65e5\u672c\u8a9e", "NoContainsJapanese!$A$1"],
			["\u65e5\u672c\u8a9ename", "\u65e5\u672c\u8a9esheet!$I$2"]
		].forEach(function(n) {(m[2] ? [ wb, wb2 ] : [ wb ]).forEach(function(wb) {
			var DN = null;
			var arr = wb?.Workbook?.Names;
			if(arr) for(var j = 0; j < arr.length; ++j) if(arr[j]?.Name == n?.[0]) DN = arr[j];
			assert.ok(DN);
			// $FlowIgnore
			assert.equal(DN?.Ref, n?.[1]);
		}); });
	}); } });

	describe('workbook codename unicode', function() {
			var wb = X.utils.book_new();
			var ws = X.utils.aoa_to_sheet([[1]]);
			X.utils.book_append_sheet(wb, ws, "Sheet1");
			wb.Workbook = { WBProps: { CodeName: "本工作簿" } };
		var exts = ['xlsx', 'xlsb'] as Array<X.BookType>; for(var i = 0; i < exts.length; ++i) { var m = exts[i]; it(m, function() {
			var bstr = X.write(wb, {type: "binary", bookType: m});
			var nwb = X.read(bstr, {type: "binary"});
			assert.equal(nwb?.Workbook?.WBProps?.CodeName, wb?.Workbook?.WBProps?.CodeName);
		}); }
	});

	describe('auto filter', function() {var af = [
		['xlsx', paths.afxlsx],
		['xlsb', paths.afxlsb],
		['xls',  paths.afxls],
		['xlml', paths.afxml],
		['ods',  paths.afods]
	]; for(var j = 0; j < af.length; ++j) { var m = af[j]; it(m[0], function() {
		var wb = X.read(fs.readFileSync(m[1]), {type:TYPE});
		assert.ok(!wb.Sheets[wb.SheetNames[0]]['!autofilter']);
		for(var i = 1; i < wb.SheetNames.length; ++i) {
			assert.ok(wb.Sheets[wb.SheetNames[i]]['!autofilter']);
			assert.equal(wb.Sheets[wb.SheetNames[i]]?.['!autofilter']?.ref,"A1:E22");
		}
	}); } });

	describe('HTML', function() {
			var ws = X.utils.aoa_to_sheet([
				["a","b","c"],
				["&","<",">","\n"]
			]);
			var wb = {SheetNames:["Sheet1"],Sheets:{Sheet1:ws}};
		var m: 'xlsx' = 'xlsx'; it(m, function() {
			var wb2 = X.read(X.write(wb, {bookType:m, type:TYPE}),{type:TYPE, cellHTML:true});
			assert.equal(get_cell(wb2.Sheets["Sheet1"], "A2").h, "&amp;");
			assert.equal(get_cell(wb2.Sheets["Sheet1"], "B2").h, "&lt;");
			assert.equal(get_cell(wb2.Sheets["Sheet1"], "C2").h, "&gt;");
			var h = get_cell(wb2.Sheets["Sheet1"], "D2").h;
			assert.ok(h == "&#x000a;" || h == "<br/>");
		});
	});

	describe('page margins', function() {
			var wbs = PMPaths.map(function(p) { return X.read(fs.readFileSync(p), {type:TYPE, WTF:true}); });
		var pm = ([
			/* Sheet Name     Margins: left   right  top bottom head foot */
			["Normal",                 [0.70, 0.70, 0.75, 0.75, 0.30, 0.30]],
			["Wide",                   [1.00, 1.00, 1.00, 1.00, 0.50, 0.50]],
			["Narrow",                 [0.25, 0.25, 0.75, 0.75, 0.30, 0.30]],
			["Custom 1 Inch Centered", [1.00, 1.00, 1.00, 1.00, 0.30, 0.30]],
			["1 Inch HF",              [0.70, 0.70, 0.75, 0.75, 1.00, 1.00]]
		] as Array<[string, number[]]>); for(var pmi = 0; pmi < pm.length; ++pmi) { let tt = pm[pmi]; it('should parse ' + tt[0] + ' margin', function() { wbs.forEach(function(wb) {
			check_margin(wb.Sheets[""+tt[0]]?.["!margins"] as any, tt[1]);
		}); }); }
	});

	describe('should correctly handle styles', function() {
			var wsxls=X.read(fs.readFileSync(paths.cssxls), {type:TYPE,cellStyles:true,WTF:true}).Sheets["Sheet1"];
			var wsxlsx=X.read(fs.readFileSync(paths.cssxlsx), {type:TYPE,cellStyles:true,WTF:true}).Sheets["Sheet1"];
			var rn = function(range: string): string[]  {
				var r = X.utils.decode_range(range);
				var out: string[] = [];
				for(var R = r.s.r; R <= r.e.r; ++R) for(var C = r.s.c; C <= r.e.c; ++C)
					out.push(X.utils.encode_cell({c:C,r:R}));
				return out;
			};
			var rn2 = function(r: string): string[] { return ([] as string[]).concat.apply(([] as string[]), r.split(",").map(rn)); };
		var ranges = [
			'A1:D1,F1:G1', 'A2:D2,F2:G2', /* rows */
			'A3:A10', 'B3:B10', 'E1:E10', 'F6:F8', /* cols */
			'H1:J4', 'H10' /* blocks */
		];
		/*eslint-disable */
		var exp/*:Array<any>*/ = [
			{ patternType: 'darkHorizontal',
			  fgColor: { theme: 9, raw_rgb: 'F79646' },
			  bgColor: { theme: 5, raw_rgb: 'C0504D' } },
			{ patternType: 'darkUp',
			  fgColor: { theme: 3, raw_rgb: '1F497D' },
			  bgColor: { theme: 7, raw_rgb: '8064A2' } },
			{ patternType: 'darkGray',
			  fgColor: { theme: 3, raw_rgb: '1F497D' },
			  bgColor: { theme: 1, raw_rgb: '000000' } },
			{ patternType: 'lightGray',
			  fgColor: { theme: 6, raw_rgb: '9BBB59' },
			  bgColor: { theme: 2, raw_rgb: 'EEECE1' } },
			{ patternType: 'lightDown',
			  fgColor: { theme: 4, raw_rgb: '4F81BD' },
			  bgColor: { theme: 7, raw_rgb: '8064A2' } },
			{ patternType: 'lightGrid',
			  fgColor: { theme: 6, raw_rgb: '9BBB59' },
			  bgColor: { theme: 9, raw_rgb: 'F79646' } },
			{ patternType: 'lightGrid',
			  fgColor: { theme: 4, raw_rgb: '4F81BD' },
			  bgColor: { theme: 2, raw_rgb: 'EEECE1' } },
			{ patternType: 'lightVertical',
			  fgColor: { theme: 3, raw_rgb: '1F497D' },
			  bgColor: { theme: 7, raw_rgb: '8064A2' } }
		];
		/*eslint-enable */
		for(var rangei = 0; rangei < ranges.length; ++rangei) { let rng = ranges[rangei];
			it('XLS  | ' + rng,function(){cmparr(rn2(rng).map(function(x){ return get_cell(wsxls,x).s; }));});
			it('XLSX | ' + rng,function(){cmparr(rn2(rng).map(function(x){ return get_cell(wsxlsx,x).s; }));});
		}
		it('different styles', function() {
			for(var i = 0; i != ranges.length-1; ++i) {
				for(var j = i+1; j != ranges.length; ++j) {
					diffsty(wsxlsx, rn2(ranges[i])[0], rn2(ranges[j])[0]);
					/* TODO: XLS */
					//diffsty(wsxls, rn2(ranges[i])[0], rn2(ranges[j])[0]);
				}
			}
		});
		it('correct styles', function() {
			//var stylesxls = ranges.map(function(r) { return rn2(r)[0]; }).map(function(r) { return get_cell(wsxls,r).s; });
			var stylesxlsx = ranges.map(function(r) { return rn2(r)[0]; }).map(function(r) { return get_cell(wsxlsx,r).s; });
			exp.forEach(function(e, i) {
				[
					"fgColor.theme","fgColor.raw_rgb",
					"bgColor.theme","bgColor.raw_rgb",
					"patternType"
				].forEach(function(k) {
					deepcmp(e, stylesxlsx[i], k, i + ":" + k, 0);
					/* TODO: XLS */
					//deepcmp(e, stylesxls[i], k, i + ":" + k, 0);
				});
			});
		});
	});

	describe('data types formats', function() {var dtf = [
		['xlsx', paths.dtfxlsx],
	]; for(var j = 0; j < dtf.length; ++j) { var m = dtf[j]; it(m[0], function() {
		var wb = X.read(fs.readFileSync(m[1]), {type: TYPE, cellDates: true});
		var ws = wb.Sheets[wb.SheetNames[0]];
		var data = X.utils.sheet_to_json<any>(ws, { header: 1, raw: true, rawNumbers: false });
		assert.ok(data[0][1] instanceof Date);
		assert.ok(data[1][1] instanceof Date);
		assert.equal(data[2][1], '$123.00');
		assert.equal(data[3][1], '98.76%');
		assert.equal(data[4][1], '456.00');
		assert.equal(data[5][1], '7,890');
	}); } });

	it('date system', function() {[
		"biff5", "ods", "slk", "xls", "xlsb", "xlsx", "xml"
	].forEach(function(ext) {
		// TODO: verify actual date values
		var wb0 = X.read(fs.readFileSync("./test_files/1904/1900." + ext), {type: TYPE, cellNF: true});
		assert.ok(!(wb0?.Workbook?.WBProps?.date1904));
		assert.equal(X.utils.sheet_to_csv(wb0.Sheets[wb0.SheetNames[0]]), [
"1,1900-01-01,1900-01-01,1900-01-01",
"11,1900-01-11,1900-01-11,1900-01-11",
"111,1900-04-20,1900-04-20,1900-04-20",
"1111,1903-01-15,1903-01-15,1903-01-15",
"11111,1930-06-02,1930-06-02,1930-06-02"
		].join("\n"));

		var wb4 = X.read(fs.readFileSync("./test_files/1904/1904." + ext), {type: TYPE, cellNF: true});
		assert.ok(wb4?.Workbook?.WBProps?.date1904);
		assert.equal(X.utils.sheet_to_csv(wb4.Sheets[wb4.SheetNames[0]]), [
"1,1904-01-02,1904-01-02,1904-01-02",
"11,1904-01-12,1904-01-12,1904-01-12",
"111,1904-04-21,1904-04-21,1904-04-21",
"1111,1907-01-16,1907-01-16,1907-01-16",
"11111,1934-06-03,1934-06-03,1934-06-03"
		].join("\n"));
	}); });

	it('bookType metadata', function() {
	([
		// TODO: keep in sync with BookType, support other formats
		"xlsx"/*, "xlsm" */, "xlsb"/* xls / xla / biff# */, "xlml", "ods", "fods"/*, "csv", "txt", */, "sylk", "html", "dif", "rtf"/*, "prn", "eth"*/, "dbf", "numbers"
	] as X.BookType[]).forEach(function(r: X.BookType) {
		if(r == "numbers" && !can_write_numbers) return;
		var ws = X.utils.aoa_to_sheet([ ["a", "b", "c"], [1, 2, 3] ]);
		var wb = X.utils.book_new(); X.utils.book_append_sheet(wb, ws, "Sheet1");
		var data = X.write(wb, {type: TYPE, bookType: r, WTF: true, numbers:XLSX_ZAHL });
		assert.equal(X.read(data, {type: TYPE, WTF: true}).bookType, r);
	}); });
});

describe('write features', function() {
	describe('props', function() {
		describe('core', function() {
			var baseprops: X.FullProperties = {
				Category: "Newspaper",
				ContentStatus: "Published",
				Keywords: "☃",
				LastAuthor: "Perry White",
				LastPrinted: "1978-12-15",
				//RevNumber: 6969, // TODO: this should actually be Revision
				//AppVersion: 69, // TODO: this should actually be a number
				Author: "Lois Lane",
				Comments: "Needs work",
				Identifier: "1d",
				Language: "English",
				Subject: "Superman",
				Title: "Man of Steel"
			};
				var ws = X.utils.aoa_to_sheet([["a","b","c"],[1,2,3]]);
			var bt1 = (['xlml', 'xlsx', 'xlsb'] as Array<X.BookType>); for(var bti = 0; bti < bt1.length; ++bti) { let w: X.BookType = bt1[bti]; it(w, function() {
				var wb: X.WorkBook = {
					Props: ({} as X.FullProperties),
					SheetNames: ["Sheet1"],
					Sheets: {Sheet1: ws}
				};
				(Object.keys(baseprops) as Array<keyof X.FullProperties>).forEach(function(k: keyof X.FullProperties) { if(wb.Props) (wb.Props as any)[k] = baseprops[k]; });
				var wb2 = X.read(X.write(wb, {bookType:w, type:TYPE}), {type:TYPE});
				(Object.keys(baseprops) as Array<keyof X.FullProperties>).forEach(function(k) { assert.equal(baseprops[k], wb2?.Props?.[k]); });
				var wb3 = X.read(X.write(wb2, {bookType:w, type:TYPE, Props: {Author:"SheetJS"}}), {type:TYPE});
				assert.equal("SheetJS", wb3?.Props?.Author);
			}); }
		});
	});
	describe('HTML', function() {
		it('should use `h` value when present', function() {
			var sheet = X.utils.aoa_to_sheet([["abc"]]);
			get_cell(sheet, "A1").h = "<b>abc</b>";
			var wb = {SheetNames:["Sheet1"], Sheets:{Sheet1:sheet}};
			var str = X.write(wb, {bookType:"html", type:"binary"});
			assert.ok(str.indexOf("<b>abc</b>") > 0);
		});
	});
	describe('sheet range limits', function() { var b = ([
		["biff2", "IV16384"],
		["biff5", "IV16384"],
		["biff8", "IV65536"],
		["xlsx", "XFD1048576"],
		["xlsb", "XFD1048576"]
	] as Array<['biff2' | 'biff5' | 'biff8' | 'xlsx' | 'xlsb', string]>); for(var j = 0; j < b.length; ++j) { var r = b[j]; it(r[0], function() {
		var C = X.utils.decode_cell(r[1]);
		var wopts: X.WritingOptions = {bookType:r[0], type:'binary', WTF:true};
		var wb = { SheetNames: ["Sheet1"], Sheets: { Sheet1: ({} as X.WorkSheet) } };

		wb.Sheets["Sheet1"]['!ref'] =  "A1:" + X.utils.encode_cell({r:0, c:C.c});
		X.write(wb, wopts);
		wb.Sheets["Sheet1"]['!ref'] =  "A" + X.utils.encode_row(C.r - 5) + ":" + X.utils.encode_cell({r:C.r, c:0});
		X.write(wb, wopts);

		wb.Sheets["Sheet1"]['!ref'] =  "A1:" + X.utils.encode_cell({r:0, c:C.c+1});
		assert.throws(function() { X.write(wb, wopts); });
		wb.Sheets["Sheet1"]['!ref'] =  "A" + X.utils.encode_row(C.r - 5) + ":" + X.utils.encode_cell({r:C.r+1, c:0});
		assert.throws(function() { X.write(wb, wopts); });
	}); } });
	it('single worksheet formats', function() {
		var wb = X.utils.book_new();
		X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet([[1,2],[3,4]]), "Sheet1");
		X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet([[5,6],[7,8]]), "Sheet2");
		assert.equal(X.write(wb, {type:"string", bookType:"csv", sheet:"Sheet1"}), "1,2\n3,4");
		assert.equal(X.write(wb, {type:"string", bookType:"csv", sheet:"Sheet2"}), "5,6\n7,8");
		assert.throws(function() { X.write(wb, {type:"string", bookType:"csv", sheet:"Sheet3"}); });
	});
	it('should create/update autofilter defined name on write', function() {([
		"xlsx", "xlsb", /* "xls", */ "xlml" /*, "ods" */
	] as Array<X.BookType>).forEach(function(fmt) {
		var wb = X.utils.book_new();
		var ws = X.utils.aoa_to_sheet([["A","B","C"],[1,2,3]]);
		ws["!autofilter"] = { ref: (ws["!ref"] as string) };
		X.utils.book_append_sheet(wb, ws, "Sheet1");
		var wb2 = X.read(X.write(wb, {bookType:fmt, type:TYPE}), {type:TYPE});
		var Name: X.DefinedName = (null as any);
		((wb2.Workbook||{}).Names || []).forEach(function(dn) { if(dn.Name == "_xlnm._FilterDatabase" && dn.Sheet == 0) Name = dn; });
		assert.ok(!!Name, "Could not find _xlnm._FilterDatabases name WR");
		assert.equal(Name.Ref, "Sheet1!$A$1:$C$2");
		X.utils.sheet_add_aoa(wb2.Sheets["Sheet1"], [[4,5,6]], { origin: -1 });
		(wb2.Sheets["Sheet1"]["!autofilter"] as any).ref = wb2.Sheets["Sheet1"]["!ref"];
		var wb3 = X.read(X.write(wb2, {bookType:fmt, type:TYPE}), {type:TYPE});
		Name = (null as any);
		((wb3.Workbook||{}).Names || []).forEach(function(dn) { if(dn.Name == "_xlnm._FilterDatabase" && dn.Sheet == 0) Name = dn; });
		assert.ok(!!Name, "Could not find _xlnm._FilterDatabases name WRWR");
		assert.equal(Name.Ref, "Sheet1!$A$1:$C$3");
		assert.equal(((wb2.Workbook as any).Names as any).length, ((wb3.Workbook as any).Names as any).length);
	}); });
});

function seq(end: number, start?: number): Array<number> {
	var s = start || 0;
	var o = new Array(end - s);
	for(var i = 0; i != o.length; ++i) o[i] = s + i;
	return o;
}

var basedate = new Date(1899, 11, 30, 0, 0, 0); // 2209161600000
function datenum(v: Date, date1904?: boolean): number {
	var epoch = v.getTime();
	if(date1904) epoch -= 1462*24*60*60*1000;
	var dnthresh = basedate.getTime() + (v.getTimezoneOffset() - basedate.getTimezoneOffset()) * 60000;
	return (epoch - dnthresh) / (24 * 60 * 60 * 1000);
}
var good_pd_date = new Date('2017-02-19T19:06:09.000Z');
if(isNaN(good_pd_date.getFullYear())) good_pd_date = new Date('2017-02-19T19:06:09');
if(isNaN(good_pd_date.getFullYear())) good_pd_date = new Date('2/19/17');
var good_pd = good_pd_date.getFullYear() == 2017;
function parseDate(str: string|Date): Date {
	var d = new Date(str);
	if(good_pd) return d;
	if(str instanceof Date) return str;
	if(good_pd_date.getFullYear() == 1917 && !isNaN(d.getFullYear())) {
		var s = d.getFullYear();
		if(str.indexOf("" + s) > -1) return d;
		d.setFullYear(d.getFullYear() + 100); return d;
	}
	var n = str.match(/\d+/g)||["2017","2","19","0","0","0"];
	return new Date(Date.UTC(+n[0], +n[1] - 1, +n[2], +n[3], +n[4], +n[5]));
}

var fixdate = browser ? parseDate("2014-02-19T14:30:00.000Z") : new Date("2014-02-19T14:30Z");

describe('roundtrip features', function() {
	describe('should preserve core properties', function() { var cp = [
		['xls', paths.cpxls],
		['xlml', paths.cpxml],
		['xlsx', paths.cpxlsx],
		['xlsb', paths.cpxlsb]
	]; for(var cpi = 0; cpi < cp.length; ++cpi) { var w = cp[cpi];
		it(w[0], function() {
			var wb1 = X.read(fs.readFileSync(w[1]), {type:TYPE});
			coreprop(wb1.Props as any);
			var wb2 = X.read(X.write(wb1, {bookType:w[0], type:TYPE}), {type:TYPE});
			coreprop(wb2.Props as any);
		});
	} });

	describe('should preserve custom properties', function() { var cp = [
		['xls', paths.cpxls],
		['xlml', paths.cpxml],
		['xlsx', paths.cpxlsx],
		['xlsb', paths.cpxlsb]
	]; for(var cpi = 0; cpi < cp.length; ++cpi) { var w = cp[cpi];
		it(w[0], function() {
			var wb1 = X.read(fs.readFileSync(w[1]), {type:TYPE});
			custprop(wb1.Custprops);
			var wb2 = X.read(X.write(wb1, {bookType:w[0], type:TYPE}), {type:TYPE});
			custprop(wb2.Custprops);
		});
	} });

	describe('should preserve merge cells', function() {
		var mcf = ["xlsx", "xlsb", "xlml", "ods", "biff8", "numbers"] as Array<X.BookType>; for(let mci = 0; mci < mcf.length; ++mci) { let f = mcf[mci]; it(f, function() {
			if(f == "numbers" && !can_write_numbers) return;
			var wb1 = X.read(fs.readFileSync(paths.mcxlsx), {type:TYPE});
			var wb2 = X.read(X.write(wb1,{bookType:f,type:'binary',numbers:XLSX_ZAHL}),{type:'binary'});
			var m1 = wb1.Sheets["Merge"]?.['!merges']?.map(X.utils.encode_range);
			var m2 = wb2.Sheets["Merge"]?.['!merges']?.map(X.utils.encode_range);
			assert.equal(m1?.length, m2?.length);
			if(m1 && m2) for(var i = 0; i < m1?.length; ++i) assert.ok(m1?.indexOf(m2?.[i]) > -1);
		}); }
	});

	describe('should preserve dates', function() {
		for(var n = 0; n < 16; ++n) {
			var d = (n & 1) ? 'd' : 'n', dk = d === 'd';
			var c = (n & 2) ? 'd' : 'n', dj = c === 'd';
			var b = (n & 4) ? 'd' : 'n', di = b === 'd';
			var a = (n & 8) ? 'd' : 'n', dh = a === 'd';
			var f: string, sheet: string, addr: string;
			if(dh) { f = paths.dtxlsx; sheet = 'Sheet1'; addr = 'B5'; }
			else { f = paths.nfxlsx; sheet = '2011'; addr = 'J36'; }
			it('[' + a + '] -> (' + b + ') -> [' + c + '] -> (' + d + ')', function() {
				var wb1 = X.read(fs.readFileSync(f), {type:TYPE, cellNF: true, cellDates: di, WTF: opts.WTF});
				var  _f = X.write(wb1, {type:'binary', cellDates:dj, WTF:opts.WTF});
				var wb2 = X.read(_f, {type:'binary', cellDates: dk, WTF: opts.WTF});
				var m = [wb1,wb2].map(function(x: X.WorkBook) { return get_cell(x.Sheets[sheet], addr); });
				assert.equal(m[0].w, m[1].w);

				assert.equal(m[0].t, b);
				assert.equal(m[1].t, d);

				if(m[0].t === 'n' && m[1].t === 'n') assert.equal(m[0].v, m[1].v);
				else if(m[0].t === 'd' && m[1].t === 'd') assert.equal(m[0].v.toString(), m[1].v.toString());
				else if(m[1].t === 'n') assert.ok(Math.abs(datenum(browser ? parseDate(m[0].v) : new Date(m[0].v)) - m[1].v) < 0.01);
			});
		}
	});

	describe('should preserve formulae', function() { var ff = [
		['xlml', paths.fstxml],
		['xlsx', paths.fstxlsx],
		['ods',  paths.fstods]
	]; for(let ii = 0; ii < ff.length; ++ii) { let w = ff[ii]; it(w[0], function() {
		var wb1 = X.read(fs.readFileSync(w[1]), {type:TYPE, cellFormula:true, WTF:true});
		var wb2 = X.read(X.write(wb1, {bookType:w[0], type:TYPE}), {cellFormula:true, type:TYPE, WTF:true});
		wb1.SheetNames.forEach(function(n) {
			assert.equal(
				X.utils.sheet_to_formulae(wb1.Sheets[n]).sort().join("\n"),
				X.utils.sheet_to_formulae(wb2.Sheets[n]).sort().join("\n")
			);
		});
	}); } });

	describe('should preserve dynamic array formulae', function() { var m19 = [
		['xlsx', paths.m19xlsx]
	]; for(var h1 = 0; h1 < m19.length; ++h1) { var w = m19[h1]; it(w[0], function() {
		var wb1 = X.read(fs.readFileSync(w[1]), {xlfn: true, type:TYPE, cellFormula:true, WTF:true});
		var wb2 = X.read(X.write(wb1, {bookType:w[0], type:TYPE}), {cellFormula:true, xlfn: true, type:TYPE, WTF:true});
		assert.equal(!!get_cell(wb2.Sheets["Sheet1"], "B3").D, true);
		assert.equal(!!get_cell(wb2.Sheets["Sheet1"], "B13").D, true);
		assert.equal(!!get_cell(wb2.Sheets["Sheet1"], "C13").D, true);

		get_cell(wb2.Sheets["Sheet1"], "B3").D = false;
		var wb3 = X.read(X.write(wb2, {bookType:w[0], type:TYPE}), {cellFormula:true, xlfn: true, type:TYPE, WTF:true});
		assert.equal(!!get_cell(wb3.Sheets["Sheet1"], "B3").D, false);
		assert.equal(!!get_cell(wb3.Sheets["Sheet1"], "B13").D, true);
		assert.equal(!!get_cell(wb3.Sheets["Sheet1"], "C13").D, true);
	}); } });

	describe('should preserve hyperlink', function() { var hl = [
		['xlml', paths.hlxml,   true],
		['xls',  paths.hlxls,   true],
		['xlsx', paths.hlxlsx,  true],
		['xlsb', paths.hlxlsb,  true],
		['xlml', paths.ilxml,  false],
		['xls',  paths.ilxls,  false],
		['xlsx', paths.ilxlsx, false],
		['xlsb', paths.ilxlsb, false],
		['ods',  paths.ilods,  false]
	]; for(var h1 = 0; h1 < hl.length; ++h1) { var w = hl[h1]; it(w[0]+" "+(w[2]?"ex":"in")+ "ternal", function() {
		var wb = X.read(fs.readFileSync(w[1]), {type:TYPE, WTF:opts.WTF});
		var hlink = (w[2] ? hlink1 : hlink2); hlink(wb.Sheets["Sheet1"]);
		wb = X.read(X.write(wb, {bookType:w[0], type:TYPE, WTF:opts.WTF}), {type:TYPE, WTF:opts.WTF});
		hlink(wb.Sheets["Sheet1"]);
	}); } });

	describe('should preserve page margins', function() { var pm = [
			['xlml', paths.pmxml],
			['xlsx', paths.pmxlsx],
			['xlsb', paths.pmxlsb]
		]; for(var p1 = 0; p1 < pm.length; ++p1) {var w = pm[p1]; it(w[0], function() {
			var wb1 = X.read(fs.readFileSync(w[1]), {type:TYPE});
			var wb2 = X.read(X.write(wb1, {bookType:w[0], type:"binary"}), {type:"binary"});
			([
				/* Sheet Name     Margins: left   right  top bottom head foot */
				["Normal",                 [0.70, 0.70, 0.75, 0.75, 0.30, 0.30]],
				["Wide",                   [1.00, 1.00, 1.00, 1.00, 0.50, 0.50]],
				["Narrow",                 [0.25, 0.25, 0.75, 0.75, 0.30, 0.30]],
				["Custom 1 Inch Centered", [1.00, 1.00, 1.00, 1.00, 0.30, 0.30]],
				["1 Inch HF",              [0.70, 0.70, 0.75, 0.75, 1.00, 1.00]]
			] as Array<[string, number[]]>).forEach(function(t) {
				check_margin(wb2.Sheets[t[0]]?.["!margins"] as X.MarginInfo, t[1]);
			});
	}); } });

	describe('should preserve sheet visibility', function() { var sv = [
			['xlml', paths.svxml],
			['xlsx', paths.svxlsx],
			['xlsb', paths.svxlsb],
			['xls', paths.svxls],
			['biff5', paths.svxls5]
			// ['ods', paths.svods]
		] as Array<[X.BookType, string]>; for(var s1 = 0; s1 < sv.length; ++s1) { var w = sv[s1];
			it(w[0], function() {
				var wb1 = X.read(fs.readFileSync(w[1]), {type:TYPE});
				var wb2 = X.read(X.write(wb1, {bookType:w[0], type:TYPE}), {type:TYPE});
				var wbs1 = wb1?.Workbook?.Sheets;
				var wbs2 = wb2?.Workbook?.Sheets;
				assert.equal(wbs1?.length, wbs2?.length);
				if(wbs1) for(var i = 0; i < wbs1?.length; ++i) {
					assert.equal(wbs1?.[i].name, wbs2?.[i].name);
					assert.equal(wbs1?.[i].Hidden, wbs2?.[i].Hidden);
				}
			});
		}
	});

	describe('should preserve column properties', function() { var bt = [
			/*'xlml',*/ /*'biff2', 'biff8', */ 'xlsx', 'xlsb', 'slk'
		] as Array<X.BookType>; for(var bti = 0; bti < bt.length; ++bti) { var w = bt[bti]; it(w, function() {
				var ws1 = X.utils.aoa_to_sheet([["hpx12", "hpt24", "hpx48", "hidden"]]);
				ws1['!cols'] = [{wch:9},{wpx:100},{width:80},{hidden:true}];
				var wb1 = {SheetNames:["Sheet1"], Sheets:{Sheet1:ws1}};
				var wb2 = X.read(X.write(wb1, {bookType:w, type:TYPE}), {type:TYPE, cellStyles:true});
				var ws2 = wb2.Sheets["Sheet1"];
				assert.equal(ws2?.['!cols']?.[3].hidden, true);
				assert.equal(ws2?.['!cols']?.[0].wch, 9);
				if(w == 'slk') return;
				assert.equal(ws2?.['!cols']?.[1].wpx, 100);
				/* xlml stores integral pixels -> approximate width */
				if(w == 'xlml') assert.equal(Math.round(ws2['!cols']?.[2]?.width||0), 80);
				else assert.equal(ws2['!cols']?.[2].width, 80);
		}); }
	});

	/* TODO: ODS and BIFF5/8 */
	describe('should preserve row properties', function() { var bt = [
			'xlml', /*'biff2', 'biff8', */ 'xlsx', 'xlsb', 'slk'
		] as Array<X.BookType>; for(var bti = 0; bti < bt.length; ++bti) { var w = bt[bti]; it(w, function() {
				var ws1 = X.utils.aoa_to_sheet([["hpx12"],["hpt24"],["hpx48"],["hidden"]]);
				ws1['!rows'] = [{hpx:12},{hpt:24},{hpx:48},{hidden:true}];
				for(var i = 0; i <= 7; ++i) ws1['!rows'].push({level:i});
				var wb1 = {SheetNames:["Sheet1"], Sheets:{Sheet1:ws1}};
				var wb2 = X.read(X.write(wb1, {bookType:w, type:TYPE, cellStyles:true}), {type:TYPE, cellStyles:true});
				var ws2 = wb2.Sheets["Sheet1"];
				assert.equal(ws2?.['!rows']?.[0]?.hpx, 12);
				assert.equal(ws2?.['!rows']?.[1]?.hpt, 24);
				assert.equal(ws2?.['!rows']?.[2]?.hpx, 48);
				assert.equal(ws2?.['!rows']?.[3]?.hidden, true);
				if(w == 'xlsb' || w == 'xlsx') for(i = 0; i <= 7; ++i) assert.equal((ws2?.['!rows']?.[4+i]||{}).level||0, i);
		}); }
	});

	/* TODO: ODS and XLS */
	describe('should preserve cell comments', function() { var cc = [
			['xlsx', paths.cstxlsx],
			['xlsb', paths.cstxlsb],
			//['xls', paths.cstxls],
			['xlml', paths.cstxml]
			//['ods', paths.cstods]
	]; for(var cci = 0; cci < cc.length; ++cci) { let w = cc[cci];
			it(w[0], function() {
				var wb1 = X.read(fs.readFileSync(w[1]), {type:TYPE});
				var wb2 = X.read(X.write(wb1, {bookType:w[0], type:TYPE}), {type:TYPE});
				check_comments(wb1);
				check_comments(wb2);
			});
		}
	});

	it('should preserve JS objects', function() {
		var data: Array<any> = [
			{a:1},
			{b:2,c:3},
			{b:"a",d:"b"},
			{a:true, c:false},
			{c:fixdate}
		];
		var o: Array<any> = X.utils.sheet_to_json(X.utils.json_to_sheet(data, {cellDates:true}));
		data.forEach(function(row, i) {
			Object.keys(row).forEach(function(k) { assert.equal(row[k], o[i][k]); });
		});
	});

	it('should preserve autofilter settings', function() {[
		['xlsx', paths.afxlsx],
		['xlsb', paths.afxlsb],
		// TODO:
		//['xls',  paths.afxls],
		['xlml', paths.afxml]
		//['ods',  paths.afods]
	].forEach(function(w) {
		var wb = X.read(fs.readFileSync(w[1]), {type:TYPE});
		var wb2 = X.read(X.write(wb, {bookType:w[0], type: TYPE}), {type:TYPE});
		assert.ok(!wb2.Sheets[wb2.SheetNames[0]]['!autofilter']);
		for(var i = 1; i < wb2.SheetNames.length; ++i) {
			assert.ok(wb2.Sheets[wb2.SheetNames[i]]['!autofilter']);
			assert.equal((wb2.Sheets[wb2.SheetNames[i]]['!autofilter'] as any).ref,"A1:E22");
		}
	});	});

	it('should preserve date system', function() {([
		"biff5", "ods", "slk", "xls", "xlsb", "xlsx", "xml"
	] as X.BookType[]).forEach(function(ext) {
		// TODO: check actual date codes and actual date values
		var wb0 = X.read(fs.readFileSync("./test_files/1904/1900." + ext), {type: TYPE});
		assert.ok(!wb0.Workbook?.WBProps?.date1904);
		var wb1 = X.read(X.write(wb0, {type: TYPE, bookType: ext}), {type: TYPE});
		assert.ok(!wb1.Workbook?.WBProps?.date1904);

		var wb2 = X.utils.book_new(); X.utils.book_append_sheet(wb2, X.utils.aoa_to_sheet([[1]]), "Sheet1");
		wb2.Workbook = { WBProps: { date1904: false } };
		assert.ok(!wb2.Workbook?.WBProps?.date1904);
		var wb3 = X.read(X.write(wb2, {type: TYPE, bookType: ext}), {type: TYPE});
		assert.ok(!wb3.Workbook?.WBProps?.date1904);

		var wb4 = X.read(fs.readFileSync("./test_files/1904/1904." + ext), {type: TYPE});
		assert.ok(wb4.Workbook?.WBProps?.date1904);
		var wb5 = X.read(X.write(wb4, {type: TYPE, bookType: ext}), {type: TYPE});
		assert.ok(wb5.Workbook?.WBProps?.date1904); // xlsb, xml

		var wb6 = X.utils.book_new(); X.utils.book_append_sheet(wb6, X.utils.aoa_to_sheet([[1]]), "Sheet1");
		wb6.Workbook = { WBProps: { date1904: true } };
		assert.ok(wb6.Workbook?.WBProps?.date1904);
		var wb7 = X.read(X.write(wb6, {type: TYPE, bookType: ext}), {type: TYPE});
		assert.ok(wb7.Workbook?.WBProps?.date1904);
	}); });

});

//function password_file(x){return x.match(/^password.*\.xls$/); }
var password_files = [
	//"password_2002_40_972000.xls",
	"password_2002_40_xor.xls"
];
describe('invalid files', function() {
	describe('parse', function() { var fl = [
		['KEY files', 'numbers/Untitled.key'],
		['PAGES files', 'numbers/Untitled.pages'],
		['password', 'apachepoi_password.xls'],
		['passwords', 'apachepoi_xor-encryption-abc.xls'],
		['DOC files', 'word_doc.doc']
	]; for(var f1 = 0; f1 < fl.length; ++f1) { let w = fl[f1]; it('should fail on ' + w[0], function() {
		assert.throws(function() { X.read(fs.readFileSync(dir + w[1], 'binary'), {type:'binary'}); });
		assert.throws(function() { X.read(fs.readFileSync(dir + w[1], 'base64'), {type:'base64'}); });
	}); } });
	describe('write', function() {
		it('should pass -> XLSX', function() { FSTPaths.forEach(function(p) {
			X.write(X.read(fs.readFileSync(p), {type:TYPE}), {type:TYPE});
		}); });
		it('should pass if a sheet is missing', function() {
			var wb = X.read(fs.readFileSync(paths.fstxlsx), {type:TYPE}); delete wb.Sheets[wb.SheetNames[0]];
			X.read(X.write(wb, {type:'binary'}), {type:'binary'});
		});
		var k1 = ['Props', 'Custprops', 'SSF']; for(var k1i = 0; k1i < k1.length; ++k1i) { var tt = k1[k1i]; it('should pass if ' + tt + ' is missing', function() {
			var wb = X.read(fs.readFileSync(paths.fstxlsx), {type:TYPE});
			assert.doesNotThrow(function() { delete (wb as any)[tt]; X.write(wb, {type:'binary'}); });
		}); }
		var k2 = ['SheetNames', 'Sheets']; for(var k2i = 0; k2i < k2.length; ++k2i) { var tt = k2[k2i]; it('should fail if ' + tt + ' is missing', function() {
			var wb = X.read(fs.readFileSync(paths.fstxlsx), {type:TYPE});
			assert.throws(function() { delete (wb as any)[tt]; X.write(wb, {type:'binary'}); });
		}); }
		it('should fail if SheetNames has duplicate entries', function() {
			var wb = X.read(fs.readFileSync(paths.fstxlsx), {type:TYPE});
			wb.SheetNames.push(wb.SheetNames[0]);
			assert.throws(function() { X.write(wb, {type:'binary'}); });
		});
	});
});


describe('json output', function() {
	function seeker(json: Array<any>, _keys: string | any[], val: any) {
		var keys: any[] = typeof _keys == "string" ? _keys.split("") : _keys;
		for(var i = 0; i != json.length; ++i) {
			for(var j = 0; j != keys.length; ++j) {
				if(json[i][keys[j]] === val) throw new Error("found " + val + " in row " + i + " key " + keys[j]);
			}
		}
	}
		var data = [
			[1,2,3],
			[true, false, null, "sheetjs"],
			["foo", "bar", fixdate, "0.3"],
			["baz", undefined, "qux"]
		];
		var ws = X.utils.aoa_to_sheet(data);
	it('should use first-row headers and full sheet by default', function() {
		var json: Array<any> = X.utils.sheet_to_json(ws, {raw: false});
		assert.equal(json.length, data.length - 1);
		assert.equal(json[0][1], "TRUE");
		assert.equal(json[1][2], "bar");
		assert.equal(json[2][3], "qux");
		assert.doesNotThrow(function() { seeker(json, [1,2,3], "sheetjs"); });
		assert.throws(function() { seeker(json, [1,2,3], "baz"); });
	});
	it('should create array of arrays if header == 1', function() {
		var json: Array<Array<any>> = X.utils.sheet_to_json(ws, {header:1, raw:false});
		assert.equal(json.length, data.length);
		assert.equal(json[1][0], "TRUE");
		assert.equal(json[2][1], "bar");
		assert.equal(json[3][2], "qux");
		assert.doesNotThrow(function() { seeker(json, [0,1,2], "sheetjs"); });
		assert.throws(function() { seeker(json, [0,1,2,3], "sheetjs"); });
		assert.throws(function() { seeker(json, [0,1,2], "baz"); });
	});
	it('should use column names if header == "A"', function() {
		var json: Array<any> = X.utils.sheet_to_json(ws, {header:'A', raw:false});
		assert.equal(json.length, data.length);
		assert.equal(json[1].A, "TRUE");
		assert.equal(json[2].B, "bar");
		assert.equal(json[3].C, "qux");
		assert.doesNotThrow(function() { seeker(json, "ABC", "sheetjs"); });
		assert.throws(function() { seeker(json, "ABCD", "sheetjs"); });
		assert.throws(function() { seeker(json, "ABC", "baz"); });
	});
	it('should use column labels if specified', function() {
		var json: Array<any> = X.utils.sheet_to_json(ws, {header:["O","D","I","N"], raw:false});
		assert.equal(json.length, data.length);
		assert.equal(json[1].O, "TRUE");
		assert.equal(json[2].D, "bar");
		assert.equal(json[3].I, "qux");
		assert.doesNotThrow(function() { seeker(json, "ODI", "sheetjs"); });
		assert.throws(function() { seeker(json, "ODIN", "sheetjs"); });
		assert.throws(function() { seeker(json, "ODIN", "baz"); });
	});
	var payload = [["string", "A2:D4"], ["numeric", 1], ["object", {s:{r:1,c:0},e:{r:3,c:3}}]];
	for(var payloadi = 0; payloadi < payload.length; ++payloadi) { var w = payload[payloadi];
		it('should accept custom ' + w[0] + ' range', function() {
			var json: any[][] = X.utils.sheet_to_json(ws, {header:1, range:w[1]});
			assert.equal(json.length, 3);
			assert.equal(json[0][0], true);
			assert.equal(json[1][1], "bar");
			assert.equal(json[2][2], "qux");
			assert.doesNotThrow(function() { seeker(json, [0,1,2], "sheetjs"); });
			assert.throws(function() { seeker(json, [0,1,2,3], "sheetjs"); });
			assert.throws(function() { seeker(json, [0,1,2], "baz"); });
		});
	}
	it('should use defval if requested', function() {
		var json: Array<any> = X.utils.sheet_to_json(ws, {defval: 'jimjin'});
		assert.equal(json.length, data.length - 1);
		assert.equal(json[0][1], true);
		assert.equal(json[1][2], "bar");
		assert.equal(json[2][3], "qux");
		assert.equal(json[2][2], "jimjin");
		assert.equal(json[0][3], "jimjin");
		assert.doesNotThrow(function() { seeker(json, [1,2,3], "sheetjs"); });
		assert.throws(function() { seeker(json, [1,2,3], "baz"); });
		X.utils.sheet_to_json(ws, {raw:true});
		X.utils.sheet_to_json(ws, {raw:true, defval: 'jimjin'});
	});
	it('should handle skipHidden for rows if requested', function() {
		var ws2 = X.utils.aoa_to_sheet(data), json: Array<any> = X.utils.sheet_to_json(ws2);
		assert.equal(json[0]["1"], true);
		assert.equal(json[2]["3"], "qux");
		ws2["!rows"] = []; ws2["!rows"][1] = {hidden:true}; json = X.utils.sheet_to_json(ws2, {skipHidden: true});
		assert.equal(json[0]["1"], "foo");
		assert.equal(json[1]["3"], "qux");
	});
	it('should handle skipHidden for columns if requested', function() {
		var ws2 = X.utils.aoa_to_sheet(data), json: Array<any> = X.utils.sheet_to_json(ws2);
		assert.equal(json[1]["2"], "bar");
		assert.equal(json[2]["3"], "qux");
		ws2["!cols"] = []; ws2["!cols"][1] = {hidden:true}; json = X.utils.sheet_to_json(ws2, {skipHidden: true});
		assert.equal(json[1]["2"], void 0);
		assert.equal(json[2]["3"], "qux");
	});
	it('should handle skipHidden when first row is hidden', function() {
		var ws2 = X.utils.aoa_to_sheet(data), json: Array<any> = X.utils.sheet_to_json(ws2);
		assert.equal(json[0]["1"], true);
		assert.equal(json[2]["3"], "qux");
		ws2["!rows"] = [{hidden:true}]; json = X.utils.sheet_to_json(ws2, {skipHidden: true});
		assert.equal(json[1]["1"], "foo");
		assert.equal(json[2]["3"], "qux");
	});
	it('should disambiguate headers', function() {
		var _data = [["S","h","e","e","t","J","S"],[1,2,3,4,5,6,7],[2,3,4,5,6,7,8]];
		var _ws = X.utils.aoa_to_sheet(_data);
		var json: any[] = X.utils.sheet_to_json(_ws);
		for(var i = 0; i < json.length; ++i) {
			assert.equal(json[i].S,   1 + i);
			assert.equal(json[i].h,   2 + i);
			assert.equal(json[i].e,   3 + i);
			assert.equal(json[i].e_1, 4 + i);
			assert.equal(json[i].t,   5 + i);
			assert.equal(json[i].J,   6 + i);
			assert.equal(json[i].S_1, 7 + i);
		}
	});
	it('should handle collisions in disambiguation', function() {
		var _data = [["a_1","a","a"],[1,2,3]];
		var _ws = X.utils.aoa_to_sheet(_data);
		var json: any = X.utils.sheet_to_json(_ws);
		assert.equal(json[0].a, 2);
		assert.equal(json[0].a_1, 1);
		assert.equal(json[0].a_2, 3);
	});
	it('should handle raw data if requested', function() {
		var _ws = X.utils.aoa_to_sheet(data, {cellDates:true});
		var json: Array<Array<any>> = X.utils.sheet_to_json(_ws, {header:1, raw:true});
		assert.equal(json.length, data.length);
		assert.equal(json[1][0], true);
		assert.ok(json[1][2] == null);
		assert.equal(json[2][1], "bar");
		assert.equal(json[2][2].getTime(), fixdate.getTime());
		assert.equal(json[3][2], "qux");
	});
	it('should include __rowNum__', function() {
		var _data = [["S","h","e","e","t","J","S"],[1,2,3,4,5,6,7],[],[2,3,4,5,6,7,8]];
		var _ws = X.utils.aoa_to_sheet(_data);
		var json: Array<any> = X.utils.sheet_to_json(_ws);
		assert.equal(json[0].__rowNum__, 1);
		assert.equal(json[1].__rowNum__, 3);
	});
	it('should handle blankrows', function() {
		var _data = [["S","h","e","e","t","J","S"],[1,2,3,4,5,6,7],[],[2,3,4,5,6,7,8]];
		var _ws = X.utils.aoa_to_sheet(_data);
		var json1: Array<any> = X.utils.sheet_to_json(_ws);
		assert.equal(json1.length, 2); // = 2 non-empty records
		var json2: Array<any> = X.utils.sheet_to_json(_ws, {header:1});
		assert.equal(json2.length, 4); // = 4 sheet rows
		var json3: Array<any> = X.utils.sheet_to_json(_ws, {blankrows:true});
		assert.equal(json3.length, 3); // = 2 records + 1 blank row
		var json4: Array<any> = X.utils.sheet_to_json(_ws, {blankrows:true, header:1});
		assert.equal(json4.length, 4); // = 4 sheet rows
		var json5: Array<any> = X.utils.sheet_to_json(_ws, {blankrows:false});
		assert.equal(json5.length, 2); // = 2 records
		var json6: Array<any> = X.utils.sheet_to_json(_ws, {blankrows:false, header:1});
		assert.equal(json6.length, 3); // = 4 sheet rows - 1 blank row
	});
	it('should have an index that starts with zero when selecting range', function() {
		var _data = [["S","h","e","e","t","J","S"],[1,2,3,4,5,6,7],[7,6,5,4,3,2,1],[2,3,4,5,6,7,8]];
		var _ws = X.utils.aoa_to_sheet(_data);
		var json1: Array<any> = X.utils.sheet_to_json(_ws, { header:1, raw: true, range: "B1:F3" });
		assert.equal(json1[0][3], "t");
		assert.equal(json1[1][0], 2);
		assert.equal(json1[2][1], 5);
		assert.equal(json1[2][3], 3);
	});
	it('should preserve values when column header is missing', function() {
		/*jshint elision:true */
		var _data = [[,"a","b",,"c"], [1,2,3,,5],[,3,4,5,6]]; // eslint-disable-line no-sparse-arrays
		/*jshint elision:false */
		var _ws = X.utils.aoa_to_sheet(_data);
		var json1: Array<any> = X.utils.sheet_to_json(_ws, { raw: true });
		assert.equal(json1[0].__EMPTY, 1);
		assert.equal(json1[1].__EMPTY_1, 5);
	});
	it('should ignore errors and support default values', function() {
		var ws = {
			A1: {t:'s', v:"Field"}, B1: {t:'s', v:"Text"},
			A2: {t:'e', v:0x2A, w:"#N/A" }, B2: {t:'s', v:"#N/A"},
			A3: {t:'e', v:0x0F }, B3: {t:'s', v:"#VALUE!"},
			A4: {t:'e', w:"#NAME?" }, B4: {t:'s', v:"#NAME?"},
			"!ref": "A1:B4" };
		seq(8).forEach(function(n) {
			var opts: X.Sheet2JSONOpts = {};
			if(n & 1) opts.header = 1;
			if(n & 2) opts.raw = true;
			if(n & 4) opts.defval = null;
			var J: Array<any> = X.utils.sheet_to_json(ws, opts);
			for(var i = 0; i < 3; ++i) {
				var k = ((n&1) ? J[i+1][0] : J[i].Field);
				assert.ok((n&4) ? (k === null) : (k !== null));
			}
		});
	});
});


var codes = [["あ 1", "\u00E3\u0081\u0082 1"]];
var plaintext_val = ([
	["A1", 'n', -0.08,  "-0.08"],
	["B1", 'n', 4001,   "4,001"],
	["C1", 's', "あ 1",  "あ 1"],
	["A2", 'n', 41.08, "$41.08"],
	["B2", 'n', 0.11,     "11%"],
	["C3", 'b', true,    "TRUE"],
	["D3", 'b', false,  "FALSE"],
	["B3", 's', " ",        " "],
	["A3"]
] as Array<[string, string, any, string] | [string] >);
function plaintext_test(wb: X.WorkBook, raw: boolean) {
	var sheet = wb.Sheets[wb.SheetNames[0]];
	plaintext_val.forEach(function(x) {
		var cell = get_cell(sheet, x[0]);
		var tcval = x[2+(!!raw ? 1 : 0)];
		var type = raw ? 's' : x[1];
		if(x.length == 1) { if(cell) { assert.equal(cell.t, 'z'); assert.ok(!cell.v); } return; }
		assert.equal(cell.v, tcval); assert.equal(cell.t, type);
	});
}
function make_html_str(idx: number) { return ["<table>",
	"<tr><td>-0.08</td><td>4,001</td><td>", codes[0][idx], "</td></tr>",
	"<tr><td>$41.08</td><td>11%</td></tr>",
	"<tr><td></td><td> \n&nbsp;</td><td>TRUE</td><td>FALSE</td></tr>",
"</table>" ].join(""); }
function make_csv_str(idx: number) { return [ (idx == 1 ? '\u00EF\u00BB\u00BF' : "") +
	'-0.08,"4,001",' + codes[0][idx] + '',
	'$41.08,11%',
	', ,TRUE,FALSE'
].join("\n"); }
var html_bstr = make_html_str(1), html_str = make_html_str(0);
var csv_bstr = make_csv_str(1), csv_str = make_csv_str(0);


describe('CSV', function() {
	describe('input', function(){
		var b = "1,2,3,\nTRUE,FALSE,,sheetjs\nfoo,bar,2/19/14,0.3\n,,,\nbaz,,qux,\n";
		it('should generate date numbers by default', function() {
			var opts: X.ParsingOptions = {type:"binary"};
			var cell = get_cell(X.read(b, opts).Sheets["Sheet1"], "C3");
			assert.equal(cell.w, '2/19/14');
			assert.equal(cell.t, 'n');
			assert.ok(typeof cell.v == "number");
		});
		it('should generate dates when requested', function() {
			var opts: X.ParsingOptions = {type:"binary", cellDates:true};
			var cell = get_cell(X.read(b, opts).Sheets["Sheet1"], "C3");
			assert.equal(cell.w, '2/19/14');
			assert.equal(cell.t, 'd');
			assert.ok(cell.v instanceof Date || typeof cell.v == "string");
		});

		it('should use US date code 14 by default', function() {
			var opts: X.ParsingOptions = { type:"binary" };
			var cell = get_cell(X.read(b, opts).Sheets["Sheet1"], "C3");
			assert.equal(cell.w, '2/19/14');
			opts.cellDates = true;
			cell = get_cell(X.read(b, opts).Sheets["Sheet1"], "C3");
			assert.equal(cell.w, '2/19/14');
		});
		it('should honor dateNF override', function() {
			var opts: X.ParsingOptions = ({type:"binary", dateNF:"YYYY-MM-DD"});
			var cell = get_cell(X.read(b, opts).Sheets["Sheet1"], "C3");
			/* NOTE: IE interprets 2-digit years as 19xx */
			assert.ok(cell.w == '2014-02-19' || cell.w == '1914-02-19');
			opts.cellDates = true; opts.dateNF = "YY-MM-DD";
			cell = get_cell(X.read(b, opts).Sheets["Sheet1"], "C3");
			assert.equal(cell.w, '14-02-19');
		});
		it('should interpret dateNF', function() {
			var bb = "1,2,3,\nTRUE,FALSE,,sheetjs\nfoo,bar,2/3/14,0.3\n,,,\nbaz,,qux,\n";
			var opts: X.ParsingOptions = {type:"binary", cellDates:true, dateNF:'m/d/yy'};
			var cell = get_cell(X.read(bb, opts).Sheets["Sheet1"], "C3");
			assert.equal(cell.v.getMonth(), 1);
			assert.equal(cell.w, "2/3/14");
			opts = {type:"binary", cellDates:true, dateNF:'d/m/yy'};
			cell = get_cell(X.read(bb, opts).Sheets["Sheet1"], "C3");
			assert.equal(cell.v.getMonth(), 2);
			assert.equal(cell.w, "2/3/14");
		});
		it('should interpret values by default', function() { plaintext_test(X.read(csv_bstr, {type:"binary"}), false); });
		it('should generate strings if raw option is passed', function() { plaintext_test(X.read(csv_str, {type:"string", raw:true}), true); });
		it('should handle formulae', function() {
			var bb = '=,=1+1,="100"';
			var sheet = X.read(bb, {type:"binary"}).Sheets["Sheet1"];
			assert.equal(get_cell(sheet, "A1").t, 's');
			assert.equal(get_cell(sheet, "A1").v, '=');
			assert.equal(get_cell(sheet, "B1").f, '1+1');
			assert.equal(get_cell(sheet, "C1").t, 's');
			assert.equal(get_cell(sheet, "C1").v, '100');
		});
		it('should interpret CRLF newlines', function() {
			var wb = X.read("sep=&\r\n1&2&3\r\n4&5&6", {type: "string"});
			assert.equal(wb.Sheets["Sheet1"]["!ref"], "A1:C2");
		});
		if(!browser || typeof cptable !== 'undefined') it('should honor codepage for binary strings', function() {
			var data = "abc,def\nghi,j\xD3l";
			([[1251, 'У'],[1252, 'Ó'], [1253, 'Σ'], [1254, 'Ó'], [1255, '׃'], [1256, 'س'], [10000, '”']] as Array<[number, string]>).forEach(function(m) {
				var ws = X.read(data, {type:"binary", codepage:m[0]}).Sheets["Sheet1"];
				assert.equal(get_cell(ws, "B2").v,  "j" + m[1] + "l");
			});
		});
		it('should parse date-less meridien time values', function() {
			var aoa = [
				["3a", "3 a", "3 a-1"],
				["3b", "3 b", "3 b-1"],
				["3p", "3 P", "3 p-1"],
			];
			var ws = X.read(aoa.map(function(row) { return row.join(","); }).join("\n"), {type: "string", cellDates: true}).Sheets.Sheet1;
			for(var R = 0; R < 3; ++R) {
				assert.equal(get_cell(ws, "A" + (R+1)).v, aoa[R][0]);
				assert.equal(get_cell(ws, "C" + (R+1)).v, aoa[R][2]);
			}
			assert.equal(get_cell(ws, "B2").v, "3 b");
			var B1 = get_cell(ws, "B1"); assert.equal(B1.t, "d"); assert.equal(B1.v.getHours(), 3);
			var B3 = get_cell(ws, "B3"); assert.equal(B3.t, "d"); assert.equal(B3.v.getHours(), 15);
			ws = X.read(aoa.map(function(row) { return row.join(","); }).join("\n"), {type: "string", cellDates: false}).Sheets.Sheet1;
			for(var R = 0; R < 3; ++R) {
				assert.equal(get_cell(ws, "A" + (R+1)).v, aoa[R][0]);
				assert.equal(get_cell(ws, "C" + (R+1)).v, aoa[R][2]);
			}
			assert.equal(get_cell(ws, "B2").v, "3 b");
			var B1 = get_cell(ws, "B1"); assert.equal(B1.t, "n"); assert.equal(B1.v * 24, 3);
			var B3 = get_cell(ws, "B3"); assert.equal(B3.t, "n"); assert.equal(B3.v * 24, 15);

		});
	});
	describe('output', function(){
			var data = [
				[1,2,3,null],
				[true, false, null, "sheetjs"],
				["foo", "bar", fixdate, "0.3"],
				[null, null, null],
				["baz", undefined, "qux"]
			];
			var ws = X.utils.aoa_to_sheet(data);
		it('should generate csv', function() {
			var baseline = "1,2,3,\nTRUE,FALSE,,sheetjs\nfoo,bar,2/19/14,0.3\n,,,\nbaz,,qux,";
			assert.equal(baseline, X.utils.sheet_to_csv(ws));
		});
		it('should handle FS', function() {
			assert.equal(X.utils.sheet_to_csv(ws, {FS:"|"}).replace(/[|]/g,","), X.utils.sheet_to_csv(ws));
			assert.equal(X.utils.sheet_to_csv(ws, {FS:";"}).replace(/[;]/g,","), X.utils.sheet_to_csv(ws));
		});
		it('should handle RS', function() {
			assert.equal(X.utils.sheet_to_csv(ws, {RS:"|"}).replace(/[|]/g,"\n"), X.utils.sheet_to_csv(ws));
			assert.equal(X.utils.sheet_to_csv(ws, {RS:";"}).replace(/[;]/g,"\n"), X.utils.sheet_to_csv(ws));
		});
		it('should handle dateNF', function() {
			var baseline = "1,2,3,\nTRUE,FALSE,,sheetjs\nfoo,bar,20140219,0.3\n,,,\nbaz,,qux,";
			var _ws =  X.utils.aoa_to_sheet(data, {cellDates:true});
			delete get_cell(_ws,"C3").w;
			delete get_cell(_ws,"C3").z;
			assert.equal(baseline, X.utils.sheet_to_csv(_ws, {dateNF:"YYYYMMDD"}));
		});
		it('should handle strip', function() {
			var baseline = "1,2,3\nTRUE,FALSE,,sheetjs\nfoo,bar,2/19/14,0.3\n\nbaz,,qux";
			assert.equal(baseline, X.utils.sheet_to_csv(ws, {strip:true}));
		});
		it('should handle blankrows', function() {
			var baseline = "1,2,3,\nTRUE,FALSE,,sheetjs\nfoo,bar,2/19/14,0.3\nbaz,,qux,";
			assert.equal(baseline, X.utils.sheet_to_csv(ws, {blankrows:false}));
		});
		it('should handle various line endings', function() {
			var data = ["1,a", "2,b", "3,c"];
			[ "\r", "\n", "\r\n" ].forEach(function(RS) {
				var wb = X.read(data.join(RS), {type:'binary'});
				assert.equal(get_cell(wb.Sheets["Sheet1"], "A1").v, 1);
				assert.equal(get_cell(wb.Sheets["Sheet1"], "B3").v, "c");
				assert.equal(wb.Sheets["Sheet1"]['!ref'], "A1:B3");
			});
		});
		it('should handle skipHidden for rows if requested', function() {
			var baseline = "1,2,3,\nTRUE,FALSE,,sheetjs\nfoo,bar,2/19/14,0.3\n,,,\nbaz,,qux,";
			delete ws["!rows"];
			assert.equal(X.utils.sheet_to_csv(ws), baseline);
			assert.equal(X.utils.sheet_to_csv(ws, {skipHidden:true}), baseline);
			ws["!rows"] = ([,{hidden:true},,] as any);
			assert.equal(X.utils.sheet_to_csv(ws), baseline);
			assert.equal(X.utils.sheet_to_csv(ws, {skipHidden:true}), "1,2,3,\nfoo,bar,2/19/14,0.3\n,,,\nbaz,,qux,");
			delete ws["!rows"];
		});
		it('should handle skipHidden for columns if requested', function() {
			var baseline = "1,2,3,\nTRUE,FALSE,,sheetjs\nfoo,bar,2/19/14,0.3\n,,,\nbaz,,qux,";
			delete ws["!cols"];
			assert.equal(X.utils.sheet_to_csv(ws), baseline);
			assert.equal(X.utils.sheet_to_csv(ws, {skipHidden:true}), baseline);
			ws["!cols"] = ([,{hidden:true},,] as any);
			assert.equal(X.utils.sheet_to_csv(ws), baseline);
			assert.equal(X.utils.sheet_to_csv(ws, {skipHidden:true}), "1,3,\nTRUE,,sheetjs\nfoo,2/19/14,0.3\n,,\nbaz,qux,");
			ws["!cols"] = ([{hidden:true},,,] as any);
			assert.equal(X.utils.sheet_to_csv(ws, {skipHidden:true}), "2,3,\nFALSE,,sheetjs\nbar,2/19/14,0.3\n,,\n,qux,");
			delete ws["!cols"];
		});
		it('should properly handle blankrows and strip options', function() {
			var _ws = X.utils.aoa_to_sheet([[""],[],["", ""]]);
			assert.equal(X.utils.sheet_to_csv(_ws, {}), ",\n,\n,");
			assert.equal(X.utils.sheet_to_csv(_ws, {strip: true}), "\n\n");
			assert.equal(X.utils.sheet_to_csv(_ws, {blankrows: false}), ",\n,");
			assert.equal(X.utils.sheet_to_csv(_ws, {blankrows: false, strip: true}), "");
		});
	});
});

describe('sylk', function() {
	describe('input', function(){
		it('codepage', function() {
			var str = "ID;PWXL;N;E\r\nC;X1;Y1;K\"a – b\"\r\nE", A1 =  "a – b";
			assert.equal(get_cell(X.read(str, {type:"string"}).Sheets["Sheet1"], "A1").v, A1);
			assert.equal(get_cell(X.read(str.replace(/–/, "\x96"), {type:"binary", codepage:1252}).Sheets["Sheet1"], "A1").v, A1);
			if(true) {
				assert.equal(get_cell(X.read(Buffer_from(str), {type:"buffer", codepage:65001}).Sheets["Sheet1"], "A1").v, A1);
				assert.equal(get_cell(X.read(Buffer_from(str.replace(/–/, "\x96"), "binary"), {type:"buffer", codepage:1252}).Sheets["Sheet1"], "A1").v, A1);
			}
		});
	});
	describe('date system', function() {
		function make_slk(d1904?: number) { return "ID;PSheetJS\nP;Pd\\/m\\/yy\nP;Pd\\/m\\/yyyy\n" + (d1904 != null ? "O;D;V" + d1904 : "") + "\nF;P0;FG0G;X1;Y1\nC;K1\nE"; }
		it('should default to 1900', function() {
			assert.equal(get_cell(X.read(make_slk(), {type: "binary"}).Sheets["Sheet1"], "A1").v, 1);
			assert.ok(get_cell(X.read(make_slk(), {type: "binary", cellDates: true}).Sheets["Sheet1"], "A1").v.getFullYear() < 1902);
			assert.equal(get_cell(X.read(make_slk(5), {type: "binary"}).Sheets["Sheet1"], "A1").v, 1);
			assert.ok(get_cell(X.read(make_slk(5), {type: "binary", cellDates: true}).Sheets["Sheet1"], "A1").v.getFullYear() < 1902);
		});
		it('should use 1904 when specified', function() {
			assert.ok(get_cell(X.read(make_slk(1), {type: "binary", cellDates: true}).Sheets["Sheet1"], "A1").v.getFullYear() > 1902);
			assert.ok(get_cell(X.read(make_slk(4), {type: "binary", cellDates: true}).Sheets["Sheet1"], "A1").v.getFullYear() > 1902);
		});
	});
});

if(typeof Uint8Array !== "undefined")
describe('numbers', function() {
	it('should parse files from Numbers 6.x', function() {
		var wb = X.read(fs.readFileSync(dir + 'numbers/types_61.numbers'), {type:TYPE, WTF:true});
		var ws = wb.Sheets["Sheet 1"];
		assert.equal(get_cell(ws, "A1").v, "Sheet");
		assert.equal(get_cell(ws, "B1").v, "JS");
		assert.equal(get_cell(ws, "B2").v, 123);
		assert.equal(get_cell(ws, "B11").v, true);
		assert.equal(get_cell(ws, "B13").v, 50);
	});
	if(can_write_numbers) it('should cap cols at 1000 (ALL)', function() {
		var aoa = [[1], [], []]; aoa[1][999] = 2; aoa[2][1000] = 3;
		var ws1 = X.utils.aoa_to_sheet(aoa);
		var wb1 = X.utils.book_new(); X.utils.book_append_sheet(wb1, ws1, "Sheet1");
		var wb2 = X.read(X.write(wb1,{bookType:"numbers",type:'binary',numbers:XLSX_ZAHL}),{type:'binary'});
		var ws2 = wb2.Sheets["Sheet1"];
		assert.equal(ws2["!ref"], "A1:ALL3");
		assert.equal(get_cell(ws2, "A1").v, 1);
		assert.equal(get_cell(ws2, "ALL2").v, 2);
	});
	it('should support icloud.com files', function() {
		var wb = X.read(fs.readFileSync(dir + 'Attendance.numbers'), {type:TYPE, WTF:true});
		var ws = wb.Sheets["Attendance"];
		assert.equal(get_cell(ws, "A1").v, "Date");
	});
});

describe('dbf', function() {
	var wbs: Array<[string, X.WorkBook]> = ([
		['d11',  dir + 'dbf/d11.dbf'],
		['vfp3', dir + 'dbf/vfp3.dbf']
	]).map(function(x) { return [x[0], X.read(fs.readFileSync(x[1]), {type:TYPE})]; });
	it(wbs[1][0], function() {
		var ws = wbs[1][1].Sheets["Sheet1"];
		([
			["A1", "v", "CHAR10"], ["A2", "v", "test1"], ["B2", "v", 123.45],
			["C2", "v", 12.345], ["D2", "v", 1234.1], ["E2", "w", "19170219"],
			/* [F2", "w", "19170219"], */ ["G2", "v", 1231.4], ["H2", "v", 123234],
			["I2", "v", true], ["L2", "v", "SheetJS"]
		] as Array<[string, string, any]>).forEach(function(r) { assert.equal(get_cell(ws, r[0])[r[1]], r[2]); });
	});
	it("Ś╫êëτ⌡ś and Š╫ěéτ⌡š", function() {
		([ [620, "Ś╫êëτ⌡ś"], [895, "Š╫ěéτ⌡š"] ] as Array<[number,string]>).forEach(function(r) {
			var book = X.utils.book_new();
			var sheet = X.utils.aoa_to_sheet([["ASCII", "encoded"], ["Test", r[1]]]);
			X.utils.book_append_sheet(book, sheet, "sheet1");
			var data = X.write(book, {type: TYPE, bookType: "dbf", codepage:r[0]});
			var wb = X.read(data, {type: TYPE});
			assert.equal(wb.Sheets.Sheet1.B2.v, r[1]);
		});
	});
});
import { JSDOM } from 'jsdom';
var domtest = false; // error: Error: Not implemented: isContext
var inserted_dom_elements = [];

function get_dom_element(html: string) {
	if(browser) {
		var domelt = document.createElement('div');
		domelt.innerHTML = html;
		if(document.body) document.body.appendChild(domelt);
		inserted_dom_elements.push(domelt);
		return domelt.children[0];
	}
	if(!JSDOM) throw new Error("Browser test fail");
	return new JSDOM(html).window.document.body.children[0];
}

describe('HTML', function() {
	describe('input string', function() {
		it('should interpret values by default', function() { plaintext_test(X.read(html_bstr, {type:"binary"}), false); });
		it('should generate strings if raw option is passed', function() { plaintext_test(X.read(html_bstr, {type:"binary", raw:true}), true); });
		it('should handle "string" type', function() { plaintext_test(X.read(html_str, {type:"string"}), false); });
		it('should handle newlines correctly', function() {
			var table = "<table><tr><td>foo<br/>bar</td><td>baz</td></tr></table>";
			var wb = X.read(table, {type:"string"});
			assert.equal(get_cell(wb.Sheets["Sheet1"], "A1").v, "foo\nbar");
		});
		it('should generate multi-sheet workbooks', function() {
			var table = "";
			for(var i = 0; i < 4; ++i) table += "<table><tr><td>" + X.utils.encode_col(i) + "</td><td>" + i + "</td></tr></table>";
			table += table; table += table;
			var wb = X.read(table, {type: "string"});
			assert.equal(wb.SheetNames.length, 16);
			assert.equal(wb.SheetNames[1], "Sheet2");
			for(var j = 0; j < 4; ++j) {
				assert.equal(get_cell(wb.Sheets["Sheet" + (j+1)], "A1").v, X.utils.encode_col(j));
				assert.equal(get_cell(wb.Sheets["Sheet" + (j+1)], "B1").v, j);
			}
		});
	});
	if(domtest) describe('input DOM', function() {
		it('should interpret values by default', function() { plaintext_test(X.utils.table_to_book(get_dom_element(html_str)), false); });
		it('should generate strings if raw option is passed', function() { plaintext_test(X.utils.table_to_book(get_dom_element(html_str), {raw:true}), true); });
		it('should handle newlines correctly', function() {
			var table = get_dom_element("<table><tr><td>foo<br/>bar</td><td>baz</td></tr></table>");
			var ws = X.utils.table_to_sheet(table);
			assert.equal(get_cell(ws, "A1").v, "foo\nbar");
		});
		it('should trim whitespace', function() {
			if(get_dom_element("foo <br> bar").innerHTML != "foo <br> bar") return;
			var table = get_dom_element("<table><tr><td>   foo  <br/>  bar   </td><td>  baz  qux  </td></tr></table>");
			var ws = X.utils.table_to_sheet(table);
			assert.equal(get_cell(ws, "A1").v.replace(/\n/g, "|"), "foo | bar");
			assert.equal(get_cell(ws, "B1").v, "baz qux");
		});
	});
	if(domtest) it('should handle entities', function() {
		var html = "<table><tr><td>A&amp;B</td><td>A&middot;B</td></tr></table>";
		var ws = X.utils.table_to_sheet(get_dom_element(html));
		assert.equal(get_cell(ws, "A1").v, "A&B");
		assert.equal(get_cell(ws, "B1").v, "A·B");
	});
	if(domtest) it('should honor sheetRows', function() {
		var html = X.utils.sheet_to_html(X.utils.aoa_to_sheet([[1,2],[3,4],[5,6]]));
		// $FlowIgnore
		html = /<body[^>]*>([\s\S]*)<\/body>/i.exec(html)?.[1] || "";
		var ws = X.utils.table_to_sheet(get_dom_element(html));
		assert.equal(ws['!ref'], "A1:B3");
		ws = X.utils.table_to_sheet(get_dom_element(html), {sheetRows:1});
		assert.equal(ws['!ref'], "A1:B1");
		assert.equal(ws['!fullref'], "A1:B3");
		ws = X.utils.table_to_sheet(get_dom_element(html), {sheetRows:2});
		assert.equal(ws['!ref'], "A1:B2");
		assert.equal(ws['!fullref'], "A1:B3");
	});
	if(domtest) it('should hide hidden rows', function() {
		var html = "<table><tr style='display: none;'><td>Foo</td></tr><tr><td style='display: none;'>Bar</td></tr><tr class='hidden'><td>Baz</td></tr></table><style>.hidden {display: none}</style>";
		var ws = X.utils.table_to_sheet(get_dom_element(html));
		var expected_rows = [];
		expected_rows[0] = expected_rows[2] = {hidden: true};
		assert.equal(ws['!ref'], "A1:A3");
		try {
			assert.deepEqual(ws['!rows'], expected_rows);
		} catch(e) {
			expected_rows[1] = {};
			assert.deepEqual(ws['!rows'], expected_rows);
		}
		assert.equal(get_cell(ws, "A1").v, "Foo");
		assert.equal(get_cell(ws, "A2").v, "Bar");
		assert.equal(get_cell(ws, "A3").v, "Baz");
	});
	if(domtest) it('should ignore hidden rows and cells when the `display` option is on', function() {
		var html = "<table><tr style='display: none;'><td>1</td><td>2</td><td>3</td></tr><tr><td class='hidden'>Foo</td><td>Bar</td><td style='display: none;'>Baz</td></tr></table><style>.hidden {display: none}</style>";
		var ws = X.utils.table_to_sheet(get_dom_element(html), {display: true});
		assert.equal(ws['!ref'], "A1");
		assert.ok(ws.hasOwnProperty('!rows') == false || !ws["!rows"]?.[0] || !ws["!rows"]?.[0]?.hidden);
		assert.equal(get_cell(ws, "A1").v, "Bar");
	});
	describe('type override', function() {
		function chk(ws: X.WorkSheet) {
			assert.equal(get_cell(ws, "A1").t, "s");
			assert.equal(get_cell(ws, "A1").v, "1234567890");
			assert.equal(get_cell(ws, "B1").t, "n");
			assert.equal(get_cell(ws, "B1").v, 1234567890);
		}
		var html = "<table><tr><td t=\"s\">1234567890</td><td>1234567890</td></tr></table>";
		it('HTML string', function() {
			var ws = X.read(html, {type:'string'}).Sheets["Sheet1"]; chk(ws);
			chk(X.read(X.utils.sheet_to_html(ws), {type:'string'}).Sheets["Sheet1"]);
		});
		if(domtest) it('DOM', function() { chk(X.utils.table_to_sheet(get_dom_element(html))); });
	});
	describe('TH/THEAD/TBODY/TFOOT elements', function() {
		var html = "<table><thead><tr><th>A</th><th>B</th></tr></thead><tbody><tr><td>1</td><td>2</td></tr><tr><td>3</td><td>4</td></tr></tbody><tfoot><tr><th>4</th><th>6</th></tr></tfoot></table>";
		it('HTML string', function() {
			var ws = X.read(html, {type:'string'}).Sheets["Sheet1"];
			assert.equal(X.utils.sheet_to_csv(ws),  "A,B\n1,2\n3,4\n4,6");
		});
		if(domtest) it('DOM', function() {
			var ws = X.utils.table_to_sheet(get_dom_element(html));
			assert.equal(X.utils.sheet_to_csv(ws),  "A,B\n1,2\n3,4\n4,6");
		});
	});
	describe('empty cell containing html element should increment cell index', function() {
		var html = "<table><tr><td>abc</td><td><b> </b></td><td>def</td></tr></table>";
		var expectedCellCount = 3;
		it('HTML string', function() {
			var ws = X.read(html, {type:'string'}).Sheets["Sheet1"];
			var range = X.utils.decode_range(ws['!ref']||"A1");
			assert.equal(range.e.c,expectedCellCount - 1);
		});
		if(domtest) it('DOM', function() {
			var ws = X.utils.table_to_sheet(get_dom_element(html));
			var range = X.utils.decode_range(ws['!ref']||"A1");
			assert.equal(range.e.c, expectedCellCount - 1);
		});
	});
});

describe('js -> file -> js', function() {
	var BIN: 'binary' ="binary";
		var ws = X.utils.aoa_to_sheet([
			["number", "bool", "string",  "date"],
			[1,        true,   "sheet"],
			[2,        false,  "dot"],
			[6.9,      false,  "JS", fixdate],
			[72.62,    true,   "0.3"]
		]);
		var wb = { SheetNames: ['Sheet1'], Sheets: {Sheet1: ws} };
	function eqcell(wb1: X.WorkBook, wb2: X.WorkBook, s: string, a: string) {
		assert.equal(get_cell(wb1.Sheets[s], a).v, get_cell(wb2.Sheets[s], a).v);
		assert.equal(get_cell(wb1.Sheets[s], a).t, get_cell(wb2.Sheets[s], a).t);
	}
	for(var ofmti = 0; ofmti < ofmt.length; ++ofmti) { var f = ofmt[ofmti];
		it(f, function() {
			var newwb = X.read(X.write(wb, {type:BIN, bookType: f}), {type:BIN});
			var cb = function(cell: string) { eqcell(wb, newwb, 'Sheet1', cell); };
			['A2', 'A3'].forEach(cb); /* int */
			['A4', 'A5'].forEach(cb); /* double */
			['B2', 'B3'].forEach(cb); /* bool */
			['C2', 'C3'].forEach(cb); /* string */
			if(!DIF_XL) cb('D4'); /* date */
			if(f != 'csv' && f != 'txt') eqcell(wb, newwb, 'Sheet1', 'C5');
		});
	}
	it('should roundtrip DIF strings', function() {
		var wb1 = X.read(X.write(wb,  {type:BIN, bookType: 'dif'}), {type:BIN});
		var wb2 = X.read(X.write(wb1, {type:BIN, bookType: 'dif'}), {type:BIN});
		eqcell(wb, wb1, 'Sheet1', 'C5');
		eqcell(wb, wb2, 'Sheet1', 'C5');
	});
});

describe('rtf', function() {
	it('roundtrip should be idempotent', function() {
		var ws = X.utils.aoa_to_sheet([
			[1,2,3],
			[true, false, null, "sheetjs"],
			["foo", "bar", fixdate, "0.3"],
			["baz", null, "q\"ux"]
		]);
		var wb1 = X.utils.book_new();
		X.utils.book_append_sheet(wb1, ws, "Sheet1");
		var rtf1 = X.write(wb1, {bookType: "rtf", type: "string"});
		var wb2 = X.read(rtf1, {type: "string"});
		var rtf2 = X.write(wb2, {bookType: "rtf", type: "string"});
		assert.equal(rtf1, rtf2);
	});
});

describe('corner cases', function() {
	it('output functions', function() {
		var ws = X.utils.aoa_to_sheet([
			[1,2,3],
			[true, false, null, "sheetjs"],
			["foo", "bar", fixdate, "0.3"],
			["baz", null, "q\"ux"]
		]);
		get_cell(ws,"A1").f = ""; get_cell(ws,"A1").w = "";
		delete get_cell(ws,"C3").w; delete get_cell(ws,"C3").z; get_cell(ws,"C3").XF = {ifmt:14};
		get_cell(ws,"A4").t = "e";
		X.utils.sheet_to_formulae(ws);
		X.utils.sheet_to_csv(ws);
		X.utils.sheet_to_json(ws);
		ws['!cols'] = [ {wch:6}, {wch:7}, {wch:10}, {wch:20} ];

		var wb = {SheetNames:['sheetjs'], Sheets:{sheetjs:ws}};
		X.write(wb, {type: "binary", bookType: 'xlsx'});
		X.write(wb, {type: TYPE, bookType: 'xlsm'});
		X.write(wb, {type: "base64", bookType: 'xlsb'});
		X.write(wb, {type: "binary", bookType: 'ods'});
		X.write(wb, {type: "binary", bookType: 'biff2'});
		X.write(wb, {type: "binary", bookType: 'biff5'});
		X.write(wb, {type: "binary", bookType: 'biff8'});
		get_cell(ws,"A2").t = "f";
		assert.throws(function() { X.utils.sheet_to_json(ws); });
	});
	it('SSF', function() {
		X.SSF.format("General", "dafuq");
		assert.throws(function() { return X.SSF.format("General", {sheet:"js"});});
		X.SSF.format("b e ddd hh AM/PM", 41722.4097222222);
		X.SSF.format("b ddd hh m", 41722.4097222222);
		["hhh","hhh A/P","hhmmm","sss","[hhh]","G eneral"].forEach(function(f) {
			assert.throws(function() { return X.SSF.format(f, 12345.6789);});
		});
		["[m]","[s]"].forEach(function(f) {
			assert.doesNotThrow(function() { return X.SSF.format(f, 12345.6789);});
		});
	});
	if(typeof JSON !== 'undefined') it('SSF oddities', function() {
		// $FlowIgnore
		var ssfdata: Array<any> = JSON.parse(fs.readFileSync('./misc/ssf.json', 'utf-8'));
		var cb = function(d: any, j: any) { return function() { return X.SSF.format(d[0], d[j][0]); }; };
		ssfdata.forEach(function(d) {
			for(var j=1;j<d.length;++j) {
				if(d[j].length == 2) {
					var expected = d[j][1], actual = X.SSF.format(d[0], d[j][0], {});
					assert.equal(actual, expected);
				} else if(d[j][2] !== "#") assert.throws(cb(d, j));
			}
		});
	});
	it('codepage', function() {
		X.read(fs.readFileSync(dir + "biff5/number_format_greek.xls"), {type:TYPE});
	});
	it('large binary files', function() {
		var data: Array<Array<string|number> > = [["Row Number"]];
		for(var j = 0; j < 19; ++j) data[0].push("Column " + j+1);
		for(var i = 0; i < 499; ++i) {
			var o: Array<string|number> = ["Row " + i];
			for(j = 0; j < 19; ++j) o.push(i + j);
			data.push(o);
		}
		var ws = X.utils.aoa_to_sheet(data);
		var wb = { Sheets:{ Sheet1: ws }, SheetNames: ["Sheet1"] };
		var type: 'binary' = "binary";
		(["xlsb", "biff8", "biff5", "biff2"] as Array<X.BookType>).forEach(function(btype) {
			void X.read(X.write(wb, {bookType:btype, type:type}), {type:type});
		});
	});
	if(fs.existsSync(dir + 'wtf_path.xlsx')) it('OPC oddities', function() {
		X.read(fs.readFileSync(dir + 'wtf_path.xlsx'), {WTF:true, type:TYPE});
		X.read(fs.readFileSync(dir + 'wtf_path.xlsb'), {WTF:true, type:TYPE});
	});
	it("should quote unicode sheet names in formulae", function() {
		var wb = X.read(fs.readFileSync(dir + "cross-sheet_formula_names.xlsb"), {WTF:true, type:TYPE});
		assert.equal(wb.Sheets["Sheet1"].A1.f, "'a-b'!A1");
		assert.equal(wb.Sheets["Sheet1"].A2.f, "'a#b'!A1");
		assert.equal(wb.Sheets["Sheet1"].A3.f, "'a^b'!A1");
		assert.equal(wb.Sheets["Sheet1"].A4.f, "'a%b'!A1");
		assert.equal(wb.Sheets["Sheet1"].A5.f, "'a\u066ab'!A1");
		assert.equal(wb.Sheets["Sheet1"].A6.f, "'☃️'!A1");
		assert.equal(wb.Sheets["Sheet1"].A7.f, "'\ud83c\udf63'!A1");
		assert.equal(wb.Sheets["Sheet1"].A8.f, "'a!!b'!A1");
		assert.equal(wb.Sheets["Sheet1"].A9.f, "'a$b'!A1");
		assert.equal(wb.Sheets["Sheet1"].A10.f, "'a!b'!A1");
		assert.equal(wb.Sheets["Sheet1"].A11.f, "'a b'!A1");
	});
	if(false) it('should parse CSV date values with preceding space', function() {
		function check_ws(ws: X.WorkSheet, dNF?: any) {
			//var d = X.SSF.parse_date_code(ws.B1.v);
			assert.equal(ws.B1.w, dNF ? '2018-03-24' : "3/23/18");
			//assert.equal(d.d, 24);
			//assert.equal(d.m, 3);
			//assert.equal(d.y, 2018);
		}
		[true, false].forEach(function(cD: boolean) {
			[void 0, 'yyyy-mm-dd'].forEach(function(dNF) {
				var ws1 = X.read(
					'7,2018-03-24',
					{cellDates: cD, dateNF: dNF, type:'string'}
				).Sheets["Sheet1"];
				check_ws(ws1, dNF);
				var ws2 = X.read(
					'7,  2018-03-24',
					{cellDates: cD, dateNF: dNF, type:'string'}
				).Sheets["Sheet1"];
				check_ws(ws2, dNF);
			});
		});
	});
	it('should handle \\r and \\n', function() {
		var base = "./test_files/crlf/";
		[
			"CRLFR9.123",
			"CRLFR9.WK1",
			"CRLFR9.WK3",
			"CRLFR9.WK4",
			"CRLFR9.XLS",
			"CRLFR9_4.XLS",
			"CRLFR9_5.XLS",
			"CRLFX5_2.XLS",
			"CRLFX5_3.XLS",
			"CRLFX5_4.XLS",
			"CRLFX5_5.XLS",
			"crlf.csv",
			"crlf.fods",
			"crlf.htm",
			"crlf.numbers",
			"crlf.ods",
			"crlf.rtf",
			"crlf.slk",
			"crlf.xls",
			"crlf.xlsb",
			"crlf.xlsx",
			"crlf.xml",
			"crlf5.xls",
			"crlfq9.qpw",
			"crlfq9.wb1",
			"crlfq9.wb2",
			"crlfq9.wb3",
			"crlfq9.wk1",
			"crlfq9.wk3",
			"crlfq9.wk4",
			"crlfq9.wks",
			"crlfq9.wq1",
			"crlfw4_2.wks",
			"crlfw4_3.wks",
			"crlfw4_4.wks"
		].map(function(path) { return base + path; }).forEach(function(w) {
			var wb = X.read(fs.readFileSync(w), {type:TYPE});
			var ws = wb.Sheets[wb.SheetNames[0]];
			var B1 = get_cell(ws, "B1"), B2 = get_cell(ws, "B2");
			var lio = w.match(/\.[^\.]*$/)?.index || 0, stem = w.slice(0, lio).toLowerCase(), ext = w.slice(lio + 1).toLowerCase();
			switch(ext) {
				case 'fm3': break;

				case '123':
					assert.equal(B1.v, "abc\ndef");
					// TODO: parse formula // assert.equal(B1.v, "abc\r\ndef");
					break;
				case 'qpw':
				case 'wb1':
				case 'wb2':
				case 'wb3':
				case 'wk1':
				case 'wk3':
				case 'wk4':
				case 'wq1':
					assert.ok(B1.v == "abcdef" || B1.v == "abc\ndef");
					// TODO: formula -> string values
					if(B2 && B2.t != "e" && B2.v != "") assert.ok(B2.v == "abcdef" || B2.v == "abc\r\ndef");
					break;

				case 'wks':
					if(stem.match(/w4/)) {
						assert.equal(B1.v, "abc\ndef");
						assert.ok(!B2 || B2.t == "z"); // Works4 did not support CODE / CHAR
					} else if(stem.match(/q9/)) {
						assert.equal(B1.v, "abcdef");
						assert.equal(B2.v, "abc\r\ndef");
					} else {
						assert.equal(B1.v, "abc\ndef");
						assert.equal(B2.v, "abc\r\ndef");
					}
					break;

				case 'xls':
					if(stem.match(/CRLFR9/i)) {
						assert.equal(B1.v, "abc\r\ndef");
					} else {
						assert.equal(B1.v, "abc\ndef");
					}
					assert.equal(B2.v, "abc\r\ndef");
					break;

				case 'rtf':
				case 'htm':
					assert.equal(B1.v, "abc\ndef");
					assert.equal(B2.v, "abc\n\ndef");
					break;

				case 'xlsx':
				case 'xlsb':
				case 'xml':
				case 'slk':
				case 'csv':
					assert.equal(B1.v, "abc\ndef");
					assert.equal(B2.v, "abc\r\ndef");
					break;
				case 'fods':
				case 'ods':
					assert.equal(B1.v, "abc\nDef");
					assert.equal(B2.v, "abc\r\ndef");
					break;
				case 'numbers':
					assert.equal(B1.v, "abc\ndef");
					// TODO: B2 should be a formula error
					break;
				default: throw ext;
			}
		});
	});
});

describe('encryption', function() {
	for(var pfi = 0; pfi < password_files.length; ++pfi) { var x = password_files[pfi];
		describe(x, function() {
			it('should throw with no password', function() {assert.throws(function() { X.read(fs.readFileSync(dir + x), {type:TYPE}); }); });
			it('should throw with wrong password', function() {
				try {
					X.read(fs.readFileSync(dir + x), {type:TYPE,password:'Password',WTF:opts.WTF});
					throw new Error("incorrect password was accepted");
				} catch(e) {
					if(e.message != "Password is incorrect") throw e;
				}
			});
			it('should recognize correct password', function() {
				try {
					X.read(fs.readFileSync(dir + x), {type:TYPE,password:'password',WTF:opts.WTF});
				} catch(e) {
					if(e.message == "Password is incorrect") throw e;
				}
			});
			if(false) it('should decrypt file', function() {
				/*var wb = */X.read(fs.readFileSync(dir + x), {type:TYPE,password:'password',WTF:opts.WTF});
			});
		});
	};
});

if(!browser || typeof cptable !== 'undefined')
describe('multiformat tests', function() {
var mfopts = opts;
var mft = fs.readFileSync('multiformat.lst','utf-8').replace(/\r/g,"").split("\n").map(function(x) { return x.trim(); });
var csv = true, formulae = false;
for(var mfti = 0; mfti < mft.length; ++mfti) { var x = mft[mfti];
	if(x.charAt(0)!="#") describe('MFT ' + x, function() {
		var f: Array<X.WorkBook> = [], r = x.split(/\s+/);
		if(r.length < 3) return;
		if(!fs.existsSync(dir + r[0] + r[1])) return;
		it('should parse all', function() {
			for(var j = 1; j < r.length; ++j) f[j-1] = X.read(fs.readFileSync(dir + r[0] + r[j]), mfopts);
		});
		it('should have the same sheetnames', function() {
			cmparr(f.map(function(x) { return x.SheetNames; }));
		});
		it('should have the same ranges', function() {
			f[0].SheetNames.forEach(function(s) {
				var ss = f.map(function(x) { return x.Sheets[s]; });
				cmparr(ss.map(function(s) { return s['!ref']; }));
			});
		});
		it('should have the same merges', function() {
			f[0].SheetNames.forEach(function(s) {
				var ss = f.map(function(x) { return x.Sheets[s]; });
				cmparr(ss.map(function(s) { return (s['!merges']||[]).map(function(y) { return X.utils.encode_range(y); }).sort(); }));
			});
		});
		if(csv) it('should have the same CSV', function() {
			cmparr(f.map(function(x) { return x.SheetNames; }));
			f[0].SheetNames.forEach(function(name) {
				cmparr(f.map(function(x) { return X.utils.sheet_to_csv(x.Sheets[name]); }));
			});
		});
		if(formulae) it('should have the same formulae', function() {
			cmparr(f.map(function(x) { return x.SheetNames; }));
			f[0].SheetNames.forEach(function(name: string) {
				cmparr(f.map(function(x) { return X.utils.sheet_to_formulae(x.Sheets[name]).sort(); }));
			});
		});
	});
	else x.split(/\s+/).forEach(function(w: string) { switch(w) {
		case "no-csv": csv = false; break;
		case "yes-csv": csv = true; break;
		case "no-formula": formulae = false; break;
		case "yes-formula": formulae = true; break;
	}});
} });
