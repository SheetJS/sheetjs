/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
/*jshint mocha:true */
/* eslint-env mocha */
/*global process, document, require */
/*global ArrayBuffer, Uint8Array */
/*::
declare type EmptyFunc = (() => void) | null;
declare type DescribeIt = { (desc:string, test:EmptyFunc):void; skip(desc:string, test:EmptyFunc):void; };
declare var describe : DescribeIt;
declare var it: DescribeIt;
declare var before:(test:EmptyFunc)=>void;
declare var cptable: any;
*/
var X;
var modp = './';
var fs = require('fs'), assert = require('assert');
describe('source',function(){it('should load',function(){X=require(modp);});});
var DIF_XL = true;

var browser = typeof document !== 'undefined';
// $FlowIgnore
if(!browser) try { require('./shim'); } catch(e) { }

var opts = ({cellNF: true}/*:any*/);
var TYPE = browser ? "binary" : "buffer";
opts.type = TYPE;
var fullex = [".xlsb", /*".xlsm",*/ ".xlsx"/*, ".xlml", ".xls"*/];
var ofmt = ["xlsb", "xlsm", "xlsx", "ods", "biff2", "biff5", "biff8", "xlml", "sylk", "dif", "dbf", "eth"];
var ex = fullex.slice(); ex = ex.concat([".ods", ".xls", ".xml", ".fods"]);
if(typeof process != 'undefined' && ((process||{}).env)) {
	opts.WTF = true;
	opts.cellStyles = true;
	if(process.env.FMTS === "full") process.env.FMTS = ex.join(":");
	if(process.env.FMTS) ex=process.env.FMTS.split(":").map(function(x){return x[0]==="."?x:"."+x;});
}
var exp = ex.map(function(x){ return x + ".pending"; });
function test_file(x){ return ex.indexOf(x.slice(-5))>=0||exp.indexOf(x.slice(-13))>=0 || ex.indexOf(x.slice(-4))>=0||exp.indexOf(x.slice(-12))>=0; }

var files = browser ? [] : (fs.existsSync('tests.lst') ? fs.readFileSync('tests.lst', 'utf-8').split("\n").map(function(x) { return x.trim(); }) : fs.readdirSync('test_files')).filter(test_file);
var fileA = browser ? [] : (fs.existsSync('tests/testA.lst') ? fs.readFileSync('tests/testA.lst', 'utf-8').split("\n").map(function(x) { return x.trim(); }) : []).filter(test_file);

/* Excel enforces 31 character sheet limit, although technical file limit is 255 */
function fixsheetname(x/*:string*/)/*:string*/ { return x.substr(0,31); }

function stripbom(x/*:string*/)/*:string*/ { return x.replace(/^\ufeff/,""); }
function fixcsv(x/*:string*/)/*:string*/ { return stripbom(x).replace(/\t/g,",").replace(/#{255}/g,"").replace(/"/g,"").replace(/[\n\r]+/g,"\n").replace(/\n*$/,""); }
function fixjson(x/*:string*/)/*:string*/ { return x.replace(/[\r\n]+$/,""); }

var dir = "./test_files/";

var dirwp = dir + "artifacts/wps/", dirqp = dir + "artifacts/quattro/";
var paths = {
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

	dtxls:  dir + 'xlsx-stream-d-date-cell.xls',
	dtxml:  dir + 'xlsx-stream-d-date-cell.xls.xml',
	dtxlsx:  dir + 'xlsx-stream-d-date-cell.xlsx',
	dtxlsb:  dir + 'xlsx-stream-d-date-cell.xlsb',

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
	svxls5: dir + 'sheet_visibility.xls',
	svxml:  dir + 'sheet_visibility.xml',
	svxlsx: dir + 'sheet_visibility.xlsx',
	svxlsb: dir + 'sheet_visibility.xlsb',

	swcxls: dir + 'apachepoi_SimpleWithComments.xls',
	swcxml: dir + '2011/apachepoi_SimpleWithComments.xls.xml',
	swcxlsx: dir + 'apachepoi_SimpleWithComments.xlsx',
	swcxlsb: dir + '2013/apachepoi_SimpleWithComments.xlsx.xlsb'
};

function pathit(p, ext) { return ext.map(function(n) { return paths[p + n]; }); }
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

function parsetest(x/*:string*/, wb/*:Workbook*/, full/*:boolean*/, ext/*:?string*/) {
	ext = (ext ? " [" + ext + "]": "");
	if(!full && ext) return;
	describe(x + ext + ' should have all bits', function() {
		var sname = dir + '2016/' + x.substr(x.lastIndexOf('/')+1) + '.sheetnames';
		if(!fs.existsSync(sname)) sname = dir + '2011/' + x.substr(x.lastIndexOf('/')+1) + '.sheetnames';
		if(!fs.existsSync(sname)) sname = dir + '2013/' + x.substr(x.lastIndexOf('/')+1) + '.sheetnames';
		it('should have all sheets', function() {
			wb.SheetNames.forEach(function(y) { assert(wb.Sheets[y], 'bad sheet ' + y); });
		});
		if(fs.existsSync(sname)) it('should have the right sheet names', function() {
			var file = fs.readFileSync(sname, 'utf-8').replace(/\r/g,"");
			var names = wb.SheetNames.map(fixsheetname).join("\n") + "\n";
			if(file.length && !x.match(/artifacts/)) assert.equal(names, file);
		});
	});
	describe(x + ext + ' should generate CSV', function() {
		wb.SheetNames.forEach(function(ws, i) {
			it('#' + i + ' (' + ws + ')', function() {
				X.utils.make_csv(wb.Sheets[ws]);
			});
		});
	});
	describe(x + ext + ' should generate JSON', function() {
		wb.SheetNames.forEach(function(ws, i) {
			it('#' + i + ' (' + ws + ')', function() {
				X.utils.sheet_to_json(wb.Sheets[ws]);
			});
		});
	});
	describe(x + ext + ' should generate formulae', function() {
		wb.SheetNames.forEach(function(ws, i) {
			it('#' + i + ' (' + ws + ')', function() {
				X.utils.get_formulae(wb.Sheets[ws]);
			});
		});
	});
	if(!full) return;
	var getfile = function(dir, x, i, type) {
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
		wb.SheetNames.forEach(function(ws, i) {
			var name = getfile(dir, x, i, ".csv");
			if(fs.existsSync(name)) it('#' + i + ' (' + ws + ')', function() {
				var file = fs.readFileSync(name, 'utf-8');
				var csv = X.utils.make_csv(wb.Sheets[ws]);
				assert.equal(fixcsv(csv), fixcsv(file), "CSV badness");
			});
		});
	});
	if(typeof JSON !== 'undefined') describe(x + ext + ' should generate correct JSON output', function() {
		wb.SheetNames.forEach(function(ws, i) {
			var rawjson = getfile(dir, x, i, ".rawjson");
			if(fs.existsSync(rawjson)) it('#' + i + ' (' + ws + ')', function() {
				var file = fs.readFileSync(rawjson, 'utf-8');
				var json = X.utils.make_json(wb.Sheets[ws],{raw:true});
				assert.equal(JSON.stringify(json), fixjson(file), "JSON badness");
			});

			var jsonf = getfile(dir, x, i, ".json");
			if(fs.existsSync(jsonf)) it('#' + i + ' (' + ws + ')', function() {
				var file = fs.readFileSync(jsonf, 'utf-8');
				var json = X.utils.make_json(wb.Sheets[ws]);
				assert.equal(JSON.stringify(json), fixjson(file), "JSON badness");
			});
		});
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

var wbtable = {};

(browser ? describe.skip : describe)('should parse test files', function() {
	files.forEach(function(x) {
		if(x.slice(-8) == ".pending" || !fs.existsSync(dir + x)) return;
		it(x, function() {
			var wb = X.readFile(dir + x, opts);
			wbtable[dir + x] = wb;
			parsetest(x, wb, true);
		});
		fullex.forEach(function(ext) {
			it(x + ' [' + ext + ']', function(){
				var wb = wbtable[dir + x];
				if(!wb) wb = X.readFile(dir + x, opts);
				wb = X.read(X.write(wb, {type:"buffer", bookType:ext.replace(/\./,"")}), {WTF:opts.WTF, cellNF: true});
				parsetest(x, wb, ext.replace(/\./,"") !== "xlsb", ext);
			});
		});
	});
	fileA.forEach(function(x) {
		if(x.slice(-8) == ".pending" || !fs.existsSync(dir + x)) return;
		it(x, function() {
			var wb = X.readFile(dir + x, {WTF:opts.WTF, sheetRows:10});
			parsetest(x, wb, false);
		});
	});
});

function get_cell(ws/*:Worksheet*/, addr/*:string*/) {
	if(!Array.isArray(ws)) return ws[addr];
	var a = X.utils.decode_cell(addr);
	return (ws[a.r]||[])[a.c];
}

function each_cell(ws, f) {
	if(Array.isArray(ws)) ws.forEach(function(row) { if(row) row.forEach(f); });
	else Object.keys(ws).forEach(function(addr) { if(addr[0] === "!" || !ws.hasOwnProperty(addr)) return; f(ws[addr]); });
}

function each_sheet(wb, f) { wb.SheetNames.forEach(function(n, i) { f(wb.Sheets[n], i); }); }

/* comments_stress_test family */
function check_comments(wb) {
	var ws0 = wb.Sheets.Sheet2;
	assert.equal(get_cell(ws0,"A1").c[0].a, 'Author');
	assert.equal(get_cell(ws0,"A1").c[0].t, 'Author:\nGod thinks this is good');
	assert.equal(get_cell(ws0,"C1").c[0].a, 'Author');
	assert.equal(get_cell(ws0,"C1").c[0].t, 'I really hope that xlsx decides not to use magic like rPr');

	var ws3 = wb.Sheets.Sheet4;
	assert.equal(get_cell(ws3,"B1").c[0].a, 'Author');
	assert.equal(get_cell(ws3,"B1").c[0].t, 'The next comment is empty');
	assert.equal(get_cell(ws3,"B2").c[0].a, 'Author');
	assert.equal(get_cell(ws3,"B2").c[0].t, '');
}

describe('parse options', function() {
	var html_cell_types = ['s'];
	var bef = (function() {
		X = require(modp);
	});
	if(typeof before != 'undefined') before(bef);
	else it('before', bef);
	describe('cell', function() {
		it('XLSX should generate HTML by default', function() {
			var wb = X.read(fs.readFileSync(paths.cstxlsx), {type:TYPE});
			var ws = wb.Sheets.Sheet1;
			each_cell(ws, function(cell) {
				assert(html_cell_types.indexOf(cell.t) === -1 || cell.h);
			});
		});
		it('XLSX should not generate HTML when requested', function() {
			var wb = X.read(fs.readFileSync(paths.cstxlsx), {type:TYPE, cellHTML:false});
			var ws = wb.Sheets.Sheet1;
			each_cell(ws, function(cell) {
				assert(typeof cell.h === 'undefined');
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
				assert(found);
			});
		});
		it('should not generate formulae when requested', function() {
			FSTPaths.forEach(function(p) {
				var wb =X.read(fs.readFileSync(p),{type:TYPE,cellFormula:false});
				wb.SheetNames.forEach(function(s) {
					each_cell(wb.Sheets[s], function(cell) {
						assert(typeof cell.f === 'undefined');
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
				assert(found);
			});
		});
		it('should not generate formatted text when requested', function() {
			FSTPaths.forEach(function(p) {
				var wb =X.read(fs.readFileSync(p),{type:TYPE, cellText:false});
				wb.SheetNames.forEach(function(s) {
					var ws = wb.Sheets[s];
					each_cell(ws, function(cell) {
						assert(typeof cell.w === 'undefined');
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
						assert(typeof cell.z === 'undefined');
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
						assert(cell.t!== 'n' || typeof cell.z !== 'undefined');
					});
				});
			});
		});
		it('should not generate cell styles by default', function() {
			CSSPaths.forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:TYPE, WTF:1});
				wb.SheetNames.forEach(function(s) {
					var ws = wb.Sheets[s];
					each_cell(ws, function(cell) {
						assert(typeof cell.s === 'undefined');
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
				assert(found);
			});
		});
		it('should not generate cell dates by default', function() {
			DTPaths.forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:TYPE});
				each_sheet(wb, function(ws/*::, i*/) { /*:: void i; */each_cell(ws, function(cell) {
					assert(cell.t !== 'd');
				}); });
			});
		});
		it('should generate cell dates when requested', function() {
			DTPaths.forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:TYPE, cellDates: true, WTF:1});
				var found = false;
				each_sheet(wb, function(ws/*::, i*/) { /*:: void i; */each_cell(ws, function(cell) {
					if(cell.t === 'd') return (found = true);
				}); });
				assert(found);
			});
		});
	});
	describe('sheet', function() {
		it('should not generate sheet stubs by default', function() {
			MCPaths.forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:TYPE});
				assert.throws(function() { return get_cell(wb.Sheets.Merge, "A2").v; });
			});
		});
		it('should generate sheet stubs when requested', function() {
			MCPaths.forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:TYPE, sheetStubs:true});
				assert(get_cell(wb.Sheets.Merge, "A2").t == 'z');
			});
		});
		it('should handle stub cells', function() {
			MCPaths.forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:TYPE, sheetStubs:true});
				X.utils.sheet_to_csv(wb.Sheets.Merge);
				X.utils.sheet_to_json(wb.Sheets.Merge);
				X.utils.sheet_to_formulae(wb.Sheets.Merge);
				ofmt.forEach(function(f) { if(f != "dbf") X.write(wb, {type:TYPE, bookType:f}); });
			});
		});
		function checkcells(wb, A46, B26, C16, D2) {
			[ ["A46", A46], ["B26", B26], ["C16", C16], ["D2", D2] ].forEach(function(r) {
				assert((typeof get_cell(wb.Sheets.Text, r[0]) !== 'undefined') == r[1]);
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
	});
	describe('book', function() {
		it('bookSheets should not generate sheets', function() {
			MCPaths.forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:TYPE, bookSheets:true});
				assert(typeof wb.Sheets === 'undefined');
			});
		});
		it('bookProps should not generate sheets', function() {
			NFPaths.forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:TYPE, bookProps:true});
				assert(typeof wb.Sheets === 'undefined');
			});
		});
		it('bookProps && bookSheets should not generate sheets', function() {
			PMPaths.forEach(function(p) {
				if(!fs.existsSync(p)) return;
				var wb = X.read(fs.readFileSync(p), {type:TYPE, bookProps:true, bookSheets:true});
				assert(typeof wb.Sheets === 'undefined');
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
				assert(typeof wb.Deps === 'undefined' || !(wb.Deps && wb.Deps.length>0));
			});
		});
		it('bookDeps should generate deps (XLSX/XLSB)', function() {
			FSTXL.forEach(function(p) {
				if(!p[1]) return;
				var wb = X.read(fs.readFileSync(p[0]), {type:TYPE, bookDeps:true});
				assert(typeof wb.Deps !== 'undefined' && wb.Deps.length > 0);
			});
		});

		var ckf = function(wb, fields, exists) { fields.forEach(function(f) { assert((typeof wb[f] !== 'undefined') == exists); }); };
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
			var wb = X.read(fs.readFileSync(p), {type:TYPE}); assert(typeof wb.vbaraw === 'undefined');
		}); });
		it('bookVBA should generate vbaraw', function() { NFVBA.forEach(function(p) {
			var wb = X.read(fs.readFileSync(p),{type: TYPE, bookVBA: true});
			assert(wb.vbaraw);
			var cfb = X.CFB.read(wb.vbaraw, {type: 'array'});
			assert(X.CFB.find(cfb, '/VBA/ThisWorkbook'));
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
	(typeof Uint8Array !== 'undefined' ? it : it.skip)('should read array', function() { artifax.forEach(function(p) {
		X.read(fs.readFileSync(p, 'binary').split("").map(function(x) { return x.charCodeAt(0); }), {type:'array'});
	}); });
	((browser || typeof Buffer === 'undefined') ? it.skip : it)('should read Buffers', function() { artifax.forEach(function(p) {
		X.read(fs.readFileSync(p), {type: 'buffer'});
	}); });
	(typeof Uint8Array !== 'undefined' ? it : it.skip)('should read ArrayBuffer / Uint8Array', function() { artifax.forEach(function(p) {
		var payload = fs.readFileSync(p, browser ? 'buffer' : null);
		var ab = new ArrayBuffer(payload.length), vu = new Uint8Array(ab);
		for(var i = 0; i < payload.length; ++i) vu[i] = payload[i];
		X.read(ab, {type: 'array'});
		X.read(vu, {type: 'array'});
	}); });
	it('should throw if format is unknown', function() { artifax.forEach(function(p) {
		assert.throws(function() { X.read(fs.readFileSync(p), {type: 'dafuq'}); });
	}); });

	var T = browser ? 'base64' : 'buffer';
	it('should default to "' + T + '" type', function() { artifax.forEach(function(p) {
		X.read(fs.readFileSync.apply(fs, browser ? [p, 'base64'] : [p]));
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
		["sylk",  false,   true],
		["eth",   false,   true],
		["html",   true,   true],
		["dif",   false,   true],
		["dbf",   false,  false],
		["prn",   false,   true]
	];
	function RT(T) {
		if(!X) X = require(modp);
		fmts.forEach(function(fmt) {
			var wb = X.utils.book_new();
			X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet([['R',"\u2603"],["\u0BEE",2]]), "Sheet1");
			if(T == 'string' && !fmt[2]) return assert.throws(function() {X.write(wb, {type: T, bookType:fmt[0], WTF:1});});
			var out = X.write(wb, {type: T, bookType:fmt[0], WTF:1});
			var nwb = X.read(out, {type: T, WTF:1});
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
	if(typeof ArrayBuffer !== 'undefined' && (typeof process == 'undefined' || !process.version.match(/v0.12/))) it('should write array buffers', function() { RT('array'); });
	if(!browser) it('should write buffers', function() { RT('buffer'); });
	it('should throw if format is unknown', function() { assert.throws(function() { RT('dafuq'); }); });
});

function eqarr(a,b) {
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
});

function coreprop(props) {
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
function custprop(props) {
	assert.equal(props['I am a boolean'], true);
	assert.equal(props['Date completed'].toISOString(), '1967-03-09T16:30:00.000Z');
	assert.equal(props.Status, 2);
	assert.equal(props.Counter, -3.14);
}

function cmparr(x){ for(var i=1;i<x.length;++i) assert.deepEqual(x[0], x[i]); }

function deepcmp(x,y,k,m,c) {
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
function diffsty(ws, r1,r2) {
	var c1 = get_cell(ws,r1).s, c2 = get_cell(ws,r2).s;
	stykeys.forEach(function(m) {
		var c = -1;
		if(styexc.indexOf(r1+"|"+r2+"|"+m) > -1) c = 1;
		else if(styexc.indexOf(r2+"|"+r1+"|"+m) > -1) c = 1;
		deepcmp(c1,c2,m,r1+","+r2,c);
	});
}

function hlink1(ws) {[
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

function hlink2(ws) { [
	["A1", "#Sheet2!A1"],
	["A2", "#WBScope"],
	["A3", "#Sheet1!WSScope1", "#Sheet1!C7:E8"],
	["A5", "#Sheet1!A5"]
].forEach(function(r) {
	if(r[2] && get_cell(ws, r[0]).l.Target == r[2]) return;
	assert.equal(get_cell(ws, r[0]).l.Target, r[1]);
}); }

function check_margin(margins, exp) {
	["left", "right", "top", "bottom", "header", "footer"].forEach(function(m,i) {
		assert.equal(margins[m],   exp[i]);
	});
}

describe('parse features', function() {
	describe('sheet visibility', function() {
		var wbs = [];
		var bef = (function() {
			wbs = SVPaths.map(function(p) { return X.read(fs.readFileSync(p), {type:TYPE}); });
		});
		if(typeof before != 'undefined') before(bef);
		else it('before', bef);

		it('should detect visible sheets', function() {
			if(!wbs.length) bef();
			wbs.forEach(function(wb) {
				assert(!wb.Workbook.Sheets[0].Hidden);
			});
		});
		it('should detect all hidden sheets', function() {
			wbs.forEach(function(wb) {
				assert(wb.Workbook.Sheets[1].Hidden);
				assert(wb.Workbook.Sheets[2].Hidden);
			});
		});
		it('should distinguish very hidden sheets', function() {
			wbs.forEach(function(wb) {
				assert.equal(wb.Workbook.Sheets[1].Hidden,1);
				assert.equal(wb.Workbook.Sheets[2].Hidden,2);
			});
		});
	});

	describe('comments', function() {
		if(fs.existsSync(paths.swcxlsx)) it('should have comment as part of cell properties', function(){
			X = require(modp);
			var sheet = 'Sheet1';
			var wb1=X.read(fs.readFileSync(paths.swcxlsx), {type:TYPE});
			var wb2=X.read(fs.readFileSync(paths.swcxlsb), {type:TYPE});
			var wb3=X.read(fs.readFileSync(paths.swcxls), {type:TYPE});
			var wb4=X.read(fs.readFileSync(paths.swcxml), {type:TYPE});

			[wb1,wb2,wb3,wb4].map(function(wb) { return wb.Sheets[sheet]; }).forEach(function(ws, i) {
				assert.equal(get_cell(ws, "B1").c.length, 1,"must have 1 comment");
				assert.equal(get_cell(ws, "B1").c[0].a, "Yegor Kozlov","must have the same author");
				assert.equal(get_cell(ws, "B1").c[0].t, "Yegor Kozlov:\nfirst cell", "must have the concatenated texts");
				if(i > 0) return;
				assert.equal(get_cell(ws, "B1").c[0].r, '<r><rPr><b/><sz val="8"/><color indexed="81"/><rFont val="Tahoma"/></rPr><t>Yegor Kozlov:</t></r><r><rPr><sz val="8"/><color indexed="81"/><rFont val="Tahoma"/></rPr><t xml:space="preserve">\r\nfirst cell</t></r>', "must have the rich text representation");
				assert.equal(get_cell(ws, "B1").c[0].h, '<span style="font-size:8;"><b>Yegor Kozlov:</b></span><span style="font-size:8;"><br/>first cell</span>', "must have the html representation");
			});
		});
		[
			['xlsx', paths.cstxlsx],
			['xlsb', paths.cstxlsb],
			['xls', paths.cstxls],
			['xlml', paths.cstxml],
			['ods', paths.cstods]
		].forEach(function(m) { it(m[0] + ' stress test', function() {
			var wb = X.read(fs.readFileSync(m[1]), {type:TYPE});
			check_comments(wb);
			var ws0 = wb.Sheets.Sheet2;
			assert.equal(get_cell(ws0,"A1").c[0].a, 'Author');
			assert.equal(get_cell(ws0,"A1").c[0].t, 'Author:\nGod thinks this is good');
			assert.equal(get_cell(ws0,"C1").c[0].a, 'Author');
			assert.equal(get_cell(ws0,"C1").c[0].t, 'I really hope that xlsx decides not to use magic like rPr');
		}); });
	});

	describe('should parse core properties and custom properties', function() {
		var wbs=[];
		var bef = (function() {
			wbs = [
				X.read(fs.readFileSync(paths.cpxlsx), {type:TYPE, WTF:1}),
				X.read(fs.readFileSync(paths.cpxlsb), {type:TYPE, WTF:1}),
				X.read(fs.readFileSync(paths.cpxls), {type:TYPE, WTF:1}),
				X.read(fs.readFileSync(paths.cpxml), {type:TYPE, WTF:1})
			];
		});
		if(typeof before != 'undefined') before(bef);
		else it('before', bef);

		['XLSX', 'XLSB', 'XLS', 'XML'].forEach(function(x, i) {
			it(x + ' should parse core properties', function() { coreprop(wbs[i].Props); });
			it(x + ' should parse custom properties', function() { custprop(wbs[i].Custprops); });
		});
		[
			["asxls",  "BIFF8", "\u2603"],
			["asxls5", "BIFF5", "_"],
			["asxml",   "XLML", "\u2603"],
			["asods",    "ODS", "God"],
			["asxlsx",  "XLSX", "\u2603"],
			["asxlsb",  "XLSB", "\u2603"]
		].forEach(function(x) {
		(fs.existsSync(paths[x[0]]) ? it : it.skip)(x[1] + ' should read ' + (x[2] == "\u2603" ? 'unicode ' : "") + 'author', function() {
			var wb = X.read(fs.readFileSync(paths[x[0]]), {type:TYPE});
			assert.equal(wb.Props.Author, x[2]);
		}); });
		var BASE = "இராமா";
		/* TODO: ODS, XLS */
		[ "xlsx", "xlsb", "xlml"/*, "ods", "xls" */].forEach(function(n) {
		it(n + ' should round-trip unicode category', function() {
			var wb = X.utils.book_new();
			X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet([["a"]]), "Sheet1");
			if(!wb.Props) wb.Props = {};
			wb.Props.Category = BASE;
			var wb2 = X.read(X.write(wb, {bookType:n, type:TYPE}), {type:TYPE});
			assert.equal(wb2.Props.Category,BASE);
		}); });
	});

	describe('sheetRows', function() {
		it('should use original range if not set', function() {
			var opts = {type:TYPE};
			FSTPaths.map(function(p) { return X.read(fs.readFileSync(p), opts); }).forEach(function(wb) {
				assert.equal(wb.Sheets.Text["!ref"],"A1:F49");
			});
		});
		it('should adjust range if set', function() {
			var opts = {type:TYPE, sheetRows:10};
			var wbs = FSTPaths.map(function(p) { return X.read(fs.readFileSync(p), opts); });
			/* TODO: XLS, XML, ODS */
			wbs.slice(0,2).forEach(function(wb) {
				assert.equal(wb.Sheets.Text["!fullref"],"A1:F49");
				assert.equal(wb.Sheets.Text["!ref"],"A1:F10");
			});
		});
		it('should not generate comment cells', function() {
			var opts = {type:TYPE, sheetRows:10};
			var wbs = CSTPaths.map(function(p) { return X.read(fs.readFileSync(p), opts); });
			/* TODO: XLS, XML, ODS */
			wbs.slice(0,2).forEach(function(wb) {
				assert.equal(wb.Sheets.Sheet7["!fullref"],"A1:N34");
				assert.equal(wb.Sheets.Sheet7["!ref"],"A1");
			});
		});
	});

	describe('column properties', function() {
		var wbs = [], wbs_no_slk = [];
		var bef = (function() {
			X = require(modp);
			wbs = CWPaths.map(function(n) { return X.read(fs.readFileSync(n), {type:TYPE, cellStyles:true}); });
			wbs_no_slk = wbs.slice(0, 5);
		});
		if(typeof before != 'undefined') before(bef);
		else it('before', bef);
		it('should have "!cols"', function() {
			wbs.forEach(function(wb) { assert(wb.Sheets.Sheet1['!cols']); });
		});
		it('should have correct widths', function() {
			/* SYLK rounds wch so skip non-integral */
			wbs_no_slk.map(function(x) { return x.Sheets.Sheet1['!cols']; }).forEach(function(x) {
				assert.equal(x[1].width, 0.1640625);
				assert.equal(x[2].width, 16.6640625);
				assert.equal(x[3].width, 1.6640625);
			});
			wbs.map(function(x) { return x.Sheets.Sheet1['!cols']; }).forEach(function(x) {
				assert.equal(x[4].width, 4.83203125);
				assert.equal(x[5].width, 8.83203125);
				assert.equal(x[6].width, 12.83203125);
				assert.equal(x[7].width, 16.83203125);
			});
		});
		it('should have correct pixels', function() {
			/* SYLK rounds wch so skip non-integral */
			wbs_no_slk.map(function(x) { return x.Sheets.Sheet1['!cols']; }).forEach(function(x) {
				assert.equal(x[1].wpx, 1);
				assert.equal(x[2].wpx, 100);
				assert.equal(x[3].wpx, 10);
			});
			wbs.map(function(x) { return x.Sheets.Sheet1['!cols']; }).forEach(function(x) {
				assert.equal(x[4].wpx, 29);
				assert.equal(x[5].wpx, 53);
				assert.equal(x[6].wpx, 77);
				assert.equal(x[7].wpx, 101);
			});
		});
	});

	describe('row properties', function() {
		var wbs = [], ols = [];
		var ol = fs.existsSync(paths.olxls);
		var bef = (function() {
			X = require(modp);
			wbs = RHPaths.map(function(p) { return X.read(fs.readFileSync(p), {type:TYPE, cellStyles:true}); });
			/* */
			if(!ol) return;
			ols = OLPaths.map(function(p) { return X.read(fs.readFileSync(p), {type:TYPE, cellStyles:true}); });
		});
		if(typeof before != 'undefined') before(bef);
		else it('before', bef);
		it('should have "!rows"', function() {
			wbs.forEach(function(wb) { assert(wb.Sheets.Sheet1['!rows']); });
		});
		it('should have correct points', function() {
			wbs.map(function(x) { return x.Sheets.Sheet1['!rows']; }).forEach(function(x) {
				assert.equal(x[1].hpt, 1);
				assert.equal(x[2].hpt, 10);
				assert.equal(x[3].hpt, 100);
			});
		});
		it('should have correct pixels', function() {
			wbs.map(function(x) { return x.Sheets.Sheet1['!rows']; }).forEach(function(x) {
				/* note: at 96 PPI hpt == hpx */
				assert.equal(x[1].hpx, 1);
				assert.equal(x[2].hpx, 10);
				assert.equal(x[3].hpx, 100);
			});
		});
		(ol ? it : it.skip)('should have correct outline levels', function() {
			ols.map(function(x) { return x.Sheets.Sheet1; }).forEach(function(ws) {
				var rows = ws['!rows'];
				for(var i = 0; i < 29; ++i) {
					var cell = get_cell(ws, "A" + X.utils.encode_row(i));
					var lvl = (rows[i]||{}).level||0;
					if(!cell || cell.t == 's') assert.equal(lvl, 0);
					else if(cell.t == 'n') {
						if(cell.v === 0) assert.equal(lvl, 0);
						else assert.equal(lvl, cell.v);
					}
				}
				assert.equal(rows[29].level, 7);
			});
		});
	});

	describe('merge cells',function() {
		var wbs=[];
		var bef = (function() {
			X = require(modp);
			wbs = MCPaths.map(function(p) { return X.read(fs.readFileSync(p), {type:TYPE}); });
		});
		if(typeof before != 'undefined') before(bef);
		else it('before', bef);
		it('should have !merges', function() {
			wbs.forEach(function(wb) {
				assert(wb.Sheets.Merge['!merges']);
			});
			var m = wbs.map(function(x) { return x.Sheets.Merge['!merges'].map(function(y) { return X.utils.encode_range(y); });});
			m.slice(1).forEach(function(x) {
				assert.deepEqual(m[0].sort(),x.sort());
			});
		});
	});

	describe('should find hyperlinks', function() {
		var wb1, wb2;
		var bef = (function() {
			X = require(modp);
			wb1 = HLPaths.map(function(p) { return X.read(fs.readFileSync(p), {type:TYPE, WTF:1}); });
			wb2 = ILPaths.map(function(p) { return X.read(fs.readFileSync(p), {type:TYPE, WTF:1}); });
		});
		if(typeof before != 'undefined') before(bef);
		else it('before', bef);

		['xlsx', 'xlsb', 'xls', 'xml'].forEach(function(x, i) {
			it(x + " external", function() { hlink1(wb1[i].Sheets.Sheet1); });
		});
		['xlsx', 'xlsb', 'xls', 'xml', 'ods'].forEach(function(x, i) {
			it(x + " internal", function() { hlink2(wb2[i].Sheets.Sheet1); });
		});
	});

	describe('should parse cells with date type (XLSX/XLSM)', function() {
		it('Must have read the date', function() {
			var wb, ws;
			var sheetName = 'Sheet1';
			wb = X.read(fs.readFileSync(paths.dtxlsx), {type:TYPE});
			ws = wb.Sheets[sheetName];
			var sheet = X.utils.sheet_to_json(ws);
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

	describe('defined names', function() {[
		/* desc     path        cmnt */
		['xlsx', paths.dnsxlsx,  true],
		['xlsb', paths.dnsxlsb,  true],
		['xls',  paths.dnsxls,   true],
		['xlml', paths.dnsxml,  false]
	].forEach(function(m) { it(m[0], function() {
		var wb = X.read(fs.readFileSync(m[1]), {type:TYPE});
		var names = wb.Workbook.Names;
		for(var i = 0; i < names.length; ++i) if(names[i].Name == "SheetJS") break;
		assert(i < names.length, "Missing name");
		assert.equal(names[i].Sheet, null);
		assert.equal(names[i].Ref, "Sheet1!$A$1");
		if(m[2]) assert.equal(names[i].Comment, "defined names just suck  excel formulae are bad  MS should feel bad");

		for(i = 0; i < names.length; ++i) if(names[i].Name == "SHEETjs") break;
		assert(i < names.length, "Missing name");
		assert.equal(names[i].Sheet, 0);
		assert.equal(names[i].Ref, "Sheet1!$A$2");
	}); }); });

	describe('auto filter', function() {[
		['xlsx', paths.afxlsx],
		['xlsb', paths.afxlsb],
		['xls',  paths.afxls],
		['xlml', paths.afxml],
		['ods',  paths.afods]
	].forEach(function(m) { it(m[0], function() {
		var wb = X.read(fs.readFileSync(m[1]), {type:TYPE});
		assert(!wb.Sheets[wb.SheetNames[0]]['!autofilter']);
		for(var i = 1; i < wb.SheetNames.length; ++i) {
			assert(wb.Sheets[wb.SheetNames[i]]['!autofilter']);
			assert.equal(wb.Sheets[wb.SheetNames[i]]['!autofilter'].ref,"A1:E22");
		}
	}); }); });

	describe('HTML', function() {
		var ws, wb;
		var bef = (function() {
			ws = X.utils.aoa_to_sheet([
				["a","b","c"],
				["&","<",">","\n"]
			]);
			wb = {SheetNames:["Sheet1"],Sheets:{Sheet1:ws}};
		});
		if(typeof before != 'undefined') before(bef);
		else it('before', bef);
		['xlsx'].forEach(function(m) { it(m, function() {
			var wb2 = X.read(X.write(wb, {bookType:m, type:TYPE}),{type:TYPE, cellHTML:true});
			assert.equal(get_cell(wb2.Sheets.Sheet1, "A2").h, "&amp;");
			assert.equal(get_cell(wb2.Sheets.Sheet1, "B2").h, "&lt;");
			assert.equal(get_cell(wb2.Sheets.Sheet1, "C2").h, "&gt;");
			assert.equal(get_cell(wb2.Sheets.Sheet1, "D2").h, "&#x000a;");
		}); });
	});

	describe('page margins', function() {
		var wbs=[];
		var bef = (function() {
			if(!fs.existsSync(paths.pmxls)) return;
			wbs = PMPaths.map(function(p) { return X.read(fs.readFileSync(p), {type:TYPE, WTF:1}); });
		});
		if(typeof before != 'undefined') before(bef);
		else it('before', bef);
		[
			/* Sheet Name     Margins: left   right  top bottom head foot */
			["Normal",                 [0.70, 0.70, 0.75, 0.75, 0.30, 0.30]],
			["Wide",                   [1.00, 1.00, 1.00, 1.00, 0.50, 0.50]],
			["Narrow",                 [0.25, 0.25, 0.75, 0.75, 0.30, 0.30]],
			["Custom 1 Inch Centered", [1.00, 1.00, 1.00, 1.00, 0.30, 0.30]],
			["1 Inch HF",              [0.70, 0.70, 0.75, 0.75, 1.00, 1.00]]
		].forEach(function(t) { it('should parse ' + t[0] + ' margin', function() { wbs.forEach(function(wb) {
			check_margin(wb.Sheets[t[0]]["!margins"], t[1]);
		}); }); });
	});

	describe('should correctly handle styles', function() {
		var wsxls, wsxlsx, rn, rn2;
		var bef = (function() {
			wsxls=X.read(fs.readFileSync(paths.cssxls), {type:TYPE,cellStyles:true,WTF:1}).Sheets.Sheet1;
			wsxlsx=X.read(fs.readFileSync(paths.cssxlsx), {type:TYPE,cellStyles:true,WTF:1}).Sheets.Sheet1;
			rn = function(range) {
				var r = X.utils.decode_range(range);
				var out = [];
				for(var R = r.s.r; R <= r.e.r; ++R) for(var C = r.s.c; C <= r.e.c; ++C)
					out.push(X.utils.encode_cell({c:C,r:R}));
				return out;
			};
			rn2 = function(r) { return [].concat.apply([], r.split(",").map(rn)); };
		});
		if(typeof before != 'undefined') before(bef);
		else it('before', bef);
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
			  fgColor: { theme: 3, raw_rgb: 'EEECE1' },
			  bgColor: { theme: 7, raw_rgb: '8064A2' } },
			{ patternType: 'darkGray',
			  fgColor: { theme: 3, raw_rgb: 'EEECE1' },
			  bgColor: { theme: 1, raw_rgb: 'FFFFFF' } },
			{ patternType: 'lightGray',
			  fgColor: { theme: 6, raw_rgb: '9BBB59' },
			  bgColor: { theme: 2, raw_rgb: '1F497D' } },
			{ patternType: 'lightDown',
			  fgColor: { theme: 4, raw_rgb: '4F81BD' },
			  bgColor: { theme: 7, raw_rgb: '8064A2' } },
			{ patternType: 'lightGrid',
			  fgColor: { theme: 6, raw_rgb: '9BBB59' },
			  bgColor: { theme: 9, raw_rgb: 'F79646' } },
			{ patternType: 'lightGrid',
			  fgColor: { theme: 4, raw_rgb: '4F81BD' },
			  bgColor: { theme: 2, raw_rgb: '1F497D' } },
			{ patternType: 'lightVertical',
			  fgColor: { theme: 3, raw_rgb: 'EEECE1' },
			  bgColor: { theme: 7, raw_rgb: '8064A2' } }
		];
		/*eslint-enable */
		ranges.forEach(function(rng) {
			it('XLS  | ' + rng,function(){cmparr(rn2(rng).map(function(x){ return get_cell(wsxls,x).s; }));});
			it('XLSX | ' + rng,function(){cmparr(rn2(rng).map(function(x){ return get_cell(wsxlsx,x).s; }));});
		});
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
});

describe('write features', function() {
	describe('props', function() {
		describe('core', function() {
			var ws;
			var baseprops = {
				Category: "Newspaper",
				ContentStatus: "Published",
				Keywords: "☃",
				LastAuthor: "Perry White",
				LastPrinted: "1978-12-15",
				RevNumber: 6969,
				AppVersion: 69,
				Author: "Lois Lane",
				Comments: "Needs work",
				Identifier: "1d",
				Language: "English",
				Subject: "Superman",
				Title: "Man of Steel"
			};
			var bef = (function() {
				X = require(modp);
				ws = X.utils.aoa_to_sheet([["a","b","c"],[1,2,3]]);
			});
			if(typeof before != 'undefined') before(bef);
			else it('before', bef);
			['xlml', 'xlsx', 'xlsb'].forEach(function(w) { it(w, function() {
				var wb = {
					Props: {},
					SheetNames: ["Sheet1"],
					Sheets: {Sheet1: ws}
				};
				Object.keys(baseprops).forEach(function(k) { wb.Props[k] = baseprops[k]; });
				var wb2 = X.read(X.write(wb, {bookType:w, type:TYPE}), {type:TYPE});
				Object.keys(baseprops).forEach(function(k) { assert.equal(baseprops[k], wb2.Props[k]); });
				var wb3 = X.read(X.write(wb2, {bookType:w, type:TYPE, Props: {Author:"SheetJS"}}), {type:TYPE});
				assert.equal("SheetJS", wb3.Props.Author);
			}); });
		});
	});
	describe('HTML', function() {
		it('should use `h` value when present', function() {
			var sheet = X.utils.aoa_to_sheet([["abc"]]);
			get_cell(sheet, "A1").h = "<b>abc</b>";
			var wb = {SheetNames:["Sheet1"], Sheets:{Sheet1:sheet}};
			var str = X.write(wb, {bookType:"html", type:"binary"});
			assert(str.indexOf("<b>abc</b>") > 0);
		});
	});
});

function seq(end/*:number*/, start/*:?number*/)/*:Array<number>*/ {
	var s = start || 0;
	var o = new Array(end - s);
	for(var i = 0; i != o.length; ++i) o[i] = s + i;
	return o;
}

var basedate = new Date(1899, 11, 30, 0, 0, 0); // 2209161600000
var dnthresh = basedate.getTime() + (new Date().getTimezoneOffset() - basedate.getTimezoneOffset()) * 60000;
function datenum(v/*:Date*/, date1904/*:?boolean*/)/*:number*/ {
	var epoch = v.getTime();
	if(date1904) epoch += 1462*24*60*60*1000;
	return (epoch - dnthresh) / (24 * 60 * 60 * 1000);
}
var good_pd_date = new Date('2017-02-19T19:06:09.000Z');
if(isNaN(good_pd_date.getFullYear())) good_pd_date = new Date('2/19/17');
var good_pd = good_pd_date.getFullYear() == 2017;
function parseDate(str/*:string|Date*/)/*:Date*/ {
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
	var bef = (function() { X = require(modp); });
	if(typeof before != 'undefined') before(bef);
	else it('before', bef);
	describe('should preserve core properties', function() { [
		['xlml', paths.cpxml],
		['xlsx', paths.cpxlsx],
		['xlsb', paths.cpxlsb]
	].forEach(function(w) {
		it(w[0], function() {
			var wb1 = X.read(fs.readFileSync(w[1]), {type:TYPE});
			coreprop(wb1.Props);
			var wb2 = X.read(X.write(wb1, {bookType:w[0], type:TYPE}), {type:TYPE});
			coreprop(wb2.Props);
		});
	}); });

	describe('should preserve custom properties', function() { [
		['xlml', paths.cpxml],
		['xlsx', paths.cpxlsx],
		['xlsb', paths.cpxlsb]
	].forEach(function(w) {
		it(w[0], function() {
			var wb1 = X.read(fs.readFileSync(w[1]), {type:TYPE});
			custprop(wb1.Custprops);
			var wb2 = X.read(X.write(wb1, {bookType:w[0], type:TYPE}), {type:TYPE});
			custprop(wb2.Custprops);
		});
	}); });

	describe('should preserve merge cells', function() {
		["xlsx", "xlsb", "xlml", "ods", "biff8"].forEach(function(f) { it(f, function() {
			var wb1 = X.read(fs.readFileSync(paths.mcxlsx), {type:TYPE});
			var wb2 = X.read(X.write(wb1,{bookType:f,type:'binary'}),{type:'binary'});
			var m1 = wb1.Sheets.Merge['!merges'].map(X.utils.encode_range);
			var m2 = wb2.Sheets.Merge['!merges'].map(X.utils.encode_range);
			assert.equal(m1.length, m2.length);
			for(var i = 0; i < m1.length; ++i) assert(m1.indexOf(m2[i]) > -1);
		}); });
	});

	describe('should preserve dates', function() {
		seq(16).forEach(function(n) {
			var d = (n & 1) ? 'd' : 'n', dk = d === 'd';
			var c = (n & 2) ? 'd' : 'n', dj = c === 'd';
			var b = (n & 4) ? 'd' : 'n', di = b === 'd';
			var a = (n & 8) ? 'd' : 'n', dh = a === 'd';
			var f, sheet, addr;
			if(dh) { f = paths.dtxlsx; sheet = 'Sheet1'; addr = 'B5'; }
			else { f = paths.nfxlsx; sheet = '2011'; addr = 'J36'; }
			it('[' + a + '] -> (' + b + ') -> [' + c + '] -> (' + d + ')', function() {
				var wb1 = X.read(fs.readFileSync(f), {type:TYPE, cellNF: true, cellDates: di, WTF: opts.WTF});
				var  _f = X.write(wb1, {type:'binary', cellDates:dj, WTF:opts.WTF});
				var wb2 = X.read(_f, {type:'binary', cellDates: dk, WTF: opts.WTF});
				var m = [wb1,wb2].map(function(x) { return get_cell(x.Sheets[sheet], addr); });
				assert.equal(m[0].w, m[1].w);

				assert.equal(m[0].t, b);
				assert.equal(m[1].t, d);

				if(m[0].t === 'n' && m[1].t === 'n') assert.equal(m[0].v, m[1].v);
				else if(m[0].t === 'd' && m[1].t === 'd') assert.equal(m[0].v.toString(), m[1].v.toString());
				else if(m[1].t === 'n') assert(Math.abs(datenum(browser ? parseDate(m[0].v) : new Date(m[0].v)) - m[1].v) < 0.01);
			});
		});
	});

	describe('should preserve formulae', function() { [
		['xlml', paths.fstxml],
		['xlsx', paths.fstxlsx],
		['ods',  paths.fstods]
	].forEach(function(w) { it(w[0], function() {
		var wb1 = X.read(fs.readFileSync(w[1]), {type:TYPE, cellFormula:true});
		var wb2 = X.read(X.write(wb1, {bookType:w[0], type:TYPE}), {cellFormula:true, type:TYPE});
		wb1.SheetNames.forEach(function(n) {
			assert.equal(
				X.utils.sheet_to_formulae(wb1.Sheets[n]).sort().join("\n"),
				X.utils.sheet_to_formulae(wb2.Sheets[n]).sort().join("\n")
			);
		});
	}); }); });

	describe('should preserve hyperlink', function() { [
		['xlml', paths.hlxml,   true],
		['xls',  paths.hlxls,   true],
		['xlsx', paths.hlxlsx,  true],
		['xlsb', paths.hlxlsb,  true],
		['xlml', paths.ilxml,  false],
		['xls',  paths.ilxls,  false],
		['xlsx', paths.ilxlsx, false],
		['xlsb', paths.ilxlsb, false],
		['ods',  paths.ilods,  false]
	].forEach(function(w) { it(w[0]+" "+(w[2]?"ex":"in")+ "ternal", function() {
		var wb = X.read(fs.readFileSync(w[1]), {type:TYPE, WTF:opts.WTF});
		var hlink = (w[2] ? hlink1 : hlink2); hlink(wb.Sheets.Sheet1);
		wb = X.read(X.write(wb, {bookType:w[0], type:TYPE, WTF:opts.WTF}), {type:TYPE, WTF:opts.WTF});
		hlink(wb.Sheets.Sheet1);
	}); }); });

	(fs.existsSync(paths.pmxlsx) ? describe : describe.skip)('should preserve page margins', function() {[
			['xlml', paths.pmxml],
			['xlsx', paths.pmxlsx],
			['xlsb', paths.pmxlsb]
		].forEach(function(w) { it(w[0], function() {
			var wb1 = X.read(fs.readFileSync(w[1]), {type:TYPE});
			var wb2 = X.read(X.write(wb1, {bookType:w[0], type:"binary"}), {type:"binary"});
			[
				/* Sheet Name     Margins: left   right  top bottom head foot */
				["Normal",                 [0.70, 0.70, 0.75, 0.75, 0.30, 0.30]],
				["Wide",                   [1.00, 1.00, 1.00, 1.00, 0.50, 0.50]],
				["Narrow",                 [0.25, 0.25, 0.75, 0.75, 0.30, 0.30]],
				["Custom 1 Inch Centered", [1.00, 1.00, 1.00, 1.00, 0.30, 0.30]],
				["1 Inch HF",              [0.70, 0.70, 0.75, 0.75, 1.00, 1.00]]
			].forEach(function(t) {
				check_margin(wb2.Sheets[t[0]]["!margins"], t[1]);
			});
	}); }); });

	describe('should preserve sheet visibility', function() { [
			['xlml', paths.svxml],
			['xlsx', paths.svxlsx],
			['xlsb', paths.svxlsb]
		].forEach(function(w) {
			it(w[0], function() {
				var wb1 = X.read(fs.readFileSync(w[1]), {type:TYPE});
				var wb2 = X.read(X.write(wb1, {bookType:w[0], type:TYPE}), {type:TYPE});
				var wbs1 = wb1.Workbook.Sheets;
				var wbs2 = wb2.Workbook.Sheets;
				assert.equal(wbs1.length, wbs2.length);
				for(var i = 0; i < wbs1.length; ++i) {
					assert.equal(wbs1[i].name, wbs2[i].name);
					assert.equal(wbs1[i].Hidden, wbs2[i].Hidden);
				}
			});
		});
	});

	describe('should preserve column properties', function() { [
			'xlml', /*'biff2', 'biff8', */ 'xlsx', 'xlsb', 'slk'
		].forEach(function(w) { it(w, function() {
				var ws1 = X.utils.aoa_to_sheet([["hpx12", "hpt24", "hpx48", "hidden"]]);
				ws1['!cols'] = [{wch:9},{wpx:100},{width:80},{hidden:true}];
				var wb1 = {SheetNames:["Sheet1"], Sheets:{Sheet1:ws1}};
				var wb2 = X.read(X.write(wb1, {bookType:w, type:TYPE}), {type:TYPE, cellStyles:true});
				var ws2 = wb2.Sheets.Sheet1;
				assert.equal(ws2['!cols'][3].hidden, true);
				assert.equal(ws2['!cols'][0].wch, 9);
				if(w == 'slk') return;
				assert.equal(ws2['!cols'][1].wpx, 100);
				/* xlml stores integral pixels -> approximate width */
				if(w == 'xlml') assert.equal(Math.round(ws2['!cols'][2].width), 80);
				else assert.equal(ws2['!cols'][2].width, 80);
		}); });
	});

	/* TODO: ODS and BIFF5/8 */
	describe('should preserve row properties', function() { [
			'xlml', /*'biff2', 'biff8', */ 'xlsx', 'xlsb', 'slk'
		].forEach(function(w) { it(w, function() {
				var ws1 = X.utils.aoa_to_sheet([["hpx12"],["hpt24"],["hpx48"],["hidden"]]);
				ws1['!rows'] = [{hpx:12},{hpt:24},{hpx:48},{hidden:true}];
				for(var i = 0; i <= 7; ++i) ws1['!rows'].push({level:i});
				var wb1 = {SheetNames:["Sheet1"], Sheets:{Sheet1:ws1}};
				var wb2 = X.read(X.write(wb1, {bookType:w, type:TYPE, cellStyles:true}), {type:TYPE, cellStyles:true});
				var ws2 = wb2.Sheets.Sheet1;
				assert.equal(ws2['!rows'][0].hpx, 12);
				assert.equal(ws2['!rows'][1].hpt, 24);
				assert.equal(ws2['!rows'][2].hpx, 48);
				assert.equal(ws2['!rows'][3].hidden, true);
				if(w == 'xlsb' || w == 'xlsx') for(i = 0; i <= 7; ++i) assert.equal((ws2['!rows'][4+i]||{}).level||0, i);
		}); });
	});

	/* TODO: ODS and XLS */
	describe('should preserve cell comments', function() { [
			['xlsx', paths.cstxlsx],
			['xlsb', paths.cstxlsb],
			//['xls', paths.cstxlsx],
			['xlml', paths.cstxml]
			//['ods', paths.cstods]
	].forEach(function(w) {
			it(w[0], function() {
				var wb1 = X.read(fs.readFileSync(w[1]), {type:TYPE});
				var wb2 = X.read(X.write(wb1, {bookType:w[0], type:TYPE}), {type:TYPE});
				check_comments(wb1);
				check_comments(wb2);
			});
		});
	});

	it('should preserve JS objects', function() {
		var data/*:Array<any>*/ = [
			{a:1},
			{b:2,c:3},
			{b:"a",d:"b"},
			{a:true, c:false},
			{c:fixdate}
		];
		var o = X.utils.sheet_to_json(X.utils.json_to_sheet(data, {cellDates:true}), {raw:true});
		data.forEach(function(row, i) {
			Object.keys(row).forEach(function(k) { assert.equal(row[k], o[i][k]); });
		});
	});
});

//function password_file(x){return x.match(/^password.*\.xls$/); }
//var password_files = fs.readdirSync('test_files').filter(password_file);
var password_files = [
	//"password_2002_40_972000.xls",
	"password_2002_40_xor.xls"
];
describe('invalid files', function() {
	describe('parse', function() { [
		['password', 'apachepoi_password.xls'],
		['passwords', 'apachepoi_xor-encryption-abc.xls'],
		['DOC files', 'word_doc.doc']
	].forEach(function(w) { it('should fail on ' + w[0], function() {
		assert.throws(function() { X.read(fs.readFileSync(dir + w[1], 'binary'), {type:'binary'}); });
		assert.throws(function() { X.read(fs.readFileSync(dir + w[1], 'base64'), {type:'base64'}); });
	}); }); });
	describe('write', function() {
		it('should pass -> XLSX', function() { FSTPaths.forEach(function(p) {
			X.write(X.read(fs.readFileSync(p), {type:TYPE}), {type:TYPE});
		}); });
		it('should pass if a sheet is missing', function() {
			var wb = X.read(fs.readFileSync(paths.fstxlsx), {type:TYPE}); delete wb.Sheets[wb.SheetNames[0]];
			X.read(X.write(wb, {type:'binary'}), {type:'binary'});
		});
		['Props', 'Custprops', 'SSF'].forEach(function(t) { it('should pass if ' + t + ' is missing', function() {
			var wb = X.read(fs.readFileSync(paths.fstxlsx), {type:TYPE});
			assert.doesNotThrow(function() { delete wb[t]; X.write(wb, {type:'binary'}); });
		}); });
		['SheetNames', 'Sheets'].forEach(function(t) { it('should fail if ' + t + ' is missing', function() {
			var wb = X.read(fs.readFileSync(paths.fstxlsx), {type:TYPE});
			assert.throws(function() { delete wb[t]; X.write(wb, {type:'binary'}); });
		}); });
		it('should fail if SheetNames has duplicate entries', function() {
			var wb = X.read(fs.readFileSync(paths.fstxlsx), {type:TYPE});
			wb.SheetNames.push(wb.SheetNames[0]);
			assert.throws(function() { X.write(wb, {type:'binary'}); });
		});
	});
});


describe('json output', function() {
	function seeker(json, keys, val) {
		if(typeof keys == "string") keys = keys.split("");
		for(var i = 0; i != json.length; ++i) {
			for(var j = 0; j != keys.length; ++j) {
				if(json[i][keys[j]] === val) throw new Error("found " + val + " in row " + i + " key " + keys[j]);
			}
		}
	}
	var data, ws;
	var bef = (function() {
		data = [
			[1,2,3],
			[true, false, null, "sheetjs"],
			["foo", "bar", fixdate, "0.3"],
			["baz", undefined, "qux"]
		];
		ws = X.utils.aoa_to_sheet(data);
	});
	if(typeof before != 'undefined') before(bef);
	else it('before', bef);
	it('should use first-row headers and full sheet by default', function() {
		var json = X.utils.sheet_to_json(ws);
		assert.equal(json.length, data.length - 1);
		assert.equal(json[0][1], "TRUE");
		assert.equal(json[1][2], "bar");
		assert.equal(json[2][3], "qux");
		assert.doesNotThrow(function() { seeker(json, [1,2,3], "sheetjs"); });
		assert.throws(function() { seeker(json, [1,2,3], "baz"); });
	});
	it('should create array of arrays if header == 1', function() {
		var json = X.utils.sheet_to_json(ws, {header:1});
		assert.equal(json.length, data.length);
		assert.equal(json[1][0], "TRUE");
		assert.equal(json[2][1], "bar");
		assert.equal(json[3][2], "qux");
		assert.doesNotThrow(function() { seeker(json, [0,1,2], "sheetjs"); });
		assert.throws(function() { seeker(json, [0,1,2,3], "sheetjs"); });
		assert.throws(function() { seeker(json, [0,1,2], "baz"); });
	});
	it('should use column names if header == "A"', function() {
		var json = X.utils.sheet_to_json(ws, {header:'A'});
		assert.equal(json.length, data.length);
		assert.equal(json[1].A, "TRUE");
		assert.equal(json[2].B, "bar");
		assert.equal(json[3].C, "qux");
		assert.doesNotThrow(function() { seeker(json, "ABC", "sheetjs"); });
		assert.throws(function() { seeker(json, "ABCD", "sheetjs"); });
		assert.throws(function() { seeker(json, "ABC", "baz"); });
	});
	it('should use column labels if specified', function() {
		var json = X.utils.sheet_to_json(ws, {header:["O","D","I","N"]});
		assert.equal(json.length, data.length);
		assert.equal(json[1].O, "TRUE");
		assert.equal(json[2].D, "bar");
		assert.equal(json[3].I, "qux");
		assert.doesNotThrow(function() { seeker(json, "ODI", "sheetjs"); });
		assert.throws(function() { seeker(json, "ODIN", "sheetjs"); });
		assert.throws(function() { seeker(json, "ODIN", "baz"); });
	});
	[["string", "A2:D4"], ["numeric", 1], ["object", {s:{r:1,c:0},e:{r:3,c:3}}]].forEach(function(w) {
		it('should accept custom ' + w[0] + ' range', function() {
			var json = X.utils.sheet_to_json(ws, {header:1, range:w[1]});
			assert.equal(json.length, 3);
			assert.equal(json[0][0], "TRUE");
			assert.equal(json[1][1], "bar");
			assert.equal(json[2][2], "qux");
			assert.doesNotThrow(function() { seeker(json, [0,1,2], "sheetjs"); });
			assert.throws(function() { seeker(json, [0,1,2,3], "sheetjs"); });
			assert.throws(function() { seeker(json, [0,1,2], "baz"); });
		});
	});
	it('should use defval if requested', function() {
		var json = X.utils.sheet_to_json(ws, {defval: 'jimjin'});
		assert.equal(json.length, data.length - 1);
		assert.equal(json[0][1], "TRUE");
		assert.equal(json[1][2], "bar");
		assert.equal(json[2][3], "qux");
		assert.equal(json[2][2], "jimjin");
		assert.equal(json[0][3], "jimjin");
		assert.doesNotThrow(function() { seeker(json, [1,2,3], "sheetjs"); });
		assert.throws(function() { seeker(json, [1,2,3], "baz"); });
		X.utils.sheet_to_json(ws, {raw:true});
		X.utils.sheet_to_json(ws, {raw:true, defval: 'jimjin'});
	});
	it('should disambiguate headers', function() {
		var _data = [["S","h","e","e","t","J","S"],[1,2,3,4,5,6,7],[2,3,4,5,6,7,8]];
		var _ws = X.utils.aoa_to_sheet(_data);
		var json = X.utils.sheet_to_json(_ws);
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
	it('should handle raw data if requested', function() {
		var _ws = X.utils.aoa_to_sheet(data, {cellDates:true});
		var json = X.utils.sheet_to_json(_ws, {header:1, raw:true});
		assert.equal(json.length, data.length);
		assert.equal(json[1][0], true);
		assert.equal(json[1][2], null);
		assert.equal(json[2][1], "bar");
		assert.equal(json[2][2].getTime(), fixdate.getTime());
		assert.equal(json[3][2], "qux");
	});
	it('should include __rowNum__', function() {
		var _data = [["S","h","e","e","t","J","S"],[1,2,3,4,5,6,7],[],[2,3,4,5,6,7,8]];
		var _ws = X.utils.aoa_to_sheet(_data);
		var json = X.utils.sheet_to_json(_ws);
		assert.equal(json[0].__rowNum__, 1);
		assert.equal(json[1].__rowNum__, 3);
	});
	it('should handle blankrows', function() {
		var _data = [["S","h","e","e","t","J","S"],[1,2,3,4,5,6,7],[],[2,3,4,5,6,7,8]];
		var _ws = X.utils.aoa_to_sheet(_data);
		var json1 = X.utils.sheet_to_json(_ws);
		assert.equal(json1.length, 2); // = 2 non-empty records
		var json2 = X.utils.sheet_to_json(_ws, {header:1});
		assert.equal(json2.length, 4); // = 4 sheet rows
		var json3 = X.utils.sheet_to_json(_ws, {blankrows:true});
		assert.equal(json3.length, 3); // = 2 records + 1 blank row
		var json4 = X.utils.sheet_to_json(_ws, {blankrows:true, header:1});
		assert.equal(json4.length, 4); // = 4 sheet rows
		var json5 = X.utils.sheet_to_json(_ws, {blankrows:false});
		assert.equal(json5.length, 2); // = 2 records
		var json6 = X.utils.sheet_to_json(_ws, {blankrows:false, header:1});
		assert.equal(json6.length, 3); // = 4 sheet rows - 1 blank row
	});
	it('should have an index that starts with zero when selecting range', function() {
		var _data = [["S","h","e","e","t","J","S"],[1,2,3,4,5,6,7],[7,6,5,4,3,2,1],[2,3,4,5,6,7,8]];
		var _ws = X.utils.aoa_to_sheet(_data);
		var json1 = X.utils.sheet_to_json(_ws, { header:1, raw: true, range: "B1:F3" });
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
		var json1 = X.utils.sheet_to_json(_ws, { raw: true });
		assert.equal(json1[0].__EMPTY, 1);
		assert.equal(json1[1].__EMPTY_1, 5);
	});
});


var codes = [["あ 1", "\u00E3\u0081\u0082 1"]];
var plaintext_val = [
	["A1", 'n', -0.08,  "-0.08"],
	["B1", 'n', 4001,   "4,001"],
	["C1", 's', "あ 1",  "あ 1"],
	["A2", 'n', 41.08, "$41.08"],
	["B2", 'n', 0.11,     "11%"],
	["C3", 'b', true,    "TRUE"],
	["D3", 'b', false,  "FALSE"],
	["B3", 's', " ",        " "],
	["A3"]
];
function plaintext_test(wb, raw) {
	var sheet = wb.Sheets[wb.SheetNames[0]];
	plaintext_val.forEach(function(x) {
		var cell = get_cell(sheet, x[0]);
		var tcval = x[2+!!raw];
		var type = raw ? 's' : x[1];
		if(x.length == 1) { if(cell) { assert.equal(cell.t, 'z'); assert(!cell.v); } return; }
		assert.equal(cell.v, tcval); assert.equal(cell.t, type);
	});
}
function make_html_str(idx) { return ["<table>",
	"<tr><td>-0.08</td><td>4,001</td><td>", codes[0][idx], "</td></tr>",
	"<tr><td>$41.08</td><td>11%</td></tr>",
	"<tr><td></td><td> \n&nbsp;</td><td>TRUE</td><td>FALSE</td></tr>",
"</table>" ].join(""); }
function make_csv_str(idx) { return [ (idx == 1 ? '\u00EF\u00BB\u00BF' : "") +
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
			var opts = {type:"binary"};
			var cell = get_cell(X.read(b, opts).Sheets.Sheet1, "C3");
			assert.equal(cell.w, '2/19/14');
			assert.equal(cell.t, 'n');
			assert(typeof cell.v == "number");
		});
		it('should generate dates when requested', function() {
			var opts = {type:"binary", cellDates:true};
			var cell = get_cell(X.read(b, opts).Sheets.Sheet1, "C3");
			assert.equal(cell.w, '2/19/14');
			assert.equal(cell.t, 'd');
			assert(cell.v instanceof Date || typeof cell.v == "string");
		});

		it('should use US date code 14 by default', function() {
			var opts = ({type:"binary"}/*:any*/);
			var cell = get_cell(X.read(b, opts).Sheets.Sheet1, "C3");
			assert.equal(cell.w, '2/19/14');
			opts.cellDates = true;
			cell = get_cell(X.read(b, opts).Sheets.Sheet1, "C3");
			assert.equal(cell.w, '2/19/14');
		});
		it('should honor dateNF override', function() {
			var opts = ({type:"binary", dateNF:"YYYY-MM-DD"}/*:any*/);
			var cell = get_cell(X.read(b, opts).Sheets.Sheet1, "C3");
			/* NOTE: IE interprets 2-digit years as 19xx */
			assert(cell.w == '2014-02-19' || cell.w == '1914-02-19');
			opts.cellDates = true; opts.dateNF = "YY-MM-DD";
			cell = get_cell(X.read(b, opts).Sheets.Sheet1, "C3");
			assert.equal(cell.w, '14-02-19');
		});
		it('should interpret dateNF', function() {
			var bb = "1,2,3,\nTRUE,FALSE,,sheetjs\nfoo,bar,2/3/14,0.3\n,,,\nbaz,,qux,\n";
			var opts = {type:"binary", cellDates:true, dateNF:'m/d/yy'};
			var cell = get_cell(X.read(bb, opts).Sheets.Sheet1, "C3");
			assert.equal(cell.v.getMonth(), 1);
			assert.equal(cell.w, "2/3/14");
			opts = {type:"binary", cellDates:true, dateNF:'d/m/yy'};
			cell = get_cell(X.read(bb, opts).Sheets.Sheet1, "C3");
			assert.equal(cell.v.getMonth(), 2);
			assert.equal(cell.w, "2/3/14");
		});
		it('should interpret values by default', function() { plaintext_test(X.read(csv_bstr, {type:"binary"}), false); });
		it('should generate strings if raw option is passed', function() { plaintext_test(X.read(csv_str, {type:"string", raw:true}), true); });
		it('should handle formulae', function() {
			var bb = '=,=1+1,="100"';
			var sheet = X.read(bb, {type:"binary"}).Sheets.Sheet1;
			assert.equal(get_cell(sheet, "A1").t, 's');
			assert.equal(get_cell(sheet, "A1").v, '=');
			assert.equal(get_cell(sheet, "B1").f, '1+1');
			assert.equal(get_cell(sheet, "C1").t, 's');
			assert.equal(get_cell(sheet, "C1").v, '100');
		});
		if(!browser || typeof cptable !== 'undefined') it('should honor codepage for binary strings', function() {
			var data = "abc,def\nghi,j\xD3l";
			[[1251, 'У'],[1252, 'Ó'], [1253, 'Σ'], [1254, 'Ó'], [1255, '׃'], [1256, 'س'], [10000, '”']].forEach(function(m) {
				var ws = X.read(data, {type:"binary", codepage:m[0]}).Sheets.Sheet1;
				assert.equal(get_cell(ws, "B2").v,  "j" + m[1] + "l");
			});
		});
	});
	describe('output', function(){
		var data, ws;
		var bef = (function() {
			data = [
				[1,2,3,null],
				[true, false, null, "sheetjs"],
				["foo", "bar", fixdate, "0.3"],
				[null, null, null],
				["baz", undefined, "qux"]
			];
			ws = X.utils.aoa_to_sheet(data);
		});
		if(typeof before != 'undefined') before(bef);
		else it('before', bef);
		it('should generate csv', function() {
			var baseline = "1,2,3,\nTRUE,FALSE,,sheetjs\nfoo,bar,2/19/14,0.3\n,,,\nbaz,,qux,\n";
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
			var baseline = "1,2,3,\nTRUE,FALSE,,sheetjs\nfoo,bar,20140219,0.3\n,,,\nbaz,,qux,\n";
			var _ws =  X.utils.aoa_to_sheet(data, {cellDates:true});
			delete get_cell(_ws,"C3").w;
			delete get_cell(_ws,"C3").z;
			assert.equal(baseline, X.utils.sheet_to_csv(_ws, {dateNF:"YYYYMMDD"}));
		});
		it('should handle strip', function() {
			var baseline = "1,2,3\nTRUE,FALSE,,sheetjs\nfoo,bar,2/19/14,0.3\n\nbaz,,qux\n";
			assert.equal(baseline, X.utils.sheet_to_csv(ws, {strip:true}));
		});
		it('should handle blankrows', function() {
			var baseline = "1,2,3,\nTRUE,FALSE,,sheetjs\nfoo,bar,2/19/14,0.3\nbaz,,qux,\n";
			assert.equal(baseline, X.utils.sheet_to_csv(ws, {blankrows:false}));
		});
		it('should handle various line endings', function() {
			var data = ["1,a", "2,b", "3,c"];
			[ "\r", "\n", "\r\n" ].forEach(function(RS) {
				var wb = X.read(data.join(RS), {type:'binary'});
				assert.equal(get_cell(wb.Sheets.Sheet1, "A1").v, 1);
				assert.equal(get_cell(wb.Sheets.Sheet1, "B3").v, "c");
				assert.equal(wb.Sheets.Sheet1['!ref'], "A1:B3");
			});
		});
		it('should handle skipHidden for rows if requested', function() {
			var baseline = "1,2,3,\nTRUE,FALSE,,sheetjs\nfoo,bar,2/19/14,0.3\n,,,\nbaz,,qux,\n";
			delete ws["!rows"];
			assert.equal(X.utils.sheet_to_csv(ws), baseline);
			assert.equal(X.utils.sheet_to_csv(ws, {skipHidden:true}), baseline);
			ws["!rows"] = [null,{hidden:true},null,null];
			assert.equal(X.utils.sheet_to_csv(ws), baseline);
			assert.equal(X.utils.sheet_to_csv(ws, {skipHidden:true}), "1,2,3,\nfoo,bar,2/19/14,0.3\n,,,\nbaz,,qux,\n");
			delete ws["!rows"];
		});
		it('should handle skipHidden for columns if requested', function() {
			var baseline = "1,2,3,\nTRUE,FALSE,,sheetjs\nfoo,bar,2/19/14,0.3\n,,,\nbaz,,qux,\n";
			delete ws["!cols"];
			assert.equal(X.utils.sheet_to_csv(ws), baseline);
			assert.equal(X.utils.sheet_to_csv(ws, {skipHidden:true}), baseline);
			ws["!cols"] = [null,{hidden:true},null,null];
			assert.equal(X.utils.sheet_to_csv(ws), baseline);
			assert.equal(X.utils.sheet_to_csv(ws, {skipHidden:true}), "1,3,\nTRUE,,sheetjs\nfoo,2/19/14,0.3\n,,\nbaz,qux,\n");
			ws["!cols"] = [{hidden:true},null,null,null];
			assert.equal(X.utils.sheet_to_csv(ws, {skipHidden:true}), "2,3,\nFALSE,,sheetjs\nbar,2/19/14,0.3\n,,\n,qux,\n");
			delete ws["!cols"];
		});
	});
});

if(fs.existsSync('./test_files/dbf/d11.dbf')) describe('dbf', function() {
	var wbs/*:Array<any>*/ = ([
		['d11',  './test_files/dbf/d11.dbf'],
		['vfp3', './test_files/dbf/vfp3.dbf']
	]/*:any*/);
	var bef = (function() {
		wbs = wbs.map(function(x) { return [x[0], X.read(fs.readFileSync(x[1]), {type:TYPE})]; });
	});
	if(typeof before != 'undefined') before(bef);
	else it('before', bef);
	it(wbs[1][0], function() {
		var ws = wbs[1][1].Sheets.Sheet1;
		[
			["A1", "v", "CHAR10"], ["A2", "v", "test1"], ["B2", "v", 123.45],
			["C2", "v", 12.345], ["D2", "v", 1234.1], ["E2", "w", "19170219"],
			/* [F2", "w", "19170219"], */ ["G2", "v", 1231.4], ["H2", "v", 123234],
			["I2", "v", true], ["L2", "v", "SheetJS"]
		].forEach(function(r) { assert.equal(get_cell(ws, r[0])[r[1]], r[2]); });
	});
});
var JSDOM = null;
// $FlowIgnore
var domtest = browser || (function(){try{return !!(JSDOM=require('jsdom').JSDOM);}catch(e){return 0;}})();

function get_dom_element(html) {
	if(browser) {
		var domelt = document.createElement('div');
		domelt.innerHTML = html;
		return domelt;
	}
	if(!JSDOM) throw new Error("Browser test fail");
	return new JSDOM(html).window.document.body.children[0];
}

describe('HTML', function() {
	describe('input string', function(){
		it('should interpret values by default', function() { plaintext_test(X.read(html_bstr, {type:"binary"}), false); });
		it('should generate strings if raw option is passed', function() { plaintext_test(X.read(html_bstr, {type:"binary", raw:true}), true); });
		it('should handle "string" type', function() { plaintext_test(X.read(html_str, {type:"string"}), false); });
		it('should handle newlines correctly', function() {
			var table = "<table><tr><td>foo<br/>bar</td><td>baz</td></tr></table>";
			var wb = X.read(table, {type:"string"});
			assert.equal(get_cell(wb.Sheets.Sheet1, "A1").v, "foo\nbar");
		});
	});
	(domtest ? describe : describe.skip)('input DOM', function() {
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
	describe('type override', function() {
		function chk(ws) {
			assert.equal(get_cell(ws, "A1").t, "s");
			assert.equal(get_cell(ws, "A1").v, "1234567890");
			assert.equal(get_cell(ws, "B1").t, "n");
			assert.equal(get_cell(ws, "B1").v, 1234567890);
		}
		var html = "<table><tr><td t=\"s\">1234567890</td><td>1234567890</td></tr></table>";
		it('HTML string', function() {
			var ws = X.read(html, {type:'string'}).Sheets.Sheet1; chk(ws);
			chk(X.read(X.utils.sheet_to_html(ws), {type:'string'}).Sheets.Sheet1);
		});
		if(domtest) it('DOM', function() { chk(X.utils.table_to_sheet(get_dom_element(html))); });
	});
});

describe('js -> file -> js', function() {
	var wb, BIN="binary";
	var bef = (function() {
		var ws = X.utils.aoa_to_sheet([
			["number", "bool", "string",  "date"],
			[1,        true,   "sheet"],
			[2,        false,  "dot"],
			[6.9,      false,  "JS", fixdate],
			[72.62,    true,   "0.3"]
		]);
		wb = { SheetNames: ['Sheet1'], Sheets: {Sheet1: ws} };
	});
	if(typeof before != 'undefined') before(bef);
	else it('before', bef);
	function eqcell(wb1, wb2, s, a) {
		assert.equal(get_cell(wb1.Sheets[s], a).v, get_cell(wb2.Sheets[s], a).v);
		assert.equal(get_cell(wb1.Sheets[s], a).t, get_cell(wb2.Sheets[s], a).t);
	}
	ofmt.forEach(function(f) {
		it(f, function() {
			var newwb = X.read(X.write(wb, {type:BIN, bookType: f}), {type:BIN});
			var cb = function(cell) { eqcell(wb, newwb, 'Sheet1', cell); };
			['A2', 'A3'].forEach(cb); /* int */
			['A4', 'A5'].forEach(cb); /* double */
			['B2', 'B3'].forEach(cb); /* bool */
			['C2', 'C3'].forEach(cb); /* string */
			if(!DIF_XL) cb('D4'); /* date */
			if(DIF_XL && f == "dif") assert.equal(get_cell(newwb.Sheets.Sheet1, 'C5').v, '=""0.3""');// dif forces string formula
			else eqcell(wb, newwb, 'Sheet1', 'C5');
		});
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
		X.utils.get_formulae(ws);
		X.utils.make_csv(ws);
		X.utils.make_json(ws);
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
		assert.throws(function() { X.utils.make_json(ws); });
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
		var ssfdata = require('./misc/ssf.json');
		var cb = function(d, j) { return function() { return X.SSF.format(d[0], d[j][0]); }; };
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
		var data = [["Row Number"]];
		for(var j = 0; j < 19; ++j) data[0].push("Column " + j+1);
		for(var i = 0; i < 499; ++i) {
			var o = ["Row " + i];
			for(j = 0; j < 19; ++j) o.push(i + j);
			data.push(o);
		}
		var ws = X.utils.aoa_to_sheet(data);
		var wb = { Sheets:{ Sheet1: ws }, SheetNames: ["Sheet1"] };
		var type = "binary";
		["xlsb", "biff8", "biff5", "biff2"].forEach(function(btype) {
			void X.read(X.write(wb, {bookType:btype, type:type}), {type:type});
		});
	});
});

describe('encryption', function() {
	password_files.forEach(function(x) {
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
			it.skip('should decrypt file', function() {
				/*var wb = */X.read(fs.readFileSync(dir + x), {type:TYPE,password:'password',WTF:opts.WTF});
			});
		});
	});
});

if(!browser || typeof cptable !== 'undefined')
describe('multiformat tests', function() {
var mfopts = opts;
var mft = fs.readFileSync('multiformat.lst','utf-8').split("\n").map(function(x) { return x.trim(); });
var csv = true, formulae = false;
mft.forEach(function(x) {
	if(x.charAt(0)!="#") describe('MFT ' + x, function() {
		var f = [], r = x.split(/\s+/);
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
		it('should have the same CSV', csv ? function() {
			cmparr(f.map(function(x) { return x.SheetNames; }));
			f[0].SheetNames.forEach(function(name) {
				cmparr(f.map(function(x) { return X.utils.sheet_to_csv(x.Sheets[name]); }));
			});
		} : null);
		it('should have the same formulae', formulae ? function() {
			cmparr(f.map(function(x) { return x.SheetNames; }));
			f[0].SheetNames.forEach(function(name) {
				cmparr(f.map(function(x) { return X.utils.sheet_to_formulae(x.Sheets[name]).sort(); }));
			});
		} : null);

	});
	else x.split(/\s+/).forEach(function(w) { switch(w) {
		case "no-csv": csv = false; break;
		case "yes-csv": csv = true; break;
		case "no-formula": formulae = false; break;
		case "yes-formula": formulae = true; break;
	}});
}); });

