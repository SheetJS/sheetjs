/* vim: set ts=2: */
/*jshint loopfunc:true, eqnull:true */
var X;
var modp = './';
//var modp = 'xlsx';
var fs = require('fs'), assert = require('assert');
describe('source',function(){it('should load',function(){X=require(modp);});});

var opts = {cellNF: true};
opts.type = "binary";
var fullex = [".xlsb", ".xlsm", ".xlsx"/*, ".xlml"*/];
var ofmt = ["xlsb", "xlsm", "xlsx", "ods", "biff2", "xlml"];
var ex = fullex.slice(); ex = ex.concat([".ods", ".xls", ".xml", ".fods"]);

/* Excel enforces 31 character sheet limit, although technical file limit is 255 */
function fixsheetname(x) { return x.substr(0,31); }

function stripbom(x) { return x.replace(/^\ufeff/,""); }
function fixcsv(x) { return stripbom(x).replace(/\t/g,",").replace(/#{255}/g,"").replace(/"/g,"").replace(/[\n\r]+/g,"\n").replace(/\n*$/,""); }
function fixjson(x) { return x.replace(/[\r\n]+$/,""); }

var dir = "./test_files/";

var paths = {
	cpxls:  dir + 'custom_properties.xls',
	cpxml:  dir + 'custom_properties.xls.xml',
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
	fstxls: dir + 'formula_stress_test.xls',
	fstxml: dir + 'formula_stress_test.xls.xml',
	fstxlsx: dir + 'formula_stress_test.xlsx',
	fstxlsb: dir + 'formula_stress_test.xlsb',
	fstods: dir + 'formula_stress_test.ods',
	hlxls:  dir + 'hyperlink_stress_test_2011.xls',
	hlxml:  dir + 'hyperlink_stress_test_2011.xml',
	hlxlsx:  dir + 'hyperlink_stress_test_2011.xlsx',
	hlxlsb:  dir + 'hyperlink_stress_test_2011.xlsb',
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
	dtxls:  dir + 'xlsx-stream-d-date-cell.xls',
	dtxml:  dir + 'xlsx-stream-d-date-cell.xls.xml',
	dtxlsx:  dir + 'xlsx-stream-d-date-cell.xlsx',
	dtxlsb:  dir + 'xlsx-stream-d-date-cell.xlsb',
	cwxls:  dir + 'column_width.xlsx',
	cwxls5:  dir + 'column_width.biff5',
	cwxml:  dir + 'column_width.xml',
	cwxlsx:  dir + 'column_width.xlsx',
	cwxlsb:  dir + 'column_width.xlsx',
	swcxls: dir + 'apachepoi_SimpleWithComments.xls',
	swcxml: dir + '2011/apachepoi_SimpleWithComments.xls.xml',
	swcxlsx: dir + 'apachepoi_SimpleWithComments.xlsx',
	swcxlsb: dir + '2013/apachepoi_SimpleWithComments.xlsx.xlsb'
};

var N1 = 'XLSX';
var N2 = 'XLSB';
var N3 = 'XLS';
var N4 = 'XML';

describe('parse options', function() {
	var html_cell_types = ['s'];
	var bef = (function() {
		X = require(modp);
	});
	if(typeof before != 'undefined') before(bef);
	else it('before', bef);
	describe('cell', function() {
		it('XLSX should generate HTML by default', function() {
			var wb = X.read(fs.readFileSync(paths.cstxlsx), {type:"binary"});
			var ws = wb.Sheets.Sheet1;
			Object.keys(ws).forEach(function(addr) {
				if(addr[0] === "!" || !ws.hasOwnProperty(addr)) return;
				assert(html_cell_types.indexOf(ws[addr].t) === -1 || ws[addr].h);
			});
		});
		it('XLSX should not generate HTML when requested', function() {
			var wb = X.read(fs.readFileSync(paths.cstxlsx), {type:"binary", cellHTML:false});
			var ws = wb.Sheets.Sheet1;
			Object.keys(ws).forEach(function(addr) {
				if(addr[0] === "!" || !ws.hasOwnProperty(addr)) return;
				assert(typeof ws[addr].h === 'undefined');
			});
		});
		it('should generate formulae by default', function() {
			[paths.fstxls, paths.fstxlsb].forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:"binary"});
				var found = false;
				wb.SheetNames.forEach(function(s) {
					var ws = wb.Sheets[s];
					Object.keys(ws).forEach(function(addr) {
						if(addr[0] === "!" || !ws.hasOwnProperty(addr)) return;
						if(typeof ws[addr].f !== 'undefined') return (found = true);
					});
				});
				assert(found);
			});
		});
		it('should not generate formulae when requested', function() {
			[paths.fstxls, paths.fstxlsb].forEach(function(p) {
				var wb =X.read(fs.readFileSync(p),{type:"binary",cellFormula:false});
				wb.SheetNames.forEach(function(s) {
					var ws = wb.Sheets[s];
					Object.keys(ws).forEach(function(addr) {
						if(addr[0] === "!" || !ws.hasOwnProperty(addr)) return;
						assert(typeof ws[addr].f === 'undefined');
					});
				});
			});
		});
		it('should not generate number formats by default', function() {
			[paths.nfxls, paths.nfxlsx].forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:"binary"});
				wb.SheetNames.forEach(function(s) {
					var ws = wb.Sheets[s];
					Object.keys(ws).forEach(function(addr) {
						if(addr[0] === "!" || !ws.hasOwnProperty(addr)) return;
						assert(typeof ws[addr].z === 'undefined');
					});
				});
			});
		});
		it('should generate number formats when requested', function() {
			[paths.nfxls, paths.nfxlsx].forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:"binary", cellNF: true});
				wb.SheetNames.forEach(function(s) {
					var ws = wb.Sheets[s];
					Object.keys(ws).forEach(function(addr) {
						if(addr[0] === "!" || !ws.hasOwnProperty(addr)) return;
						assert(ws[addr].t!== 'n' || typeof ws[addr].z !== 'undefined');
					});
				});
			});
		});
		it('should not generate cell styles by default', function() {
			[paths.cssxlsx, paths.cssxls, paths.cssxml].forEach(function(p) {
			var wb = X.read(fs.readFileSync(p), {type:"binary", WTF:1});
			wb.SheetNames.forEach(function(s) {
				var ws = wb.Sheets[s];
				Object.keys(ws).forEach(function(addr) {
					if(addr[0] === "!" || !ws.hasOwnProperty(addr)) return;
					assert(typeof ws[addr].s === 'undefined');
				});
			});
			});
		});
		it('should generate cell styles when requested', function() {
			/* TODO: XLS / XLML */
			[paths.cssxlsx /*,paths.cssxls, paths.cssxml*/].forEach(function(p) {
			var wb = X.read(fs.readFileSync(p), {type:"binary", cellStyles:true});
			var found = false;
			wb.SheetNames.forEach(function(s) {
				var ws = wb.Sheets[s];
				Object.keys(ws).forEach(function(addr) {
					if(addr[0] === "!" || !ws.hasOwnProperty(addr)) return;
					if(typeof ws[addr].s !== 'undefined') return (found = true);
				});
			});
			assert(found);
			});
		});
		it('should not generate cell dates by default', function() {
			var wb = X.read(fs.readFileSync(paths.dtxlsx), {type:"binary"});
			wb.SheetNames.forEach(function(s) {
				var ws = wb.Sheets[s];
				Object.keys(ws).forEach(function(addr) {
					if(addr[0] === "!" || !ws.hasOwnProperty(addr)) return;
					assert(ws[addr].t !== 'd');
				});
			});
		});
		it('XLSX should generate cell dates when requested', function() {
			var wb = X.read(fs.readFileSync(paths.dtxlsx), {type:"binary", cellDates: true, WTF:1});
			var found = false;
			wb.SheetNames.forEach(function(s) {
				var ws = wb.Sheets[s];
				Object.keys(ws).forEach(function(addr) {
					if(addr[0] === "!" || !ws.hasOwnProperty(addr)) return;
					if(ws[addr].t === 'd') return (found = true);
				});
			});
			assert(found);
		});
	});
	describe('sheet', function() {
		it('should not generate sheet stubs by default', function() {
			[paths.mcxlsx, paths.mcxlsb, paths.mcods, paths.mcxls, paths.mcxml].forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:"binary"});
				assert.throws(function() { return wb.Sheets.Merge.A2.v; });
			});
		});
		it('should generate sheet stubs when requested', function() {
			[paths.mcxlsx, paths.mcxlsb, paths.mcods, paths.mcxls, paths.mcxml].forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:"binary", sheetStubs:true});
				assert(wb.Sheets.Merge.A2.t == 'z');
			});
		});
		it('should handle stub cells', function() {
			[paths.mcxlsx, paths.mcxlsb, paths.mcods, paths.mcxls, paths.mcxml].forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:"binary", sheetStubs:true});
				X.utils.sheet_to_csv(wb.Sheets.Merge);
				X.utils.sheet_to_json(wb.Sheets.Merge);
				X.utils.sheet_to_formulae(wb.Sheets.Merge);
				ofmt.forEach(function(f) { X.write(wb, {type:"binary", bookType:f}); });
			});
		});
		function checkcells(wb, A46, B26, C16, D2) {
			assert((typeof wb.Sheets.Text.A46 !== 'undefined') == A46);
			assert((typeof wb.Sheets.Text.B26 !== 'undefined') == B26);
			assert((typeof wb.Sheets.Text.C16 !== 'undefined') == C16);
			assert((typeof wb.Sheets.Text.D2  !== 'undefined') == D2);
		}
		it('should read all cells by default', function() {
			[paths.fstxlsx, paths.fstxlsb, paths.fstods, paths.fstxls, paths.fstxml].forEach(function(p) {
				checkcells(X.read(fs.readFileSync(p), {type:"binary"}), true, true, true, true);
			});
		});
		it('sheetRows n=20', function() {
			[paths.fstxlsx, paths.fstxlsb, paths.fstods, paths.fstxls, paths.fstxml].forEach(function(p) {
				checkcells(X.read(fs.readFileSync(p), {type:"binary", sheetRows:20}), false, false, true, true);
			});
		});
		it('sheetRows n=10', function() {
			[paths.fstxlsx, paths.fstxlsb, paths.fstods, paths.fstxls, paths.fstxml].forEach(function(p) {
				checkcells(X.read(fs.readFileSync(p), {type:"binary", sheetRows:10}), false, false, false, true);
			});
		});
	});
	describe('book', function() {
		it('bookSheets should not generate sheets', function() {
			[paths.mcxlsx, paths.mcxlsb, paths.mcxls, paths.mcxml].forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:"binary", bookSheets:true});
				assert(typeof wb.Sheets === 'undefined');
			});
		});
		it('bookProps should not generate sheets', function() {
			[paths.nfxlsx, paths.nfxlsb, paths.nfxls, paths.nfxml].forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:"binary", bookProps:true});
				assert(typeof wb.Sheets === 'undefined');
			});
		});
		it('bookProps && bookSheets should not generate sheets', function() {
			[paths.lonxlsx, paths.lonxls].forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:"binary", bookProps:true, bookSheets:true});
				assert(typeof wb.Sheets === 'undefined');
			});
		});
		it('should not generate deps by default', function() {
			[paths.fstxlsx, paths.fstxlsb, paths.fstxls, paths.fstxml].forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:"binary"});
				assert(typeof wb.Deps === 'undefined' || !(wb.Deps && wb.Deps.length>0));
			});
		});
		it('bookDeps should generate deps (XLSX/XLSB)', function() {
			[paths.fstxlsx, paths.fstxlsb].forEach(function(p) {
				var wb = X.read(fs.readFileSync(p), {type:"binary", bookDeps:true});
				assert(typeof wb.Deps !== 'undefined' && wb.Deps.length > 0);
			});
		});
		var ckf = function(wb, fields, exists) { fields.forEach(function(f) {
			assert((typeof wb[f] !== 'undefined') == exists);
		}); };
		it('should not generate book files by default', function() {
			var wb;
			wb = X.read(fs.readFileSync(paths.fstxlsx), {type:"binary"});
			ckf(wb, ['files', 'keys'], false);
			wb = X.read(fs.readFileSync(paths.fstxlsb), {type:"binary"});
			ckf(wb, ['files', 'keys'], false);
			wb = X.read(fs.readFileSync(paths.fstxls), {type:"binary"});
			ckf(wb, ['cfb'], false);
		});
		it('bookFiles should generate book files', function() {
			var wb;
			wb = X.read(fs.readFileSync(paths.fstxlsx), {type:"binary", bookFiles:true});
			ckf(wb, ['files', 'keys'], true);
			wb = X.read(fs.readFileSync(paths.fstxlsb), {type:"binary", bookFiles:true});
			ckf(wb, ['files', 'keys'], true);
			wb = X.read(fs.readFileSync(paths.fstxls), {type:"binary", bookFiles:true});
			ckf(wb, ['cfb'], true);
		});
		it('should not generate VBA by default', function() {
			var wb = X.read(fs.readFileSync(paths.nfxlsx), {type:"binary"});
			assert(typeof wb.vbaraw === 'undefined');
			wb = X.read(fs.readFileSync(paths.nfxlsb), {type:"binary"});
			assert(typeof wb.vbaraw === 'undefined');
		});
		it('bookVBA should generate vbaraw (XLSX/XLSB)', function() {
			var wb = X.read(fs.readFileSync(paths.nfxlsx),{type:"binary", bookVBA:true});
			assert(wb.vbaraw);
			wb = X.read(fs.readFileSync(paths.nfxlsb),{type:"binary", bookVBA:true});
			assert(wb.vbaraw);
		});
	});
});

describe('input formats', function() {
	it('should read binary strings', function() {
		X.read(fs.readFileSync(paths.cstxlsx, 'binary'), {type: 'binary'});
		X.read(fs.readFileSync(paths.cstxlsb, 'binary'), {type: 'binary'});
		X.read(fs.readFileSync(paths.cstxls, 'binary'), {type: 'binary'});
		X.read(fs.readFileSync(paths.cstxml, 'binary'), {type: 'binary'});
	});
	it('should read base64 strings', function() {
		X.read(fs.readFileSync(paths.cstxls, 'base64'), {type: 'base64'});
		X.read(fs.readFileSync(paths.cstxml, 'base64'), {type: 'base64'});
		X.read(fs.readFileSync(paths.cstxlsx, 'base64'), {type: 'base64'});
		X.read(fs.readFileSync(paths.cstxlsb, 'base64'), {type: 'base64'});
	});
	(typeof UInt8Array !== 'undefined' ? it : it.skip)('should read arrays', function() {
		X.read(fs.readFileSync(paths.cstxls, 'buffer'), {type: 'array'});
		X.read(fs.readFileSync(paths.cstxml, 'buffer'), {type: 'array'});
		X.read(fs.readFileSync(paths.cstxlsx, 'buffer'), {type: 'array'});
		X.read(fs.readFileSync(paths.cstxlsb, 'buffer'), {type: 'array'});
	});
	(typeof UInt8Array !== 'undefined' ? it : it.skip)('should read array', function() {
		X.read(fs.readFileSync(paths.mcxls, 'binary').split("").map(function(x) { return x.charCodeAt(0); }), {type:'array'});
		X.read(fs.readFileSync(paths.mcxml, 'binary').split("").map(function(x) { return x.charCodeAt(0); }), {type:'array'});
		X.read(fs.readFileSync(paths.mcxlsx, 'binary').split("").map(function(x) { return x.charCodeAt(0); }), {type:'array'});
		X.read(fs.readFileSync(paths.mcxlsb, 'binary').split("").map(function(x) { return x.charCodeAt(0); }), {type:'array'});
		X.read(fs.readFileSync(paths.mcods, 'binary').split("").map(function(x) { return x.charCodeAt(0); }), {type:'array'});
	});
	it('should throw if format is unknown', function() {
		assert.throws(function() { X.read(fs.readFileSync(paths.cstxls), {type: 'dafuq'}); });
		assert.throws(function() { X.read(fs.readFileSync(paths.cstxml), {type: 'dafuq'}); });
		assert.throws(function() { X.read(fs.readFileSync(paths.cstxlsx), {type: 'dafuq'}); });
		assert.throws(function() { X.read(fs.readFileSync(paths.cstxlsb), {type: 'dafuq'}); });
	});
	it('should default to base64 type', function() {
		assert.throws(function() { X.read(fs.readFileSync(paths.cstxls, 'binary')); });
		assert.throws(function() { X.read(fs.readFileSync(paths.cstxml, 'binary')); });
		assert.throws(function() { X.read(fs.readFileSync(paths.cstxlsx, 'binary')); });
		assert.throws(function() { X.read(fs.readFileSync(paths.cstxlsb, 'binary')); });
		X.read(fs.readFileSync(paths.cstxls, 'base64'));
		X.read(fs.readFileSync(paths.cstxml, 'base64'));
		X.read(fs.readFileSync(paths.cstxlsx, 'base64'));
		X.read(fs.readFileSync(paths.cstxlsb, 'base64'));
	});
});

describe('output formats', function() {
	var wb1, wb2, wb3, wb4;
	var bef = (function() {
		X = require(modp);
		wb1 = X.read(fs.readFileSync(paths.cpxlsx), {type:"binary"});
		wb2 = X.read(fs.readFileSync(paths.cpxlsb), {type:"binary"});
		wb3 = X.read(fs.readFileSync(paths.cpxls), {type:"binary"});
		wb4 = X.read(fs.readFileSync(paths.cpxml), {type:"binary"});
	});
	if(typeof before != 'undefined') before(bef);
	else it('before', bef);
	it('should write binary strings', function() {
		if(!wb1) {
			wb1 = X.read(fs.readFileSync(paths.cpxlsx), {type:"binary"});
			wb2 = X.read(fs.readFileSync(paths.cpxlsb), {type:"binary"});
			wb3 = X.read(fs.readFileSync(paths.cpxls), {type:"binary"});
			wb4 = X.read(fs.readFileSync(paths.cpxml), {type:"binary"});
		}
		X.write(wb1, {type: 'binary'});
		X.write(wb2, {type: 'binary'});
		X.write(wb3, {type: 'binary'});
		X.write(wb4, {type: 'binary'});
		X.read(X.write(wb1, {type: 'binary'}), {type: 'binary'});
		X.read(X.write(wb2, {type: 'binary'}), {type: 'binary'});
		X.read(X.write(wb3, {type: 'binary'}), {type: 'binary'});
		X.read(X.write(wb4, {type: 'binary'}), {type: 'binary'});
	});
	it('should write base64 strings', function() {
		X.write(wb1, {type: 'base64'});
		X.write(wb2, {type: 'base64'});
		X.write(wb3, {type: 'base64'});
		X.write(wb4, {type: 'base64'});
		X.read(X.write(wb1, {type: 'base64'}), {type: 'base64'});
		X.read(X.write(wb2, {type: 'base64'}), {type: 'base64'});
		X.read(X.write(wb3, {type: 'base64'}), {type: 'base64'});
		X.read(X.write(wb4, {type: 'base64'}), {type: 'base64'});
	});
	it('should throw if format is unknown', function() {
		assert.throws(function() { X.write(wb1, {type: 'dafuq'}); });
		assert.throws(function() { X.write(wb2, {type: 'dafuq'}); });
		assert.throws(function() { X.write(wb3, {type: 'dafuq'}); });
		assert.throws(function() { X.write(wb4, {type: 'dafuq'}); });
	});
});

function coreprop(wb) {
	assert.equal(wb.Props.Title, 'Example with properties');
	assert.equal(wb.Props.Subject, 'Test it before you code it');
	assert.equal(wb.Props.Author, 'Pony Foo');
	assert.equal(wb.Props.Manager, 'Despicable Drew');
	assert.equal(wb.Props.Company, 'Vector Inc');
	assert.equal(wb.Props.Category, 'Quirky');
	assert.equal(wb.Props.Keywords, 'example humor');
	assert.equal(wb.Props.Comments, 'some comments');
	assert.equal(wb.Props.LastAuthor, 'Hugues');
}
function custprop(wb) {
	assert.equal(wb.Custprops['I am a boolean'], true);
	assert.equal(wb.Custprops['Date completed'].toISOString(), '1967-03-09T16:30:00.000Z');
	assert.equal(wb.Custprops.Status, 2);
	assert.equal(wb.Custprops.Counter, -3.14);
}

function cmparr(x){ for(var i=1;i!=x.length;++i) assert.deepEqual(x[0], x[i]); }

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
	var c1 = ws[r1].s, c2 = ws[r2].s;
	stykeys.forEach(function(m) {
		var c = -1;
		if(styexc.indexOf(r1+"|"+r2+"|"+m) > -1) c = 1;
		else if(styexc.indexOf(r2+"|"+r1+"|"+m) > -1) c = 1;
		deepcmp(c1,c2,m,r1+","+r2,c);
	});
}

describe('parse features', function() {
	if(fs.existsSync(paths.swcxlsx)) it('should have comment as part of cell properties', function(){
		var X = require(modp);
		var sheet = 'Sheet1';
		var wb1=X.read(fs.readFileSync(paths.swcxlsx), {type:"binary"});
		var wb2=X.read(fs.readFileSync(paths.swcxlsb), {type:"binary"});
		var wb3=X.read(fs.readFileSync(paths.swcxls), {type:"binary"});
		var wb4=X.read(fs.readFileSync(paths.swcxml), {type:"binary"});

		[wb1,wb2,wb3,wb4].map(function(wb) { return wb.Sheets[sheet]; }).forEach(function(ws, i) {
			assert.equal(ws.B1.c.length, 1,"must have 1 comment");
			assert.equal(ws.B1.c[0].a, "Yegor Kozlov","must have the same author");
			assert.equal(ws.B1.c[0].t.replace(/\r\n/g,"\n").replace(/\r/g,"\n"), "Yegor Kozlov:\nfirst cell", "must have the concatenated texts");
			if(i > 0) return;
			assert.equal(ws.B1.c[0].r, '<r><rPr><b/><sz val="8"/><color indexed="81"/><rFont val="Tahoma"/></rPr><t>Yegor Kozlov:</t></r><r><rPr><sz val="8"/><color indexed="81"/><rFont val="Tahoma"/></rPr><t xml:space="preserve">\r\nfirst cell</t></r>', "must have the rich text representation");
			assert.equal(ws.B1.c[0].h, '<span style="font-weight: bold;">Yegor Kozlov:</span><span style=""><br/>first cell</span>', "must have the html representation");
		});
	});

	describe('should parse core properties and custom properties', function() {
		var wb1, wb2, wb3, wb4;
		var bef = (function() {
			wb1 = X.read(fs.readFileSync(paths.cpxlsx), {type:"binary"});
			wb2 = X.read(fs.readFileSync(paths.cpxlsb), {type:"binary"});
			wb3 = X.read(fs.readFileSync(paths.cpxls), {type:"binary"});
			wb4 = X.read(fs.readFileSync(paths.cpxml), {type:"binary"});
		});
		if(typeof before != 'undefined') before(bef);
		else it('before', bef);

		it(N1 + ' should parse core properties', function() { coreprop(wb1); });
		it(N2 + ' should parse core properties', function() { coreprop(wb2); });
		it(N3 + ' should parse core properties', function() { coreprop(wb3); });
		it(N4 + ' should parse core properties', function() { coreprop(wb4); });
		it(N1 + ' should parse custom properties', function() { custprop(wb1); });
		it(N2 + ' should parse custom properties', function() { custprop(wb2); });
		it(N3 + ' should parse custom properties', function() { custprop(wb3); });
		it(N4 + ' should parse custom properties', function() { custprop(wb4); });
	});

	describe('sheetRows', function() {
		it('should use original range if not set', function() {
			var opts = {type:"binary"};
			var wb1 = X.read(fs.readFileSync(paths.fstxlsx), opts);
			var wb2 = X.read(fs.readFileSync(paths.fstxlsb), opts);
			var wb3 = X.read(fs.readFileSync(paths.fstxls), opts);
			var wb4 = X.read(fs.readFileSync(paths.fstxml), opts);
			[wb1, wb2, wb3, wb4].forEach(function(wb) {
				assert.equal(wb.Sheets.Text["!ref"],"A1:F49");
			});
		});
		it('should adjust range if set', function() {
			var opts = {type:"binary", sheetRows:10};
			var wb1 = X.read(fs.readFileSync(paths.fstxlsx), opts);
			var wb2 = X.read(fs.readFileSync(paths.fstxlsb), opts);
			var wb3 = X.read(fs.readFileSync(paths.fstxls), opts);
			var wb4 = X.read(fs.readFileSync(paths.fstxml), opts);
			/* TODO */
			[wb1, wb2 /*, wb3, wb4 */].forEach(function(wb) {
				assert.equal(wb.Sheets.Text["!fullref"],"A1:F49");
				assert.equal(wb.Sheets.Text["!ref"],"A1:F10");
			});
		});
		it('should not generate comment cells', function() {
			var opts = {type:"binary", sheetRows:10};
			var wb1 = X.read(fs.readFileSync(paths.cstxlsx), opts);
			var wb2 = X.read(fs.readFileSync(paths.cstxlsb), opts);
			var wb3 = X.read(fs.readFileSync(paths.cstxls), opts);
			var wb4 = X.read(fs.readFileSync(paths.cstxml), opts);
			/* TODO */
			[wb1, wb2 /*, wb3, wb4 */].forEach(function(wb) {
				assert.equal(wb.Sheets.Sheet7["!fullref"],"A1:N34");
				assert.equal(wb.Sheets.Sheet7["!ref"],"A1");
			});
		});
	});

	describe('column properties', function() {
		var wb1, wb2, wb3, wb4, wb5;
		var bef = (function() {
			X = require(modp);
			wb1 = X.read(fs.readFileSync(paths.cwxlsx), {type:"binary", cellStyles:true});
			wb2 = X.read(fs.readFileSync(paths.cwxlsb), {type:"binary", cellStyles:true});
			wb3 = X.read(fs.readFileSync(paths.cwxls), {type:"binary", cellStyles:true});
			wb4 = X.read(fs.readFileSync(paths.cwxls5), {type:"binary", cellStyles:true});
			wb5 = X.read(fs.readFileSync(paths.cwxml), {type:"binary", cellStyles:true});
		});
		if(typeof before != 'undefined') before(bef);
		else it('before', bef);
		it('should have "!cols"', function() {
			assert(wb1.Sheets.Sheet1['!cols']);
			assert(wb2.Sheets.Sheet1['!cols']);
			assert(wb3.Sheets.Sheet1['!cols']);
			assert(wb4.Sheets.Sheet1['!cols']);
			assert(wb5.Sheets.Sheet1['!cols']);
		});
		it('should have correct widths', function() {
			[wb1, wb2, wb3, wb4, wb5].map(function(x) { return x.Sheets.Sheet1['!cols']; }).forEach(function(x) {
				assert.equal(x[1].width, 0.1640625);
				assert.equal(x[2].width, 16.6640625);
				assert.equal(x[3].width, 1.6640625);
				assert.equal(x[4].width, 4.83203125);
				assert.equal(x[5].width, 8.83203125);
				assert.equal(x[6].width, 12.83203125);
				assert.equal(x[7].width, 16.83203125);
			});
		});
		it('should have correct pixels', function() {
			[wb1, wb2, wb3, wb4, wb5].map(function(x) { return x.Sheets.Sheet1['!cols']; }).forEach(function(x) {
				assert.equal(x[1].wpx, 1);
				assert.equal(x[2].wpx, 100);
				assert.equal(x[3].wpx, 10);
				assert.equal(x[4].wpx, 29);
				assert.equal(x[5].wpx, 53);
				assert.equal(x[6].wpx, 77);
				assert.equal(x[7].wpx, 101);
			});
		});
	});

	describe('merge cells',function() {
		var wb1, wb2, wb3, wb4, wb5;
		var bef = (function() {
			X = require(modp);
			wb1 = X.read(fs.readFileSync(paths.mcxlsx), {type:"binary"});
			wb2 = X.read(fs.readFileSync(paths.mcxlsb), {type:"binary"});
			wb3 = X.read(fs.readFileSync(paths.mcods), {type:"binary"});
			wb4 = X.read(fs.readFileSync(paths.mcxls), {type:"binary"});
			wb5 = X.read(fs.readFileSync(paths.mcxml), {type:"binary"});
		});
		if(typeof before != 'undefined') before(bef);
		else it('before', bef);
		it('should have !merges', function() {
			assert(wb1.Sheets.Merge['!merges']);
			assert(wb2.Sheets.Merge['!merges']);
			assert(wb3.Sheets.Merge['!merges']);
			assert(wb4.Sheets.Merge['!merges']);
			assert(wb5.Sheets.Merge['!merges']);
			var m = [wb1, wb2, wb3, wb4, wb5].map(function(x) { return x.Sheets.Merge['!merges'].map(function(y) { return X.utils.encode_range(y); });});
			assert.deepEqual(m[0].sort(),m[1].sort());
			assert.deepEqual(m[0].sort(),m[2].sort());
			assert.deepEqual(m[0].sort(),m[3].sort());
			assert.deepEqual(m[0].sort(),m[4].sort());
		});
	});

	describe('should find hyperlinks', function() {
		var wb1, wb2, wb3, wb4;
		var bef = (function() {
			X = require(modp);
			wb1 = X.read(fs.readFileSync(paths.hlxlsx), {type:"binary"});
			wb2 = X.read(fs.readFileSync(paths.hlxlsb), {type:"binary"});
			wb3 = X.read(fs.readFileSync(paths.hlxls), {type:"binary"});
			wb4 = X.read(fs.readFileSync(paths.hlxml), {type:"binary"});
		});
		if(typeof before != 'undefined') before(bef);
		else it('before', bef);

		function hlink(wb) {
			var ws = wb.Sheets.Sheet1;
			assert.equal(ws.A1.l.Target, "http://www.sheetjs.com");
			assert.equal(ws.A2.l.Target, "http://oss.sheetjs.com");
			assert.equal(ws.A3.l.Target, "http://oss.sheetjs.com#foo");
			assert.equal(ws.A4.l.Target, "mailto:dev@sheetjs.com");
			assert.equal(ws.A5.l.Target, "mailto:dev@sheetjs.com?subject=hyperlink");
			assert.equal(ws.A6.l.Target, "../../sheetjs/Documents/Test.xlsx");
			assert.equal(ws.A7.l.Target, "http://sheetjs.com");
		}

		it(N1, function() { hlink(wb1); });
		it(N2, function() { hlink(wb2); });
		it(N3, function() { hlink(wb3); });
		it(N4, function() { hlink(wb4); });
	});

	describe('should parse cells with date type (XLSX/XLSM)', function() {
		it('Must have read the date', function() {
			var wb, ws;
			var sheetName = 'Sheet1';
			wb = X.read(fs.readFileSync(paths.dtxlsx), {type:"binary"});
			ws = wb.Sheets[sheetName];
			var sheet = X.utils.sheet_to_row_object_array(ws);
			assert.equal(sheet[3]['てすと'], '2/14/14');
		});
		it('cellDates should not affect formatted text', function() {
			var wb1, ws1, wb2, ws2;
			var sheetName = 'Sheet1';
			wb1 = X.read(fs.readFileSync(paths.dtxlsx), {type:"binary"});
			ws1 = wb1.Sheets[sheetName];
			wb2 = X.read(fs.readFileSync(paths.dtxlsb), {type:"binary"});
			ws2 = wb2.Sheets[sheetName];
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
			wb = X.read(fs.readFileSync(f[1]), {type:"binary"});
			ws = wb.Sheets[f[2]];
			assert.equal(ws[f[3]].w, f[4]);
			assert.equal(ws[f[3]].t, 'n');
		}); });
		it('should generate date cells if cellDates is true', function() { fmts.forEach(function(f) {
			var wb, ws;
			wb = X.read(fs.readFileSync(f[1]), {type:"binary", cellDates:true});
			ws = wb.Sheets[f[2]];
			assert.equal(ws[f[3]].w, f[4]);
			assert.equal(ws[f[3]].t, 'd');
		}); });
	});

	describe('should correctly handle styles', function() {
		var wsxls, wsxlsx, rn, rn2;
		var bef = (function() {
			wsxls=X.read(fs.readFileSync(paths.cssxls), {type:"binary",cellStyles:true,WTF:1}).Sheets.Sheet1;
			wsxlsx=X.read(fs.readFileSync(paths.cssxlsx), {type:"binary",cellStyles:true,WTF:1}).Sheets.Sheet1;
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
		var exp = [
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
		ranges.forEach(function(rng) {
			it('XLS  | ' + rng,function(){cmparr(rn2(rng).map(function(x){ return wsxls[x].s; }));});
			it('XLSX | ' + rng,function(){cmparr(rn2(rng).map(function(x){ return wsxlsx[x].s; }));});
		});
		it('different styles', function() {
			for(var i = 0; i != ranges.length-1; ++i) {
				for(var j = i+1; j != ranges.length; ++j) {
					diffsty(wsxlsx, rn2(ranges[i])[0], rn2(ranges[j])[0]);
					/* TODO */
					//diffsty(wsxls, rn2(ranges[i])[0], rn2(ranges[j])[0]);
				}
			}
		});
		it('correct styles', function() {
			var stylesxls = ranges.map(function(r) { return rn2(r)[0]; }).map(function(r) { return wsxls[r].s; });
			var stylesxlsx = ranges.map(function(r) { return rn2(r)[0]; }).map(function(r) { return wsxlsx[r].s; });
			for(var i = 0; i != exp.length; ++i) {
				[
					"fgColor.theme","fgColor.raw_rgb",
					"bgColor.theme","bgColor.raw_rgb",
					"patternType"
				].forEach(function(k) {
					deepcmp(exp[i], stylesxlsx[i], k, i + ":"+k);
					/* TODO */
					//deepcmp(exp[i], stylesxls[i], k, i + ":"+k);
				});
			}
		});
	});
});

function seq(end, start) {
	var s = start || 0;
	var o = new Array(end - s);
	for(var i = 0; i != o.length; ++i) o[i] = s + i;
	return o;
}

describe('roundtrip features', function() {
	var bef = (function() { X = require(modp); });
	if(typeof before != 'undefined') before(bef);
	else it('before', bef);
	describe('should parse core properties and custom properties', function() {
		var wb1, wb2, base = './tmp/cp';
		var bef = (function() {
			wb1 = X.read(fs.readFileSync(paths.cpxlsx), {type:"binary"});
			wb2 = X.read(fs.readFileSync(paths.cpxlsb), {type:"binary"});
			fullex.forEach(function(p) {
				fs.writeFileSync(base + '.xlsm' + p, X.write(wb1, {type:"base64", bookType:p.replace(/\./g,"")}));
				fs.writeFileSync(base + '.xlsb' + p, X.write(wb2, {type:"base64", bookType:p.replace(/\./g,"")}));
			});
		});
		if(typeof before != 'undefined') before(bef);
		else it('before', bef);
		fullex.forEach(function(p) { ['.xlsm','.xlsb'].forEach(function(q) {
			it(q + p + ' should roundtrip core and custom properties', function() {
				var wb = X.read(fs.readFileSync(base + q + p), {type:"binary"});
				coreprop(wb);
				custprop(wb);
			}); });
		});
	});

	describe('should preserve features', function() {
		it('merge cells', function() {
		["xlsx", "xlsb", "xlml", "ods"].forEach(function(f) {
			var wb1 = X.read(fs.readFileSync(paths.mcxlsx), {type:"binary"});
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
				var wb1 = X.read(fs.readFileSync(f), {type:"binary", cellNF: true, cellDates: di, WTF: opts.WTF || true});
				var wb2 = X.read(X.write(wb1, {type:'binary', cellDates:dj, WTF:opts.WTF||true}), {type:'binary', cellDates: dk, WTF: opts.WTF||true});
				var m = [wb1,wb2].map(function(x) { return x.Sheets[sheet][addr]; });
				assert.equal(m[0].w, m[1].w);

				assert.equal(m[0].t, b);
				assert.equal(m[1].t, d);

				if(m[0].t === 'n' && m[1].t === 'n') assert.equal(m[0].v, m[1].v);
				else if(m[0].t === 'd' && m[1].t === 'd') assert.equal(m[0].v.toString(), m[1].v.toString());
				else if(m[1].t === 'n') assert(Math.abs(datenum(parseDate(m[0].v)) - m[1].v) < 0.01); /* TODO: 1sec adjustment */
			});
		});
	});

	describe('should preserve formulae', function() { [
			['xlml', paths.fstxml],
			['xlsx', paths.fstxlsx],
			['ods',  paths.fstods]
		].forEach(function(w) {
			it(w[0], function() {
				var wb1 = X.read(fs.readFileSync(w[1]), {type:"binary", cellFormula:true});
				var wb2 = X.read(X.write(wb1, {bookType:w[0], type:"binary"}), {cellFormula:true, type:"binary"});
				wb1.SheetNames.forEach(function(n) {
					assert.deepEqual( X.utils.sheet_to_formulae(wb1.Sheets[n]).sort(), X.utils.sheet_to_formulae(wb2.Sheets[n]).sort() );
				});
			});
		});
	});
});

function password_file(x){return x.match(/^password.*\.xls$/); }
var password_files = fs.readdirSync('test_files').filter(password_file);
describe('invalid files', function() {
	describe('parse', function() { [
			['password', 'apachepoi_password.xls'],
			['passwords', 'apachepoi_xor-encryption-abc.xls'],
			['DOC files', 'word_doc.doc']
		].forEach(function(w) { it('should fail on ' + w[0], function() {
			assert.throws(function() { X.read(fs.readFileSync(dir + w[1]), {type:"binary"}); });
			assert.throws(function() { X.read(fs.readFileSync(dir+w[1], 'base64'), {type:'base64'}); });
		}); });
	});
	describe('write', function() {
		it('should pass -> XLSX', function() {
			X.write(X.read(fs.readFileSync(paths.fstxlsb), {type:'binary'}), {type:'binary'});
			X.write(X.read(fs.readFileSync(paths.fstxlsx), {type:'binary'}), {type:'binary'});
			X.write(X.read(fs.readFileSync(paths.fstxls), {type:'binary'}), {type:'binary'});
			X.write(X.read(fs.readFileSync(paths.fstxml), {type:'binary'}), {type:'binary'});
		});
		it('should pass if a sheet is missing', function() {
			var wb = X.read(fs.readFileSync(paths.fstxlsx), {type:"binary"}); delete wb.Sheets[wb.SheetNames[0]];
			X.read(X.write(wb, {type:'binary'}), {type:'binary'});
		});
		['Props', 'Custprops', 'SSF'].forEach(function(t) {
			it('should pass if ' + t + ' is missing', function() {
				var wb = X.read(fs.readFileSync(paths.fstxlsx), {type:"binary"});
				assert.doesNotThrow(function() {
					delete wb[t];
					X.write(wb, {type:'binary'});
				});
			});
		});
		['SheetNames', 'Sheets'].forEach(function(t) {
			it('should fail if ' + t + ' is missing', function() {
				var wb = X.read(fs.readFileSync(paths.fstxlsx), {type:"binary"});
				assert.throws(function() {
					delete wb[t];
					X.write(wb, {type:'binary'});
				});
			});
		});
		it('should fail if SheetNames has duplicate entries', function() {
			var wb = X.read(fs.readFileSync(paths.fstxlsx), {type:"binary"});
			wb.SheetNames.push(wb.SheetNames[0]);
			assert.throws(function() {
				X.write(wb, {type:'binary'});
			});
		});
	});
});

function datenum(v/*:Date*/, date1904/*:?boolean*/)/*:number*/ {
	var epoch = v.getTime();
	if(date1904) epoch += 1462*24*60*60*1000;
	return (epoch + 2209161600000) / (24 * 60 * 60 * 1000);
}
var good_pd_date = new Date('2017-02-19T19:06:09.000Z');
var good_pd = good_pd_date.getFullYear() == 2017;
function parseDate(str/*:string|Date*/)/*:Date*/ {
	if(good_pd) return new Date(str);
	if(str instanceof Date) return str;
	var n = str.match(/\d+/g);
	return new Date(Date.UTC(+n[0], n[1] - 1, +n[2], +n[3], +n[4], +n[5]));
}

describe('json output', function() {
	function seeker(json, keys, val) {
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
			["foo","bar", parseDate("2014-02-19T14:30:00.000Z"), "0.3"],
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
		var json = X.utils.sheet_to_json(ws, {raw:true});
		var json = X.utils.sheet_to_json(ws, {raw:true, defval: 'jimjin'});
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
		assert.equal(json[2][2].getTime(), parseDate("2014-02-19T14:30:00.000Z").getTime());
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
		var json2 = X.utils.sheet_to_json(_ws, {header:1});
		var json3 = X.utils.sheet_to_json(_ws, {blankrows:true});
		var json4 = X.utils.sheet_to_json(_ws, {blankrows:true, header:1});
		var json5 = X.utils.sheet_to_json(_ws, {blankrows:false});
		var json6 = X.utils.sheet_to_json(_ws, {blankrows:false, header:1});
		assert.equal(json1.length, 2); // = 2 non-empty records
		assert.equal(json2.length, 4); // = 4 sheet rows
		assert.equal(json3.length, 3); // = 2 records + 1 blank row
		assert.equal(json4.length, 4); // = 4 sheet rows
		assert.equal(json5.length, 2); // = 2 records
		assert.equal(json6.length, 3); // = 4 sheet rows - 1 blank row
	});
});

describe('csv output', function() {
	var data, ws;
	var bef = (function() {
		data = [
			[1,2,3,null],
			[true, false, null, "sheetjs"],
			["foo", "bar", parseDate("2014-02-19T14:30:00.000Z"), "0.3"],
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
		delete _ws.C3.w;
		delete _ws.C3.z;
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
});

describe('js -> file -> js', function() {
	var data, ws, wb, BIN="binary";
	var bef = (function() {
		data = [
			[1,2,3],
			[true, false, null, "sheetjs"],
			["foo","bar", parseDate("2014-02-19T14:30:00.000Z"), "0.3"],
			["baz", 6.9, "qux"]
		];
		ws = X.utils.aoa_to_sheet(data);
		wb = { SheetNames: ['Sheet1'], Sheets: {Sheet1: ws} };
	});
	if(typeof before != 'undefined') before(bef);
	else it('before', bef);
	function eqcell(wb1, wb2, s, a) {
		assert.equal(wb1.Sheets[s][a].v, wb2.Sheets[s][a].v);
		assert.equal(wb1.Sheets[s][a].t, wb2.Sheets[s][a].t);
	}
	ofmt.forEach(function(f) {
		it(f, function() {
			var newwb = X.read(X.write(wb, {type:BIN, bookType: f}), {type:BIN});
			/* int */
			eqcell(wb, newwb, 'Sheet1', 'A1');
			eqcell(wb, newwb, 'Sheet1', 'B1');
			eqcell(wb, newwb, 'Sheet1', 'C1');
			/* double */
			eqcell(wb, newwb, 'Sheet1', 'B4');
			/* bool */
			eqcell(wb, newwb, 'Sheet1', 'A2');
			eqcell(wb, newwb, 'Sheet1', 'B2');
			/* string */
			eqcell(wb, newwb, 'Sheet1', 'D2');
			eqcell(wb, newwb, 'Sheet1', 'A3');
			eqcell(wb, newwb, 'Sheet1', 'B3');
			eqcell(wb, newwb, 'Sheet1', 'D3');
			eqcell(wb, newwb, 'Sheet1', 'A4');
			eqcell(wb, newwb, 'Sheet1', 'C4');
			/* date */
			eqcell(wb, newwb, 'Sheet1', 'C3');
		});
	});
});

describe('corner cases', function() {
	it('output functions', function() {
		var data = [
			[1,2,3],
			[true, false, null, "sheetjs"],
			["foo","bar", parseDate("2014-02-19T14:30:00.000Z"), "0.3"],
			["baz", null, "q\"ux"]
		];
		var ws = X.utils.aoa_to_sheet(data);
		ws.A1.f = ""; ws.A1.w = "";
		delete ws.C3.w; delete ws.C3.z; ws.C3.XF = {ifmt:14};
		ws.A4.t = "e";
		X.utils.get_formulae(ws);
		X.utils.make_csv(ws);
		X.utils.make_json(ws);
		ws['!cols'] = [ {wch:6}, {wch:7}, {wch:10}, {wch:20} ];

		var wb = {SheetNames:['sheetjs'], Sheets:{sheetjs:ws}};
		X.write(wb, {type: "binary", bookType: 'xlsx'});
		X.write(wb, {type: "base64", bookType: 'xlsm'});
		X.write(wb, {type: "base64", bookType: 'xlsb'});
		X.write(wb, {type: "binary", bookType: 'ods'});
		X.write(wb, {type: "binary", bookType: 'biff2'});
		ws.A2.t = "f";
		assert.throws(function() { X.utils.make_json(ws); });
	});
	it('SSF', function() {
		X.SSF.format("General", "dafuq");
		assert.throws(function(x) { return X.SSF.format("General", {sheet:"js"});});
		X.SSF.format("b e ddd hh AM/PM", 41722.4097222222);
		X.SSF.format("b ddd hh m", 41722.4097222222);
		["hhh","hhh A/P","hhmmm","sss","[hhh]","G eneral"].forEach(function(f) {
			assert.throws(function(x) { return X.SSF.format(f, 12345.6789);});
		});
		["[m]","[s]"].forEach(function(f) {
			assert.doesNotThrow(function(x) { return X.SSF.format(f, 12345.6789);});
		});
	});
	it('SSF oddities', function() {
		var ssfdata = require('./misc/ssf.json');
		ssfdata.forEach(function(d) {
			for(var j=1;j<d.length;++j) {
				if(d[j].length == 2) {
					var expected = d[j][1], actual = X.SSF.format(d[0], d[j][0], {});
					assert.equal(actual, expected);
				} else if(d[j][2] !== "#") assert.throws(function() { SSF.format(d[0], d[j][0]); });
			}
		});
	});
	it('codepage', function() {
		X.read(fs.readFileSync(dir + "biff5/number_format_greek.xls"), {type:"binary"});
	});
});

describe('encryption', function() {
	password_files.forEach(function(x) {
		describe(x, function() {
			it('should throw with no password', function() {assert.throws(function() { X.read(fs.readFileSync(dir + x), {type:"binary"}); }); });
			it('should throw with wrong password', function() {assert.throws(function() { X.read(fs.readFileSync(dir + x), {type:"binary",password:'passwor',WTF:opts.WTF}); }); });
			it('should recognize correct password', function() {
				try { X.read(fs.readFileSync(dir + x), {type:"binary",password:'password',WTF:opts.WTF}); }
				catch(e) { if(e.message == "Password is incorrect") throw e; }
			});
			it.skip('should decrypt file', function() {
				var wb = X.read(fs.readFileSync(dir + x), {type:"binary",password:'password',WTF:opts.WTF});
			});
		});
	});
});

describe.skip('multiformat tests', function() {
var mfopts = opts;
var mft = fs.readFileSync('multiformat.lst','utf-8').split("\n");
var csv = true, formulae = false;
mft.forEach(function(x) {
	if(x[0]!="#") describe('MFT ' + x, function() {
		var fil = {}, f = [], r = x.split(/\s+/);
		if(r.length < 3) return;
		it('should parse all', function() {
			for(var j = 1; j != r.length; ++j) f[j-1] = X.read(fs.readFileSync(dir + r[0] + r[j]), mfopts);
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
			var names = f[0].SheetNames;
			names.forEach(function(name) {
				cmparr(f.map(function(x) { return X.utils.sheet_to_csv(x.Sheets[name]); }));
			});
		} : null);
		it('should have the same formulae', formulae ? function() {
			cmparr(f.map(function(x) { return x.SheetNames; }));
			var names = f[0].SheetNames;
			names.forEach(function(name) {
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
});
});

