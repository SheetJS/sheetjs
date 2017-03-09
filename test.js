/* vim: set ts=2: */
/*jshint loopfunc:true, eqnull:true */
var X;
var modp = './';
//var modp = 'xlsx';
var fs = require('fs'), assert = require('assert');
describe('source',function(){it('should load',function(){X=require(modp);});});

var opts = {cellNF: true};
if(process.env.WTF) {
	opts.WTF = true;
	opts.cellStyles = true;
}
var fullex = [".xlsb", ".xlsm", ".xlsx"];
var ofmt = ["xlsb", "xlsm", "xlsx", "ods", "biff2"];
var ex = fullex.slice(); ex.push(".ods"); ex.push(".xls"); ex.push("xml");
if(process.env.FMTS === "full") process.env.FMTS = ex.join(":");
if(process.env.FMTS) ex=process.env.FMTS.split(":").map(function(x){return x[0]==="."?x:"."+x;});
var exp = ex.map(function(x){ return x + ".pending"; });
function test_file(x){ return ex.indexOf(x.substr(-5))>=0||exp.indexOf(x.substr(-13))>=0 || ex.indexOf(x.substr(-4))>=0||exp.indexOf(x.substr(-12))>=0; }

var files = (fs.existsSync('tests.lst') ? fs.readFileSync('tests.lst', 'utf-8').split("\n") : fs.readdirSync('test_files')).filter(test_file);
var fileA = (fs.existsSync('testA.lst') ? fs.readFileSync('testA.lst', 'utf-8').split("\n") : []).filter(test_file);

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
	dtxlsx:  dir + 'xlsx-stream-d-date-cell.xlsx',
	dtxlsb:  dir + 'xlsx-stream-d-date-cell.xlsb',
	swcxls: dir + 'apachepoi_SimpleWithComments.xls',
	swcxml: dir + '2011/apachepoi_SimpleWithComments.xls.xml',
	swcxlsx: dir + 'apachepoi_SimpleWithComments.xlsx',
	swcxlsb: dir + '2013/apachepoi_SimpleWithComments.xlsx.xlsb'
};

var N1 = 'XLSX';
var N2 = 'XLSB';
var N3 = 'XLS';
var N4 = 'XML';

function parsetest(x, wb, full, ext) {
	ext = (ext ? " [" + ext + "]": "");
	if(!full && ext) return;
	describe(x + ext + ' should have all bits', function() {
		var sname = dir + '2011/' + x.substr(x.lastIndexOf('/')+1) + '.sheetnames';
		it('should have all sheets', function() {
			wb.SheetNames.forEach(function(y) { assert(wb.Sheets[y], 'bad sheet ' + y); });
		});
		it('should have the right sheet names', fs.existsSync(sname) ? function() {
			var file = fs.readFileSync(sname, 'utf-8').replace(/\r/g,"");
			var names = wb.SheetNames.map(fixsheetname).join("\n") + "\n";
			assert.equal(names, file);
		} : null);
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
				X.utils.sheet_to_row_object_array(wb.Sheets[ws]);
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
			it('#' + i + ' (' + ws + ')', fs.existsSync(name) ? function() {
				var file = fs.readFileSync(name, 'utf-8');
				var csv = X.utils.make_csv(wb.Sheets[ws]);
				assert.equal(fixcsv(csv), fixcsv(file), "CSV badness");
			} : null);
		});
	});
	describe(x + ext + ' should generate correct JSON output', function() {
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
			var wb = X.readFile(dir + '2011/' + x + '.xml', opts);
		});
	});
	if(fs.existsSync(dir + '2013/' + x + '.xlsb'))
	describe(x + ext + '.xlsb from 2013', function() {
		it('should parse', function() {
			var wb = X.readFile(dir + '2013/' + x + '.xlsb', opts);
		});
	});
	if(fs.existsSync(dir + x + '.xml' + ext))
	describe(x + '.xml', function() {
		it('should parse', function() {
			var wb = X.readFile(dir + x + '.xml', opts);
		});
	});
}

var wbtable = {};

describe('should parse test files', function() {
	files.forEach(function(x) {
		if(!fs.existsSync(dir + x)) return;
		it(x, x.substr(-8) == ".pending" ? null : function() {
			var wb = X.readFile(dir + x, opts);
			wbtable[dir + x] = wb;
			parsetest(x, wb, true);
		});
		fullex.forEach(function(ext, idx) {
			it(x + ' [' + ext + ']', function(){
				var wb = wbtable[dir + x];
				if(!wb) wb = X.readFile(dir + x, opts);
				parsetest(x, X.read(X.write(wb, {type:"buffer", bookType:ext.replace(/\./,"")}), {WTF:opts.WTF}), ext.replace(/\./,"") !== "xlsb", ext);
			});
		});
	});
	fileA.forEach(function(x) {
		if(!fs.existsSync(dir + x)) return;
		it(x, x.substr(-8) == ".pending" ? null : function() {
			var wb = X.readFile(dir + x, {WTF:opts.wtf, sheetRows:10});
			parsetest(x, wb, false);
		});
	});
});

describe('parse options', function() {
	var html_cell_types = ['s'];
	before(function() {
		X = require(modp);
	});
	describe('cell', function() {
		it('XLSX should generate HTML by default', function() {
			var wb = X.readFile(paths.cstxlsx);
			var ws = wb.Sheets.Sheet1;
			Object.keys(ws).forEach(function(addr) {
				if(addr[0] === "!" || !ws.hasOwnProperty(addr)) return;
				assert(html_cell_types.indexOf(ws[addr].t) === -1 || ws[addr].h);
			});
		});
		it('XLSX should not generate HTML when requested', function() {
			var wb = X.readFile(paths.cstxlsx, {cellHTML:false});
			var ws = wb.Sheets.Sheet1;
			Object.keys(ws).forEach(function(addr) {
				if(addr[0] === "!" || !ws.hasOwnProperty(addr)) return;
				assert(typeof ws[addr].h === 'undefined');
			});
		});
		it('should generate formulae by default', function() {
			[paths.fstxls, paths.fstxlsb].forEach(function(p) {
				var wb = X.readFile(p);
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
				var wb =X.readFile(p,{cellFormula:false});
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
				var wb = X.readFile(p);
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
				var wb = X.readFile(p, {cellNF: true});
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
			var wb = X.readFile(p);
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
			[paths.cssxlsx, /*paths.cssxls, paths.cssxml*/].forEach(function(p) {
			var wb = X.readFile(p, {cellStyles:true});
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
			var wb = X.readFile(paths.dtxlsx);
			wb.SheetNames.forEach(function(s) {
				var ws = wb.Sheets[s];
				Object.keys(ws).forEach(function(addr) {
					if(addr[0] === "!" || !ws.hasOwnProperty(addr)) return;
					assert(ws[addr].t !== 'd');
				});
			});
		});
		it('XLSB should not generate cell dates', function() {
			var wb = X.readFile(paths.dtxlsb, {cellDates: true});
			wb.SheetNames.forEach(function(s) {
				var ws = wb.Sheets[s];
				Object.keys(ws).forEach(function(addr) {
					if(addr[0] === "!" || !ws.hasOwnProperty(addr)) return;
					assert(ws[addr].t !== 'd');
				});
			});
		});
		it('XLSX should generate cell dates when requested', function() {
			var wb = X.readFile(paths.dtxlsx, {cellDates: true});
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
				var wb = X.readFile(p);
				assert.throws(function() { return wb.Sheets.Merge.A2.v; });
			});
		});
		it('should generate sheet stubs when requested', function() {
			/* TODO: ODS/XLS/XML */
			[paths.mcxlsx, paths.mcxlsb /*, paths.mcods, paths.mcxls, paths.mcxml*/].forEach(function(p) {
				var wb = X.readFile(p, {sheetStubs:true});
				assert(typeof wb.Sheets.Merge.A2.t !== 'undefined');
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
				checkcells(X.readFile(p), true, true, true, true);
			});
		});
		it('sheetRows n=20', function() {
			[paths.fstxlsx, paths.fstxlsb, paths.fstods, paths.fstxls, paths.fstxml].forEach(function(p) {
				checkcells(X.readFile(p, {sheetRows:20}), false, false, true, true);
			});
		});
		it('sheetRows n=10', function() {
			[paths.fstxlsx, paths.fstxlsb, paths.fstods, paths.fstxls, paths.fstxml].forEach(function(p) {
				checkcells(X.readFile(p, {sheetRows:10}), false, false, false, true);
			});
		});
	});
	describe('book', function() {
		it('bookSheets should not generate sheets', function() {
			[paths.mcxlsx, paths.mcxlsb, paths.mcxls, paths.mcxml].forEach(function(p) {
				var wb = X.readFile(p, {bookSheets:true});
				assert(typeof wb.Sheets === 'undefined');
			});
		});
		it('bookProps should not generate sheets', function() {
			[paths.nfxlsx, paths.nfxlsb, paths.nfxls, paths.nfxml].forEach(function(p) {
				var wb = X.readFile(p, {bookProps:true});
				assert(typeof wb.Sheets === 'undefined');
			});
		});
		it('bookProps && bookSheets should not generate sheets', function() {
			[paths.lonxlsx, paths.lonxls].forEach(function(p) {
				var wb = X.readFile(p, {bookProps:true, bookSheets:true});
				assert(typeof wb.Sheets === 'undefined');
			});
		});
		it('should not generate deps by default', function() {
			[paths.fstxlsx, paths.fstxlsb, paths.fstxls, paths.fstxml].forEach(function(p) {
				var wb = X.readFile(p);
				assert(typeof wb.Deps === 'undefined' || !(wb.Deps && wb.Deps.length>0));
			});
		});
		it('bookDeps should generate deps (XLSX/XLSB)', function() {
			[paths.fstxlsx, paths.fstxlsb].forEach(function(p) {
				var wb = X.readFile(p, {bookDeps:true});
				assert(typeof wb.Deps !== 'undefined' && wb.Deps.length > 0);
			});
		});
		var ckf = function(wb, fields, exists) { fields.forEach(function(f) {
			assert((typeof wb[f] !== 'undefined') == exists);
		}); };
		it('should not generate book files by default', function() {
			var wb;
			wb = X.readFile(paths.fstxlsx);
			ckf(wb, ['files', 'keys'], false);
			wb = X.readFile(paths.fstxlsb);
			ckf(wb, ['files', 'keys'], false);
			wb = X.readFile(paths.fstxls);
			ckf(wb, ['cfb'], false);
		});
		it('bookFiles should generate book files', function() {
			var wb;
			wb = X.readFile(paths.fstxlsx, {bookFiles:true});
			ckf(wb, ['files', 'keys'], true);
			wb = X.readFile(paths.fstxlsb, {bookFiles:true});
			ckf(wb, ['files', 'keys'], true);
			wb = X.readFile(paths.fstxls, {bookFiles:true});
			ckf(wb, ['cfb'], true);
		});
		it('should not generate VBA by default', function() {
			var wb = X.readFile(paths.nfxlsx);
			assert(typeof wb.vbaraw === 'undefined');
			wb = X.readFile(paths.nfxlsb);
			assert(typeof wb.vbaraw === 'undefined');
		});
		it('bookVBA should generate vbaraw (XLSX/XLSB)', function() {
			var wb = X.readFile(paths.nfxlsx,{bookVBA:true});
			assert(typeof wb.vbaraw !== 'undefined');
			wb = X.readFile(paths.nfxlsb,{bookVBA:true});
			assert(typeof wb.vbaraw !== 'undefined');
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
	it('should read buffers', function() {
		X.read(fs.readFileSync(paths.cstxls), {type: 'buffer'});
		X.read(fs.readFileSync(paths.cstxml), {type: 'buffer'});
		X.read(fs.readFileSync(paths.cstxlsx), {type: 'buffer'});
		X.read(fs.readFileSync(paths.cstxlsb), {type: 'buffer'});
	});
	it('should read array', function() {
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
	it('should infer buffer type', function() {
		X.read(fs.readFileSync(paths.cstxls));
		X.read(fs.readFileSync(paths.cstxml));
		X.read(fs.readFileSync(paths.cstxlsx));
		X.read(fs.readFileSync(paths.cstxlsb));
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
	before(function() {
		X = require(modp);
		wb1 = X.readFile(paths.cpxlsx);
		wb2 = X.readFile(paths.cpxlsb);
		wb3 = X.readFile(paths.cpxls);
		wb4 = X.readFile(paths.cpxml);
	});
	it('should write binary strings', function() {
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
	it('should write buffers', function() {
		X.write(wb1, {type: 'buffer'});
		X.write(wb2, {type: 'buffer'});
		X.write(wb3, {type: 'buffer'});
		X.write(wb4, {type: 'buffer'});
		X.read(X.write(wb1, {type: 'buffer'}), {type: 'buffer'});
		X.read(X.write(wb2, {type: 'buffer'}), {type: 'buffer'});
		X.read(X.write(wb3, {type: 'buffer'}), {type: 'buffer'});
		X.read(X.write(wb4, {type: 'buffer'}), {type: 'buffer'});
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
	it('should have comment as part of cell properties', function(){
		var X = require(modp);
		var sheet = 'Sheet1';
		var wb1=X.readFile(paths.swcxlsx);
		var wb2=X.readFile(paths.swcxlsb);
		var wb3=X.readFile(paths.swcxls);
		var wb4=X.readFile(paths.swcxml);

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
		var wb1, wb2;
		before(function() {
			wb1 = X.readFile(paths.cpxlsx);
			wb2 = X.readFile(paths.cpxlsb);
			wb3 = X.readFile(paths.cpxls);
			wb4 = X.readFile(paths.cpxml);
		});

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
			var opts = {};
			var wb1 = X.readFile(paths.fstxlsx, opts);
			var wb2 = X.readFile(paths.fstxlsb, opts);
			var wb3 = X.readFile(paths.fstxls, opts);
			var wb4 = X.readFile(paths.fstxml, opts);
			[wb1, wb2, wb3, wb4].forEach(function(wb) {
				assert.equal(wb.Sheets.Text["!ref"],"A1:F49");
			});
		});
		it('should adjust range if set', function() {
			var opts = {sheetRows:10};
			var wb1 = X.readFile(paths.fstxlsx, opts);
			var wb2 = X.readFile(paths.fstxlsb, opts);
			var wb3 = X.readFile(paths.fstxls, opts);
			var wb4 = X.readFile(paths.fstxml, opts);
			/* TODO */
			[wb1, wb2 /*, wb3, wb4 */].forEach(function(wb) {
				assert.equal(wb.Sheets.Text["!fullref"],"A1:F49");
				assert.equal(wb.Sheets.Text["!ref"],"A1:F10");
			});
		});
		it('should not generate comment cells', function() {
			var opts = {sheetRows:10};
			var wb1 = X.readFile(paths.cstxlsx, opts);
			var wb2 = X.readFile(paths.cstxlsb, opts);
			var wb3 = X.readFile(paths.cstxls, opts);
			var wb4 = X.readFile(paths.cstxml, opts);
			/* TODO */
			[wb1, wb2 /*, wb3, wb4 */].forEach(function(wb) {
				assert.equal(wb.Sheets.Sheet7["!fullref"],"A1:N34");
				assert.equal(wb.Sheets.Sheet7["!ref"],"A1");
			});
		});
	});

	describe('merge cells',function() {
		var wb1, wb2;
		before(function() {
			X = require(modp);
			wb1 = X.readFile(paths.mcxlsx);
			wb2 = X.readFile(paths.mcxlsb);
			wb3 = X.readFile(paths.mcods);
			wb4 = X.readFile(paths.mcxls);
			wb5 = X.readFile(paths.mcxml);
		});
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
		var wb1, wb2;
		before(function() {
			X = require(modp);
			wb1 = X.readFile(paths.hlxlsx);
			wb2 = X.readFile(paths.hlxlsb);
			wb3 = X.readFile(paths.hlxls);
			wb4 = X.readFile(paths.hlxml);
		});

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
			wb = X.readFile(paths.dtxlsx);
			ws = wb.Sheets[sheetName];
			var sheet = X.utils.sheet_to_row_object_array(ws);
			assert.equal(sheet[3]['てすと'], '2/14/14');
		});
		it('cellDates should not affect formatted text', function() {
			var wb1, ws1, wb2, ws2;
			var sheetName = 'Sheet1';
			wb1 = X.readFile(paths.dtxlsx);
			ws1 = wb1.Sheets[sheetName];
			wb2 = X.readFile(paths.dtxlsb);
			ws2 = wb2.Sheets[sheetName];
			assert.equal(X.utils.sheet_to_csv(ws1),X.utils.sheet_to_csv(ws2));
		});
	});

	describe('should correctly handle styles', function() {
		var wsxls, wsxlsx, rn, rn2;
		before(function() {
			wsxls=X.readFile(paths.cssxls, {cellStyles:true,WTF:1}).Sheets.Sheet1;
			wsxlsx=X.readFile(paths.cssxlsx, {cellStyles:true,WTF:1}).Sheets.Sheet1;
			rn = function(range) {
				var r = X.utils.decode_range(range);
				var out = [];
				for(var R = r.s.r; R <= r.e.r; ++R) for(var C = r.s.c; C <= r.e.c; ++C)
					out.push(X.utils.encode_cell({c:C,r:R}));
				return out;
			};
			rn2 = function(r) { return [].concat.apply([], r.split(",").map(rn)); };
		});
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
	before(function() {
		X = require(modp);
	});
	describe('should parse core properties and custom properties', function() {
		var wb1, wb2, base = './tmp/cp';
		before(function() {
			wb1 = X.readFile(paths.cpxlsx);
			wb2 = X.readFile(paths.cpxlsb);
			fullex.forEach(function(p) {
				X.writeFile(wb1, base + '.xlsm' + p);
				X.writeFile(wb2, base + '.xlsb' + p);
			});
		});
		fullex.forEach(function(p) { ['.xlsm','.xlsb'].forEach(function(q) {
			it(q + p + ' should roundtrip core and custom properties', function() {
				var wb = X.readFile(base + q + p);
				coreprop(wb);
				custprop(wb);
			}); });
		});
	});

	describe('should preserve features', function() {
		it('merge cells', function() {
			var wb1 = X.readFile(paths.mcxlsx);
			var wb2 = X.read(X.write(wb1, {type:'binary'}), {type:'binary'});
			var m1 = wb1.Sheets.Merge['!merges'].map(X.utils.encode_range);
			var m2 = wb2.Sheets.Merge['!merges'].map(X.utils.encode_range);
			assert.equal(m1.length, m2.length);
			for(var i = 0; i < m1.length; ++i) assert.equal(m1[i], m2[i]);
		});
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
				var wb1 = X.readFile(f, {cellNF: true, cellDates: di, WTF: opts.WTF});
				var wb2 = X.read(X.write(wb1, {type:'binary', cellDates:dj, WTF:opts.WTF}), {type:'binary', cellDates: dk, WTF: opts.WTF});
				var m = [wb1,wb2].map(function(x) { return x.Sheets[sheet][addr]; });
				assert.equal(m[0].w, m[1].w);

				/* wb1 cellDates */
				if(dh && di) assert.equal(m[0].t, 'd');
				//else if(a !== 'd' && di) assert.equal(m[0].t, 'd'); /* TODO */
				else assert.equal(m[0].t, 'n');
				/* wb2 cellDates */
				if(dh && di && dj && dk) assert.equal(m[1].t, 'd');
				else if(dj && dk && !di); /* TODO: convert to date */
				else assert.equal(m[1].t, 'n');

				if(m[0].t === 'n' && m[1].t === 'n') assert.equal(m[0].v, m[1].v);
				else if(m[0].t === 'd' && m[1].t === 'd') assert.equal(m[0].v.toString(), m[1].v.toString());
				else if(m[1].t === 'n') assert(Math.abs(datenum(new Date(m[0].v)) - m[1].v) < 0.01); /* TODO: 1sec adjustment */
			});
		});
	});

	describe('xls to xlsx conversions', function() { [
			['XLS', 'formula_stress_test.xls'],
			['XML', 'formula_stress_test.xls.xml']
		].forEach(function(w) {
			it('should be able to write ' + w[0] + ' files', function() {
				var xls = X.readFile('./test_files/' + w[1], {cellNF:true});
				X.writeFile(xls, './tmp/' + w[1] + '.xlsx', {bookSST:true});
				X.writeFile(xls, './tmp/' + w[1] + '.xlsb', {bookSST:true});
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
			assert.throws(function() { X.readFile(dir + w[1]); });
			assert.throws(function() { X.read(fs.readFileSync(dir+w[1], 'base64'), {type:'base64'}); });
		}); });
	});
	describe('write', function() {
		it('should pass -> XLSX', function() {
			X.write(X.readFile(paths.fstxlsb), {type:'binary'});
			X.write(X.readFile(paths.fstxlsx), {type:'binary'});
			X.write(X.readFile(paths.fstxls), {type:'binary'});
			X.write(X.readFile(paths.fstxml), {type:'binary'});
		});
		it('should pass if a sheet is missing', function() {
			var wb = X.readFile(paths.fstxlsx); delete wb.Sheets[wb.SheetNames[0]];
			X.read(X.write(wb, {type:'binary'}), {type:'binary'});
		});
		['Props', 'Custprops', 'SSF'].forEach(function(t) {
			it('should pass if ' + t + ' is missing', function() {
				var wb = X.readFile(paths.fstxlsx);
				assert.doesNotThrow(function() {
					delete wb[t];
					X.write(wb, {type:'binary'});
				});
			});
		});
		['SheetNames', 'Sheets'].forEach(function(t) {
			it('should fail if ' + t + ' is missing', function() {
				var wb = X.readFile(paths.fstxlsx);
				assert.throws(function() {
					delete wb[t];
					X.write(wb, {type:'binary'});
				});
			});
		});
	});
});

function datenum(v/*:Date*/, date1904/*:?boolean*/)/*:number*/ {
	var epoch = v.getTime();
	if(date1904) epoch += 1462*24*60*60*1000;
	return (epoch + 2209161600000) / (24 * 60 * 60 * 1000);
}
function sheet_from_array_of_arrays(data, opts) {
	var ws = {};
	var range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
	for(var R = 0; R != data.length; ++R) {
		for(var C = 0; C != data[R].length; ++C) {
			if(range.s.r > R) range.s.r = R;
			if(range.s.c > C) range.s.c = C;
			if(range.e.r < R) range.e.r = R;
			if(range.e.c < C) range.e.c = C;
			var cell = {v: data[R][C] };
			if(cell.v == null) continue;
			var cell_ref = X.utils.encode_cell({c:C,r:R});
			if(typeof cell.v === 'number') cell.t = 'n';
			else if(typeof cell.v === 'boolean') cell.t = 'b';
			else if(cell.v instanceof Date) {
				cell.t = 'n'; cell.z = X.SSF._table[14];
				cell.v = datenum(cell.v);
			}
			else cell.t = 's';
			ws[cell_ref] = cell;
		}
	}
	if(range.s.c < 10000000) ws['!ref'] = X.utils.encode_range(range);
	return ws;
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
	before(function() {
		data = [
			[1,2,3],
			[true, false, null, "sheetjs"],
			["foo","bar",new Date("2014-02-19T14:30Z"), "0.3"],
			["baz", null, "qux"]
		];
		ws = sheet_from_array_of_arrays(data);
	});
	it('should use first-row headers and full sheet by default', function() {
		var json = X.utils.sheet_to_json(ws);
		assert.equal(json.length, data.length - 1);
		assert.equal(json[0][1], true);
		assert.equal(json[1][2], "bar");
		assert.equal(json[2][3], "qux");
		assert.doesNotThrow(function() { seeker(json, [1,2,3], "sheetjs"); });
		assert.throws(function() { seeker(json, [1,2,3], "baz"); });
	});
	it('should create array of arrays if header == 1', function() {
		var json = X.utils.sheet_to_json(ws, {header:1});
		assert.equal(json.length, data.length);
		assert.equal(json[1][0], true);
		assert.equal(json[2][1], "bar");
		assert.equal(json[3][2], "qux");
		assert.doesNotThrow(function() { seeker(json, [0,1,2], "sheetjs"); });
		assert.throws(function() { seeker(json, [0,1,2,3], "sheetjs"); });
		assert.throws(function() { seeker(json, [0,1,2], "baz"); });
	});
	it('should use column names if header == "A"', function() {
		var json = X.utils.sheet_to_json(ws, {header:'A'});
		assert.equal(json.length, data.length);
		assert.equal(json[1].A, true);
		assert.equal(json[2].B, "bar");
		assert.equal(json[3].C, "qux");
		assert.doesNotThrow(function() { seeker(json, "ABC", "sheetjs"); });
		assert.throws(function() { seeker(json, "ABCD", "sheetjs"); });
		assert.throws(function() { seeker(json, "ABC", "baz"); });
	});
	it('should use column labels if specified', function() {
		var json = X.utils.sheet_to_json(ws, {header:["O","D","I","N"]});
		assert.equal(json.length, data.length);
		assert.equal(json[1].O, true);
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
			assert.equal(json[0][0], true);
			assert.equal(json[1][1], "bar");
			assert.equal(json[2][2], "qux");
			assert.doesNotThrow(function() { seeker(json, [0,1,2], "sheetjs"); });
			assert.throws(function() { seeker(json, [0,1,2,3], "sheetjs"); });
			assert.throws(function() { seeker(json, [0,1,2], "baz"); });
		});
	});
});

describe('js -> file -> js', function() {
	var data, ws, wb, BIN="binary";
	before(function() {
		data = [
			[1,2,3],
			[true, false, null, "sheetjs"],
			["foo","bar",new Date("2014-02-19T14:30Z"), "0.3"],
			["baz", 6.9, "qux"]
		];
		ws = sheet_from_array_of_arrays(data);
		wb = { SheetNames: ['Sheet1'], Sheets: {Sheet1: ws} };
	});
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
			["foo","bar",new Date("2014-02-19T14:30Z"), "0.3"],
			["baz", null, "q\"ux"]
		];
		ws = sheet_from_array_of_arrays(data);
		ws.A1.f = ""; ws.A1.w = "";
		delete ws.C3.w; delete ws.C3.z; ws.C3.XF = {ifmt:14};
		ws.A4.t = "e";
		X.utils.get_formulae(ws);
		X.utils.make_csv(ws);
		X.utils.make_json(ws);
		ws['!cols'] = [ {wch:6}, {wch:7}, {wch:10}, {wch:20} ];

		var wb = {SheetNames:['sheetjs'], Sheets:{sheetjs:ws}};
		X.write(wb, {type: "binary", bookType: 'xlsx'});
		X.write(wb, {type: "buffer", bookType: 'xlsm'});
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
	it('CFB', function() {
		var cfb = X.CFB.read(paths.swcxls, {type:"file"});
		var xls = X.parse_xlscfb(cfb);
	});
	it('codepage', function() {
		X.readFile(dir + "biff5/number_format_greek.xls");
	});
});

describe('encryption', function() {
	password_files.forEach(function(x) {
		describe(x, function() {
			it('should throw with no password', function() {assert.throws(function() { X.readFile(dir + x); }); });
			it('should throw with wrong password', function() {assert.throws(function() { X.readFile(dir + x, {password:'passwor',WTF:opts.WTF}); }); });
			it('should recognize correct password', function() {
				try { X.readFile(dir + x, {password:'password',WTF:opts.WTF}); }
				catch(e) { if(e.message == "Password is incorrect") throw e; }
			});
			it.skip('should decrypt file', function() {
				var wb = X.readFile(dir + x, {password:'password',WTF:opts.WTF});
			});
		});
	});
});
