/* vim: set ts=2: */
var XLSX;
var fs = require('fs'), assert = require('assert');
describe('source',function(){it('should load',function(){XLSX=require('./');});});

var opts = {};
if(process.env.WTF) opts.WTF = true;

var ex = [".xlsb", ".xlsm", ".xlsx"];
if(process.env.FMTS) ex=process.env.FMTS.split(":").map(function(x){return x[0]==="."?x:"."+x;});
var exp = ex.map(function(x){ return x + ".pending"; });
function test_file(x){return ex.indexOf(x.substr(-5))>=0||exp.indexOf(x.substr(-13))>=0;}

var files = (fs.existsSync('tests.lst') ? fs.readFileSync('tests.lst', 'utf-8').split("\n") : fs.readdirSync('test_files')).filter(test_file);

/* Excel enforces 31 character sheet limit, although technical file limit is 255 */
function fixsheetname(x) { return x.substr(0,31); }

function normalizecsv(x) { return x.replace(/\t/g,",").replace(/#{255}/g,"").replace(/"/g,"").replace(/[\n\r]+/g,"\n").replace(/\n*$/,""); }

var dir = "./test_files/";

function parsetest(x, wb) {
	describe(x + ' should have all bits', function() {
		var sname = dir + '2011/' + x + '.sheetnames';
		it('should have all sheets', function() {
			wb.SheetNames.forEach(function(y) { assert(wb.Sheets[y], 'bad sheet ' + y); });
		});
		it('should have the right sheet names', fs.existsSync(sname) ? function() {
			var file = fs.readFileSync(sname, 'utf-8');
			var names = wb.SheetNames.map(fixsheetname).join("\n") + "\n";
			assert.equal(names, file);
		} : null);
	});
	describe(x + ' should generate CSV', function() {
		wb.SheetNames.forEach(function(ws, i) {
			it('#' + i + ' (' + ws + ')', function() {
				var csv = XLSX.utils.make_csv(wb.Sheets[ws]);
			});
		});
	});
	describe(x + ' should generate JSON', function() {
		wb.SheetNames.forEach(function(ws, i) {
			it('#' + i + ' (' + ws + ')', function() {
				var json = XLSX.utils.sheet_to_row_object_array(wb.Sheets[ws]);
			});
		});
	});
	describe(x + ' should generate formulae', function() {
		wb.SheetNames.forEach(function(ws, i) {
			it('#' + i + ' (' + ws + ')', function() {
				var json = XLSX.utils.get_formulae(wb.Sheets[ws]);
			});
		});
	});
	describe(x + ' should generate correct output', function() {
		wb.SheetNames.forEach(function(ws, i) {
			var name = (dir + x + '.' + i + '.csv');
			it('#' + i + ' (' + ws + ')', fs.existsSync(name) ? function() {
				var file = fs.readFileSync(name, 'utf-8');
				var csv = XLSX.utils.make_csv(wb.Sheets[ws]);
				assert.equal(normalizecsv(csv), normalizecsv(file), "CSV badness");
			} : null);
		});
	});
	if(!fs.existsSync(dir + '2013/' + x + '.xlsb')) return;
	describe(x + '.xlsb from 2013', function() {
		it('should parse', function() {
			var xlsb = XLSX.readFile(dir + '2013/' + x + '.xlsb', opts);
		});
	});
}

describe('should parse test files', function() {
	files.forEach(function(x) {
		it(x, x.substr(-8) == ".pending" ? null : function() {
			var wb = XLSX.readFile(dir + x, opts);
			parsetest(x, wb);
		});
	});
});

describe('options', function() {
	var html_cell_types = ['s'];
	before(function() {
		XLSX = require('./');
	});
	describe('cell', function() {
		it('should generate HTML by default', function() {
			var wb = XLSX.readFile(dir + 'comments_stress_test.xlsx');
			var ws = wb.Sheets.Sheet1;
			Object.keys(ws).forEach(function(addr) {
				if(addr[0] === "!" || !ws.hasOwnProperty(addr)) return;
				assert(html_cell_types.indexOf(ws[addr].t) === -1 || ws[addr].h);
			});
		});
		it('should not generate HTML when requested', function() {
			var wb = XLSX.readFile(dir+'comments_stress_test.xlsx', {cellHTML:false});
			var ws = wb.Sheets.Sheet1;
			Object.keys(ws).forEach(function(addr) {
				if(addr[0] === "!" || !ws.hasOwnProperty(addr)) return;
				assert(typeof ws[addr].h === 'undefined');
			});
		});
		it('should generate formulae by default', function() {
			var wb = XLSX.readFile(dir + 'formula_stress_test.xlsb');
			var found = false;
			wb.SheetNames.forEach(function(s) {
				var ws = wb.Sheets[s];
				Object.keys(ws).forEach(function(addr) {
					if(addr[0] === "!" || !ws.hasOwnProperty(addr)) return;
					if(typeof ws[addr].f !== 'undefined') return found = true;
				});
			});
			assert(found);
		});
		it('should not generate formulae when requested', function() {
			var wb =XLSX.readFile(dir+'formula_stress_test.xlsb',{cellFormula:false});
			wb.SheetNames.forEach(function(s) {
				var ws = wb.Sheets[s];
				Object.keys(ws).forEach(function(addr) {
					if(addr[0] === "!" || !ws.hasOwnProperty(addr)) return;
					assert(typeof ws[addr].f === 'undefined');
				});
			});
		});
		it('should not generate number formats by default', function() {
			var wb = XLSX.readFile(dir+'number_format.xlsm');
			wb.SheetNames.forEach(function(s) {
				var ws = wb.Sheets[s];
				Object.keys(ws).forEach(function(addr) {
					if(addr[0] === "!" || !ws.hasOwnProperty(addr)) return;
					assert(typeof ws[addr].z === 'undefined');
				});
			});
		});
		it('should generate number formats when requested', function() {
			var wb = XLSX.readFile(dir+'number_format.xlsm', {cellNF: true});
			wb.SheetNames.forEach(function(s) {
				var ws = wb.Sheets[s];
				Object.keys(ws).forEach(function(addr) {
					if(addr[0] === "!" || !ws.hasOwnProperty(addr)) return;
					assert(typeof ws[addr].t!== 'n' || typeof ws[addr].z !== 'undefined');
				});
			});
		});
	});
	describe('sheet', function() {
		it('should not generate sheet stubs by default', function() {
			var wb = XLSX.readFile(dir+'merge_cells.xlsx');
			assert.throws(function() { wb.Sheets.Merge.A2.v; });
		});
		it('should generate sheet stubs when requested', function() {
			var wb = XLSX.readFile(dir+'merge_cells.xlsx', {sheetStubs:true});
			assert(typeof wb.Sheets.Merge.A2.t !== 'undefined');
		});
	});
	describe('book', function() {
		it('bookSheets should not generate sheets', function() {
			var wb = XLSX.readFile(dir+'merge_cells.xlsx', {bookSheets:true});
			assert(typeof wb.Sheets === 'undefined');
		});
		it('bookProps should not generate sheets', function() {
			var wb = XLSX.readFile(dir+'number_format.xlsb', {bookProps:true});
			assert(typeof wb.Sheets === 'undefined');
		});
		it('bookProps && bookSheets should not generate sheets', function() {
			var wb = XLSX.readFile(dir+'LONumbers.xlsx', {bookProps:true, bookSheets:true});
			assert(typeof wb.Sheets === 'undefined');
		});
		it('should not generate deps by default', function() {
			var wb = XLSX.readFile(dir+'formula_stress_test.xlsx');
			assert(typeof wb.Deps === 'undefined' || !(wb.Deps.length>0));
		});
		it('bookDeps should generate deps', function() {
			var wb = XLSX.readFile(dir+'formula_stress_test.xlsx', {bookDeps:true});
			assert(typeof wb.Deps !== 'undefined' && wb.Deps.length > 0);
		});
		it('should not generate files or keys by default', function() {
			var wb = XLSX.readFile(dir+'formula_stress_test.xlsx');
			assert(typeof wb.files === 'undefined');
			assert(typeof wb.keys === 'undefined');
			wb = XLSX.readFile(dir+'formula_stress_test.xlsb');
			assert(typeof wb.files === 'undefined');
			assert(typeof wb.keys === 'undefined');
		});
		it('bookFiles should generate files and keys', function() {
			var wb = XLSX.readFile(dir+'formula_stress_test.xlsx', {bookFiles:true});
			assert(typeof wb.files !== 'undefined');
			assert(typeof wb.keys !== 'undefined');
			wb = XLSX.readFile(dir+'formula_stress_test.xlsb', {bookFiles:true});
			assert(typeof wb.files !== 'undefined');
			assert(typeof wb.keys !== 'undefined');
		});
	});
});

describe('input formats', function() {
	it('should read binary strings', function() {
		XLSX.read(fs.readFileSync(dir+'comments_stress_test.xlsb', 'binary'), {type: 'binary'});
		XLSX.read(fs.readFileSync(dir+'comments_stress_test.xlsx', 'binary'), {type: 'binary'});
	});
	it('should read base64 strings', function() {
		XLSX.read(fs.readFileSync(dir+'comments_stress_test.xlsb', 'base64'), {type: 'base64'});
		XLSX.read(fs.readFileSync(dir+'comments_stress_test.xlsx', 'base64'), {type: 'base64'});
	});
});

describe('features', function() {
	describe('should have comment as part of cell properties', function(){
		var ws;
		before(function() {
			XLSX = require('./');
			var wb = XLSX.readFile(dir+'apachepoi_SimpleWithComments.xlsx');
			var sheetName = 'Sheet1';
			ws = wb.Sheets[sheetName];
		});
		it('Parse comments.xml and insert into cell',function(){
			assert.equal(ws.B1.c.length, 1,"must have 1 comment");
			assert.equal(ws.B1.c[0].t, "Yegor Kozlov:\r\nfirst cell", "must have the concatenated texts");
			assert.equal(ws.B1.c[0].h, '<span style="font-weight: bold;">Yegor Kozlov:</span><span style=""><br/>first cell</span>', "must have the html representation");
			assert.equal(ws.B1.c[0].r, '<r><rPr><b/><sz val="8"/><color indexed="81"/><rFont val="Tahoma"/></rPr><t>Yegor Kozlov:</t></r><r><rPr><sz val="8"/><color indexed="81"/><rFont val="Tahoma"/></rPr><t xml:space="preserve">\r\nfirst cell</t></r>', "must have the rich text representation");
			assert.equal(ws.B1.c[0].a, "Yegor Kozlov","must have the same author");
		});
	});

	describe('should have core properties and custom properties parsed', function() {
		var wb;
		before(function() {
			XLSX = require('./');
			wb = XLSX.readFile(dir+'custom_properties.xlsx');
		});
		it('Must have read the core properties', function() {
			assert.equal(wb.Props.Company, 'Vector Inc');
			assert.equal(wb.Props.Creator, 'Pony Foo');
		});
		it('Must have read the custom properties', function() {
			assert.equal(wb.Custprops['I am a boolean'], true);
			assert.equal(wb.Custprops['Date completed'], '1967-03-09T16:30:00Z');
			assert.equal(wb.Custprops.Status, 2);
			assert.equal(wb.Custprops.Counter, -3.14);
		});
	});

	describe('should parse cells with date type', function() {
		var wb, ws;
		before(function() {
			XLSX = require('./');
			wb = XLSX.readFile(dir+'xlsx-stream-d-date-cell.xlsx');
			var sheetName = 'Sheet1';
			ws = wb.Sheets[sheetName];
		});
		it('Must have read the date', function() {
			var sheet = XLSX.utils.sheet_to_row_object_array(ws);
			assert.equal(sheet[3]['てすと'], '2/14/14');
		});
	});
});
