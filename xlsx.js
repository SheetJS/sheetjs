/* vim: set ts=2:*/
/*jshint eqnull:true */
var XLSX = (function(){
var debug = 0;
var ct2type = {
	"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml": "workbooks",
	"application/vnd.openxmlformats-package.core-properties+xml": "coreprops",
	"application/vnd.openxmlformats-officedocument.extended-properties+xml": "extprops",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml": "calcchains",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml":"sheets",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml": "strs",	
	"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml":"styles",
	"application/vnd.openxmlformats-officedocument.theme+xml":"themes",
	"foo": "bar"
};

var WBPropsDef = {
	allowRefreshQuery: '0',
	autoCompressPictures: '1',
	backupFile: '0',
	checkCompatibility: '0',
	codeName: '',
	date1904: '0',
	dateCompatibility: '1',
	//defaultThemeVersion: '0',
	filterPrivacy: '0',
	hidePivotFieldList: '0',
	promptedSolutions: '0',
	publishItems: '0',
	refreshAllConnections: false,
	saveExternalLinkValues: '1',
	showBorderUnselectedTables: '1',
	showInkAnnotation: '1',
	showObjects: 'all',
	showPivotChartFilter: '0'
	//updateLinks: 'userSet'
};

var WBViewDef = {
	activeTab: '0',
	autoFilterDateGrouping: '1',
	firstSheet: '0',
	minimized: '0',
	showHorizontalScroll: '1',
	showSheetTabs: '1',
	showVerticalScroll: '1',
	tabRatio: '600',
	visibility: 'visible'
	//window{Height,Width}, {x,y}Window
};

var SheetDef = {
	state: 'visible'
};

var CalcPrDef = {
	calcCompleted: '1',
	calcMode: 'auto',
	calcOnSave: '1',
	concurrentCalc: '1',
	fullCalcOnLoad: '0',
	iterate: 'false',
	iterateCount: '100',
	iterateDelta: '0.001',
	refMode: 'A1'
};

var XMLNS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
var XMLNS_WB = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

var encodings = {
	'&gt;': '>',
	'&lt;': '<',
	'&amp;': '&'
};

function unescapexml(text){
	var s = text + '';
	for(var y in encodings) s = s.replace(new RegExp(y,'g'), encodings[y]);
	return s;
}

function parsexmltag(tag) {
	var words = tag.split(/\s+/);
	var z = {'0': words[0]};
	if(words.length === 1) return z;
	tag.match(/(\w+)="([^"]*)"/g).map(
		function(x){var y=x.match(/(\w+)="([^"]*)"/); z[y[1]] = y[2]; });
	return z; 
}


var strs = {}; // shared strings


function parseSheet(data) { //TODO: use a real xml parser
	var s = {};
	var ref = data.match(/<dimension ref="([^"]*)"\s*\/>/);
	if(ref) s["!ref"] = ref[1];
	var refguess = {s: {r:1000000, c:1000000}, e: {r:0, c:0} };
	//s.rows = {};
	//s.cells = {};
	var q = ["v","f"];
	if(!data.match(/<sheetData *\/>/)) 
	data.match(/<sheetData>(.*)<\/sheetData>/)[1].split("</row>").forEach(function(x) { if(x === "") return;
		var row = parsexmltag(x.match(/<row[^>]*>/)[0]); //s.rows[row.r]=row.spans;
		if(refguess.s.r > row.r - 1) refguess.s.r = row.r - 1; 
		if(refguess.e.r < row.r - 1) refguess.e.r = row.r - 1;
		var cells = x.substr(x.indexOf('>')+1).split(/<\/c>|\/>/);
		cells.forEach(function(c, idx) { if(c === "") return;
			if(refguess.s.c > idx) refguess.s.c = idx; 
			if(refguess.e.c < idx) refguess.e.c = idx;
			var cell = parsexmltag((c.match(/<c[^>]*>/)||[c])[0]); delete cell[0];
			var d = c.substr(c.indexOf('>')+1);
			var p = {};
			q.forEach(function(f){var x=d.match(matchtag(f));if(x)p[f]=unescapexml(x[1]);});
			/* SCHEMA IS ACTUALLY INCORRECT HERE.  IF A CELL HAS NO T, EMIT "" */
			if(cell.t === undefined && p.v === undefined) { p.t = "str"; p.v = undefined; }
			else p.t = (cell.t ? cell.t : "n"); // default is "n" in schema
			switch(p.t) {
				case 'n': p.v = parseFloat(p.v); break;
				case 's': p.v = strs[parseInt(p.v, 10)].t; break;
				case 'str': break; // normal string
				default: throw "Unrecognized cell type: " + p.t;
			}
			//s.cells[cell.r] = p;
			s[cell.r] = p;
		});
	});
	if(!s["!ref"]) s["!ref"] = encode_range(refguess);

	if(debug) s.rawdata = data;
	return s;
}

// matches <foo>...</foo> extracts content
function matchtag(f,g) {return new RegExp('<' + f + '>([\\s\\S]*)</' + f + '>',g||"");}

function parseStrs(data) { 
	var s = [];
	var sst = data.match(new RegExp("<sst ([^>]*)>([\\s\\S]*)<\/sst>","m"));
	if(sst) {
		s = sst[2].replace(/<si>/g,"").split(/<\/si>/).map(function(x) { var z = {};
			var y=x.match(/<(.*)>([\s\S]*)<\/.*/); if(y) z[y[1]]=unescapexml(y[2]); return z;});
	
		sst = parsexmltag(sst[1]); s.count = sst.count; s.uniqueCount = sst.uniqueCount;
	}
	if(debug) s.rawdata = data;
	return s;
}

function parseProps(data) {
	var p = { Company:'' }, q = {};
	var strings = ["Application", "DocSecurity", "Company", "AppVersion"];
	var bools = ["HyperlinksChanged","SharedDoc","LinksUpToDate","ScaleCrop"];
	var xtra = ["HeadingPairs", "TitlesOfParts","dc:creator","cp:lastModifiedBy","dcterms:created", "dcterms:modified"];
	
	strings.forEach(function(f){p[f] = (data.match(matchtag(f))||[])[1];});
	bools.forEach(function(f){p[f] = (data.match(matchtag(f))||[])[1] == "true";});
	xtra.forEach(function(f) {
		var cur = data.match(new RegExp("<" + f + "[^>]*>(.*)<\/" + f + ">"));
		if(cur && cur.length > 0) q[f] = cur[1];
	});

	if(q["HeadingPairs"]) p["Worksheets"] = parseInt(q["HeadingPairs"].match(new RegExp("<vt:i4>(.*)<\/vt:i4>"))[1], 10); 
	if(q["TitlesOfParts"]) p["SheetNames"] = q["TitlesOfParts"].match(new RegExp("<vt:lpstr>([^<]*)<\/vt:lpstr>","g")).map(function(x){return x.match(new RegExp("<vt:lpstr>([^<]*)<\/vt:lpstr>"))[1];});
	p["Creator"] = q["dc:creator"];
	p["LastModifiedBy"] = q["cp:lastModifiedBy"];
	p["CreatedDate"] = new Date(q["dcterms:created"]);
	p["ModifiedDate"] = new Date(q["dcterms:modified"]);
	
	if(debug) p.rawdata = data;
	return p;
}

function parseDeps(data) {
	var d = [];
	var l = 0, i = 1;
	data.match(/<[^>]*>/g).forEach(function(x) {
		var y = parsexmltag(x);
		switch(y[0]) {
			case '<?xml': break;
			case '<calcChain': break;
			case '<c': delete y[0]; if(y.i) i = y.i; else y.i = i; d.push(y); break;
		}
	});
	if(debug) d.rawdata = data;
	return d;
}

var ctext = {};

function parseCT(data) {
	var ct = { workbooks: [], sheets: [], calcchains: [], themes: [], styles: [], 
		coreprops: [], extprops: [], strs:[], xmlns: "" };
	if(data == null) return data;
	data.match(/<[^>]*>/g).forEach(function(x) {
		var y = parsexmltag(x);
		switch(y[0]) {
			case '<?xml': break;
			case '<Types': ct.xmlns = y.xmlns; break;
			case '<Default': ctext[y.Extension] = y.ContentType; break;
			case '<Override': 
				if(y.ContentType in ct2type)ct[ct2type[y.ContentType]].push(y.PartName);
				break;
		}
	});
	if(ct.xmlns !== XMLNS_CT) throw "Unknown Namespace: " + ct.xmlns;
	ct.calcchain = ct.calcchains.length > 0 ? ct.calcchains[0] : "";
	delete ct.calcchains;
	if(debug) ct.rawdata = data;
	return ct;
}


function parseWB(data) {
	var wb = { AppVersion:{}, WBProps:{}, WBView:[], Sheets:[], CalcPr:{}, xmlns: "" };
	var pass = false;
	data.match(/<[^>]*>/g).forEach(function(x) {
		var y = parsexmltag(x);
		switch(y[0]) {
			case '<?xml': break;
			case '<workbook': wb.xmlns = y.xmlns; break;
			case '<fileVersion':
				//if(y.appName != "xl") throw "Unexpected workbook.appName: "+y.appName;
				delete y[0]; wb.AppVersion = y; break;
			case '<workbookPr': delete y[0]; wb.WBProps = y; break;
			case '<workbookPr/>': delete y[0]; wb.WBProps = y; break;
			case '<bookViews>': case '</bookViews>': break; // aggregate workbookView
			case '<workbookView': delete y[0]; wb.WBView.push(y); break;
			case '<sheets>': case '</sheets>': break; // aggregate sheet
			case '<sheet': delete y[0]; wb.Sheets.push(y); break; 
			case '</extLst>': case '</workbook>': break;
			case '<workbookProtection/>': break; // LibreOffice 
			case '<extLst>': break; 
			case '<calcPr': delete y[0]; wb.CalcPr = y; break;
			case '<calcPr/>': delete y[0]; wb.CalcPr = y; break;
			
			case '<definedNames/>': break;
			case '<mx:ArchID': break;
			case '<ext': pass=true; break; //TODO: check with versions of excel
			case '</ext>': pass=false; break; 
			
			/* Introduced for Excel2013 Baseline */
			case '<mc:AlternateContent': pass=true; break; // TODO: do something 
			case '</mc:AlternateContent>': pass=false; break; // TODO: do something
			default: if(!pass) console.error("WB Tag",x,y);
		}
	});
	if(wb.xmlns !== XMLNS_WB) throw "Unknown Namespace: " + wb.xmlns;
	
	var z;
	for(z in WBPropsDef) if(null == wb.WBProps[z]) wb.WBProps[z] = WBPropsDef[z];
	wb.WBView.forEach(function(w){for(var z in WBViewDef) if(null==w[z]) w[z]=WBViewDef[z]; });
	for(z in CalcPrDef) if(null == wb.CalcPr[z]) wb.CalcPr[z] = CalcPrDef[z];
	wb.Sheets.forEach(function(w){for(var z in SheetDef) if(null==w[z]) w[z]=SheetDef[z]; }); 
	if(debug) wb.rawdata = data;
	return wb;
}

function parseZip(zip) {
	var entries = Object.keys(zip.files);
	var keys = entries.filter(function(x){return x.substr(-1) != '/';}).sort();
	var dir = parseCT((zip.files['[Content_Types].xml']||{}).data);
	var wb = parseWB(zip.files[dir.workbooks[0].replace(/^\//,'')].data);
	var propdata = dir.coreprops.length !== 0 ? zip.files[dir.coreprops[0].replace(/^\//,'')].data : "";
	propdata += dir.extprops.length !== 0 ? zip.files[dir.extprops[0].replace(/^\//,'')].data : "";
	var props = propdata !== "" ? parseProps(propdata) : {};
	var deps = {};
	if(dir.calcchain) deps=parseDeps(zip.files[dir.calcchain.replace(/^\//,'')].data);
	if(dir.strs[0]) strs=parseStrs(zip.files[dir.strs[0].replace(/^\//,'')].data);
	var sheets = {};
	if(!props.Worksheets) {
		var wbsheets = wb.Sheets;
		props.Worksheets = wbsheets.length;
		props.SheetNames = [];
		for(var j = 0; j != wbsheets.length; ++j) {
			props.SheetNames[j] = wbsheets[j].name;
		}
	}
	for(var i = 0; i != props.Worksheets; ++i) {
		sheets[props.SheetNames[i]]=parseSheet(zip.files[dir.sheets[i].replace(/^\//,'')].data);
	}

	return {
		Directory: dir,
		Workbook: wb,
		Props: props,
		Deps: deps,
		Sheets: sheets,
		SheetNames: props.SheetNames,
		Strings: strs,
		keys: keys,
		files: zip.files
	};
}

var fs, jszip;
if(typeof JSZip !== "undefined") jszip = JSZip;
if(typeof require !== "undefined") {
	if(typeof jszip === 'undefined') jszip = require('./jszip').JSZip;
	fs = require('fs');
}

function readSync(data, options) {
	var zip, d = data;
	var o = options||{};
	switch((o.type||"base64")){
		case "file": d = fs.readFileSync(data).toString('base64');
			/* falls through */
		case "base64": zip = new jszip(d, { base64:true }); break;
		case "binary": zip = new jszip(d, { base64:false }); break;
	}
	return parseZip(zip);
}

function readFileSync(data, options) {
	var o = options||{}; o.type = 'file';
	return readSync(data, o);
}

this.read = readSync;
this.readFile = readFileSync;
this.parseZip = parseZip;
return this;

})();

function encode_col(col) { var s=""; for(++col; col; col=Math.floor((col-1)/26)) s = String.fromCharCode(((col-1)%26) + 65) + s; return s; }
function encode_row(row) { return "" + (row + 1); }
function encode_cell(cell) { return encode_col(cell.c) + encode_row(cell.r); }

function decode_col(c) { var d = 0, i = 0; for(; i !== c.length; ++i) d = 26*d + c.charCodeAt(i) - 64; return d - 1; }
function decode_row(rowstr) { return Number(rowstr) - 1; }
function split_cell(cstr) { return cstr.replace(/(\$?[A-Z]*)(\$?[0-9]*)/,"$1,$2").split(","); }
function decode_cell(cstr) { var splt = split_cell(cstr); return { c:decode_col(splt[0]), r:decode_row(splt[1]) }; }
function decode_range(range) { var x =range.split(":").map(decode_cell); return {s:x[0],e:x[x.length-1]}; }
function encode_range(range) { return encode_cell(range.s) + ":" + encode_cell(range.e); }
/**
 * Convert a sheet into an array of objects where the column headers are keys.
 **/
function sheet_to_row_object_array(sheet){
	var val, rowObject, range, columnHeaders, emptyRow, C;
	var outSheet = [];
	if (sheet["!ref"]) {
		range = decode_range(sheet["!ref"]);

		columnHeaders = {};
		for (C = range.s.c; C <= range.e.c; ++C) {
			val = sheet[encode_cell({
				c: C,
				r: range.s.r
			})];
			if(val){
				if(val.t === "s"){
					columnHeaders[C] = val.v;
				}
			}
		}

		for (var R = range.s.r + 1; R <= range.e.r; ++R) {
			emptyRow = true;
			//Row number is recorded in the prototype
			//so that it doesn't appear when stringified.
			rowObject = Object.create({ __rowNum__ : R });
			for (C = range.s.c; C <= range.e.c; ++C) {
				val = sheet[encode_cell({
					c: C,
					r: R
				})];
				if(val !== undefined) switch(val.t){
					case 's': case 'str':
						if(val.v !== undefined) {
							rowObject[columnHeaders[C]] = val.v;
							emptyRow = false;
						}
						break;
					default: throw 'unrecognized type ' + val.t;
				}
			}
			if(!emptyRow) {
				outSheet.push(rowObject);
			}
		}
	}
	return outSheet;
}

function sheet_to_csv(sheet) {
	var stringify = function stringify(val) {
		switch(val.t){
			case 'n': return val.v;
			case 's': case 'str': return JSON.stringify(val.v);
			default: throw 'unrecognized type ' + val.t;
		}
	};
	var out = "";
	if(sheet["!ref"]) {
		var r = utils.decode_range(sheet["!ref"]);
		for(var R = r.s.r; R <= r.e.r; ++R) { 
			var row = [];
			for(var C = r.s.c; C <= r.e.c; ++C) {
				var val = sheet[utils.encode_cell({c:C,r:R})];
				row.push(val ? stringify(val) : "");
			}
			out += row.join(",") + "\n";
		}
	}
	return out;
}

var utils = {
	encode_col: encode_col,
	encode_row: encode_row,
	encode_cell: encode_cell,
	encode_range: encode_range,
	decode_col: decode_col,
	decode_row: decode_row,
	split_cell: split_cell,
	decode_cell: decode_cell,
	decode_range: decode_range,
	sheet_to_csv: sheet_to_csv,
	sheet_to_row_object_array: sheet_to_row_object_array
};

if(typeof require !== 'undefined' && typeof exports !== 'undefined') {
	exports.read = XLSX.read;
	exports.readFile = XLSX.readFile;
	exports.utils = utils;
	exports.main = function(args) {
		var zip = XLSX.read(args[0], {type:'file'});
		console.log(zip.Sheets);
	};
if(typeof module !== 'undefined' && require.main === module) 
	exports.main(process.argv.slice(2));
}
