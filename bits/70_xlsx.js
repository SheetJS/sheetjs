
var ct2type = {
	"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml": "workbooks",
	"application/vnd.openxmlformats-package.core-properties+xml": "coreprops",
	"application/vnd.openxmlformats-officedocument.extended-properties+xml": "extprops",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml": "calcchains",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml":"sheets",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml": "strs",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml":"styles",
	"application/vnd.openxmlformats-officedocument.theme+xml":"themes",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml": "comments",
	"foo": "bar"
};

/* 18.2.28 (CT_WorkbookProtection) Defaults */
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

/* 18.2.30 (CT_BookView) Defaults */
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

/* 18.2.19 (CT_Sheet) Defaults */
var SheetDef = {
	state: 'visible'
};

/* 18.2.2  (CT_CalcPr) Defaults */
var CalcPrDef = {
	calcCompleted: 'true',
	calcMode: 'auto',
	calcOnSave: 'true',
	concurrentCalc: 'true',
	fullCalcOnLoad: 'false',
	fullPrecision: 'true',
	iterate: 'false',
	iterateCount: '100',
	iterateDelta: '0.001',
	refMode: 'A1'
};

/* 18.2.3 (CT_CustomWorkbookView) Defaults */
var CustomWBViewDef = {
	autoUpdate: 'false',
	changesSavedWin: 'false',
	includeHiddenRowCol: 'true',
	includePrintSettings: 'true',
	maximized: 'false',
	minimized: 'false',
	onlySync: 'false',
	personalView: 'false',
	showComments: 'commIndicator',
	showFormulaBar: 'true',
	showHorizontalScroll: 'true',
	showObjects: 'all',
	showSheetTabs: 'true',
	showStatusbar: 'true',
	showVerticalScroll: 'true',
	tabRatio: '600',
	xWindow: '0',
	yWindow: '0'
};

var XMLNS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
var XMLNS_WB = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

var strs = {}; // shared strings
var styles = {}; // shared styles
var _ssfopts = {}; // spreadsheet formatting options

/* 18.3 Worksheets */
function parseSheet(data) {
	if(!data) return data;
	/* 18.3.1.99 worksheet CT_Worksheet */
	var s = {};

	/* 18.3.1.35 dimension CT_SheetDimension ? */
	var ref = data.match(/<dimension ref="([^"]*)"\s*\/>/);
	if(ref && ref.length == 2 && ref[1].indexOf(":") !== -1) s["!ref"] = ref[1];

	var refguess = {s: {r:1000000, c:1000000}, e: {r:0, c:0} };
	var q = ["v","f"];
	var sidx = 0;
	/* 18.3.1.80 sheetData CT_SheetData ? */
	if(!data.match(/<sheetData *\/>/))
	data.match(/<sheetData>([^\u2603]*)<\/sheetData>/m)[1].split("</row>").forEach(function(x) {
		if(x === "" || x.trim() === "") return;

		/* 18.3.1.73 row CT_Row */
		var row = parsexmltag(x.match(/<row[^>]*>/)[0]);
		if(refguess.s.r > row.r - 1) refguess.s.r = row.r - 1;
		if(refguess.e.r < row.r - 1) refguess.e.r = row.r - 1;

		/* 18.3.1.4 c CT_Cell */
		var cells = x.substr(x.indexOf('>')+1).split(/<c/);
		cells.forEach(function(c, idx) { if(c === "" || c.trim() === "") return;
			var cref = c.match(/r="([^"]*)"/);
			c = "<c" + c;
			if(cref && cref.length == 2) {
				var cref_cell = decode_cell(cref[1]);
				idx = cref_cell.c;
			}
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
				case 's': {
					sidx = parseInt(p.v, 10);
					p.v = strs[sidx].t;
					p.r = strs[sidx].r;
				} break;
				case 'str': if(p.v) p.v = utf8read(p.v); break; // normal string
				case 'inlineStr':
					p.t = 'str'; p.v = unescapexml((d.match(matchtag('t'))||["",""])[1]);
					break; // inline string
				case 'b':
					switch(p.v) {
						case '0': case 'FALSE': case "false": case false: p.v=false; break;
						case '1': case 'TRUE':  case "true":  case true:  p.v=true;  break;
						default: throw "Unrecognized boolean: " + p.v;
					} break;
				/* in case of error, stick value in .raw */
				case 'e': p.raw = p.v; p.v = undefined; break;
				default: throw "Unrecognized cell type: " + p.t;
			}

			/* formatting */
			if(cell.s && styles.CellXf) { /* TODO: second check is a hacked guard */
				var cf = styles.CellXf[cell.s];
				if(cf && cf.numFmtId && cf.numFmtId !== 0) {
					p.raw = p.v;
					p.rawt = p.t;
					try {
						p.v = SSF.format(cf.numFmtId,p.v,_ssfopts);
						p.t = 'str';
					} catch(e) { p.v = p.raw; }
				}
			}

			s[cell.r] = p;
		});
	});
	if(!s["!ref"]) s["!ref"] = encode_range(refguess);
	return s;
}

function parseProps(data) {
	var p = { Company:'' }, q = {};
	var strings = ["Application", "DocSecurity", "Company", "AppVersion"];
	var bools = ["HyperlinksChanged","SharedDoc","LinksUpToDate","ScaleCrop"];
	var xtra = ["HeadingPairs", "TitlesOfParts"];
	var xtracp = ["category", "contentStatus", "lastModifiedBy", "lastPrinted", "revision", "version"];
	var xtradc = ["creator", "description", "identifier", "language", "subject", "title"];
	var xtradcterms = ["created", "modified"];
	xtra = xtra.concat(xtracp.map(function(x) { return "cp:" + x; }));
	xtra = xtra.concat(xtradc.map(function(x) { return "dc:" + x; }));
	xtra = xtra.concat(xtradcterms.map(function(x) { return "dcterms:" + x; }));


	strings.forEach(function(f){p[f] = (data.match(matchtag(f))||[])[1];});
	bools.forEach(function(f){p[f] = (data.match(matchtag(f))||[])[1] == "true";});
	xtra.forEach(function(f) {
		var cur = data.match(new RegExp("<" + f + "[^>]*>(.*)<\/" + f + ">"));
		if(cur && cur.length > 0) q[f] = cur[1];
	});

	if(q.HeadingPairs && q.TitlesOfParts) {
		var v = parseVector(q.HeadingPairs);
		var j = 0, widx = 0;
		for(var i = 0; i !== v.length; ++i) {
			switch(v[i].v) {
				case "Worksheets": widx = j; p.Worksheets = +v[++i]; break;
				case "Named Ranges": ++i; break; // TODO: Handle Named Ranges
			}
		}
		var parts = parseVector(q.TitlesOfParts).map(utf8read);
		p.SheetNames = parts.slice(widx, widx + p.Worksheets);
	}
	p.Creator = q["dc:creator"];
	p.LastModifiedBy = q["cp:lastModifiedBy"];
	p.CreatedDate = new Date(q["dcterms:created"]);
	p.ModifiedDate = new Date(q["dcterms:modified"]);
	return p;
}

/* 18.6 Calculation Chain */
function parseDeps(data) {
	var d = [];
	var l = 0, i = 1;
	(data.match(/<[^>]*>/g)||[]).forEach(function(x) {
		var y = parsexmltag(x);
		switch(y[0]) {
			case '<?xml': break;
			/* 18.6.2  calcChain CT_CalcChain 1 */
			case '<calcChain': case '<calcChain>': case '</calcChain>': break;
			/* 18.6.1  c CT_CalcCell 1 */
			case '<c': delete y[0]; if(y.i) i = y.i; else y.i = i; d.push(y); break;
		}
	});
	return d;
}

var ctext = {};

function parseCT(data) {
	if(!data || !data.match) return data;
	var ct = { workbooks: [], sheets: [], calcchains: [], themes: [], styles: [],
		coreprops: [], extprops: [], strs:[], comments: [], xmlns: "" };
	(data.match(/<[^>]*>/g)||[]).forEach(function(x) {
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
	if(ct.xmlns !== XMLNS_CT) throw new Error("Unknown Namespace: " + ct.xmlns);
	ct.calcchain = ct.calcchains.length > 0 ? ct.calcchains[0] : "";
	ct.sst = ct.strs.length > 0 ? ct.strs[0] : "";
	ct.style = ct.styles.length > 0 ? ct.styles[0] : "";
	delete ct.calcchains;
	return ct;
}


/* 18.2 Workbook */
function parseWB(data) {
	var wb = { AppVersion:{}, WBProps:{}, WBView:[], Sheets:[], CalcPr:{}, xmlns: "" };
	var pass = false;
	data.match(/<[^>]*>/g).forEach(function(x) {
		var y = parsexmltag(x);
		switch(y[0]) {
			case '<?xml': break;

			/* 18.2.27 workbook CT_Workbook 1 */
			case '<workbook': wb.xmlns = y.xmlns; break;
			case '</workbook>': break;

			/* 18.2.13 fileVersion CT_FileVersion ? */
			case '<fileVersion': delete y[0]; wb.AppVersion = y; break;
			case '<fileVersion/>': break;

			/* 18.2.12 fileSharing CT_FileSharing ? */
			case '<fileSharing': case '<fileSharing/>': break;

			/* 18.2.28 workbookPr CT_WorkbookPr ? */
			case '<workbookPr': delete y[0]; wb.WBProps = y; break;
			case '<workbookPr/>': delete y[0]; wb.WBProps = y; break;

			/* 18.2.29 workbookProtection CT_WorkbookProtection ? */
			case '<workbookProtection/>': break;

			/* 18.2.1  bookViews CT_BookViews ? */
			case '<bookViews>': case '</bookViews>': break;
			/* 18.2.30   workbookView CT_BookView + */
			case '<workbookView': delete y[0]; wb.WBView.push(y); break;

			/* 18.2.20 sheets CT_Sheets 1 */
			case '<sheets>': case '</sheets>': break; // aggregate sheet
			/* 18.2.19   sheet CT_Sheet + */
			case '<sheet': delete y[0]; y.name = utf8read(y.name); wb.Sheets.push(y); break;

			/* 18.2.15 functionGroups CT_FunctionGroups ? */
			case '<functionGroups': case '<functionGroups/>': break;
			/* 18.2.14   functionGroup CT_FunctionGroup + */
			case '<functionGroup': break;

			/* 18.2.9  externalReferences CT_ExternalReferences ? */
			case '<externalReferences': case '</externalReferences>': break;
			/* 18.2.8    externalReference CT_ExternalReference + */
			case '<externalReference': break;

			/* 18.2.6  definedNames CT_DefinedNames ? */
			case '<definedNames/>': break;
			case '<definedNames>': pass=true; break;
			case '</definedNames>': pass=false; break;
			/* 18.2.5    definedName CT_DefinedName + */
			case '<definedName': case '<definedName/>': case '</definedName>': break;

			/* 18.2.2  calcPr CT_CalcPr ? */
			case '<calcPr': delete y[0]; wb.CalcPr = y; break;
			case '<calcPr/>': delete y[0]; wb.CalcPr = y; break;

			/* 18.2.16 oleSize CT_OleSize ? (ref required) */
			case '<oleSize': break;

			/* 18.2.4  customWorkbookViews CT_CustomWorkbookViews ? */
			case '<customWorkbookViews>': case '</customWorkbookViews>': case '<customWorkbookViews': break;
			/* 18.2.3    customWorkbookView CT_CustomWorkbookView + */
			case '<customWorkbookView': case '</customWorkbookView>': break;

			/* 18.2.18 pivotCaches CT_PivotCaches ? */
			case '<pivotCaches>': case '</pivotCaches>': case '<pivotCaches': break;
			/* 18.2.17 pivotCache CT_PivotCache ? */
			case '<pivotCache': break;

			/* 18.2.21 smartTagPr CT_SmartTagPr ? */
			case '<smartTagPr': case '<smartTagPr/>': break;

			/* 18.2.23 smartTagTypes CT_SmartTagTypes ? */
			case '<smartTagTypes': case '<smartTagTypes>': case '</smartTagTypes>': break;
			/* 18.2.22   smartTagType CT_SmartTagType ? */
			case '<smartTagType': break;

			/* 18.2.24 webPublishing CT_WebPublishing ? */
			case '<webPublishing': case '<webPublishing/>': break;

			/* 18.2.11 fileRecoveryPr CT_FileRecoveryPr ? */
			case '<fileRecoveryPr': case '<fileRecoveryPr/>': break;

			/* 18.2.26 webPublishObjects CT_WebPublishObjects ? */
			case '<webPublishObjects>': case '<webPublishObjects': case '</webPublishObjects>': break;
			/* 18.2.25 webPublishObject CT_WebPublishObject ? */
			case '<webPublishObject': break;

			/* 18.2.10 extLst CT_ExtensionList ? */
			case '<extLst>': case '</extLst>': case '<extLst/>': break;
			/* 18.2.7    ext CT_Extension + */
			case '<ext': pass=true; break; //TODO: check with versions of excel
			case '</ext>': pass=false; break;

			/* Others */
			case '<mx:ArchID': break;
			case '<mc:AlternateContent': pass=true; break;
			case '</mc:AlternateContent>': pass=false; break;
		}
	});
	if(wb.xmlns !== XMLNS_WB) throw new Error("Unknown Namespace: " + wb.xmlns);

	var z;
	/* defaults */
	for(z in WBPropsDef) if(typeof wb.WBProps[z] === 'undefined') wb.WBProps[z] = WBPropsDef[z];
	for(z in CalcPrDef) if(typeof wb.CalcPr[z] === 'undefined') wb.CalcPr[z] = CalcPrDef[z];

	wb.WBView.forEach(function(w){for(var z in WBViewDef) if(typeof w[z] === 'undefined') w[z]=WBViewDef[z]; });
	wb.Sheets.forEach(function(w){for(var z in SheetDef) if(typeof w[z] === 'undefined') w[z]=SheetDef[z]; });

	_ssfopts.date1904 = parsexmlbool(wb.WBProps.date1904, 'date1904');

	return wb;
}

/* 18.8.31 numFmts CT_NumFmts */
function parseNumFmts(t) {
	styles.NumberFmt = [];
	for(var y in SSF._table) styles.NumberFmt[y] = SSF._table[y];
	t[0].match(/<[^>]*>/g).forEach(function(x) {
		var y = parsexmltag(x);
		switch(y[0]) {
			case '<numFmts': case '</numFmts>': case '<numFmts/>': break;
			case '<numFmt': {
				var f=unescapexml(y.formatCode), i=parseInt(y.numFmtId,10);
				styles.NumberFmt[i] = f; SSF.load(f,i);
			} break;
			default: throw 'unrecognized ' + y[0] + ' in numFmts';
		}
	});
}

/* 18.8.10 cellXfs CT_CellXfs */
function parseCXfs(t) {
	styles.CellXf = [];
	t[0].match(/<[^>]*>/g).forEach(function(x) {
		var y = parsexmltag(x);
		switch(y[0]) {
			case '<cellXfs': case '<cellXfs/>': case '</cellXfs>': break;

			/* 18.8.45 xf CT_Xf */
			case '<xf': if(y.numFmtId) y.numFmtId = parseInt(y.numFmtId, 10);
				styles.CellXf.push(y); break;
			case '</xf>': break;

			/* 18.8.1 alignment CT_CellAlignment */
			case '<alignment': break;

			/* 18.8.33 protection CT_CellProtection */
			case '<protection': case '</protection>': case '<protection/>': break;

			case '<extLst': case '</extLst>': break;
			case '<ext': break;
			default: throw 'unrecognized ' + y[0] + ' in cellXfs';
		}
	});
}

/* 18.8 Styles CT_Stylesheet*/
function parseStyles(data) {
	/* 18.8.39 styleSheet CT_Stylesheet */
	var t;

	/* numFmts CT_NumFmts ? */
	if((t=data.match(/<numFmts([^>]*)>.*<\/numFmts>/))) parseNumFmts(t);

	/* fonts CT_Fonts ? */
	/* fills CT_Fills ? */
	/* borders CT_Borders ? */
	/* cellStyleXfs CT_CellStyleXfs ? */

	/* cellXfs CT_CellXfs ? */
	if((t=data.match(/<cellXfs([^>]*)>.*<\/cellXfs>/))) parseCXfs(t);

	/* dxfs CT_Dxfs ? */
	/* tableStyles CT_TableStyles ? */
	/* colors CT_Colors ? */
	/* extLst CT_ExtensionList ? */

	return styles;
}

/* 9.3.2 OPC Relationships Markup */
function parseRels(data, currentFilePath) {
	if (!data) return data;
	if (currentFilePath.charAt(0) !== '/') {
		currentFilePath = '/'+currentFilePath;
	}
	var rels = {};

	var resolveRelativePathIntoAbsolute = function (to) {
	    var toksFrom = currentFilePath.split('/');
	 	toksFrom.pop(); // folder path
	    var toksTo = to.split('/');
	    var reversed = [];
	    while (toksTo.length !== 0) {
	        var tokTo = toksTo.shift();
	        if (tokTo === '..') {
	            toksFrom.pop();
	        } else if (tokTo !== '.') {
	            toksFrom.push(tokTo);
	        }
	    }
	    return toksFrom.join('/');
	}

	data.match(/<[^>]*>/g).forEach(function(x) {
		var y = parsexmltag(x);
		/* 9.3.2.2 OPC_Relationships */
		if (y[0] === '<Relationship') {
			var rel = {}; rel.Type = y.Type; rel.Target = y.Target; rel.Id = y.Id; rel.TargetMode = y.TargetMode;
			var canonictarget = resolveRelativePathIntoAbsolute(y.Target);
			rels[canonictarget] = rel;
		}
	});

	return rels;
}

/* 18.7.3 CT_Comment */
function parseComments(data) {
	if(data.match(/<comments *\/>/)) {
		throw new Error('Not a valid comments xml');
	}
	var authors = [];
	var commentList = [];
	data.match(/<authors>([^\u2603]*)<\/authors>/m)[1].split('</author>').forEach(function(x) {
		if(x === "" || x.trim() === "") return;
		authors.push(x.match(/<author[^>]*>(.*)/)[1]);
	});
	data.match(/<commentList>([^\u2603]*)<\/commentList>/m)[1].split('</comment>').forEach(function(x, index) {
		if(x === "" || x.trim() === "") return;
		var y = parsexmltag(x.match(/<comment[^>]*>/)[0]);
		var comment = { author: y.authorId && authors[y.authorId] ? authors[y.authorId] : undefined, ref: y.ref, guid: y.guid };
		var textMatch = x.match(/<text>([^\u2603]*)<\/text>/m);
		if (!textMatch || !textMatch[1]) return; // a comment may contain an empty text tag.
		var rt = parse_si(textMatch[1]);
		comment.raw = rt.raw;
		comment.t = rt.t;
		comment.r = rt.r;
		commentList.push(comment);
	});
	return commentList;
}

function parseCommentsAddToSheets(zip, dirComments, sheets, sheetRels) {
	for(var i = 0; i != dirComments.length; ++i) {
		var canonicalpath=dirComments[i];
		var comments=parseComments(getdata(getzipfile(zip, canonicalpath.replace(/^\//,''))));
		// find the sheets targeted by these comments
		var sheetNames = Object.keys(sheets);
		for(var j = 0; j != sheetNames.length; ++j) {
			var sheetName = sheetNames[j];
			var rels = sheetRels[sheetName];
			if (rels) {
				var rel = rels[canonicalpath];
				if (rel) {
					insertCommentsIntoSheet(sheetName, sheets[sheetName], comments);
				}
			}
		}
	}	
}

function insertCommentsIntoSheet(sheetName, sheet, comments) {
	comments.forEach(function(comment) {
		var cell = sheet[comment.ref];
		if (!cell) {
			cell = {};
			sheet[comment.ref] = cell;
			var range = decode_range(sheet["!ref"]);
			var thisCell = decode_cell(comment.ref);
			if(range.s.r > thisCell.r) range.s.r = thisCell.r;
			if(range.e.r < thisCell.r) range.e.r = thisCell.r;
			if(range.s.c > thisCell.c) range.s.c = thisCell.c;
			if(range.e.c < thisCell.c) range.e.c = thisCell.c;
			var encoded = encode_range(range);
			if (encoded !== sheet["!ref"]) sheet["!ref"] = encoded;
		}

		if (!cell.c) {
			cell.c = [];
		}
		cell.c.push({a: comment.author, t: comment.t, raw: comment.raw, r: comment.r});
	});
}

function getdata(data) {
	if(!data) return null;
	if(data.data) return data.data;
	if(data._data && data._data.getContent) return Array.prototype.slice.call(data._data.getContent(),0).map(function(x) { return String.fromCharCode(x); }).join("");
	return null;
}

function getzipfile(zip, file) {
	var f = file; if(zip.files[f]) return zip.files[f];
	f = file.toLowerCase(); if(zip.files[f]) return zip.files[f];
	f = f.replace(/\//g,'\\'); if(zip.files[f]) return zip.files[f];
	throw new Error("Cannot find file " + file + " in zip");
}

function parseZip(zip) {
	var entries = Object.keys(zip.files);
	var keys = entries.filter(function(x){return x.substr(-1) != '/';}).sort();
	var dir = parseCT(getdata(getzipfile(zip, '[Content_Types].xml')));
	if(dir.workbooks.length === 0) throw new Error("Could not find workbook entry");
	strs = {};
	if(dir.sst) strs=parse_sst(getdata(getzipfile(zip, dir.sst.replace(/^\//,''))));

	styles = {};
	if(dir.style) styles = parseStyles(getdata(getzipfile(zip, dir.style.replace(/^\//,''))));

	var wb = parseWB(getdata(getzipfile(zip, dir.workbooks[0].replace(/^\//,''))));
	var propdata = dir.coreprops.length !== 0 ? getdata(getzipfile(zip, dir.coreprops[0].replace(/^\//,''))) : "";
	propdata += dir.extprops.length !== 0 ? getdata(getzipfile(zip, dir.extprops[0].replace(/^\//,''))) : "";
	var props = propdata !== "" ? parseProps(propdata) : {};
	var deps = {};
	if(dir.calcchain) deps=parseDeps(getdata(getzipfile(zip, dir.calcchain.replace(/^\//,''))));
	var sheets = {}, i=0;
	var sheetRels = {};	
	if(!props.Worksheets) {
		/* Google Docs doesn't generate the appropriate metadata, so we impute: */
		var wbsheets = wb.Sheets;
		props.Worksheets = wbsheets.length;
		props.SheetNames = [];
		for(var j = 0; j != wbsheets.length; ++j) {
			props.SheetNames[j] = wbsheets[j].name;
		}
		for(i = 0; i != props.Worksheets; ++i) {
			try { /* TODO: remove these guards */
				var path = 'xl/worksheets/sheet' + (i+1) + '.xml';
				var relsPath = path.replace(/^(.*)(\/)([^\/]*)$/, "$1/_rels/$3.rels");
				sheets[props.SheetNames[i]]=parseSheet(getdata(getzipfile(zip, path)));
				sheetRels[props.SheetNames[i]]=parseRels(getdata(getzipfile(zip, relsPath)), path);
			} catch(e) {}
		}
	} else {
		for(i = 0; i != props.Worksheets; ++i) {
			try {
				var path = dir.sheets[i].replace(/^\//,'');
				var relsPath = path.replace(/^(.*)(\/)([^\/]*)$/, "$1/_rels/$3.rels");
				sheets[props.SheetNames[i]]=parseSheet(getdata(getzipfile(zip, path)));
				sheetRels[props.SheetNames[i]]=parseRels(getdata(getzipfile(zip, relsPath)), path);
			} catch(e) {}
		}
	}

	if(dir.comments) parseCommentsAddToSheets(zip, dir.comments, sheets, sheetRels);

	return {
		Directory: dir,
		Workbook: wb,
		Props: props,
		Deps: deps,
		Sheets: sheets,
		SheetNames: props.SheetNames,
		Strings: strs,
		Styles: styles,
		keys: keys,
		files: zip.files
	};
}

var _fs, jszip;
if(typeof JSZip !== 'undefined') jszip = JSZip;
if (typeof exports !== 'undefined') {
	if (typeof module !== 'undefined' && module.exports) {
		if(typeof jszip === 'undefined') jszip = require('./jszip').JSZip;
		_fs = require('fs');
	}
}

function readSync(data, options) {
	var zip, d = data;
	var o = options||{};
	switch((o.type||"base64")){
		case "file": d = _fs.readFileSync(data).toString('base64');
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

XLSX.read = readSync;
XLSX.readFile = readFileSync;
XLSX.parseZip = parseZip;
