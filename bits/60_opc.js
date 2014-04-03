/* Parts enumerated in OPC spec, MS-XLSB and MS-XLSX */
/* 12.3 Part Summary <SpreadsheetML> */
/* 14.2 Part Summary <DrawingML> */
/* [MS-XLSX] 2.1 Part Enumerations */
/* [MS-XLSB] 2.1.7 Part Enumeration */
var ct2type = {
	/* Workbook */
	"application/vnd.ms-excel.main": "workbooks",
	"application/vnd.ms-excel.sheet.macroEnabled.main+xml": "workbooks",
	"application/vnd.ms-excel.sheet.binary.macroEnabled.main": "workbooks",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml": "workbooks",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml": "TODO", /* Template */

	/* Worksheet */
	"application/vnd.ms-excel.worksheet": "sheets",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml": "sheets",
	"application/vnd.ms-excel.binIndexWs": "TODO", /* Binary Index */

	/* Chartsheet */
	"application/vnd.ms-excel.chartsheet": "TODO",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml": "TODO",

	/* Dialogsheet */
	"application/vnd.ms-excel.dialogsheet": "TODO",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml": "TODO",

	/* Macrosheet */
	"application/vnd.ms-excel.macrosheet": "TODO",
	"application/vnd.ms-excel.macrosheet+xml": "TODO",
	"application/vnd.ms-excel.intlmacrosheet": "TODO",
	"application/vnd.ms-excel.binIndexMs": "TODO", /* Binary Index */

	/* Shared Strings */
	"application/vnd.ms-excel.sharedStrings": "strs",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml": "strs",

	/* Styles */
	"application/vnd.ms-excel.styles": "styles",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml": "styles",

	/* File Properties */
	"application/vnd.openxmlformats-package.core-properties+xml": "coreprops",
	"application/vnd.openxmlformats-officedocument.custom-properties+xml": "custprops",
	"application/vnd.openxmlformats-officedocument.extended-properties+xml": "extprops",

	/* Custom Data Properties */
	"application/vnd.openxmlformats-officedocument.customXmlProperties+xml": "TODO",

	/* Comments */
	"application/vnd.ms-excel.comments": "comments",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml": "comments",

	/* PivotTable */
	"application/vnd.ms-excel.pivotTable": "TODO",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml": "TODO",

	/* Calculation Chain */
	"application/vnd.ms-excel.calcChain": "calcchains",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml": "calcchains",

	/* Printer Settings */
	"application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings": "TODO",

	/* ActiveX */
	"application/vnd.ms-office.activeX": "TODO",
	"application/vnd.ms-office.activeX+xml": "TODO",

	/* Custom Toolbars */
	"application/vnd.ms-excel.attachedToolbars": "TODO",

	/* External Data Connections */
	"application/vnd.ms-excel.connections": "TODO",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml": "TODO",

	/* External Links */
	"application/vnd.ms-excel.externalLink": "TODO",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml": "TODO",

	/* Metadata */
	"application/vnd.ms-excel.sheetMetadata": "TODO",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml": "TODO",

	/* PivotCache */
	"application/vnd.ms-excel.pivotCacheDefinition": "TODO",
	"application/vnd.ms-excel.pivotCacheRecords": "TODO",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml": "TODO",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml": "TODO",

	/* Query Table */
	"application/vnd.ms-excel.queryTable": "TODO",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml": "TODO",

	/* Shared Workbook */
	"application/vnd.ms-excel.userNames": "TODO",
	"application/vnd.ms-excel.revisionHeaders": "TODO",
	"application/vnd.ms-excel.revisionLog": "TODO",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.revisionHeaders+xml": "TODO",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.revisionLog+xml": "TODO",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.userNames+xml": "TODO",

	/* Single Cell Table */
	"application/vnd.ms-excel.tableSingleCells": "TODO",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.tableSingleCells+xml": "TODO",

	/* Slicer */
	"application/vnd.ms-excel.slicer": "TODO",
	"application/vnd.ms-excel.slicerCache": "TODO",
	"application/vnd.ms-excel.slicer+xml": "TODO",
	"application/vnd.ms-excel.slicerCache+xml": "TODO",

	/* Sort Map */
	"application/vnd.ms-excel.wsSortMap": "TODO",

	/* Table */
	"application/vnd.ms-excel.table": "TODO",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml": "TODO",

	/* Themes */
	"application/vnd.openxmlformats-officedocument.theme+xml": "themes",

	/* Timeline */
	"application/vnd.ms-excel.Timeline+xml": "TODO", /* verify */
	"application/vnd.ms-excel.TimelineCache+xml": "TODO", /* verify */

	/* VBA */
	"application/vnd.ms-office.vbaProject": "vba",
	"application/vnd.ms-office.vbaProjectSignature": "vba",

	/* Volatile Dependencies */
	"application/vnd.ms-office.volatileDependencies": "TODO",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.volatileDependencies+xml": "TODO",

	/* Control Properties */
	"application/vnd.ms-excel.controlproperties+xml": "TODO",

	/* Data Model */
	"application/vnd.openxmlformats-officedocument.model+data": "TODO",

	/* Survey */
	"application/vnd.ms-excel.Survey+xml": "TODO",

	/* Drawing */
	"application/vnd.openxmlformats-officedocument.drawing+xml": "TODO",
	"application/vnd.openxmlformats-officedocument.drawingml.chart+xml": "TODO",
	"application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml": "TODO",
	"application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml": "TODO",
	"application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml": "TODO",
	"application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml": "TODO",
	"application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml": "TODO",

	/* VML */
	"application/vnd.openxmlformats-officedocument.vmlDrawing": "TODO",

	"application/vnd.openxmlformats-package.relationships+xml": "TODO",
	"application/vnd.openxmlformats-officedocument.oleObject": "TODO",

	"foo": "bar"
};

var XMLNS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';

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
				case "Worksheets": widx = j; p.Worksheets = +(v[++i].v); break;
				case "Named Ranges": ++i; break; // TODO: Handle Named Ranges
			}
		}
		var parts = parseVector(q.TitlesOfParts).map(function(x) { return utf8read(x.v); });
		p.SheetNames = parts.slice(widx, widx + p.Worksheets);
	}
	p.Creator = q["dc:creator"];
	p.LastModifiedBy = q["cp:lastModifiedBy"];
	p.CreatedDate = new Date(q["dcterms:created"]);
	p.ModifiedDate = new Date(q["dcterms:modified"]);
	return p;
}

/* 15.2.12.2 Custom File Properties Part */
function parseCustomProps(data) {
	var p = {}, name;
	data.match(/<[^>]+>([^<]*)/g).forEach(function(x) {
		var y = parsexmltag(x);
		switch(y[0]) {
			case '<property': name = y.name; break;
			case '</property>': name = null; break;
			default: if (x.indexOf('<vt:') === 0) {
				var toks = x.split('>');
				var type = toks[0].substring(4), text = toks[1];
				/* 22.4.2.32 (CT_Variant). Omit the binary types from 22.4 (Variant Types) */
				switch(type) {
					case 'lpstr': case 'lpwstr': case 'bstr': case 'lpwstr':
						p[name] = unescapexml(text);
						break;
					case 'bool':
						p[name] = parsexmlbool(text, '<vt:bool>');
						break;
					case 'i1': case 'i2': case 'i4': case 'i8': case 'int': case 'uint':
						p[name] = parseInt(text, 10);
						break;
					case 'r4': case 'r8': case 'decimal':
						p[name] = parseFloat(text);
						break;
					case 'filetime': case 'date':
						p[name] = text; // should we make this into a date?
						break;
					case 'cy': case 'error':
						p[name] = unescapexml(text);
						break;
					default:
						console.warn('Unexpected', x, type, toks);
				}
			}
		}
	});
	return p;
}

var ctext = {};
function parseCT(data, opts) {
	if(!data || !data.match) return data;
	var ct = { workbooks: [], sheets: [], calcchains: [], themes: [], styles: [],
		coreprops: [], extprops: [], custprops: [], strs:[], comments: [], vba: [],
		TODO:[], xmlns: "" };
	(data.match(/<[^>]*>/g)||[]).forEach(function(x) {
		var y = parsexmltag(x);
		switch(y[0]) {
			case '<?xml': break;
			case '<Types': ct.xmlns = y.xmlns; break;
			case '<Default': ctext[y.Extension] = y.ContentType; break;
			case '<Override':
				if(y.ContentType in ct2type)ct[ct2type[y.ContentType]].push(y.PartName);
				else if(opts.WTF) console.error(y.ContentType);
				break;
		}
	});
	if(ct.xmlns !== XMLNS_CT) throw new Error("Unknown Namespace: " + ct.xmlns);
	ct.calcchain = ct.calcchains.length > 0 ? ct.calcchains[0] : "";
	ct.sst = ct.strs.length > 0 ? ct.strs[0] : "";
	ct.style = ct.styles.length > 0 ? ct.styles[0] : "";
	ct.defaults = ctext;
	delete ct.calcchains;
	return ct;
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
	};

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


