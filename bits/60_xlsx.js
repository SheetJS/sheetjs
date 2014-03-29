
var ct2type = {
	"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml": "workbooks", /*XLSX*/
	"application/vnd.ms-excel.sheet.macroEnabled.main+xml":"workbooks", /*XLSM*/
	"application/vnd.ms-excel.sheet.binary.macroEnabled.main":"workbooks", /*XLSB*/

	"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml":"sheets", /*XLS[XM]*/
	"application/vnd.ms-excel.worksheet":"sheets", /*XLSB*/

	"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml":"styles", /*XLS[XM]*/
	"application/vnd.ms-excel.styles":"styles", /*XLSB*/

	"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml": "strs", /*XLS[XM]*/
	"application/vnd.ms-excel.sharedStrings": "strs", /*XLSB*/

	"application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml": "calcchains", /*XLS[XM]*/
	//"application/vnd.ms-excel.calcChain": "calcchains", /*XLSB*/

	"application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml": "comments", /*XLS[XM]*/
	"application/vnd.ms-excel.comments": "comments", /*XLSB*/

	"application/vnd.openxmlformats-package.core-properties+xml": "coreprops",
	"application/vnd.openxmlformats-officedocument.extended-properties+xml": "extprops",
	"application/vnd.openxmlformats-officedocument.custom-properties+xml": "custprops",
	"application/vnd.openxmlformats-officedocument.theme+xml":"themes",
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
		coreprops: [], extprops: [], custprops: [], strs:[], comments: [], xmlns: "" };
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


