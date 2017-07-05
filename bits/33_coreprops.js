/* ECMA-376 Part II 11.1 Core Properties Part */
/* [MS-OSHARED] 2.3.3.2.[1-2].1 (PIDSI/PIDDSI) */
var CORE_PROPS/*:Array<Array<string> >*/ = [
	["cp:category", "Category"],
	["cp:contentStatus", "ContentStatus"],
	["cp:keywords", "Keywords"],
	["cp:lastModifiedBy", "LastAuthor"],
	["cp:lastPrinted", "LastPrinted"],
	["cp:revision", "RevNumber"],
	["cp:version", "Version"],
	["dc:creator", "Author"],
	["dc:description", "Comments"],
	["dc:identifier", "Identifier"],
	["dc:language", "Language"],
	["dc:subject", "Subject"],
	["dc:title", "Title"],
	["dcterms:created", "CreatedDate", 'date'],
	["dcterms:modified", "ModifiedDate", 'date']
];

XMLNS.CORE_PROPS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
RELS.CORE_PROPS  = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';

var CORE_PROPS_REGEX/*:Array<RegExp>*/ = (function() {
	var r = new Array(CORE_PROPS.length);
	for(var i = 0; i < CORE_PROPS.length; ++i) {
		var f = CORE_PROPS[i];
		var g = "(?:"+ f[0].substr(0,f[0].indexOf(":")) +":)"+ f[0].substr(f[0].indexOf(":")+1);
		r[i] = new RegExp("<" + g + "[^>]*>([\\s\\S]*?)<\/" + g + ">");
	}
	return r;
})();

function parse_core_props(data) {
	var p = {};

	for(var i = 0; i < CORE_PROPS.length; ++i) {
		var f = CORE_PROPS[i], cur = data.match(CORE_PROPS_REGEX[i]);
		if(cur != null && cur.length > 0) p[f[1]] = cur[1];
		if(f[2] === 'date' && p[f[1]]) p[f[1]] = parseDate(p[f[1]]);
	}

	return p;
}

var CORE_PROPS_XML_ROOT = writextag('cp:coreProperties', null, {
	//'xmlns': XMLNS.CORE_PROPS,
	'xmlns:cp': XMLNS.CORE_PROPS,
	'xmlns:dc': XMLNS.dc,
	'xmlns:dcterms': XMLNS.dcterms,
	'xmlns:dcmitype': XMLNS.dcmitype,
	'xmlns:xsi': XMLNS.xsi
});

function cp_doit(f, g, h, o, p) {
	if(p[f] != null || g == null || g === "") return;
	p[f] = g;
	o[o.length] = (h ? writextag(f,g,h) : writetag(f,g));
}

function write_core_props(cp, _opts) {
	var opts = _opts || {};
	var o = [XML_HEADER, CORE_PROPS_XML_ROOT], p = {};
	if(!cp && !opts.Props) return o.join("");

	if(cp) {
		if(cp.CreatedDate != null) cp_doit("dcterms:created", typeof cp.CreatedDate === "string" ? cp.CreatedDate : write_w3cdtf(cp.CreatedDate, opts.WTF), {"xsi:type":"dcterms:W3CDTF"}, o, p);
		if(cp.ModifiedDate != null) cp_doit("dcterms:modified", typeof cp.ModifiedDate === "string" ? cp.ModifiedDate : write_w3cdtf(cp.ModifiedDate, opts.WTF), {"xsi:type":"dcterms:W3CDTF"}, o, p);
	}

	for(var i = 0; i != CORE_PROPS.length; ++i) {
		var f = CORE_PROPS[i];
		var v = opts.Props && opts.Props[f[1]] != null ? opts.Props[f[1]] : cp ? cp[f[1]] : null;
		if(v === true) v = "1";
		else if(v === false) v = "0";
		else if(typeof v == "number") v = String(v);
		if(v != null) cp_doit(f[0], v, null, o, p);
	}
	if(o.length>2){ o[o.length] = ('</cp:coreProperties>'); o[1]=o[1].replace("/>",">"); }
	return o.join("");
}
