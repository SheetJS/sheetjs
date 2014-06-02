/* ECMA-376 Part II 11.1 Core Properties Part */
/* [MS-OSHARED] 2.3.3.2.[1-2].1 (PIDSI/PIDDSI) */
var CORE_PROPS = [
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


function parse_core_props(data) {
	var p = {};

	CORE_PROPS.forEach(function(f) {
		var g = "(?:"+ f[0].substr(0,f[0].indexOf(":")) +":)"+ f[0].substr(f[0].indexOf(":")+1);
		var cur = data.match(new RegExp("<" + g + "[^>]*>(.*)<\/" + g + ">"));
		if(cur && cur.length > 0) p[f[1]] = cur[1];
		if(f[2] === 'date' && p[f[1]]) p[f[1]] = new Date(p[f[1]]);
	});

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

function write_core_props(cp, opts) {
	var o = [], p = {};
	o.push(XML_HEADER);
	o.push(CORE_PROPS_XML_ROOT);
	if(!cp) return o.join("");

	var doit = function(f, g, h) {
		if(p[f] || typeof g === 'undefined' || g === "") return;
		if(typeof g !== 'string') g = String(g); /* TODO: remove */
		p[f] = g;
		o.push(h ? writextag(f,g,h) : writetag(f,g));
	};

	if(typeof cp.CreatedDate !== 'undefined') doit("dcterms:created", typeof cp.CreatedDate === "string" ? cp.CreatedDate : write_w3cdtf(cp.CreatedDate, opts.WTF), {"xsi:type":"dcterms:W3CDTF"});
	if(typeof cp.ModifiedDate !== 'undefined') doit("dcterms:modified", typeof cp.ModifiedDate === "string" ? cp.ModifiedDate : write_w3cdtf(cp.ModifiedDate, opts.WTF), {"xsi:type":"dcterms:W3CDTF"});

	CORE_PROPS.forEach(function(f) { doit(f[0], cp[f[1]]); });
	if(o.length>2){ o.push('</cp:coreProperties>'); o[1]=o[1].replace("/>",">"); }
	return o.join("");
}
