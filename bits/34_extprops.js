/* 15.2.12.3 Extended File Properties Part */
/* [MS-OSHARED] 2.3.3.2.[1-2].1 (PIDSI/PIDDSI) */
var EXT_PROPS/*:Array<Array<string> >*/ = [
	["Application", "Application", "string"],
	["AppVersion", "AppVersion", "string"],
	["Company", "Company", "string"],
	["DocSecurity", "DocSecurity", "string"],
	["Manager", "Manager", "string"],
	["HyperlinksChanged", "HyperlinksChanged", "bool"],
	["SharedDoc", "SharedDoc", "bool"],
	["LinksUpToDate", "LinksUpToDate", "bool"],
	["ScaleCrop", "ScaleCrop", "bool"],
	["HeadingPairs", "HeadingPairs", "raw"],
	["TitlesOfParts", "TitlesOfParts", "raw"]
];

var PseudoPropsPairs = [
	"Worksheets",  "SheetNames",
	"NamedRanges", "DefinedNames",
	"Chartsheets", "ChartNames"
];
function load_props_pairs(HP/*:string|Array<Array<any>>*/, TOP, props, opts) {
	var v = [];
	if(typeof HP == "string") v = parseVector(HP, opts);
	else for(var j = 0; j < HP.length; ++j) v = v.concat(HP[j].map(function(hp) { return {v:hp}; }));
	var parts = (typeof TOP == "string") ? parseVector(TOP, opts).map(function (x) { return x.v; }) : TOP;
	var idx = 0, len = 0;
	if(parts.length > 0) for(var i = 0; i !== v.length; i += 2) {
		len = +(v[i+1].v);
		switch(v[i].v) {
			case "Worksheets":
			case "工作表":
			case "Листы":
			case "أوراق العمل":
			case "ワークシート":
			case "גליונות עבודה":
			case "Arbeitsblätter":
			case "Çalışma Sayfaları":
			case "Feuilles de calcul":
			case "Fogli di lavoro":
			case "Folhas de cálculo":
			case "Planilhas":
			case "Regneark":
			case "Hojas de cálculo":
			case "Werkbladen":
				props.Worksheets = len;
				props.SheetNames = parts.slice(idx, idx + len);
				break;

			case "Named Ranges":
			case "Rangos con nombre":
			case "名前付き一覧":
			case "Benannte Bereiche":
			case "Navngivne områder":
				props.NamedRanges = len;
				props.DefinedNames = parts.slice(idx, idx + len);
				break;

			case "Charts":
			case "Diagramme":
				props.Chartsheets = len;
				props.ChartNames = parts.slice(idx, idx + len);
				break;
		}
		idx += len;
	}
}

function parse_ext_props(data, p, opts) {
	var q = {}; if(!p) p = {};
	data = utf8read(data);

	EXT_PROPS.forEach(function(f) {
		var xml = (data.match(matchtag(f[0]))||[])[1];
		switch(f[2]) {
			case "string": if(xml) p[f[1]] = unescapexml(xml); break;
			case "bool": p[f[1]] = xml === "true"; break;
			case "raw":
				var cur = data.match(new RegExp("<" + f[0] + "[^>]*>([\\s\\S]*?)<\/" + f[0] + ">"));
				if(cur && cur.length > 0) q[f[1]] = cur[1];
				break;
		}
	});

	if(q.HeadingPairs && q.TitlesOfParts) load_props_pairs(q.HeadingPairs, q.TitlesOfParts, p, opts);

	return p;
}

function write_ext_props(cp/*::, opts*/)/*:string*/ {
	var o/*:Array<string>*/ = [], W = writextag;
	if(!cp) cp = {};
	cp.Application = "SheetJS";
	o[o.length] = (XML_HEADER);
	o[o.length] = (writextag('Properties', null, {
		'xmlns': XMLNS.EXT_PROPS,
		'xmlns:vt': XMLNS.vt
	}));

	EXT_PROPS.forEach(function(f) {
		if(cp[f[1]] === undefined) return;
		var v;
		switch(f[2]) {
			case 'string': v = escapexml(String(cp[f[1]])); break;
			case 'bool': v = cp[f[1]] ? 'true' : 'false'; break;
		}
		if(v !== undefined) o[o.length] = (W(f[0], v));
	});

	/* TODO: HeadingPairs, TitlesOfParts */
	o[o.length] = (W('HeadingPairs', W('vt:vector', W('vt:variant', '<vt:lpstr>Worksheets</vt:lpstr>')+W('vt:variant', W('vt:i4', String(cp.Worksheets))), {size:2, baseType:"variant"})));
	o[o.length] = (W('TitlesOfParts', W('vt:vector', cp.SheetNames.map(function(s) { return "<vt:lpstr>" + escapexml(s) + "</vt:lpstr>"; }).join(""), {size: cp.Worksheets, baseType:"lpstr"})));
	if(o.length>2){ o[o.length] = ('</Properties>'); o[1]=o[1].replace("/>",">"); }
	return o.join("");
}
