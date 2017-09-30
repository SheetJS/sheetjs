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

XMLNS.EXT_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
RELS.EXT_PROPS  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties';

function parse_ext_props(data, p, opts) {
	var q = {}; if(!p) p = {};
	data = utf8read(data);

	EXT_PROPS.forEach(function(f) {
		switch(f[2]) {
			case "string": p[f[1]] = (data.match(matchtag(f[0]))||[])[1]; break;
			case "bool": p[f[1]] = (data.match(matchtag(f[0]))||[])[1] === "true"; break;
			case "raw":
				var cur = data.match(new RegExp("<" + f[0] + "[^>]*>([\\s\\S]*?)<\/" + f[0] + ">"));
				if(cur && cur.length > 0) q[f[1]] = cur[1];
				break;
		}
	});

	if(q.HeadingPairs && q.TitlesOfParts) {
		var v = parseVector(q.HeadingPairs, opts);
		var parts = parseVector(q.TitlesOfParts, opts).map(function (x) { return x.v; });
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
				case "Werkbladen":
					p.Worksheets = len;
					p.SheetNames = parts.slice(idx, idx + len);
					break;

				case "Named Ranges":
				case "名前付き一覧":
				case "Benannte Bereiche":
				case "Navngivne områder":
					p.NamedRanges = len;
					p.DefinedNames = parts.slice(idx, idx + len);
					break;

				case "Charts":
				case "Diagramme":
					p.Chartsheets = len;
					p.ChartNames = parts.slice(idx, idx + len);
					break;
			}
			idx += len;
		}
	}

	return p;
}

var EXT_PROPS_XML_ROOT = writextag('Properties', null, {
	'xmlns': XMLNS.EXT_PROPS,
	'xmlns:vt': XMLNS.vt
});

function write_ext_props(cp, opts)/*:string*/ {
	var o = [], p = {}, W = writextag;
	if(!cp) cp = {};
	cp.Application = "SheetJS";
	o[o.length] = (XML_HEADER);
	o[o.length] = (EXT_PROPS_XML_ROOT);

	EXT_PROPS.forEach(function(f) {
		if(cp[f[1]] === undefined) return;
		var v;
		switch(f[2]) {
			case 'string': v = String(cp[f[1]]); break;
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
