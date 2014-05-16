/* 18.8.31 numFmts CT_NumFmts */
function parse_numFmts(t, opts) {
	styles.NumberFmt = [];
	for(var y in SSF._table) styles.NumberFmt[y] = SSF._table[y];
	t[0].match(/<[^>]*>/g).forEach(function(x) {
		var y = parsexmltag(x);
		switch(y[0]) {
			case '<numFmts': case '</numFmts>': case '<numFmts/>': break;
			case '<numFmt': {
				var f=utf8read(unescapexml(y.formatCode)), i=parseInt(y.numFmtId,10);
				styles.NumberFmt[i] = f; if(i>0) SSF.load(f,i);
			} break;
			default: if(opts.WTF) throw 'unrecognized ' + y[0] + ' in numFmts';
		}
	});
}

function write_numFmts(NF, opts) {
	var o = [];
	o.push("<numFmts>");
	[[5,8],[23,26],[41,44],[63,66],[164,392]].forEach(function(r) {
		for(var i = r[0]; i <= r[1]; ++i) if(NF[i]) 
		o.push(writextag('numFmt',null,{numFmtId:i,formatCode:escapexml(NF[i])}));
	});
	o.push("</numFmts>");
	if(o.length === 2) return "";
	o[0] = writextag('numFmts', null, { count:o.length-2 }).replace("/>", ">");
	return o.join("");
}

/* 18.8.10 cellXfs CT_CellXfs */
function parse_cellXfs(t, opts) {
	styles.CellXf = [];
	t[0].match(/<[^>]*>/g).forEach(function(x) {
		var y = parsexmltag(x);
		switch(y[0]) {
			case '<cellXfs': case '<cellXfs/>': case '</cellXfs>': break;

			/* 18.8.45 xf CT_Xf */
			case '<xf': delete y[0];
				if(y.numFmtId) y.numFmtId = parseInt(y.numFmtId, 10);
				styles.CellXf.push(y); break;
			case '</xf>': break;

			/* 18.8.1 alignment CT_CellAlignment */
			case '<alignment': case '<alignment/>': break;

			/* 18.8.33 protection CT_CellProtection */
			case '<protection': case '</protection>': case '<protection/>': break;

			case '<extLst': case '</extLst>': break;
			case '<ext': break;
			default: if(opts.WTF) throw 'unrecognized ' + y[0] + ' in cellXfs';
		}
	});
}

function write_cellXfs(cellXfs) {
	var o = [];
	o.push(writextag('cellXfs',null));
	cellXfs.forEach(function(c) { o.push(writextag('xf', null, c)); });
	o.push("</cellXfs>");
	if(o.length === 2) return "";
	o[0] = writextag('cellXfs',null, {count:o.length-2}).replace("/>",">");
	return o.join("");
}

/* 18.8 Styles CT_Stylesheet*/
function parse_sty_xml(data, opts) {
	/* 18.8.39 styleSheet CT_Stylesheet */
	var t;

	/* numFmts CT_NumFmts ? */
	if((t=data.match(/<numFmts([^>]*)>.*<\/numFmts>/))) parse_numFmts(t, opts);

	/* fonts CT_Fonts ? */
	/* fills CT_Fills ? */
	/* borders CT_Borders ? */
	/* cellStyleXfs CT_CellStyleXfs ? */

	/* cellXfs CT_CellXfs ? */
	if((t=data.match(/<cellXfs([^>]*)>.*<\/cellXfs>/))) parse_cellXfs(t, opts);

	/* dxfs CT_Dxfs ? */
	/* tableStyles CT_TableStyles ? */
	/* colors CT_Colors ? */
	/* extLst CT_ExtensionList ? */

	return styles;
}

var STYLES_XML_ROOT = writextag('styleSheet', null, {
	'xmlns': XMLNS.main[0],
	'xmlns:vt': XMLNS.vt
});

RELS.STY = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

function write_sty_xml(wb, opts) {
	var o = [], p = {}, W = writextag, w;
	o.push(XML_HEADER);
	o.push(STYLES_XML_ROOT);
	if((w = write_numFmts(wb.SSF))) o.push(w);
  o.push('<fonts count="1"><font><sz val="12"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>');
  o.push('<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>');
	o.push('<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>');
	o.push('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>');
	if((w = write_cellXfs(opts.cellXfs))) o.push(w);
	o.push('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>');
	o.push('<dxfs count="0"/>');
	o.push('<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleMedium4"/>');

	if(o.length>2){ o.push('</styleSheet>'); o[1]=o[1].replace("/>",">"); }
	return o.join("");
}
