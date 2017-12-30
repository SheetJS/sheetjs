/* 18.8.5 borders CT_Borders */
function parse_borders(t, styles, themes, opts) {
	styles.Borders = [];
	var border = {}, sub_border = {};
	t[0].match(tagregex).forEach(function(x) {
		var y = parsexmltag(x);
		switch (y[0]) {
			case '<borders': case '<borders>': case '</borders>': break;

			/* 18.8.4 border CT_Border */
			case '<border': case '<border>': case '<border/>':
				border = {};
				if (y.diagonalUp) { border.diagonalUp = y.diagonalUp; }
				if (y.diagonalDown) { border.diagonalDown = y.diagonalDown; }
				styles.Borders.push(border);
				break;
			case '</border>': break;

			/* note: not in spec, appears to be CT_BorderPr */
			case '<left/>': break;
			case '<left': case '<left>': break;
			case '</left>': break;

			/* note: not in spec, appears to be CT_BorderPr */
			case '<right/>': break;
			case '<right': case '<right>': break;
			case '</right>': break;

			/* 18.8.43 top CT_BorderPr */
			case '<top/>': break;
			case '<top': case '<top>': break;
			case '</top>': break;

			/* 18.8.6 bottom CT_BorderPr */
			case '<bottom/>': break;
			case '<bottom': case '<bottom>': break;
			case '</bottom>': break;

			/* 18.8.13 diagonal CT_BorderPr */
			case '<diagonal': case '<diagonal>': case '<diagonal/>': break;
			case '</diagonal>': break;

			/* 18.8.25 horizontal CT_BorderPr */
			case '<horizontal': case '<horizontal>': case '<horizontal/>': break;
			case '</horizontal>': break;

			/* 18.8.44 vertical CT_BorderPr */
			case '<vertical': case '<vertical>': case '<vertical/>': break;
			case '</vertical>': break;

			/* 18.8.37 start CT_BorderPr */
			case '<start': case '<start>': case '<start/>': break;
			case '</start>': break;

			/* 18.8.16 end CT_BorderPr */
			case '<end': case '<end>': case '<end/>': break;
			case '</end>': break;

			/* 18.8.? color CT_Color */
			case '<color': case '<color>': break;
			case '<color/>': case '</color>': break;

			default: if(opts && opts.WTF) throw new Error('unrecognized ' + y[0] + ' in borders');
		}
	});
}

/* 18.8.21 fills CT_Fills */
function parse_fills(t, styles, themes, opts) {
	styles.Fills = [];
	var fill = {};
	t[0].match(tagregex).forEach(function(x) {
		var y = parsexmltag(x);
		switch(y[0]) {
			case '<fills': case '<fills>': case '</fills>': break;

			/* 18.8.20 fill CT_Fill */
			case '<fill>': case '<fill': case '<fill/>':
				fill = {}; styles.Fills.push(fill); break;
			case '</fill>': break;

			/* 18.8.24 gradientFill CT_GradientFill */
			case '<gradientFill>': break;
			case '<gradientFill':
			case '</gradientFill>': styles.Fills.push(fill); fill = {}; break;

			/* 18.8.32 patternFill CT_PatternFill */
			case '<patternFill': case '<patternFill>':
				if(y.patternType) fill.patternType = y.patternType;
				break;
			case '<patternFill/>': case '</patternFill>': break;

			/* 18.8.3 bgColor CT_Color */
			case '<bgColor':
				if(!fill.bgColor) fill.bgColor = {};
				if(y.indexed) fill.bgColor.indexed = parseInt(y.indexed, 10);
				if(y.theme) fill.bgColor.theme = parseInt(y.theme, 10);
				if(y.tint) fill.bgColor.tint = parseFloat(y.tint);
				/* Excel uses ARGB strings */
				if(y.rgb) fill.bgColor.rgb = y.rgb.slice(-6);
				break;
			case '<bgColor/>': case '</bgColor>': break;

			/* 18.8.19 fgColor CT_Color */
			case '<fgColor':
				if(!fill.fgColor) fill.fgColor = {};
				if(y.theme) fill.fgColor.theme = parseInt(y.theme, 10);
				if(y.tint) fill.fgColor.tint = parseFloat(y.tint);
				/* Excel uses ARGB strings */
				if(y.rgb) fill.fgColor.rgb = y.rgb.slice(-6);
				break;
			case '<fgColor/>': case '</fgColor>': break;

			/* 18.8.38 stop CT_GradientStop */
			case '<stop': case '<stop/>': break;
			case '</stop>': break;

			/* 18.8.? color CT_Color */
			case '<color': case '<color/>': break;
			case '</color>': break;

			default: if(opts && opts.WTF) throw new Error('unrecognized ' + y[0] + ' in fills');
		}
	});
}

/* 18.8.23 fonts CT_Fonts */
function parse_fonts(t, styles, themes, opts) {
	styles.Fonts = [];
	var font = {};
	t[0].match(tagregex).forEach(function(x) {
		var y = parsexmltag(x);
		switch (y[0]) {
			case '<fonts': case '<fonts>': case '</fonts>': break;

			/* 18.8.22 font CT_Font */
			case '<font': case '<font>': break;
			case '</font>': case '<font/>':
				styles.Fonts.push(font);
				font = {};
				break;

			/* 18.8.29 name CT_FontName */
			case '<name': if(y.val) font.name = y.val; break;
			case '<name/>': case '</name>': break;

			/* 18.8.2  b CT_BooleanProperty */
			case '<b': font.bold = y.val ? parsexmlbool(y.val) : 1; break;
			case '<b/>': font.bold = 1; break;

			/* 18.8.26 i CT_BooleanProperty */
			case '<i': font.italic = y.val ? parsexmlbool(y.val) : 1; break;
			case '<i/>': font.italic = 1; break;

			/* 18.4.13 u CT_UnderlineProperty */
			case '<u':
				switch(y.val) {
					case "none": font.underline = 0x00; break;
					case "single": font.underline = 0x01; break;
					case "double": font.underline = 0x02; break;
					case "singleAccounting": font.underline = 0x21; break;
					case "doubleAccounting": font.underline = 0x22; break;
				} break;
			case '<u/>': font.underline = 1; break;

			/* 18.4.10 strike CT_BooleanProperty */
			case '<strike': font.strike = y.val ? parsexmlbool(y.val) : 1; break;
			case '<strike/>': font.strike = 1; break;

			/* 18.4.2  outline CT_BooleanProperty */
			case '<outline': font.outline = y.val ? parsexmlbool(y.val) : 1; break;
			case '<outline/>': font.outline = 1; break;

			/* 18.8.36 shadow CT_BooleanProperty */
			case '<shadow': font.shadow = y.val ? parsexmlbool(y.val) : 1; break;
			case '<shadow/>': font.shadow = 1; break;

			/* 18.8.12 condense CT_BooleanProperty */
			case '<condense': font.condense = y.val ? parsexmlbool(y.val) : 1; break;
			case '<condense/>': font.condense = 1; break;

			/* 18.8.17 extend CT_BooleanProperty */
			case '<extend': font.extend = y.val ? parsexmlbool(y.val) : 1; break;
			case '<extend/>': font.extend = 1; break;

			/* 18.4.11 sz CT_FontSize */
			case '<sz': if(y.val) font.sz = +y.val; break;
			case '<sz/>': case '</sz>': break;

			/* 18.4.14 vertAlign CT_VerticalAlignFontProperty */
			case '<vertAlign': if(y.val) font.vertAlign = y.val; break;
			case '<vertAlign/>': case '</vertAlign>': break;

			/* 18.8.18 family CT_FontFamily */
			case '<family': if(y.val) font.family = parseInt(y.val,10); break;
			case '<family/>': case '</family>': break;

			/* 18.8.35 scheme CT_FontScheme */
			case '<scheme': if(y.val) font.scheme = y.val; break;
			case '<scheme/>': case '</scheme>': break;

			/* 18.4.1 charset CT_IntProperty */
			case '<charset':
				if(y.val == '1') break;
				y.codepage = CS2CP[parseInt(y.val, 10)];
				break;

			/* 18.?.? color CT_Color */
			case '<color':
				if(!font.color) font.color = {};
				if(y.auto) font.color.auto = parsexmlbool(y.auto);

				if(y.rgb) font.color.rgb = y.rgb.slice(-6);
				else if(y.indexed) {
					font.color.index = parseInt(y.indexed, 10);
					var icv = XLSIcv[font.color.index];
					if(font.color.index == 81) icv = XLSIcv[1];
					if(!icv) throw new Error(x);
					font.color.rgb = icv[0].toString(16) + icv[1].toString(16) + icv[2].toString(16);
				} else if(y.theme) {
					font.color.theme = parseInt(y.theme, 10);
					if(y.tint) font.color.tint = parseFloat(y.tint);
					if(y.theme && themes.themeElements && themes.themeElements.clrScheme) {
						font.color.rgb = rgb_tint(themes.themeElements.clrScheme[font.color.theme].rgb, font.color.tint || 0);
					}
				}

				break;
			case '<color/>': case '</color>': break;

			default: if(opts && opts.WTF) throw new Error('unrecognized ' + y[0] + ' in fonts');
		}
	});
}

/* 18.8.31 numFmts CT_NumFmts */
function parse_numFmts(t, styles, opts) {
	styles.NumberFmt = [];
	var k/*Array<number>*/ = (keys(SSF._table)/*:any*/);
	for(var i=0; i < k.length; ++i) styles.NumberFmt[k[i]] = SSF._table[k[i]];
	var m = t[0].match(tagregex);
	if(!m) return;
	for(i=0; i < m.length; ++i) {
		var y = parsexmltag(m[i]);
		switch(y[0]) {
			case '<numFmts': case '</numFmts>': case '<numFmts/>': case '<numFmts>': break;
			case '<numFmt': {
				var f=unescapexml(utf8read(y.formatCode)), j=parseInt(y.numFmtId,10);
				styles.NumberFmt[j] = f;
				if(j>0) {
					if(j > 0x188) {
						for(j = 0x188; j > 0x3c; --j) if(styles.NumberFmt[j] == null) break;
						styles.NumberFmt[j] = f;
					}
					SSF.load(f,j);
				}
			} break;
			case '</numFmt>': break;
			default: if(opts.WTF) throw new Error('unrecognized ' + y[0] + ' in numFmts');
		}
	}
}

function write_numFmts(NF/*:{[n:number|string]:string}*/, opts) {
	var o = ["<numFmts>"];
	[[5,8],[23,26],[41,44],[/*63*/50,/*66],[164,*/392]].forEach(function(r) {
		for(var i = r[0]; i <= r[1]; ++i) if(NF[i] != null) o[o.length] = (writextag('numFmt',null,{numFmtId:i,formatCode:escapexml(NF[i])}));
	});
	if(o.length === 1) return "";
	o[o.length] = ("</numFmts>");
	o[0] = writextag('numFmts', null, { count:o.length-2 }).replace("/>", ">");
	return o.join("");
}

/* 18.8.10 cellXfs CT_CellXfs */
var cellXF_uint = [ "numFmtId", "fillId", "fontId", "borderId", "xfId" ];
var cellXF_bool = [ "applyAlignment", "applyBorder", "applyFill", "applyFont", "applyNumberFormat", "applyProtection", "pivotButton", "quotePrefix" ];
function parse_cellXfs(t, styles, opts) {
	styles.CellXf = [];
	var xf;
	t[0].match(tagregex).forEach(function(x) {
		var y = parsexmltag(x), i = 0;
		switch(y[0]) {
			case '<cellXfs': case '<cellXfs>': case '<cellXfs/>': case '</cellXfs>': break;

			/* 18.8.45 xf CT_Xf */
			case '<xf': case '<xf/>':
				xf = y;
				delete xf[0];
				for(i = 0; i < cellXF_uint.length; ++i) if(xf[cellXF_uint[i]])
					xf[cellXF_uint[i]] = parseInt(xf[cellXF_uint[i]], 10);
				for(i = 0; i < cellXF_bool.length; ++i) if(xf[cellXF_bool[i]])
					xf[cellXF_bool[i]] = parsexmlbool(xf[cellXF_bool[i]], "");
				if(xf.numFmtId > 0x188) {
					for(i = 0x188; i > 0x3c; --i) if(styles.NumberFmt[xf.numFmtId] == styles.NumberFmt[i]) { xf.numFmtId = i; break; }
				}
				styles.CellXf.push(xf); break;
			case '</xf>': break;

			/* 18.8.1 alignment CT_CellAlignment */
			case '<alignment': case '<alignment/>':
				var alignment = {};
				if(y.vertical) alignment.vertical = y.vertical;
				if(y.horizontal) alignment.horizontal = y.horizontal;
				if(y.textRotation != null) alignment.textRotation = y.textRotation;
				if(y.indent) alignment.indent = y.indent;
				if(y.wrapText) alignment.wrapText = y.wrapText;
				xf.alignment = alignment;
				break;
			case '</alignment>': break;

			/* 18.8.33 protection CT_CellProtection */
			case '<protection': case '</protection>': case '<protection/>': break;

			/* 18.2.10 extLst CT_ExtensionList ? */
			case '<extLst': case '</extLst>': break;
			case '<ext': break;
			default: if(opts.WTF) throw new Error('unrecognized ' + y[0] + ' in cellXfs');
		}
	});
}

function write_cellXfs(cellXfs)/*:string*/ {
	var o/*:Array<string>*/ = [];
	o[o.length] = (writextag('cellXfs',null));
	cellXfs.forEach(function(c) { o[o.length] = (writextag('xf', null, c)); });
	o[o.length] = ("</cellXfs>");
	if(o.length === 2) return "";
	o[0] = writextag('cellXfs',null, {count:o.length-2}).replace("/>",">");
	return o.join("");
}

/* 18.8 Styles CT_Stylesheet*/
var parse_sty_xml= (function make_pstyx() {
var numFmtRegex = /<numFmts([^>]*)>[\S\s]*?<\/numFmts>/;
var cellXfRegex = /<cellXfs([^>]*)>[\S\s]*?<\/cellXfs>/;
var fillsRegex = /<fills([^>]*)>[\S\s]*?<\/fills>/;
var fontsRegex = /<fonts([^>]*)>[\S\s]*?<\/fonts>/;
var bordersRegex = /<borders([^>]*)>[\S\s]*?<\/borders>/;

return function parse_sty_xml(data, themes, opts) {
	var styles = {};
	if(!data) return styles;
	data = data.replace(/<!--([\s\S]*?)-->/mg,"").replace(/<!DOCTYPE[^\[]*\[[^\]]*\]>/gm,"");
	/* 18.8.39 styleSheet CT_Stylesheet */
	var t;

	/* 18.8.31 numFmts CT_NumFmts ? */
	if((t=data.match(numFmtRegex))) parse_numFmts(t, styles, opts);

	/* 18.8.23 fonts CT_Fonts ? */
	if((t=data.match(fontsRegex))) parse_fonts(t, styles, themes, opts);

	/* 18.8.21 fills CT_Fills ? */
	if((t=data.match(fillsRegex))) parse_fills(t, styles, themes, opts);

	/* 18.8.5  borders CT_Borders ? */
	if((t=data.match(bordersRegex))) parse_borders(t, styles, themes, opts);

	/* 18.8.9  cellStyleXfs CT_CellStyleXfs ? */

	/* 18.8.10 cellXfs CT_CellXfs ? */
	if((t=data.match(cellXfRegex))) parse_cellXfs(t, styles, opts);

	/* 18.8.8  cellStyles CT_CellStyles ? */
	/* 18.8.15 dxfs CT_Dxfs ? */
	/* 18.8.42 tableStyles CT_TableStyles ? */
	/* 18.8.11 colors CT_Colors ? */
	/* 18.2.10 extLst CT_ExtensionList ? */

	return styles;
};
})();

var STYLES_XML_ROOT = writextag('styleSheet', null, {
	'xmlns': XMLNS.main[0],
	'xmlns:vt': XMLNS.vt
});

RELS.STY = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

function write_sty_xml(wb/*:Workbook*/, opts)/*:string*/ {
	var o = [XML_HEADER, STYLES_XML_ROOT], w;
	if(wb.SSF && (w = write_numFmts(wb.SSF)) != null) o[o.length] = w;
	o[o.length] = ('<fonts count="1"><font><sz val="12"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>');
	o[o.length] = ('<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>');
	o[o.length] = ('<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>');
	o[o.length] = ('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>');
	if((w = write_cellXfs(opts.cellXfs))) o[o.length] = (w);
	o[o.length] = ('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>');
	o[o.length] = ('<dxfs count="0"/>');
	o[o.length] = ('<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleMedium4"/>');

	if(o.length>2){ o[o.length] = ('</styleSheet>'); o[1]=o[1].replace("/>",">"); }
	return o.join("");
}
