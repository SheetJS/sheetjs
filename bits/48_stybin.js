/* [MS-XLSB] 2.4.651 BrtFmt */
function parse_BrtFmt(data, length/*:number*/) {
	var numFmtId = data.read_shift(2);
	var stFmtCode = parse_XLWideString(data,length-2);
	return [numFmtId, stFmtCode];
}
function write_BrtFmt(i/*:number*/, f/*:string*/, o) {
	if(!o) o = new_buf(6 + 4 * f.length);
	o.write_shift(2, i);
	write_XLWideString(f, o);
	var out = (o.length > o.l) ? o.slice(0, o.l) : o;
	if(o.l == null) o.l = o.length;
	return out;
}

/* [MS-XLSB] 2.4.653 BrtFont TODO */
function parse_BrtFont(data, length/*:number*/, opts) {
	var out = ({}/*:any*/);

	out.sz = data.read_shift(2) / 20;

	var grbit = parse_FontFlags(data, 2, opts);
	if(grbit.fCondense) out.condense = 1;
	if(grbit.fExtend) out.extend = 1;
	if(grbit.fShadow) out.shadow = 1;
	if(grbit.fOutline) out.outline = 1;
	if(grbit.fStrikeout) out.strike = 1;
	if(grbit.fItalic) out.italic = 1;

	var bls = data.read_shift(2);
	if(bls === 0x02BC) out.bold = 1;

	switch(data.read_shift(2)) {
		/* case 0: out.vertAlign = "baseline"; break; */
		case 1: out.vertAlign = "superscript"; break;
		case 2: out.vertAlign = "subscript"; break;
	}

	var underline = data.read_shift(1);
	if(underline != 0) out.underline = underline;

	var family = data.read_shift(1);
	if(family > 0) out.family = family;

	var bCharSet = data.read_shift(1);
	if(bCharSet > 0) out.charset = bCharSet;

	data.l++;
	out.color = parse_BrtColor(data, 8);

	switch(data.read_shift(1)) {
		/* case 0: out.scheme = "none": break; */
		case 1: out.scheme = "major"; break;
		case 2: out.scheme = "minor"; break;
	}

	out.name = parse_XLWideString(data, length - 21);

	return out;
}
function write_BrtFont(font, o) {
	if(!o) o = new_buf(25+4*32);
	o.write_shift(2, font.sz * 20);
	write_FontFlags(font, o);
	o.write_shift(2, font.bold ? 0x02BC : 0x0190);
	var sss = 0;
	if(font.vertAlign == "superscript") sss = 1;
	else if(font.vertAlign == "subscript") sss = 2;
	o.write_shift(2, sss);
	o.write_shift(1, font.underline || 0);
	o.write_shift(1, font.family || 0);
	o.write_shift(1, font.charset || 0);
	o.write_shift(1, 0);
	write_BrtColor(font.color, o);
	var scheme = 0;
	if(font.scheme == "major") scheme = 1;
	if(font.scheme == "minor") scheme = 2;
	o.write_shift(1, scheme);
	write_XLWideString(font.name, o);
	return o.length > o.l ? o.slice(0, o.l) : o;
}

/* [MS-XLSB] 2.4.644 BrtFill */
var XLSBFillPTNames = [
	"none",
	"solid",
	"mediumGray",
	"darkGray",
	"lightGray",
	"darkHorizontal",
	"darkVertical",
	"darkDown",
	"darkUp",
	"darkGrid",
	"darkTrellis",
	"lightHorizontal",
	"lightVertical",
	"lightDown",
	"lightUp",
	"lightGrid",
	"lightTrellis",
	"gray125",
	"gray0625"
];
var rev_XLSBFillPTNames/*:EvertNumType*/ = (evert(XLSBFillPTNames)/*:any*/);
/* TODO: gradient fill representation */
var parse_BrtFill = parsenoop;
function write_BrtFill(fill, o) {
	if(!o) o = new_buf(4*3 + 8*7 + 16*1);
	var fls/*:number*/ = rev_XLSBFillPTNames[fill.patternType];
	if(fls == null) fls = 0x28;
	o.write_shift(4, fls);
	var j = 0;
	if(fls != 0x28) {
		/* TODO: custom FG Color */
		write_BrtColor({auto:1}, o);
		/* TODO: custom BG Color */
		write_BrtColor({auto:1}, o);

		for(; j < 12; ++j) o.write_shift(4, 0);
	} else {
		for(; j < 4; ++j) o.write_shift(4, 0);

		for(; j < 12; ++j) o.write_shift(4, 0); /* TODO */
		/* iGradientType */
		/* xnumDegree */
		/* xnumFillToLeft */
		/* xnumFillToRight */
		/* xnumFillToTop */
		/* xnumFillToBottom */
		/* cNumStop */
		/* xfillGradientStop */
	}
	return o.length > o.l ? o.slice(0, o.l) : o;
}

/* [MS-XLSB] 2.4.816 BrtXF */
function parse_BrtXF(data, length/*:number*/) {
	var tgt = data.l + length;
	var ixfeParent = data.read_shift(2);
	var ifmt = data.read_shift(2);
	data.l = tgt;
	return {ixfe:ixfeParent, numFmtId:ifmt };
}
function write_BrtXF(data, ixfeP, o) {
	if(!o) o = new_buf(16);
	o.write_shift(2, ixfeP||0);
	o.write_shift(2, data.numFmtId||0);
	o.write_shift(2, 0); /* iFont */
	o.write_shift(2, 0); /* iFill */
	o.write_shift(2, 0); /* ixBorder */
	o.write_shift(1, 0); /* trot */
	o.write_shift(1, 0); /* indent */
	o.write_shift(1, 0); /* flags */
	o.write_shift(1, 0); /* flags */
	o.write_shift(1, 0); /* xfGrbitAtr */
	o.write_shift(1, 0);
	return o;
}

/* [MS-XLSB] 2.5.4 Blxf TODO */
function write_Blxf(data, o) {
	if(!o) o = new_buf(10);
	o.write_shift(1, 0); /* dg */
	o.write_shift(1, 0);
	o.write_shift(4, 0); /* color */
	o.write_shift(4, 0); /* color */
	return o;
}
/* [MS-XLSB] 2.4.299 BrtBorder TODO */
var parse_BrtBorder = parsenoop;
function write_BrtBorder(border, o) {
	if(!o) o = new_buf(51);
	o.write_shift(1, 0); /* diagonal */
	write_Blxf(null, o); /* top */
	write_Blxf(null, o); /* bottom */
	write_Blxf(null, o); /* left */
	write_Blxf(null, o); /* right */
	write_Blxf(null, o); /* diag */
	return o.length > o.l ? o.slice(0, o.l) : o;
}

/* [MS-XLSB] 2.4.755 BrtStyle TODO */
function write_BrtStyle(style, o) {
	if(!o) o = new_buf(12+4*10);
	o.write_shift(4, style.xfId);
	o.write_shift(2, 1);
	o.write_shift(1, +style.builtinId);
	o.write_shift(1, 0); /* iLevel */
	write_XLNullableWideString(style.name || "", o);
	return o.length > o.l ? o.slice(0, o.l) : o;
}

/* [MS-XLSB] 2.4.269 BrtBeginTableStyles */
function write_BrtBeginTableStyles(cnt, defTableStyle, defPivotStyle) {
	var o = new_buf(4+256*2*4);
	o.write_shift(4, cnt);
	write_XLNullableWideString(defTableStyle, o);
	write_XLNullableWideString(defPivotStyle, o);
	return o.length > o.l ? o.slice(0, o.l) : o;
}

/* [MS-XLSB] 2.1.7.50 Styles */
function parse_sty_bin(data, themes, opts) {
	var styles = {};
	styles.NumberFmt = ([]/*:any*/);
	for(var y in SSF._table) styles.NumberFmt[y] = SSF._table[y];

	styles.CellXf = [];
	styles.Fonts = [];
	var state/*:Array<string>*/ = [];
	var pass = false;
	recordhopper(data, function hopper_sty(val, R_n, RT) {
		switch(RT) {
			case 0x002C: /* 'BrtFmt' */
				styles.NumberFmt[val[0]] = val[1]; SSF.load(val[1], val[0]);
				break;
			case 0x002B: /* 'BrtFont' */
				styles.Fonts.push(val);
				if(val.color.theme != null && themes && themes.themeElements && themes.themeElements.clrScheme) {
					val.color.rgb = rgb_tint(themes.themeElements.clrScheme[val.color.theme].rgb, val.color.tint || 0);
				}
				break;
			case 0x0401: /* 'BrtKnownFonts' */ break;
			case 0x002D: /* 'BrtFill' */ break;
			case 0x002E: /* 'BrtBorder' */ break;
			case 0x002F: /* 'BrtXF' */
				if(state[state.length - 1] == "BrtBeginCellXFs") {
					styles.CellXf.push(val);
				}
				break;
			case 0x0030: /* 'BrtStyle' */
			case 0x01FB: /* 'BrtDXF' */
			case 0x023C: /* 'BrtMRUColor' */
			case 0x01DB: /* 'BrtIndexedColor': */
				break;

			case 0x0493: /* 'BrtDXF14' */
			case 0x0836: /* 'BrtDXF15' */
			case 0x046A: /* 'BrtSlicerStyleElement' */
			case 0x0200: /* 'BrtTableStyleElement' */
			case 0x082F: /* 'BrtTimelineStyleElement' */
			/* case 'BrtUid' */
				break;

			case 0x0023: /* 'BrtFRTBegin' */
				pass = true; break;
			case 0x0024: /* 'BrtFRTEnd' */
				pass = false; break;
			case 0x0025: /* 'BrtACBegin' */
				state.push(R_n); break;
			case 0x0026: /* 'BrtACEnd' */
				state.pop(); break;

			default:
				if((R_n||"").indexOf("Begin") > 0) state.push(R_n);
				else if((R_n||"").indexOf("End") > 0) state.pop();
				else if(!pass || opts.WTF) throw new Error("Unexpected record " + RT + " " + R_n);
		}
	});
	return styles;
}

function write_FMTS_bin(ba, NF/*:?SSFTable*/) {
	if(!NF) return;
	var cnt = 0;
	[[5,8],[23,26],[41,44],[/*63*/50,/*66],[164,*/392]].forEach(function(r) {
		/*:: if(!NF) return; */
		for(var i = r[0]; i <= r[1]; ++i) if(NF[i] != null) ++cnt;
	});

	if(cnt == 0) return;
	write_record(ba, "BrtBeginFmts", write_UInt32LE(cnt));
	[[5,8],[23,26],[41,44],[/*63*/50,/*66],[164,*/392]].forEach(function(r) {
		/*:: if(!NF) return; */
		for(var i = r[0]; i <= r[1]; ++i) if(NF[i] != null) write_record(ba, "BrtFmt", write_BrtFmt(i, NF[i]));
	});
	write_record(ba, "BrtEndFmts");
}

function write_FONTS_bin(ba/*::, data*/) {
	var cnt = 1;

	if(cnt == 0) return;
	write_record(ba, "BrtBeginFonts", write_UInt32LE(cnt));
	write_record(ba, "BrtFont", write_BrtFont({
		sz:12,
		color: {theme:1},
		name: "Calibri",
		family: 2,
		scheme: "minor"
	}));
	/* 1*65491BrtFont [ACFONTS] */
	write_record(ba, "BrtEndFonts");
}

function write_FILLS_bin(ba/*::, data*/) {
	var cnt = 2;

	if(cnt == 0) return;
	write_record(ba, "BrtBeginFills", write_UInt32LE(cnt));
	write_record(ba, "BrtFill", write_BrtFill({patternType:"none"}));
	write_record(ba, "BrtFill", write_BrtFill({patternType:"gray125"}));
	/* 1*65431BrtFill */
	write_record(ba, "BrtEndFills");
}

function write_BORDERS_bin(ba/*::, data*/) {
	var cnt = 1;

	if(cnt == 0) return;
	write_record(ba, "BrtBeginBorders", write_UInt32LE(cnt));
	write_record(ba, "BrtBorder", write_BrtBorder({}));
	/* 1*65430BrtBorder */
	write_record(ba, "BrtEndBorders");
}

function write_CELLSTYLEXFS_bin(ba/*::, data*/) {
	var cnt = 1;
	write_record(ba, "BrtBeginCellStyleXFs", write_UInt32LE(cnt));
	write_record(ba, "BrtXF", write_BrtXF({
		numFmtId:0,
		fontId:0,
		fillId:0,
		borderId:0
	}, 0xFFFF));
	/* 1*65430(BrtXF *FRT) */
	write_record(ba, "BrtEndCellStyleXFs");
}

function write_CELLXFS_bin(ba, data) {
	write_record(ba, "BrtBeginCellXFs", write_UInt32LE(data.length));
	data.forEach(function(c) { write_record(ba, "BrtXF", write_BrtXF(c,0)); });
	/* 1*65430(BrtXF *FRT) */
	write_record(ba, "BrtEndCellXFs");
}

function write_STYLES_bin(ba/*::, data*/) {
	var cnt = 1;

	write_record(ba, "BrtBeginStyles", write_UInt32LE(cnt));
	write_record(ba, "BrtStyle", write_BrtStyle({
		xfId:0,
		builtinId:0,
		name:"Normal"
	}));
	/* 1*65430(BrtStyle *FRT) */
	write_record(ba, "BrtEndStyles");
}

function write_DXFS_bin(ba/*::, data*/) {
	var cnt = 0;

	write_record(ba, "BrtBeginDXFs", write_UInt32LE(cnt));
	/* *2147483647(BrtDXF *FRT) */
	write_record(ba, "BrtEndDXFs");
}

function write_TABLESTYLES_bin(ba/*::, data*/) {
	var cnt = 0;

	write_record(ba, "BrtBeginTableStyles", write_BrtBeginTableStyles(cnt, "TableStyleMedium9", "PivotStyleMedium4"));
	/* *TABLESTYLE */
	write_record(ba, "BrtEndTableStyles");
}

function write_COLORPALETTE_bin(/*::ba, data*/) {
	return;
	/* BrtBeginColorPalette [INDEXEDCOLORS] [MRUCOLORS] BrtEndColorPalette */
}

/* [MS-XLSB] 2.1.7.50 Styles */
function write_sty_bin(wb, opts) {
	var ba = buf_array();
	write_record(ba, "BrtBeginStyleSheet");
	write_FMTS_bin(ba, wb.SSF);
	write_FONTS_bin(ba, wb);
	write_FILLS_bin(ba, wb);
	write_BORDERS_bin(ba, wb);
	write_CELLSTYLEXFS_bin(ba, wb);
	write_CELLXFS_bin(ba, opts.cellXfs);
	write_STYLES_bin(ba, wb);
	write_DXFS_bin(ba, wb);
	write_TABLESTYLES_bin(ba, wb);
	write_COLORPALETTE_bin(ba, wb);
	/* FRTSTYLESHEET*/
	write_record(ba, "BrtEndStyleSheet");
	return ba.end();
}
