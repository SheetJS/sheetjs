/* [MS-XLSB] 2.4.651 BrtFmt */
function parse_BrtFmt(data, length) {
	var ifmt = data.read_shift(2);
	var stFmtCode = parse_XLWideString(data,length-2);
	return [ifmt, stFmtCode];
}

/* [MS-XLSB] 2.4.653 BrtFont TODO */
function parse_BrtFont(data, length) {
	var out = {flags:{}};
	out.dyHeight = data.read_shift(2);
	out.grbit = parse_FontFlags(data, 2);
	out.bls = data.read_shift(2);
	out.sss = data.read_shift(2);
	out.uls = data.read_shift(1);
	out.bFamily = data.read_shift(1);
	out.bCharSet = data.read_shift(1);
	data.l++;
	out.brtColor = parse_BrtColor(data, 8);
	out.bFontScheme = data.read_shift(1);
	out.name = parse_XLWideString(data, length - 21);

	out.flags.Bold = out.bls === 0x02BC;
	out.flags.Italic = out.grbit.fItalic;
	out.flags.Strikeout = out.grbit.fStrikeout;
	out.flags.Outline = out.grbit.fOutline;
	out.flags.Shadow = out.grbit.fShadow;
	out.flags.Condense = out.grbit.fCondense;
	out.flags.Extend = out.grbit.fExtend;
	out.flags.Sub = out.sss & 0x2;
	out.flags.Sup = out.sss & 0x1;
	return out;
}

/* [MS-XLSB] 2.4.816 BrtXF */
function parse_BrtXF(data, length) {
	var ixfeParent = data.read_shift(2);
	var ifmt = data.read_shift(2);
	parsenoop(data, length-4);
	return {ixfe:ixfeParent, ifmt:ifmt };
}

/* [MS-XLSB] 2.1.7.50 Styles */
function parse_sty_bin(data, opts) {
	styles.NumberFmt = [];
	for(var y in SSF._table) styles.NumberFmt[y] = SSF._table[y];

	styles.CellXf = [];
	var state = ""; /* TODO: this should be a stack */
	var pass = false;
	recordhopper(data, function hopper_sty(val, R, RT) {
		switch(R.n) {
			case 'BrtFmt':
				styles.NumberFmt[val[0]] = val[1]; SSF.load(val[1], val[0]);
				break;
			case 'BrtFont': break; /* TODO */
			case 'BrtKnownFonts': break; /* TODO */
			case 'BrtFill': break; /* TODO */
			case 'BrtBorder': break; /* TODO */
			case 'BrtXF':
				if(state === "CELLXFS") {
					styles.CellXf.push(val);
				}
				break; /* TODO */
			case 'BrtStyle': break; /* TODO */
			case 'BrtDXF': break; /* TODO */
			case 'BrtMRUColor': break; /* TODO */
			case 'BrtIndexedColor': break; /* TODO */
			case 'BrtBeginStyleSheet': break;
			case 'BrtEndStyleSheet': break;
			case 'BrtBeginTableStyle': break;
			case 'BrtTableStyleElement': break;
			case 'BrtEndTableStyle': break;
			case 'BrtBeginFmts': state = "FMTS"; break;
			case 'BrtEndFmts': state = ""; break;
			case 'BrtBeginFonts': state = "FONTS"; break;
			case 'BrtEndFonts': state = ""; break;
			case 'BrtACBegin': state = "ACFONTS"; break;
			case 'BrtACEnd': state = ""; break;
			case 'BrtBeginFills': state = "FILLS"; break;
			case 'BrtEndFills': state = ""; break;
			case 'BrtBeginBorders': state = "BORDERS"; break;
			case 'BrtEndBorders': state = ""; break;
			case 'BrtBeginCellStyleXFs': state = "CELLSTYLEXFS"; break;
			case 'BrtEndCellStyleXFs': state = ""; break;
			case 'BrtBeginCellXFs': state = "CELLXFS"; break;
			case 'BrtEndCellXFs': state = ""; break;
			case 'BrtBeginStyles': state = "STYLES"; break;
			case 'BrtEndStyles': state = ""; break;
			case 'BrtBeginDXFs': state = "DXFS"; break;
			case 'BrtEndDXFs': state = ""; break;
			case 'BrtBeginTableStyles': state = "TABLESTYLES"; break;
			case 'BrtEndTableStyles': state = ""; break;
			case 'BrtBeginColorPalette': state = "COLORPALETTE"; break;
			case 'BrtEndColorPalette': state = ""; break;
			case 'BrtBeginIndexedColors': state = "INDEXEDCOLORS"; break;
			case 'BrtEndIndexedColors': state = ""; break;
			case 'BrtBeginMRUColors': state = "MRUCOLORS"; break;
			case 'BrtEndMRUColors': state = ""; break;
			case 'BrtFRTBegin': pass = true; break;
			case 'BrtFRTEnd': pass = false; break;
			case 'BrtBeginStyleSheetExt14': break;
			case 'BrtBeginSlicerStyles': break;
			case 'BrtEndSlicerStyles': break;
			case 'BrtBeginTimelineStylesheetExt15': break;
			case 'BrtEndTimelineStylesheetExt15': break;
			case 'BrtBeginTimelineStyles': break;
			case 'BrtEndTimelineStyles': break;
			case 'BrtEndStyleSheetExt14': break;
			default: if(!pass || opts.WTF) throw new Error("Unexpected record " + RT + " " + R.n);
		}
	});
	return styles;
}

/* [MS-XLSB] 2.1.7.50 Styles */
function write_sty_bin(data, opts) {
	var ba = buf_array();
	write_record(ba, "BrtBeginStyleSheet");
	/* [FMTS] */
	/* [FONTS] */
	/* [FILLS] */
	/* [BORDERS] */
	/* CELLSTYLEXFS */
	/* CELLXFS*/
	/* STYLES */
	/* DXFS */
	/* TABLESTYLES */
	/* [COLORPALETTE] */
	/* FRTSTYLESHEET*/
	write_record(ba, "BrtEndStyleSheet");
	return ba.end();
}
