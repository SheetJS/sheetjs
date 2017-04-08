/* [MS-XLSB] 2.4.651 BrtFmt */
function parse_BrtFmt(data, length/*:number*/) {
	var ifmt = data.read_shift(2);
	var stFmtCode = parse_XLWideString(data,length-2);
	return [ifmt, stFmtCode];
}

/* [MS-XLSB] 2.4.653 BrtFont TODO */
function parse_BrtFont(data, length/*:number*/) {
	var out = ({flags:{}}/*:any*/);
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
function parse_BrtXF(data, length/*:number*/) {
	var ixfeParent = data.read_shift(2);
	var ifmt = data.read_shift(2);
	parsenoop(data, length-4);
	return {ixfe:ixfeParent, ifmt:ifmt };
}

/* [MS-XLSB] 2.1.7.50 Styles */
function parse_sty_bin(data, themes, opts) {
	var styles = {};
	styles.NumberFmt = ([]/*:any*/);
	for(var y in SSF._table) styles.NumberFmt[y] = SSF._table[y];

	styles.CellXf = [];
	var state = [];
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
				if(state[state.length - 1] == "BrtBeginCellXFs") {
					styles.CellXf.push(val);
				}
				break; /* TODO */
			case 'BrtStyle': break; /* TODO */
			case 'BrtDXF': break; /* TODO */
			case 'BrtMRUColor': break; /* TODO */
			case 'BrtIndexedColor': break; /* TODO */

			case 'BrtDXF14': break;
			case 'BrtDXF15': break;
			case 'BrtUid': break;
			case 'BrtSlicerStyleElement': break;
			case 'BrtTableStyleElement': break;
			case 'BrtTimelineStyleElement': break;

			case 'BrtFRTBegin': pass = true; break;
			case 'BrtFRTEnd': pass = false; break;
			case 'BrtACBegin': state.push(R.n); break;
			case 'BrtACEnd': state.pop(); break;

			default:
				if((R.n||"").indexOf("Begin") > 0) state.push(R.n);
				else if((R.n||"").indexOf("End") > 0) state.pop();
				else if(!pass || opts.WTF) throw new Error("Unexpected record " + RT + " " + R.n);
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
