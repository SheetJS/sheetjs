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
	styles.Fonts = [];
	var state = [];
	var pass = false;
	recordhopper(data, function hopper_sty(val, R_n, RT) {
		switch(RT) {
			case 0x002C: /* 'BrtFmt' */
				styles.NumberFmt[val[0]] = val[1]; SSF.load(val[1], val[0]);
				break;
			case 0x002B: /* 'BrtFont' */
				styles.Fonts.push(val);
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
