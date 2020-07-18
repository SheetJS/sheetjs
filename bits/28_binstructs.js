function write_UInt32LE(x/*:number*/, o) {
	if (!o) o = new_buf(4);
	o.write_shift(4, x);
	return o;
}

/* [MS-XLSB] 2.5.168 */
function parse_XLWideString(data/*::, length*/)/*:string*/ {
	var cchCharacters = data.read_shift(4);
	return cchCharacters === 0 ? "" : data.read_shift(cchCharacters, 'dbcs');
}
function write_XLWideString(data/*:string*/, o) {
	var _null = false; if (o == null) { _null = true; o = new_buf(4 + 2 * data.length); }
	o.write_shift(4, data.length);
	if (data.length > 0) o.write_shift(0, data, 'dbcs');
	return _null ? o.slice(0, o.l) : o;
}

/* [MS-XLSB] 2.5.91 */
//function parse_LPWideString(data/*::, length*/)/*:string*/ {
//	var cchCharacters = data.read_shift(2);
//	return cchCharacters === 0 ? "" : data.read_shift(cchCharacters, "utf16le");
//}

/* [MS-XLSB] 2.5.143 */
function parse_StrRun(data) {
	return { ich: data.read_shift(2), ifnt: data.read_shift(2) };
}
function write_StrRun(run, o) {
	if (!o) o = new_buf(4);
	o.write_shift(2, run.ich || 0);
	o.write_shift(2, run.ifnt || 0);
	return o;
}

/* [MS-XLSB] 2.5.121 */
function parse_RichStr(data, length/*:number*/)/*:XLString*/ {
	var start = data.l;
	var flags = data.read_shift(1);
	var str = parse_XLWideString(data);
	var rgsStrRun = [];
	var z = ({ t: str, h: str }/*:any*/);
	if ((flags & 1) !== 0) { /* fRichStr */
		/* TODO: formatted string */
		var dwSizeStrRun = data.read_shift(4);
		for (var i = 0; i != dwSizeStrRun; ++i) rgsStrRun.push(parse_StrRun(data));
		z.r = rgsStrRun;
	}
	else z.r = [{ ich: 0, ifnt: 0 }];
	//if((flags & 2) !== 0) { /* fExtStr */
	//	/* TODO: phonetic string */
	//}
	data.l = start + length;
	return z;
}
function write_RichStr(str/*:XLString*/, o/*:?Block*/)/*:Block*/ {
	/* TODO: formatted string */
	var _null = false; if (o == null) { _null = true; o = new_buf(15 + 4 * str.t.length); }
	o.write_shift(1, 0);
	write_XLWideString(str.t, o);
	return _null ? o.slice(0, o.l) : o;
}
/* [MS-XLSB] 2.4.328 BrtCommentText (RichStr w/1 run) */
var parse_BrtCommentText = parse_RichStr;
function write_BrtCommentText(str/*:XLString*/, o/*:?Block*/)/*:Block*/ {
	/* TODO: formatted string */
	var _null = false; if (o == null) { _null = true; o = new_buf(23 + 4 * str.t.length); }
	o.write_shift(1, 1);
	write_XLWideString(str.t, o);
	o.write_shift(4, 1);
	write_StrRun({ ich: 0, ifnt: 0 }, o);
	return _null ? o.slice(0, o.l) : o;
}

/* [MS-XLSB] 2.5.9 */
function parse_XLSBCell(data)/*:any*/ {
	var col = data.read_shift(4);
	var iStyleRef = data.read_shift(2);
	iStyleRef += data.read_shift(1) << 16;
	data.l++; //var fPhShow = data.read_shift(1);
	return { c: col, iStyleRef: iStyleRef };
}
function write_XLSBCell(cell/*:any*/, o/*:?Block*/) {
	if (o == null) o = new_buf(8);
	o.write_shift(-4, cell.c);
	o.write_shift(3, cell.iStyleRef || cell.s);
	o.write_shift(1, 0); /* fPhShow */
	return o;
}


/* [MS-XLSB] 2.5.21 */
var parse_XLSBCodeName = parse_XLWideString;
var write_XLSBCodeName = write_XLWideString;

/* [MS-XLSB] 2.5.166 */
function parse_XLNullableWideString(data/*::, length*/)/*:string*/ {
	var cchCharacters = data.read_shift(4);
	return cchCharacters === 0 || cchCharacters === 0xFFFFFFFF ? "" : data.read_shift(cchCharacters, 'dbcs');
}
function write_XLNullableWideString(data/*:string*/, o) {
	var _null = false; if (o == null) { _null = true; o = new_buf(127); }
	o.write_shift(4, data.length > 0 ? data.length : 0xFFFFFFFF);
	if (data.length > 0) o.write_shift(0, data, 'dbcs');
	return _null ? o.slice(0, o.l) : o;
}

/* [MS-XLSB] 2.5.165 */
var parse_XLNameWideString = parse_XLWideString;
//var write_XLNameWideString = write_XLWideString;

/* [MS-XLSB] 2.5.114 */
var parse_RelID = parse_XLNullableWideString;
var write_RelID = write_XLNullableWideString;


/* [MS-XLS] 2.5.217 ; [MS-XLSB] 2.5.122 */
function parse_RkNumber(data)/*:number*/ {
	var b = data.slice(data.l, data.l + 4);
	var fX100 = (b[0] & 1), fInt = (b[0] & 2);
	data.l += 4;
	b[0] &= 0xFC; // b[0] &= ~3;
	var RK = fInt === 0 ? __double([0, 0, 0, 0, b[0], b[1], b[2], b[3]], 0) : __readInt32LE(b, 0) >> 2;
	return fX100 ? (RK / 100) : RK;
}
function write_RkNumber(data/*:number*/, o) {
	if (o == null) o = new_buf(4);
	var fX100 = 0, fInt = 0, d100 = data * 100;
	if ((data == (data | 0)) && (data >= -(1 << 29)) && (data < (1 << 29))) { fInt = 1; }
	else if ((d100 == (d100 | 0)) && (d100 >= -(1 << 29)) && (d100 < (1 << 29))) { fInt = 1; fX100 = 1; }
	if (fInt) o.write_shift(-4, ((fX100 ? d100 : data) << 2) + (fX100 + 2));
	else throw new Error("unsupported RkNumber " + data); // TODO
}


/* [MS-XLSB] 2.5.117 RfX */
function parse_RfX(data /*::, length*/)/*:Range*/ {
	var cell/*:Range*/ = ({ s: {}, e: {} }/*:any*/);
	cell.s.r = data.read_shift(4);
	cell.e.r = data.read_shift(4);
	cell.s.c = data.read_shift(4);
	cell.e.c = data.read_shift(4);
	return cell;
}
function write_RfX(r/*:Range*/, o) {
	if (!o) o = new_buf(16);
	o.write_shift(4, r.s.r);
	o.write_shift(4, r.e.r);
	o.write_shift(4, r.s.c);
	o.write_shift(4, r.e.c);
	return o;
}

/* [MS-XLSB] 2.5.153 UncheckedRfX */
var parse_UncheckedRfX = parse_RfX;
var write_UncheckedRfX = write_RfX;

/* [MS-XLSB] 2.5.155 UncheckedSqRfX */
//function parse_UncheckedSqRfX(data) {
//	var cnt = data.read_shift(4);
//	var out = [];
//	for(var i = 0; i < cnt; ++i) {
//		var rng = parse_UncheckedRfX(data);
//		out.push(encode_range(rng));
//	}
//	return out.join(",");
//}
//function write_UncheckedSqRfX(sqrfx/*:string*/) {
//	var parts = sqrfx.split(/\s*,\s*/);
//	var o = new_buf(4); o.write_shift(4, parts.length);
//	var out = [o];
//	parts.forEach(function(rng) {
//		out.push(write_UncheckedRfX(safe_decode_range(rng)));
//	});
//	return bconcat(out);
//}

/* [MS-XLS] 2.5.342 ; [MS-XLSB] 2.5.171 */
/* TODO: error checking, NaN and Infinity values are not valid Xnum */
function parse_Xnum(data/*::, length*/) { return data.read_shift(8, 'f'); }
function write_Xnum(data, o) { return (o || new_buf(8)).write_shift(8, data, 'f'); }

/* [MS-XLSB] 2.4.324 BrtColor */
function parse_BrtColor(data/*::, length*/) {
	var out = {};
	var d = data.read_shift(1);

	//var fValidRGB = d & 1;
	var xColorType = d >>> 1;

	var index = data.read_shift(1);
	var nTS = data.read_shift(2, 'i');
	var bR = data.read_shift(1);
	var bG = data.read_shift(1);
	var bB = data.read_shift(1);
	data.l++; //var bAlpha = data.read_shift(1);

	switch (xColorType) {
		case 0: out.auto = 1; break;
		case 1:
			out.index = index;
			var icv = XLSIcv[index];
			/* automatic pseudo index 81 */
			if (icv) out.rgb = rgb2Hex(icv);
			break;
		case 2:
			/* if(!fValidRGB) throw new Error("invalid"); */
			out.rgb = rgb2Hex([bR, bG, bB]);
			break;
		case 3: out.theme = index; break;
	}
	if (nTS != 0) out.tint = nTS > 0 ? nTS / 32767 : nTS / 32768;

	return out;
}
function write_BrtColor(color, o) {
	if (!o) o = new_buf(8);
	if (!color || color.auto) { o.write_shift(4, 0); o.write_shift(4, 0); return o; }
	if (color.index != null) {
		o.write_shift(1, 0x02);
		o.write_shift(1, color.index);
	} else if (color.theme != null) {
		o.write_shift(1, 0x06);
		o.write_shift(1, color.theme);
	} else {
		o.write_shift(1, 0x05);
		o.write_shift(1, 0);
	}
	var nTS = color.tint || 0;
	if (nTS > 0) nTS *= 32767;
	else if (nTS < 0) nTS *= 32768;
	o.write_shift(2, nTS);
	if (!color.rgb || color.theme != null) {
		o.write_shift(2, 0);
		o.write_shift(1, 0);
		o.write_shift(1, 0);
	} else {
		var rgb = (color.rgb || 'FFFFFF');
		if (typeof rgb == 'number') rgb = ("000000" + rgb.toString(16)).slice(-6);
		o.write_shift(1, parseInt(rgb.slice(0, 2), 16));
		o.write_shift(1, parseInt(rgb.slice(2, 4), 16));
		o.write_shift(1, parseInt(rgb.slice(4, 6), 16));
		o.write_shift(1, 0xFF);
	}
	return o;
}

/* [MS-XLSB] 2.5.52 */
function parse_FontFlags(data/*::, length, opts*/) {
	var d = data.read_shift(1);
	data.l++;
	var out = {
		fBold: d & 0x01,
		fItalic: d & 0x02,
		fUnderline: d & 0x04,
		fStrikeout: d & 0x08,
		fOutline: d & 0x10,
		fShadow: d & 0x20,
		fCondense: d & 0x40,
		fExtend: d & 0x80
	};
	return out;
}
function write_FontFlags(font, o) {
	if (!o) o = new_buf(2);
	var grbit =
		(font.italic ? 0x02 : 0) |
		(font.strike ? 0x08 : 0) |
		(font.outline ? 0x10 : 0) |
		(font.shadow ? 0x20 : 0) |
		(font.condense ? 0x40 : 0) |
		(font.extend ? 0x80 : 0);
	o.write_shift(1, grbit);
	o.write_shift(1, 0);
	return o;
}

/* [MS-OLEDS] 2.3.1 and 2.3.2 */
function parse_ClipboardFormatOrString(o, w/*:number*/)/*:string*/ {
	// $FlowIgnore
	var ClipFmt = { 2: "BITMAP", 3: "METAFILEPICT", 8: "DIB", 14: "ENHMETAFILE" };
	var m/*:number*/ = o.read_shift(4);
	switch (m) {
		case 0x00000000: return "";
		case 0xffffffff: case 0xfffffffe: return ClipFmt[o.read_shift(4)] || "";
	}
	if (m > 0x190) throw new Error("Unsupported Clipboard: " + m.toString(16));
	o.l -= 4;
	return o.read_shift(0, w == 1 ? "lpstr" : "lpwstr");
}
function parse_ClipboardFormatOrAnsiString(o) { return parse_ClipboardFormatOrString(o, 1); }
function parse_ClipboardFormatOrUnicodeString(o) { return parse_ClipboardFormatOrString(o, 2); }

