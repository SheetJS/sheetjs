/* --- MS-XLS --- */

/* 2.5.19 */
function parse_XLSCell(blob, length)/*:Cell*/ {
	var rw = blob.read_shift(2); // 0-indexed
	var col = blob.read_shift(2);
	var ixfe = blob.read_shift(2);
	return ({r:rw, c:col, ixfe:ixfe}/*:any*/);
}

/* 2.5.134 */
function parse_frtHeader(blob) {
	var rt = blob.read_shift(2);
	var flags = blob.read_shift(2); // TODO: parse these flags
	blob.l += 8;
	return {type: rt, flags: flags};
}



function parse_OptXLUnicodeString(blob, length, opts) { return length === 0 ? "" : parse_XLUnicodeString2(blob, length, opts); }

/* 2.5.158 */
//var HIDEOBJENUM = ['SHOWALL', 'SHOWPLACEHOLDER', 'HIDEALL'];
var parse_HideObjEnum = parseuint16;

/* 2.5.344 */
function parse_XTI(blob, length, opts) {
	var w = opts.biff > 8 ? 4 : 2;
	var iSupBook = blob.read_shift(w), itabFirst = blob.read_shift(w,'i'), itabLast = blob.read_shift(w,'i');
	return [iSupBook, itabFirst, itabLast];
}

/* 2.5.218 */
function parse_RkRec(blob, length) {
	var ixfe = blob.read_shift(2);
	var RK = parse_RkNumber(blob);
	return [ixfe, RK];
}

/* 2.5.1 */
function parse_AddinUdf(blob, length, opts) {
	blob.l += 4; length -= 4;
	var l = blob.l + length;
	var udfName = parse_ShortXLUnicodeString(blob, length, opts);
	var cb = blob.read_shift(2);
	l -= blob.l;
	if(cb !== l) throw new Error("Malformed AddinUdf: padding = " + l + " != " + cb);
	blob.l += cb;
	return udfName;
}

/* 2.5.209 TODO: Check sizes */
function parse_Ref8U(blob, length) {
	var rwFirst = blob.read_shift(2);
	var rwLast = blob.read_shift(2);
	var colFirst = blob.read_shift(2);
	var colLast = blob.read_shift(2);
	return {s:{c:colFirst, r:rwFirst}, e:{c:colLast,r:rwLast}};
}

/* 2.5.211 */
function parse_RefU(blob, length) {
	var rwFirst = blob.read_shift(2);
	var rwLast = blob.read_shift(2);
	var colFirst = blob.read_shift(1);
	var colLast = blob.read_shift(1);
	return {s:{c:colFirst, r:rwFirst}, e:{c:colLast,r:rwLast}};
}

/* 2.5.207 */
var parse_Ref = parse_RefU;

/* 2.5.143 */
function parse_FtCmo(blob, length) {
	blob.l += 4;
	var ot = blob.read_shift(2);
	var id = blob.read_shift(2);
	var flags = blob.read_shift(2);
	blob.l+=12;
	return [id, ot, flags];
}

/* 2.5.149 */
function parse_FtNts(blob, length) {
	var out = {};
	blob.l += 4;
	blob.l += 16; // GUID TODO
	out.fSharedNote = blob.read_shift(2);
	blob.l += 4;
	return out;
}

/* 2.5.142 */
function parse_FtCf(blob, length) {
	var out = {};
	blob.l += 4;
	blob.cf = blob.read_shift(2);
	return out;
}

/* 2.5.140 - 2.5.154 and friends */
function parse_FtSkip(blob, length) { blob.l += 2; blob.l += blob.read_shift(2); }
var FtTab = {
	/*::[*/0x00/*::]*/: parse_FtSkip,      /* FtEnd */
	/*::[*/0x04/*::]*/: parse_FtSkip,      /* FtMacro */
	/*::[*/0x05/*::]*/: parse_FtSkip,      /* FtButton */
	/*::[*/0x06/*::]*/: parse_FtSkip,      /* FtGmo */
	/*::[*/0x07/*::]*/: parse_FtCf,        /* FtCf */
	/*::[*/0x08/*::]*/: parse_FtSkip,      /* FtPioGrbit */
	/*::[*/0x09/*::]*/: parse_FtSkip,      /* FtPictFmla */
	/*::[*/0x0A/*::]*/: parse_FtSkip,      /* FtCbls */
	/*::[*/0x0B/*::]*/: parse_FtSkip,      /* FtRbo */
	/*::[*/0x0C/*::]*/: parse_FtSkip,      /* FtSbs */
	/*::[*/0x0D/*::]*/: parse_FtNts,       /* FtNts */
	/*::[*/0x0E/*::]*/: parse_FtSkip,      /* FtSbsFmla */
	/*::[*/0x0F/*::]*/: parse_FtSkip,      /* FtGboData */
	/*::[*/0x10/*::]*/: parse_FtSkip,      /* FtEdoData */
	/*::[*/0x11/*::]*/: parse_FtSkip,      /* FtRboData */
	/*::[*/0x12/*::]*/: parse_FtSkip,      /* FtCblsData */
	/*::[*/0x13/*::]*/: parse_FtSkip,      /* FtLbsData */
	/*::[*/0x14/*::]*/: parse_FtSkip,      /* FtCblsFmla */
	/*::[*/0x15/*::]*/: parse_FtCmo
};
function parse_FtArray(blob, length, ot) {
	var tgt = blob.l + length;
	var fts = [];
	while(blob.l < tgt) {
		var ft = blob.read_shift(2);
		blob.l-=2;
		try {
			fts.push(FtTab[ft](blob, tgt - blob.l));
		} catch(e) { blob.l = tgt; return fts; }
	}
	if(blob.l != tgt) blob.l = tgt; //throw new Error("bad Object Ft-sequence");
	return fts;
}

/* 2.5.129 */
var parse_FontIndex = parseuint16;

/* --- 2.4 Records --- */

/* 2.4.21 */
function parse_BOF(blob, length) {
	var o = {BIFFVer:0, dt:0};
	o.BIFFVer = blob.read_shift(2); length -= 2;
	if(length >= 2) { o.dt = blob.read_shift(2); blob.l -= 2; }
	switch(o.BIFFVer) {
		case 0x0600: /* BIFF8 */
		case 0x0500: /* BIFF5 */
		case 0x0002: case 0x0007: /* BIFF2 */
			break;
		default: if(length > 6) throw new Error("Unexpected BIFF Ver " + o.BIFFVer);
	}

	blob.read_shift(length);
	return o;
}


/* 2.4.146 */
function parse_InterfaceHdr(blob, length) {
	if(length === 0) return 0x04b0;
	var q;
	if((q=blob.read_shift(2))!==0x04b0){/* empty */}
	return 0x04b0;
}


/* 2.4.349 */
function parse_WriteAccess(blob, length, opts) {
	if(opts.enc) { blob.l += length; return ""; }
	var l = blob.l;
	// TODO: make sure XLUnicodeString doesnt overrun
	var UserName = parse_XLUnicodeString(blob, 0, opts);
	blob.read_shift(length + l - blob.l);
	return UserName;
}

/* 2.4.28 */
function parse_BoundSheet8(blob, length, opts) {
	var pos = blob.read_shift(4);
	var hidden = blob.read_shift(1) & 0x03;
	var dt = blob.read_shift(1);
	switch(dt) {
		case 0: dt = 'Worksheet'; break;
		case 1: dt = 'Macrosheet'; break;
		case 2: dt = 'Chartsheet'; break;
		case 6: dt = 'VBAModule'; break;
	}
	var name = parse_ShortXLUnicodeString(blob, 0, opts);
	if(name.length === 0) name = "Sheet1";
	return { pos:pos, hs:hidden, dt:dt, name:name };
}

/* 2.4.265 TODO */
function parse_SST(blob, length)/*:SST*/ {
	var end = blob.l + length;
	var cnt = blob.read_shift(4);
	var ucnt = blob.read_shift(4);
	var strs/*:SST*/ = ([]/*:any*/);
	for(var i = 0; i != ucnt && blob.l < end; ++i) {
		strs.push(parse_XLUnicodeRichExtendedString(blob));
	}
	strs.Count = cnt; strs.Unique = ucnt;
	return strs;
}

/* 2.4.107 */
function parse_ExtSST(blob, length) {
	var extsst = {};
	extsst.dsst = blob.read_shift(2);
	blob.l += length-2;
	return extsst;
}


/* 2.4.221 TODO: check BIFF2-4 */
function parse_Row(blob, length) {
	var z = ({}/*:any*/);
	z.r = blob.read_shift(2);
	z.c = blob.read_shift(2);
	z.cnt = blob.read_shift(2) - z.c;
	var miyRw = blob.read_shift(2);
	blob.l += 4; // reserved(2), unused(2)
	var flags = blob.read_shift(1); // various flags
	blob.l += 3; // reserved(8), ixfe(12), flags(4)
	if(flags & 0x07) z.level = flags & 0x07;
	// collapsed: flags & 0x10
	if(flags & 0x20) z.hidden = true;
	if(flags & 0x40) z.hpt = miyRw / 20;
	return z;
}


/* 2.4.125 */
function parse_ForceFullCalculation(blob, length) {
	var header = parse_frtHeader(blob);
	if(header.type != 0x08A3) throw new Error("Invalid Future Record " + header.type);
	var fullcalc = blob.read_shift(4);
	return fullcalc !== 0x0;
}


var parse_CompressPictures = parsenoop2; /* 2.4.55 Not interesting */



/* 2.4.215 rt */
function parse_RecalcId(blob, length) {
	blob.read_shift(2);
	return blob.read_shift(4);
}

/* 2.4.87 */
function parse_DefaultRowHeight(blob, length, opts) {
	var f = 0;
	if(!(opts && opts.biff == 2)) {
		f = blob.read_shift(2);
	}
	var miyRw = blob.read_shift(2);
	if((opts && opts.biff == 2)) {
		f = 1 - (miyRw >> 15); miyRw &= 0x7fff;
	}
	var fl = {Unsynced:f&1,DyZero:(f&2)>>1,ExAsc:(f&4)>>2,ExDsc:(f&8)>>3};
	return [fl, miyRw];
}

/* 2.4.345 TODO */
function parse_Window1(blob, length) {
	var xWn = blob.read_shift(2), yWn = blob.read_shift(2), dxWn = blob.read_shift(2), dyWn = blob.read_shift(2);
	var flags = blob.read_shift(2), iTabCur = blob.read_shift(2), iTabFirst = blob.read_shift(2);
	var ctabSel = blob.read_shift(2), wTabRatio = blob.read_shift(2);
	return { Pos: [xWn, yWn], Dim: [dxWn, dyWn], Flags: flags, CurTab: iTabCur,
		FirstTab: iTabFirst, Selected: ctabSel, TabRatio: wTabRatio };
}

/* 2.4.122 TODO */
function parse_Font(blob, length, opts) {
	var o/*:any*/ = {
		dyHeight: blob.read_shift(2),
		fl: blob.read_shift(2)
	};
	switch(opts && opts.biff || 8) {
		case 2: break;
		case 3: case 4: blob.l += 2; break;
		default: blob.l += 10; break;
	}
	o.name = parse_ShortXLUnicodeString(blob, 0, opts);
	return o;
}

/* 2.4.149 */
function parse_LabelSst(blob, length) {
	var cell = parse_XLSCell(blob);
	cell.isst = blob.read_shift(4);
	return cell;
}

/* 2.4.148 */
function parse_Label(blob, length, opts) {
	var target = blob.l + length;
	var cell = parse_XLSCell(blob, 6);
	if(opts.biff == 2) blob.l++;
	var str = parse_XLUnicodeString(blob, target - blob.l, opts);
	cell.val = str;
	return cell;
}

/* 2.4.126 Number Formats */
function parse_Format(blob, length, opts) {
	var numFmtId = blob.read_shift(2);
	var fmtstr = parse_XLUnicodeString2(blob, 0, opts);
	return [numFmtId, fmtstr];
}
var parse_BIFF2Format = parse_XLUnicodeString2;

/* 2.4.90 */
function parse_Dimensions(blob, length, opts) {
	var end = blob.l + length;
	var w = opts.biff == 8 || !opts.biff ? 4 : 2;
	var r = blob.read_shift(w), R = blob.read_shift(w);
	var c = blob.read_shift(2), C = blob.read_shift(2);
	blob.l = end;
	return {s: {r:r, c:c}, e: {r:R, c:C}};
}

/* 2.4.220 */
function parse_RK(blob, length) {
	var rw = blob.read_shift(2), col = blob.read_shift(2);
	var rkrec = parse_RkRec(blob);
	return {r:rw, c:col, ixfe:rkrec[0], rknum:rkrec[1]};
}

/* 2.4.175 */
function parse_MulRk(blob, length) {
	var target = blob.l + length - 2;
	var rw = blob.read_shift(2), col = blob.read_shift(2);
	var rkrecs = [];
	while(blob.l < target) rkrecs.push(parse_RkRec(blob));
	if(blob.l !== target) throw new Error("MulRK read error");
	var lastcol = blob.read_shift(2);
	if(rkrecs.length != lastcol - col + 1) throw new Error("MulRK length mismatch");
	return {r:rw, c:col, C:lastcol, rkrec:rkrecs};
}
/* 2.4.174 */
function parse_MulBlank(blob, length) {
	var target = blob.l + length - 2;
	var rw = blob.read_shift(2), col = blob.read_shift(2);
	var ixfes = [];
	while(blob.l < target) ixfes.push(blob.read_shift(2));
	if(blob.l !== target) throw new Error("MulBlank read error");
	var lastcol = blob.read_shift(2);
	if(ixfes.length != lastcol - col + 1) throw new Error("MulBlank length mismatch");
	return {r:rw, c:col, C:lastcol, ixfe:ixfes};
}

/* 2.5.20 2.5.249 TODO: interpret values here */
function parse_CellStyleXF(blob, length, style, opts) {
	var o = {};
	var a = blob.read_shift(4), b = blob.read_shift(4);
	var c = blob.read_shift(4), d = blob.read_shift(2);
	o.patternType = XLSFillPattern[c >> 26];

	if(!opts.cellStyles) return o;
	o.alc = a & 0x07;
	o.fWrap = (a >> 3) & 0x01;
	o.alcV = (a >> 4) & 0x07;
	o.fJustLast = (a >> 7) & 0x01;
	o.trot = (a >> 8) & 0xFF;
	o.cIndent = (a >> 16) & 0x0F;
	o.fShrinkToFit = (a >> 20) & 0x01;
	o.iReadOrder = (a >> 22) & 0x02;
	o.fAtrNum = (a >> 26) & 0x01;
	o.fAtrFnt = (a >> 27) & 0x01;
	o.fAtrAlc = (a >> 28) & 0x01;
	o.fAtrBdr = (a >> 29) & 0x01;
	o.fAtrPat = (a >> 30) & 0x01;
	o.fAtrProt = (a >> 31) & 0x01;

	o.dgLeft = b & 0x0F;
	o.dgRight = (b >> 4) & 0x0F;
	o.dgTop = (b >> 8) & 0x0F;
	o.dgBottom = (b >> 12) & 0x0F;
	o.icvLeft = (b >> 16) & 0x7F;
	o.icvRight = (b >> 23) & 0x7F;
	o.grbitDiag = (b >> 30) & 0x03;

	o.icvTop = c & 0x7F;
	o.icvBottom = (c >> 7) & 0x7F;
	o.icvDiag = (c >> 14) & 0x7F;
	o.dgDiag = (c >> 21) & 0x0F;

	o.icvFore = d & 0x7F;
	o.icvBack = (d >> 7) & 0x7F;
	o.fsxButton = (d >> 14) & 0x01;
	return o;
}
function parse_CellXF(blob, length, opts) {return parse_CellStyleXF(blob,length,0, opts);}
function parse_StyleXF(blob, length, opts) {return parse_CellStyleXF(blob,length,1, opts);}

/* 2.4.353 TODO: actually do this right */
function parse_XF(blob, length, opts) {
	var o = {};
	o.ifnt = blob.read_shift(2); o.numFmtId = blob.read_shift(2); o.flags = blob.read_shift(2);
	o.fStyle = (o.flags >> 2) & 0x01;
	length -= 6;
	o.data = parse_CellStyleXF(blob, length, o.fStyle, opts);
	return o;
}

/* 2.4.134 */
function parse_Guts(blob, length) {
	blob.l += 4;
	var out = [blob.read_shift(2), blob.read_shift(2)];
	if(out[0] !== 0) out[0]--;
	if(out[1] !== 0) out[1]--;
	if(out[0] > 7 || out[1] > 7) throw new Error("Bad Gutters: " + out.join("|"));
	return out;
}

/* 2.4.24 */
function parse_BoolErr(blob, length, opts) {
	var cell = parse_XLSCell(blob, 6);
	if(opts.biff == 2) ++blob.l;
	var val = parse_Bes(blob, 2);
	cell.val = val;
	cell.t = (val === true || val === false) ? 'b' : 'e';
	return cell;
}

/* 2.4.180 Number */
function parse_Number(blob, length) {
	var cell = parse_XLSCell(blob, 6);
	var xnum = parse_Xnum(blob, 8);
	cell.val = xnum;
	return cell;
}

var parse_XLHeaderFooter = parse_OptXLUnicodeString; // TODO: parse 2.4.136

/* 2.4.271 */
function parse_SupBook(blob, length, opts) {
	var end = blob.l + length;
	var ctab = blob.read_shift(2);
	var cch = blob.read_shift(2);
	opts.sbcch = cch;
	if(cch == 0x0401 || cch == 0x3A01) return [cch, ctab];
	if(cch < 0x01 || cch >0xff) throw new Error("Unexpected SupBook type: "+cch);
	var virtPath = parse_XLUnicodeStringNoCch(blob, cch);
	/* TODO: 2.5.277 Virtual Path */
	var rgst = [];
	while(end > blob.l) rgst.push(parse_XLUnicodeString(blob));
	return [cch, ctab, virtPath, rgst];
}

/* 2.4.105 TODO */
function parse_ExternName(blob, length, opts) {
	var flags = blob.read_shift(2);
	var body;
	var o = ({
		fBuiltIn: flags & 0x01,
		fWantAdvise: (flags >>> 1) & 0x01,
		fWantPict: (flags >>> 2) & 0x01,
		fOle: (flags >>> 3) & 0x01,
		fOleLink: (flags >>> 4) & 0x01,
		cf: (flags >>> 5) & 0x3FF,
		fIcon: flags >>> 15 & 0x01
	}/*:any*/);
	if(opts.sbcch === 0x3A01) body = parse_AddinUdf(blob, length-2, opts);
	//else throw new Error("unsupported SupBook cch: " + opts.sbcch);
	o.body = body || blob.read_shift(length-2);
	if(typeof body === "string") o.Name = body;
	return o;
}

/* 2.4.150 TODO */
var XLSLblBuiltIn = [
	"_xlnm.Consolidate_Area",
	"_xlnm.Auto_Open",
	"_xlnm.Auto_Close",
	"_xlnm.Extract",
	"_xlnm.Database",
	"_xlnm.Criteria",
	"_xlnm.Print_Area",
	"_xlnm.Print_Titles",
	"_xlnm.Recorder",
	"_xlnm.Data_Form",
	"_xlnm.Auto_Activate",
	"_xlnm.Auto_Deactivate",
	"_xlnm.Sheet_Title",
	"_xlnm._FilterDatabase"
];
function parse_Lbl(blob, length, opts) {
	var target = blob.l + length;
	var flags = blob.read_shift(2);
	var chKey = blob.read_shift(1);
	var cch = blob.read_shift(1);
	var cce = blob.read_shift(opts && opts.biff == 2 ? 1 : 2);
	var itab = 0;
	if(!opts || opts.biff >= 5) {
		blob.l += 2;
		itab = blob.read_shift(2);
		blob.l += 4;
	}
	var name = parse_XLUnicodeStringNoCch(blob, cch, opts);
	if(flags & 0x20) name = XLSLblBuiltIn[name.charCodeAt(0)];
	var npflen = target - blob.l; if(opts && opts.biff == 2) --npflen;
	var rgce = target == blob.l || cce == 0 ? [] : parse_NameParsedFormula(blob, npflen, opts, cce);
	return {
		chKey: chKey,
		Name: name,
		itab: itab,
		rgce: rgce
	};
}

/* 2.4.106 TODO: verify filename encoding */
function parse_ExternSheet(blob, length, opts) {
	if(opts.biff < 8) return parse_BIFF5ExternSheet(blob, length, opts);
	var o = [], target = blob.l + length, len = blob.read_shift(opts.biff > 8 ? 4 : 2);
	while(len-- !== 0) o.push(parse_XTI(blob, opts.biff > 8 ? 12 : 6, opts));
		// [iSupBook, itabFirst, itabLast];
	var oo = [];
	return o;
}
function parse_BIFF5ExternSheet(blob, length, opts) {
	if(blob[blob.l + 1] == 0x03) blob[blob.l]++;
	var o = parse_ShortXLUnicodeString(blob, length, opts);
	return o.charCodeAt(0) == 0x03 ? o.slice(1) : o;
}
var parse_ExternCount = parseuint16;

/* 2.4.176 TODO: check older biff */
function parse_NameCmt(blob, length, opts) {
	if(opts.biff < 8) { blob.l += length; return; }
	var cchName = blob.read_shift(2);
	var cchComment = blob.read_shift(2);
	var name = parse_XLUnicodeStringNoCch(blob, cchName, opts);
	var comment = parse_XLUnicodeStringNoCch(blob, cchComment, opts);
	return [name, comment];
}

/* 2.4.260 */
function parse_ShrFmla(blob, length, opts) {
	var ref = parse_RefU(blob, 6);
	blob.l++;
	var cUse = blob.read_shift(1);
	length -= 8;
	return [parse_SharedParsedFormula(blob, length, opts), cUse];
}

/* 2.4.4 TODO */
function parse_Array(blob, length, opts) {
	var ref = parse_Ref(blob, 6);
	/* TODO: fAlwaysCalc */
	switch(opts.biff) {
		case 2: blob.l ++; length -= 7; break;
		case 3: case 4: blob.l += 2; length -= 8; break;
		default: blob.l += 6; length -= 12;
	}
	return [ref, parse_ArrayParsedFormula(blob, length, opts, ref)];
}

/* 2.4.173 */
function parse_MTRSettings(blob, length) {
	var fMTREnabled = blob.read_shift(4) !== 0x00;
	var fUserSetThreadCount = blob.read_shift(4) !== 0x00;
	var cUserThreadCount = blob.read_shift(4);
	return [fMTREnabled, fUserSetThreadCount, cUserThreadCount];
}

/* 2.5.186 TODO: BIFF5 */
function parse_NoteSh(blob, length, opts) {
	if(opts.biff < 8) return;
	var row = blob.read_shift(2), col = blob.read_shift(2);
	var flags = blob.read_shift(2), idObj = blob.read_shift(2);
	var stAuthor = parse_XLUnicodeString2(blob, 0, opts);
	if(opts.biff < 8) blob.read_shift(1);
	return [{r:row,c:col}, stAuthor, idObj, flags];
}

/* 2.4.179 */
function parse_Note(blob, length, opts) {
	/* TODO: Support revisions */
	return parse_NoteSh(blob, length, opts);
}

/* 2.4.168 */
function parse_MergeCells(blob, length) {
	var merges = [];
	var cmcs = blob.read_shift(2);
	while (cmcs--) merges.push(parse_Ref8U(blob,length));
	return merges;
}

/* 2.4.181 TODO: parse all the things! */
function parse_Obj(blob, length, opts) {
	if(opts && opts.biff < 8) return parse_BIFF5Obj(blob, length, opts);
	var cmo = parse_FtCmo(blob, 22); // id, ot, flags
	var fts = parse_FtArray(blob, length-22, cmo[1]);
	return { cmo: cmo, ft:fts };
}
/* from older spec */
var parse_BIFF5OT = [];
parse_BIFF5OT[0x08] = function(blob, length, opts) {
	var tgt = blob.l + length;
	blob.l += 10; // todo
	var cf = blob.read_shift(2);
	blob.l += 4;
	var cbPictFmla = blob.read_shift(2);
	blob.l += 2;
	var grbit = blob.read_shift(2);
	blob.l += 4;
	var cchName = blob.read_shift(1);
	blob.l += cchName; // TODO: stName
	blob.l = tgt; // TODO: fmla
	return { fmt:cf };
};

function parse_BIFF5Obj(blob, length, opts) {
	var cnt = blob.read_shift(4);
	var ot = blob.read_shift(2);
	var id = blob.read_shift(2);
	var grbit = blob.read_shift(2);
	var colL = blob.read_shift(2);
	var dxL = blob.read_shift(2);
	var rwT = blob.read_shift(2);
	var dyT = blob.read_shift(2);
	var colR = blob.read_shift(2);
	var dxR = blob.read_shift(2);
	var rwB = blob.read_shift(2);
	var dyB = blob.read_shift(2);
	var cbMacro = blob.read_shift(2);
	blob.l += 6;
	length -= 36;
	var fts = [];
	fts.push((parse_BIFF5OT[ot]||parsenoop)(blob, length, opts));
	return { cmo: [id, ot, grbit], ft:fts };
}

/* 2.4.329 TODO: parse properly */
function parse_TxO(blob, length, opts) {
	var s = blob.l;
	var texts = "";
try {
	blob.l += 4;
	var ot = (opts.lastobj||{cmo:[0,0]}).cmo[1];
	var controlInfo;
	if([0,5,7,11,12,14].indexOf(ot) == -1) blob.l += 6;
	else controlInfo = parse_ControlInfo(blob, 6, opts);
	var cchText = blob.read_shift(2);
	var cbRuns = blob.read_shift(2);
	var ifntEmpty = parse_FontIndex(blob, 2);
	var len = blob.read_shift(2);
	blob.l += len;
	//var fmla = parse_ObjFmla(blob, s + length - blob.l);

	for(var i = 1; i < blob.lens.length-1; ++i) {
		if(blob.l-s != blob.lens[i]) throw new Error("TxO: bad continue record");
		var hdr = blob[blob.l];
		var t = parse_XLUnicodeStringNoCch(blob, blob.lens[i+1]-blob.lens[i]-1);
		texts += t;
		if(texts.length >= (hdr ? cchText : 2*cchText)) break;
	}
	if(texts.length !== cchText && texts.length !== cchText*2) {
		throw new Error("cchText: " + cchText + " != " + texts.length);
	}

	blob.l = s + length;
	/* 2.5.272 TxORuns */
//	var rgTxoRuns = [];
//	for(var j = 0; j != cbRuns/8-1; ++j) blob.l += 8;
//	var cchText2 = blob.read_shift(2);
//	if(cchText2 !== cchText) throw new Error("TxOLastRun mismatch: " + cchText2 + " " + cchText);
//	blob.l += 6;
//	if(s + length != blob.l) throw new Error("TxO " + (s + length) + ", at " + blob.l);
	return { t: texts };
} catch(e) { blob.l = s + length; return { t: texts }; }
}

/* 2.4.140 */
var parse_HLink = function(blob, length) {
	var ref = parse_Ref8U(blob, 8);
	blob.l += 16; /* CLSID */
	var hlink = parse_Hyperlink(blob, length-24);
	return [ref, hlink];
};

/* 2.4.141 */
var parse_HLinkTooltip = function(blob, length) {
	var end = blob.l + length;
	blob.read_shift(2);
	var ref = parse_Ref8U(blob, 8);
	var wzTooltip = blob.read_shift((length-10)/2, 'dbcs-cont');
	wzTooltip = wzTooltip.replace(chr0,"");
	return [ref, wzTooltip];
};

/* 2.4.63 */
function parse_Country(blob, length) {
	var o = [], d;
	d = blob.read_shift(2); o[0] = CountryEnum[d] || d;
	d = blob.read_shift(2); o[1] = CountryEnum[d] || d;
	return o;
}

/* 2.4.50 ClrtClient */
function parse_ClrtClient(blob, length) {
	var ccv = blob.read_shift(2);
	var o = [];
	while(ccv-->0) o.push(parse_LongRGB(blob, 8));
	return o;
}

/* 2.4.188 */
function parse_Palette(blob, length) {
	var ccv = blob.read_shift(2);
	var o = [];
	while(ccv-->0) o.push(parse_LongRGB(blob, 8));
	return o;
}

/* 2.4.354 */
function parse_XFCRC(blob, length) {
	blob.l += 2;
	var o = {cxfs:0, crc:0};
	o.cxfs = blob.read_shift(2);
	o.crc = blob.read_shift(4);
	return o;
}

/* 2.4.53 TODO: parse flags */
/* [MS-XLSB] 2.4.323 TODO: parse flags */
function parse_ColInfo(blob, length, opts) {
	if(!opts.cellStyles) return parsenoop(blob, length);
	var w = opts && opts.biff >= 12 ? 4 : 2;
	var colFirst = blob.read_shift(w);
	var colLast = blob.read_shift(w);
	var coldx = blob.read_shift(w);
	var ixfe = blob.read_shift(w);
	var flags = blob.read_shift(2);
	if(w == 2) blob.l += 2;
	return {s:colFirst, e:colLast, w:coldx, ixfe:ixfe, flags:flags};
}

/* 2.4.257 */
function parse_Setup(blob, length, opts) {
	var o = {};
	blob.l += 16;
	o.header = parse_Xnum(blob, 8);
	o.footer = parse_Xnum(blob, 8);
	blob.l += 2;
	return o;
}

/* 2.4.261 */
function parse_ShtProps(blob, length, opts) {
	var def = {area:false};
	if(opts.biff != 5) { blob.l += length; return def; }
	var d = blob.read_shift(1); blob.l += 3;
	if((d & 0x10)) def.area = true;
	return def;
}

var parse_Style = parsenoop;
var parse_StyleExt = parsenoop;

var parse_Window2 = parsenoop;

var parse_Backup = parsebool; /* 2.4.14 */
var parse_Blank = parse_XLSCell; /* 2.4.20 Just the cell */
var parse_BottomMargin = parse_Xnum; /* 2.4.27 */
var parse_BuiltInFnGroupCount = parseuint16; /* 2.4.30 0x0E or 0x10 but excel 2011 generates 0x11? */
var parse_CalcCount = parseuint16; /* 2.4.31 #Iterations */
var parse_CalcDelta = parse_Xnum; /* 2.4.32 */
var parse_CalcIter = parsebool;  /* 2.4.33 1=iterative calc */
var parse_CalcMode = parseuint16; /* 2.4.34 0=manual, 1=auto (def), 2=table */
var parse_CalcPrecision = parsebool; /* 2.4.35 */
var parse_CalcRefMode = parsenoop2; /* 2.4.36 */
var parse_CalcSaveRecalc = parsebool; /* 2.4.37 */
var parse_CodePage = parseuint16; /* 2.4.52 */
var parse_Compat12 = parsebool; /* 2.4.54 true = no compatibility check */
var parse_Date1904 = parsebool; /* 2.4.77 - 1=1904,0=1900 */
var parse_DefColWidth = parseuint16; /* 2.4.89 */
var parse_DSF = parsenoop2; /* 2.4.94 -- MUST be ignored */
var parse_EntExU2 = parsenoop2; /* 2.4.102 -- Explicitly says to ignore */
var parse_EOF = parsenoop2; /* 2.4.103 */
var parse_Excel9File = parsenoop2; /* 2.4.104 -- Optional and unused */
var parse_FeatHdr = parsenoop2; /* 2.4.112 */
var parse_FontX = parseuint16; /* 2.4.123 */
var parse_Footer = parse_XLHeaderFooter; /* 2.4.124 */
var parse_GridSet = parseuint16; /* 2.4.132, =1 */
var parse_HCenter = parsebool; /* 2.4.135 sheet centered horizontal on print */
var parse_Header = parse_XLHeaderFooter; /* 2.4.136 */
var parse_HideObj = parse_HideObjEnum; /* 2.4.139 */
var parse_InterfaceEnd = parsenoop2; /* 2.4.145 -- noop */
var parse_LeftMargin = parse_Xnum; /* 2.4.151 */
var parse_Mms = parsenoop2; /* 2.4.169 -- Explicitly says to ignore */
var parse_ObjProtect = parsebool; /* 2.4.183 -- must be 1 if present */
var parse_Password = parseuint16; /* 2.4.191 */
var parse_PrintGrid = parsebool; /* 2.4.202 */
var parse_PrintRowCol = parsebool; /* 2.4.203 */
var parse_PrintSize = parseuint16; /* 2.4.204 0:3 */
var parse_Prot4Rev = parsebool; /* 2.4.205 */
var parse_Prot4RevPass = parseuint16; /* 2.4.206 */
var parse_Protect = parsebool; /* 2.4.207 */
var parse_RefreshAll = parsebool; /* 2.4.217 -- must be 0 if not template */
var parse_RightMargin = parse_Xnum; /* 2.4.219 */
var parse_RRTabId = parseuint16a; /* 2.4.241 */
var parse_ScenarioProtect = parsebool; /* 2.4.245 */
var parse_Scl = parseuint16a; /* 2.4.247 num, den */
var parse_String = parse_XLUnicodeString; /* 2.4.268 */
var parse_SxBool = parsebool; /* 2.4.274 */
var parse_TopMargin = parse_Xnum; /* 2.4.328 */
var parse_UsesELFs = parsebool; /* 2.4.337 -- should be 0 */
var parse_VCenter = parsebool; /* 2.4.342 */
var parse_WinProtect = parsebool; /* 2.4.347 */
var parse_WriteProtect = parsenoop; /* 2.4.350 empty record */


/* ---- */
var parse_VerticalPageBreaks = parsenoop;
var parse_HorizontalPageBreaks = parsenoop;
var parse_Selection = parsenoop;
var parse_Continue = parsenoop;
var parse_Pane = parsenoop;
var parse_Pls = parsenoop;
var parse_DCon = parsenoop;
var parse_DConRef = parsenoop;
var parse_DConName = parsenoop;
var parse_XCT = parsenoop;
var parse_CRN = parsenoop;
var parse_FileSharing = parsenoop;
var parse_Uncalced = parsenoop;
var parse_Template = parsenoop;
var parse_Intl = parsenoop;
var parse_WsBool = parsenoop;
var parse_Sort = parsenoop;
var parse_Sync = parsenoop;
var parse_LPr = parsenoop;
var parse_DxGCol = parsenoop;
var parse_FnGroupName = parsenoop;
var parse_FilterMode = parsenoop;
var parse_AutoFilterInfo = parsenoop;
var parse_AutoFilter = parsenoop;
var parse_ScenMan = parsenoop;
var parse_SCENARIO = parsenoop;
var parse_SxView = parsenoop;
var parse_Sxvd = parsenoop;
var parse_SXVI = parsenoop;
var parse_SxIvd = parsenoop;
var parse_SXLI = parsenoop;
var parse_SXPI = parsenoop;
var parse_DocRoute = parsenoop;
var parse_RecipName = parsenoop;
var parse_SXDI = parsenoop;
var parse_SXDB = parsenoop;
var parse_SXFDB = parsenoop;
var parse_SXDBB = parsenoop;
var parse_SXNum = parsenoop;
var parse_SxErr = parsenoop;
var parse_SXInt = parsenoop;
var parse_SXString = parsenoop;
var parse_SXDtr = parsenoop;
var parse_SxNil = parsenoop;
var parse_SXTbl = parsenoop;
var parse_SXTBRGIITM = parsenoop;
var parse_SxTbpg = parsenoop;
var parse_ObProj = parsenoop;
var parse_SXStreamID = parsenoop;
var parse_DBCell = parsenoop;
var parse_SXRng = parsenoop;
var parse_SxIsxoper = parsenoop;
var parse_BookBool = parsenoop;
var parse_DbOrParamQry = parsenoop;
var parse_OleObjectSize = parsenoop;
var parse_SXVS = parsenoop;
var parse_BkHim = parsenoop;
var parse_MsoDrawingGroup = parsenoop;
var parse_MsoDrawing = parsenoop;
var parse_MsoDrawingSelection = parsenoop;
var parse_PhoneticInfo = parsenoop;
var parse_SxRule = parsenoop;
var parse_SXEx = parsenoop;
var parse_SxFilt = parsenoop;
var parse_SxDXF = parsenoop;
var parse_SxItm = parsenoop;
var parse_SxName = parsenoop;
var parse_SxSelect = parsenoop;
var parse_SXPair = parsenoop;
var parse_SxFmla = parsenoop;
var parse_SxFormat = parsenoop;
var parse_SXVDEx = parsenoop;
var parse_SXFormula = parsenoop;
var parse_SXDBEx = parsenoop;
var parse_RRDInsDel = parsenoop;
var parse_RRDHead = parsenoop;
var parse_RRDChgCell = parsenoop;
var parse_RRDRenSheet = parsenoop;
var parse_RRSort = parsenoop;
var parse_RRDMove = parsenoop;
var parse_RRFormat = parsenoop;
var parse_RRAutoFmt = parsenoop;
var parse_RRInsertSh = parsenoop;
var parse_RRDMoveBegin = parsenoop;
var parse_RRDMoveEnd = parsenoop;
var parse_RRDInsDelBegin = parsenoop;
var parse_RRDInsDelEnd = parsenoop;
var parse_RRDConflict = parsenoop;
var parse_RRDDefName = parsenoop;
var parse_RRDRstEtxp = parsenoop;
var parse_LRng = parsenoop;
var parse_CUsr = parsenoop;
var parse_CbUsr = parsenoop;
var parse_UsrInfo = parsenoop;
var parse_UsrExcl = parsenoop;
var parse_FileLock = parsenoop;
var parse_RRDInfo = parsenoop;
var parse_BCUsrs = parsenoop;
var parse_UsrChk = parsenoop;
var parse_UserBView = parsenoop;
var parse_UserSViewBegin = parsenoop; // overloaded
var parse_UserSViewEnd = parsenoop;
var parse_RRDUserView = parsenoop;
var parse_Qsi = parsenoop;
var parse_CondFmt = parsenoop;
var parse_CF = parsenoop;
var parse_DVal = parsenoop;
var parse_DConBin = parsenoop;
var parse_Lel = parsenoop;
var parse_XLSCodeName = parse_XLUnicodeString;
var parse_SXFDBType = parsenoop;
var parse_ObNoMacros = parsenoop;
var parse_Dv = parsenoop;
var parse_Index = parsenoop;
var parse_Table = parsenoop;
var parse_BigName = parsenoop;
var parse_ContinueBigName = parsenoop;
var parse_WebPub = parsenoop;
var parse_QsiSXTag = parsenoop;
var parse_DBQueryExt = parsenoop;
var parse_ExtString = parsenoop;
var parse_TxtQry = parsenoop;
var parse_Qsir = parsenoop;
var parse_Qsif = parsenoop;
var parse_RRDTQSIF = parsenoop;
var parse_OleDbConn = parsenoop;
var parse_WOpt = parsenoop;
var parse_SXViewEx = parsenoop;
var parse_SXTH = parsenoop;
var parse_SXPIEx = parsenoop;
var parse_SXVDTEx = parsenoop;
var parse_SXViewEx9 = parsenoop;
var parse_ContinueFrt = parsenoop;
var parse_RealTimeData = parsenoop;
var parse_ChartFrtInfo = parsenoop;
var parse_FrtWrapper = parsenoop;
var parse_StartBlock = parsenoop;
var parse_EndBlock = parsenoop;
var parse_StartObject = parsenoop;
var parse_EndObject = parsenoop;
var parse_CatLab = parsenoop;
var parse_YMult = parsenoop;
var parse_SXViewLink = parsenoop;
var parse_PivotChartBits = parsenoop;
var parse_FrtFontList = parsenoop;
var parse_SheetExt = parsenoop;
var parse_BookExt = parsenoop;
var parse_SXAddl = parsenoop;
var parse_CrErr = parsenoop;
var parse_HFPicture = parsenoop;
var parse_Feat = parsenoop;
var parse_DataLabExt = parsenoop;
var parse_DataLabExtContents = parsenoop;
var parse_CellWatch = parsenoop;
var parse_FeatHdr11 = parsenoop;
var parse_Feature11 = parsenoop;
var parse_DropDownObjIds = parsenoop;
var parse_ContinueFrt11 = parsenoop;
var parse_DConn = parsenoop;
var parse_List12 = parsenoop;
var parse_Feature12 = parsenoop;
var parse_CondFmt12 = parsenoop;
var parse_CF12 = parsenoop;
var parse_CFEx = parsenoop;
var parse_AutoFilter12 = parsenoop;
var parse_ContinueFrt12 = parsenoop;
var parse_MDTInfo = parsenoop;
var parse_MDXStr = parsenoop;
var parse_MDXTuple = parsenoop;
var parse_MDXSet = parsenoop;
var parse_MDXProp = parsenoop;
var parse_MDXKPI = parsenoop;
var parse_MDB = parsenoop;
var parse_PLV = parsenoop;
var parse_DXF = parsenoop;
var parse_TableStyles = parsenoop;
var parse_TableStyle = parsenoop;
var parse_TableStyleElement = parsenoop;
var parse_NamePublish = parsenoop;
var parse_SortData = parsenoop;
var parse_GUIDTypeLib = parsenoop;
var parse_FnGrp12 = parsenoop;
var parse_NameFnGrp12 = parsenoop;
var parse_HeaderFooter = parsenoop;
var parse_CrtLayout12 = parsenoop;
var parse_CrtMlFrt = parsenoop;
var parse_CrtMlFrtContinue = parsenoop;
var parse_ShapePropsStream = parsenoop;
var parse_TextPropsStream = parsenoop;
var parse_RichTextStream = parsenoop;
var parse_CrtLayout12A = parsenoop;
var parse_Units = parsenoop;
var parse_Chart = parsenoop;
var parse_Series = parsenoop;
var parse_DataFormat = parsenoop;
var parse_LineFormat = parsenoop;
var parse_MarkerFormat = parsenoop;
var parse_AreaFormat = parsenoop;
var parse_PieFormat = parsenoop;
var parse_AttachedLabel = parsenoop;
var parse_SeriesText = parsenoop;
var parse_ChartFormat = parsenoop;
var parse_Legend = parsenoop;
var parse_SeriesList = parsenoop;
var parse_Bar = parsenoop;
var parse_Line = parsenoop;
var parse_Pie = parsenoop;
var parse_Area = parsenoop;
var parse_Scatter = parsenoop;
var parse_CrtLine = parsenoop;
var parse_Axis = parsenoop;
var parse_Tick = parsenoop;
var parse_ValueRange = parsenoop;
var parse_CatSerRange = parsenoop;
var parse_AxisLine = parsenoop;
var parse_CrtLink = parsenoop;
var parse_DefaultText = parsenoop;
var parse_Text = parsenoop;
var parse_ObjectLink = parsenoop;
var parse_Frame = parsenoop;
var parse_Begin = parsenoop;
var parse_End = parsenoop;
var parse_PlotArea = parsenoop;
var parse_Chart3d = parsenoop;
var parse_PicF = parsenoop;
var parse_DropBar = parsenoop;
var parse_Radar = parsenoop;
var parse_Surf = parsenoop;
var parse_RadarArea = parsenoop;
var parse_AxisParent = parsenoop;
var parse_LegendException = parsenoop;
var parse_SerToCrt = parsenoop;
var parse_AxesUsed = parsenoop;
var parse_SBaseRef = parsenoop;
var parse_SerParent = parsenoop;
var parse_SerAuxTrend = parsenoop;
var parse_IFmtRecord = parsenoop;
var parse_Pos = parsenoop;
var parse_AlRuns = parsenoop;
var parse_BRAI = parsenoop;
var parse_SerAuxErrBar = parsenoop;
var parse_SerFmt = parsenoop;
var parse_Chart3DBarShape = parsenoop;
var parse_Fbi = parsenoop;
var parse_BopPop = parsenoop;
var parse_AxcExt = parsenoop;
var parse_Dat = parsenoop;
var parse_PlotGrowth = parsenoop;
var parse_SIIndex = parsenoop;
var parse_GelFrame = parsenoop;
var parse_BopPopCustom = parsenoop;
var parse_Fbi2 = parsenoop;

/* --- Specific to versions before BIFF8 --- */
function parse_ImData(blob, length, opts) {
	var tgt = blob.l + length;
	var cf = blob.read_shift(2);
	var env = blob.read_shift(2);
	var lcb = blob.read_shift(4);
	var o = {fmt:cf, env:env, len:lcb, data:blob.slice(blob.l,blob.l+lcb)};
	blob.l += lcb;
	return o;
}

function parse_BIFF5String(blob) {
	var len = blob.read_shift(1);
	return blob.read_shift(len, 'sbcs-cont');
}

/* BIFF2_??? where ??? is the name from [XLS] */
function parse_BIFF2STR(blob, length, opts) {
	var cell = parse_XLSCell(blob, 6);
	++blob.l;
	var str = parse_XLUnicodeString2(blob, length-7, opts);
	cell.t = 'str';
	cell.val = str;
	return cell;
}

function parse_BIFF2NUM(blob, length, opts) {
	var cell = parse_XLSCell(blob, 6);
	++blob.l;
	var num = parse_Xnum(blob, 8);
	cell.t = 'n';
	cell.val = num;
	return cell;
}

function parse_BIFF2INT(blob, length) {
	var cell = parse_XLSCell(blob, 6);
	++blob.l;
	var num = blob.read_shift(2);
	cell.t = 'n';
	cell.val = num;
	return cell;
}

function parse_BIFF2STRING(blob, length) {
	var cch = blob.read_shift(1);
	if(cch === 0) { blob.l++; return ""; }
	return blob.read_shift(cch, 'sbcs-cont');
}

/* TODO: convert to BIFF8 font struct */
function parse_BIFF2FONTXTRA(blob, length) {
	blob.l += 6; // unknown
	blob.l += 2; // font weight "bls"
	blob.l += 1; // charset
	blob.l += 3; // unknown
	blob.l += 1; // font family
	blob.l += length - 13;
}

/* TODO: parse rich text runs */
function parse_RString(blob, length, opts) {
	var end = blob.l + length;
	var cell = parse_XLSCell(blob, 6);
	var cch = blob.read_shift(2);
	var str = parse_XLUnicodeStringNoCch(blob, cch, opts);
	blob.l = end;
	cell.t = 'str';
	cell.val = str;
	return cell;
}
