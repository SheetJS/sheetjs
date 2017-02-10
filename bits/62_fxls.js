/* --- formula references point to MS-XLS --- */
/* Small helpers */
function parseread(l) { return function(blob, length) { blob.l+=l; return; }; }
function parseread1(blob, length) { blob.l+=1; return; }

/* Rgce Helpers */

/* 2.5.51 */
function parse_ColRelU(blob, length) {
	var c = blob.read_shift(2);
	return [c & 0x3FFF, (c >> 14) & 1, (c >> 15) & 1];
}

/* 2.5.198.105 */
function parse_RgceArea(blob, length) {
	var r=blob.read_shift(2), R=blob.read_shift(2);
	var c=parse_ColRelU(blob, 2);
	var C=parse_ColRelU(blob, 2);
	return { s:{r:r, c:c[0], cRel:c[1], rRel:c[2]}, e:{r:R, c:C[0], cRel:C[1], rRel:C[2]} };
}

/* 2.5.198.105 TODO */
function parse_RgceAreaRel(blob, length) {
	var r=blob.read_shift(2), R=blob.read_shift(2);
	var c=parse_ColRelU(blob, 2);
	var C=parse_ColRelU(blob, 2);
	return { s:{r:r, c:c[0], cRel:c[1], rRel:c[2]}, e:{r:R, c:C[0], cRel:C[1], rRel:C[2]} };
}

/* 2.5.198.109 */
function parse_RgceLoc(blob, length) {
	var r = blob.read_shift(2);
	var c = parse_ColRelU(blob, 2);
	return {r:r, c:c[0], cRel:c[1], rRel:c[2]};
}

/* 2.5.198.111 */
function parse_RgceLocRel(blob, length) {
	var r = blob.read_shift(2);
	var cl = blob.read_shift(2);
	var cRel = (cl & 0x8000) >> 15, rRel = (cl & 0x4000) >> 14;
	cl &= 0x3FFF;
	if(cRel !== 0) while(cl >= 0x100) cl -= 0x100;
	return {r:r,c:cl,cRel:cRel,rRel:rRel};
}

/* Ptg Tokens */

/* 2.5.198.27 */
function parse_PtgArea(blob, length) {
	var type = (blob[blob.l++] & 0x60) >> 5;
	var area = parse_RgceArea(blob, 8);
	return [type, area];
}

/* 2.5.198.28 */
function parse_PtgArea3d(blob, length) {
	var type = (blob[blob.l++] & 0x60) >> 5;
	var ixti = blob.read_shift(2);
	var area = parse_RgceArea(blob, 8);
	return [type, ixti, area];
}

/* 2.5.198.29 */
function parse_PtgAreaErr(blob, length) {
	var type = (blob[blob.l++] & 0x60) >> 5;
	blob.l += 8;
	return [type];
}
/* 2.5.198.30 */
function parse_PtgAreaErr3d(blob, length) {
	var type = (blob[blob.l++] & 0x60) >> 5;
	var ixti = blob.read_shift(2);
	blob.l += 8;
	return [type, ixti];
}

/* 2.5.198.31 */
function parse_PtgAreaN(blob, length) {
	var type = (blob[blob.l++] & 0x60) >> 5;
	var area = parse_RgceAreaRel(blob, 8);
	return [type, area];
}

/* 2.5.198.32 -- ignore this and look in PtgExtraArray for shape + values */
function parse_PtgArray(blob, length) {
	var type = (blob[blob.l++] & 0x60) >> 5;
	blob.l += 7;
	return [type];
}

/* 2.5.198.33 */
function parse_PtgAttrBaxcel(blob, length) {
	var bitSemi = blob[blob.l+1] & 0x01; /* 1 = volatile */
	var bitBaxcel = 1;
	blob.l += 4;
	return [bitSemi, bitBaxcel];
}

/* 2.5.198.34 */
function parse_PtgAttrChoose(blob, length) {
	blob.l +=2;
	var offset = blob.read_shift(2);
	var o = [];
	/* offset is 1 less than the number of elements */
	for(var i = 0; i <= offset; ++i) o.push(blob.read_shift(2));
	return o;
}

/* 2.5.198.35 */
function parse_PtgAttrGoto(blob, length) {
	var bitGoto = (blob[blob.l+1] & 0xFF) ? 1 : 0;
	blob.l += 2;
	return [bitGoto, blob.read_shift(2)];
}

/* 2.5.198.36 */
function parse_PtgAttrIf(blob, length) {
	var bitIf = (blob[blob.l+1] & 0xFF) ? 1 : 0;
	blob.l += 2;
	return [bitIf, blob.read_shift(2)];
}

/* 2.5.198.37 */
function parse_PtgAttrSemi(blob, length) {
	var bitSemi = (blob[blob.l+1] & 0xFF) ? 1 : 0;
	blob.l += 4;
	return [bitSemi];
}

/* 2.5.198.40 (used by PtgAttrSpace and PtgAttrSpaceSemi) */
function parse_PtgAttrSpaceType(blob, length) {
	var type = blob.read_shift(1), cch = blob.read_shift(1);
	return [type, cch];
}

/* 2.5.198.38 */
function parse_PtgAttrSpace(blob, length) {
	blob.read_shift(2);
	return parse_PtgAttrSpaceType(blob, 2);
}

/* 2.5.198.39 */
function parse_PtgAttrSpaceSemi(blob, length) {
	blob.read_shift(2);
	return parse_PtgAttrSpaceType(blob, 2);
}

/* 2.5.198.84 TODO */
function parse_PtgRef(blob, length) {
	var ptg = blob[blob.l] & 0x1F;
	var type = (blob[blob.l] & 0x60)>>5;
	blob.l += 1;
	var loc = parse_RgceLoc(blob,4);
	return [type, loc];
}

/* 2.5.198.88 TODO */
function parse_PtgRefN(blob, length) {
	var ptg = blob[blob.l] & 0x1F;
	var type = (blob[blob.l] & 0x60)>>5;
	blob.l += 1;
	var loc = parse_RgceLocRel(blob,4);
	return [type, loc];
}

/* 2.5.198.85 TODO */
function parse_PtgRef3d(blob, length) {
	var ptg = blob[blob.l] & 0x1F;
	var type = (blob[blob.l] & 0x60)>>5;
	blob.l += 1;
	var ixti = blob.read_shift(2); // XtiIndex
	var loc = parse_RgceLoc(blob,4);
	return [type, ixti, loc];
}


/* 2.5.198.62 TODO */
function parse_PtgFunc(blob, length) {
	var ptg = blob[blob.l] & 0x1F;
	var type = (blob[blob.l] & 0x60)>>5;
	blob.l += 1;
	var iftab = blob.read_shift(2);
	return [FtabArgc[iftab], Ftab[iftab]];
}
/* 2.5.198.63 TODO */
function parse_PtgFuncVar(blob, length) {
	blob.l++;
	var cparams = blob.read_shift(1), tab = parsetab(blob);
	return [cparams, (tab[0] === 0 ? Ftab : Cetab)[tab[1]]];
}

function parsetab(blob, length) {
	return [blob[blob.l+1]>>7, blob.read_shift(2) & 0x7FFF];
}

/* 2.5.198.41 */
var parse_PtgAttrSum = parseread(4);
/* 2.5.198.43 */
var parse_PtgConcat = parseread1;

/* 2.5.198.58 */
function parse_PtgExp(blob, length) {
	blob.l++;
	var row = blob.read_shift(2);
	var col = blob.read_shift(2);
	return [row, col];
}

/* 2.5.198.57 */
function parse_PtgErr(blob, length) { blob.l++; return BErr[blob.read_shift(1)]; }

/* 2.5.198.66 TODO */
function parse_PtgInt(blob, length) { blob.l++; return blob.read_shift(2); }

/* 2.5.198.42 */
function parse_PtgBool(blob, length) { blob.l++; return blob.read_shift(1)!==0;}

/* 2.5.198.79 */
function parse_PtgNum(blob, length) { blob.l++; return parse_Xnum(blob, 8); }

/* 2.5.198.89 */
function parse_PtgStr(blob, length) { blob.l++; return parse_ShortXLUnicodeString(blob); }

/* 2.5.192.112 + 2.5.192.11{3,4,5,6,7} */
function parse_SerAr(blob) {
	var val = [];
	switch((val[0] = blob.read_shift(1))) {
		/* 2.5.192.113 */
		case 0x04: /* SerBool -- boolean */
			val[1] = parsebool(blob, 1) ? 'TRUE' : 'FALSE';
			blob.l += 7; break;
		/* 2.5.192.114 */
		case 0x10: /* SerErr -- error */
			val[1] = BErr[blob[blob.l]];
			blob.l += 8; break;
		/* 2.5.192.115 */
		case 0x00: /* SerNil -- honestly, I'm not sure how to reproduce this */
			blob.l += 8; break;
		/* 2.5.192.116 */
		case 0x01: /* SerNum -- Xnum */
			val[1] = parse_Xnum(blob, 8); break;
		/* 2.5.192.117 */
		case 0x02: /* SerStr -- XLUnicodeString (<256 chars) */
			val[1] = parse_XLUnicodeString(blob); break;
		// default: throw "Bad SerAr: " + val[0]; /* Unreachable */
	}
	return val;
}

/* 2.5.198.61 */
function parse_PtgExtraMem(blob, cce) {
	var count = blob.read_shift(2);
	var out = [];
	for(var i = 0; i != count; ++i) out.push(parse_Ref8U(blob, 8));
	return out;
}

/* 2.5.198.59 */
function parse_PtgExtraArray(blob) {
	var cols = 1 + blob.read_shift(1); //DColByteU
	var rows = 1 + blob.read_shift(2); //DRw
	for(var i = 0, o=[]; i != rows && (o[i] = []); ++i)
		for(var j = 0; j != cols; ++j) o[i][j] = parse_SerAr(blob);
	return o;
}

/* 2.5.198.76 */
function parse_PtgName(blob, length) {
	var type = (blob.read_shift(1) >>> 5) & 0x03;
	var nameindex = blob.read_shift(4);
	return [type, 0, nameindex];
}

/* 2.5.198.77 */
function parse_PtgNameX(blob, length) {
	var type = (blob.read_shift(1) >>> 5) & 0x03;
	var ixti = blob.read_shift(2); // XtiIndex
	var nameindex = blob.read_shift(4);
	return [type, ixti, nameindex];
}

/* 2.5.198.70 */
function parse_PtgMemArea(blob, length) {
	var type = (blob.read_shift(1) >>> 5) & 0x03;
	blob.l += 4;
	var cce = blob.read_shift(2);
	return [type, cce];
}

/* 2.5.198.72 */
function parse_PtgMemFunc(blob, length) {
	var type = (blob.read_shift(1) >>> 5) & 0x03;
	var cce = blob.read_shift(2);
	return [type, cce];
}


/* 2.5.198.86 */
function parse_PtgRefErr(blob, length) {
	var type = (blob.read_shift(1) >>> 5) & 0x03;
	blob.l += 4;
	return [type];
}

/* 2.5.198.26 */
var parse_PtgAdd = parseread1;
/* 2.5.198.45 */
var parse_PtgDiv = parseread1;
/* 2.5.198.56 */
var parse_PtgEq = parseread1;
/* 2.5.198.64 */
var parse_PtgGe = parseread1;
/* 2.5.198.65 */
var parse_PtgGt = parseread1;
/* 2.5.198.67 */
var parse_PtgIsect = parseread1;
/* 2.5.198.68 */
var parse_PtgLe = parseread1;
/* 2.5.198.69 */
var parse_PtgLt = parseread1;
/* 2.5.198.74 */
var parse_PtgMissArg = parseread1;
/* 2.5.198.75 */
var parse_PtgMul = parseread1;
/* 2.5.198.78 */
var parse_PtgNe = parseread1;
/* 2.5.198.80 */
var parse_PtgParen = parseread1;
/* 2.5.198.81 */
var parse_PtgPercent = parseread1;
/* 2.5.198.82 */
var parse_PtgPower = parseread1;
/* 2.5.198.83 */
var parse_PtgRange = parseread1;
/* 2.5.198.90 */
var parse_PtgSub = parseread1;
/* 2.5.198.93 */
var parse_PtgUminus = parseread1;
/* 2.5.198.94 */
var parse_PtgUnion = parseread1;
/* 2.5.198.95 */
var parse_PtgUplus = parseread1;

/* 2.5.198.71 */
var parse_PtgMemErr = parsenoop;
/* 2.5.198.73 */
var parse_PtgMemNoMem = parsenoop;
/* 2.5.198.87 */
var parse_PtgRefErr3d = parsenoop;
/* 2.5.198.92 */
var parse_PtgTbl = parsenoop;

/* 2.5.198.25 */
var PtgTypes = {
	/*::[*/0x01/*::]*/: { n:'PtgExp', f:parse_PtgExp },
	/*::[*/0x02/*::]*/: { n:'PtgTbl', f:parse_PtgTbl },
	/*::[*/0x03/*::]*/: { n:'PtgAdd', f:parse_PtgAdd },
	/*::[*/0x04/*::]*/: { n:'PtgSub', f:parse_PtgSub },
	/*::[*/0x05/*::]*/: { n:'PtgMul', f:parse_PtgMul },
	/*::[*/0x06/*::]*/: { n:'PtgDiv', f:parse_PtgDiv },
	/*::[*/0x07/*::]*/: { n:'PtgPower', f:parse_PtgPower },
	/*::[*/0x08/*::]*/: { n:'PtgConcat', f:parse_PtgConcat },
	/*::[*/0x09/*::]*/: { n:'PtgLt', f:parse_PtgLt },
	/*::[*/0x0A/*::]*/: { n:'PtgLe', f:parse_PtgLe },
	/*::[*/0x0B/*::]*/: { n:'PtgEq', f:parse_PtgEq },
	/*::[*/0x0C/*::]*/: { n:'PtgGe', f:parse_PtgGe },
	/*::[*/0x0D/*::]*/: { n:'PtgGt', f:parse_PtgGt },
	/*::[*/0x0E/*::]*/: { n:'PtgNe', f:parse_PtgNe },
	/*::[*/0x0F/*::]*/: { n:'PtgIsect', f:parse_PtgIsect },
	/*::[*/0x10/*::]*/: { n:'PtgUnion', f:parse_PtgUnion },
	/*::[*/0x11/*::]*/: { n:'PtgRange', f:parse_PtgRange },
	/*::[*/0x12/*::]*/: { n:'PtgUplus', f:parse_PtgUplus },
	/*::[*/0x13/*::]*/: { n:'PtgUminus', f:parse_PtgUminus },
	/*::[*/0x14/*::]*/: { n:'PtgPercent', f:parse_PtgPercent },
	/*::[*/0x15/*::]*/: { n:'PtgParen', f:parse_PtgParen },
	/*::[*/0x16/*::]*/: { n:'PtgMissArg', f:parse_PtgMissArg },
	/*::[*/0x17/*::]*/: { n:'PtgStr', f:parse_PtgStr },
	/*::[*/0x1C/*::]*/: { n:'PtgErr', f:parse_PtgErr },
	/*::[*/0x1D/*::]*/: { n:'PtgBool', f:parse_PtgBool },
	/*::[*/0x1E/*::]*/: { n:'PtgInt', f:parse_PtgInt },
	/*::[*/0x1F/*::]*/: { n:'PtgNum', f:parse_PtgNum },
	/*::[*/0x20/*::]*/: { n:'PtgArray', f:parse_PtgArray },
	/*::[*/0x21/*::]*/: { n:'PtgFunc', f:parse_PtgFunc },
	/*::[*/0x22/*::]*/: { n:'PtgFuncVar', f:parse_PtgFuncVar },
	/*::[*/0x23/*::]*/: { n:'PtgName', f:parse_PtgName },
	/*::[*/0x24/*::]*/: { n:'PtgRef', f:parse_PtgRef },
	/*::[*/0x25/*::]*/: { n:'PtgArea', f:parse_PtgArea },
	/*::[*/0x26/*::]*/: { n:'PtgMemArea', f:parse_PtgMemArea },
	/*::[*/0x27/*::]*/: { n:'PtgMemErr', f:parse_PtgMemErr },
	/*::[*/0x28/*::]*/: { n:'PtgMemNoMem', f:parse_PtgMemNoMem },
	/*::[*/0x29/*::]*/: { n:'PtgMemFunc', f:parse_PtgMemFunc },
	/*::[*/0x2A/*::]*/: { n:'PtgRefErr', f:parse_PtgRefErr },
	/*::[*/0x2B/*::]*/: { n:'PtgAreaErr', f:parse_PtgAreaErr },
	/*::[*/0x2C/*::]*/: { n:'PtgRefN', f:parse_PtgRefN },
	/*::[*/0x2D/*::]*/: { n:'PtgAreaN', f:parse_PtgAreaN },
	/*::[*/0x39/*::]*/: { n:'PtgNameX', f:parse_PtgNameX },
	/*::[*/0x3A/*::]*/: { n:'PtgRef3d', f:parse_PtgRef3d },
	/*::[*/0x3B/*::]*/: { n:'PtgArea3d', f:parse_PtgArea3d },
	/*::[*/0x3C/*::]*/: { n:'PtgRefErr3d', f:parse_PtgRefErr3d },
	/*::[*/0x3D/*::]*/: { n:'PtgAreaErr3d', f:parse_PtgAreaErr3d },
	/*::[*/0xFF/*::]*/: {}
};
/* These are duplicated in the PtgTypes table */
var PtgDupes = {
	/*::[*/0x40/*::]*/: 0x20, /*::[*/0x60/*::]*/: 0x20,
	/*::[*/0x41/*::]*/: 0x21, /*::[*/0x61/*::]*/: 0x21,
	/*::[*/0x42/*::]*/: 0x22, /*::[*/0x62/*::]*/: 0x22,
	/*::[*/0x43/*::]*/: 0x23, /*::[*/0x63/*::]*/: 0x23,
	/*::[*/0x44/*::]*/: 0x24, /*::[*/0x64/*::]*/: 0x24,
	/*::[*/0x45/*::]*/: 0x25, /*::[*/0x65/*::]*/: 0x25,
	/*::[*/0x46/*::]*/: 0x26, /*::[*/0x66/*::]*/: 0x26,
	/*::[*/0x47/*::]*/: 0x27, /*::[*/0x67/*::]*/: 0x27,
	/*::[*/0x48/*::]*/: 0x28, /*::[*/0x68/*::]*/: 0x28,
	/*::[*/0x49/*::]*/: 0x29, /*::[*/0x69/*::]*/: 0x29,
	/*::[*/0x4A/*::]*/: 0x2A, /*::[*/0x6A/*::]*/: 0x2A,
	/*::[*/0x4B/*::]*/: 0x2B, /*::[*/0x6B/*::]*/: 0x2B,
	/*::[*/0x4C/*::]*/: 0x2C, /*::[*/0x6C/*::]*/: 0x2C,
	/*::[*/0x4D/*::]*/: 0x2D, /*::[*/0x6D/*::]*/: 0x2D,
	/*::[*/0x59/*::]*/: 0x39, /*::[*/0x79/*::]*/: 0x39,
	/*::[*/0x5A/*::]*/: 0x3A, /*::[*/0x7A/*::]*/: 0x3A,
	/*::[*/0x5B/*::]*/: 0x3B, /*::[*/0x7B/*::]*/: 0x3B,
	/*::[*/0x5C/*::]*/: 0x3C, /*::[*/0x7C/*::]*/: 0x3C,
	/*::[*/0x5D/*::]*/: 0x3D, /*::[*/0x7D/*::]*/: 0x3D
};
(function(){for(var y in PtgDupes) PtgTypes[y] = PtgTypes[PtgDupes[y]];})();

var Ptg18 = {};
var Ptg19 = {
	/*::[*/0x01/*::]*/: { n:'PtgAttrSemi', f:parse_PtgAttrSemi },
	/*::[*/0x02/*::]*/: { n:'PtgAttrIf', f:parse_PtgAttrIf },
	/*::[*/0x04/*::]*/: { n:'PtgAttrChoose', f:parse_PtgAttrChoose },
	/*::[*/0x08/*::]*/: { n:'PtgAttrGoto', f:parse_PtgAttrGoto },
	/*::[*/0x10/*::]*/: { n:'PtgAttrSum', f:parse_PtgAttrSum },
	/*::[*/0x20/*::]*/: { n:'PtgAttrBaxcel', f:parse_PtgAttrBaxcel },
	/*::[*/0x40/*::]*/: { n:'PtgAttrSpace', f:parse_PtgAttrSpace },
	/*::[*/0x41/*::]*/: { n:'PtgAttrSpaceSemi', f:parse_PtgAttrSpaceSemi },
	/*::[*/0xFF/*::]*/: {}
};

/* 2.4.127 TODO */
function parse_Formula(blob, length, opts) {
	var cell = parse_XLSCell(blob, 6);
	if(opts.biff == 2) ++blob.l;
	var val = parse_FormulaValue(blob,8);
	var flags = blob.read_shift(1);
	if(opts.biff != 2) {
		blob.read_shift(1);
		if(opts.biff >= 5) {
			var chn = blob.read_shift(4);
		}
	}
	var cbf = "";
	if(opts.biff < 5) blob.l += length-16;
	else if(opts.biff === 5) blob.l += length-20;
	else cbf = parse_XLSCellParsedFormula(blob, length-20, opts);
	return {cell:cell, val:val[0], formula:cbf, shared: (flags >> 3) & 1, tt:val[1]};
}

/* 2.5.133 TODO: how to emit empty strings? */
function parse_FormulaValue(blob) {
	var b;
	if(__readUInt16LE(blob,blob.l + 6) !== 0xFFFF) return [parse_Xnum(blob),'n'];
	switch(blob[blob.l]) {
		case 0x00: blob.l += 8; return ["String", 's'];
		case 0x01: b = blob[blob.l+2] === 0x1; blob.l += 8; return [b,'b'];
		case 0x02: b = blob[blob.l+2]; blob.l += 8; return [b,'e'];
		case 0x03: blob.l += 8; return ["",'s'];
	}
	return [];
}

/* 2.5.198.103 */
function parse_RgbExtra(blob, length, rgce, opts) {
	if(opts.biff < 8) return parsenoop(blob, length);
	var target = blob.l + length;
	var o = [];
	for(var i = 0; i !== rgce.length; ++i) {
		switch(rgce[i][0]) {
			case 'PtgArray': /* PtgArray -> PtgExtraArray */
				rgce[i][1] = parse_PtgExtraArray(blob);
				o.push(rgce[i][1]);
				break;
			case 'PtgMemArea': /* PtgMemArea -> PtgExtraMem */
				rgce[i][2] = parse_PtgExtraMem(blob, rgce[i][1]);
				o.push(rgce[i][2]);
				break;
			default: break;
		}
	}
	length = target - blob.l;
	if(length !== 0) o.push(parsenoop(blob, length));
	return o;
}

/* 2.5.198.21 */
function parse_NameParsedFormula(blob, length, opts, cce) {
	var target = blob.l + length;
	var rgce = parse_Rgce(blob, cce);
	var rgcb;
	if(target !== blob.l) rgcb = parse_RgbExtra(blob, target - blob.l, rgce, opts);
	return [rgce, rgcb];
}

/* 2.5.198.3 TODO */
function parse_XLSCellParsedFormula(blob, length, opts) {
	var target = blob.l + length;
	var rgcb, cce = blob.read_shift(2); // length of rgce
	if(cce == 0xFFFF) return [[],parsenoop(blob, length-2)];
	var rgce = parse_Rgce(blob, cce);
	if(length !== cce + 2) rgcb = parse_RgbExtra(blob, length - cce - 2, rgce, opts);
	return [rgce, rgcb];
}

/* 2.5.198.118 TODO */
function parse_SharedParsedFormula(blob, length, opts) {
	var target = blob.l + length;
	var rgcb, cce = blob.read_shift(2); // length of rgce
	var rgce = parse_Rgce(blob, cce);
	if(cce == 0xFFFF) return [[],parsenoop(blob, length-2)];
	if(length !== cce + 2) rgcb = parse_RgbExtra(blob, target - cce - 2, rgce, opts);
	return [rgce, rgcb];
}

/* 2.5.198.1 TODO */
function parse_ArrayParsedFormula(blob, length, opts, ref) {
	var target = blob.l + length;
	var rgcb, cce = blob.read_shift(2); // length of rgce
	if(cce == 0xFFFF) return [[],parsenoop(blob, length-2)];
	var rgce = parse_Rgce(blob, cce);
	if(length !== cce + 2) rgcb = parse_RgbExtra(blob, target - cce - 2, rgce, opts);
	return [rgce, rgcb];
}

/* 2.5.198.104 */
function parse_Rgce(blob, length) {
	var target = blob.l + length;
	var R, id, ptgs = [];
	while(target != blob.l) {
		length = target - blob.l;
		id = blob[blob.l];
		R = PtgTypes[id];
		//console.log("ptg", id, R)
		if(id === 0x18 || id === 0x19) {
			id = blob[blob.l + 1];
			R = (id === 0x18 ? Ptg18 : Ptg19)[id];
		}
		if(!R || !R.f) { ptgs.push(parsenoop(blob, length)); }
		else { ptgs.push([R.n, R.f(blob, length)]); }
	}
	return ptgs;
}

function mapper(x) { return x.map(function f2(y) { return y[1];}).join(",");}

/* 2.2.2 + Magic TODO */
function stringify_formula(formula, range, cell, supbooks, opts) {
	if(opts !== undefined && opts.biff === 5) return "BIFF5??";
	var _range = range !== undefined ? range : {s:{c:0, r:0}};
	var stack = [], e1, e2, type, c, ixti, nameidx, r;
	if(!formula[0] || !formula[0][0]) return "";
	//console.log("--",cell,formula[0])
	for(var ff = 0, fflen = formula[0].length; ff < fflen; ++ff) {
		var f = formula[0][ff];
		//console.log("++",f, stack)
		switch(f[0]) {
		/* 2.2.2.1 Unary Operator Tokens */
			/* 2.5.198.93 */
			case 'PtgUminus': stack.push("-" + stack.pop()); break;
			/* 2.5.198.95 */
			case 'PtgUplus': stack.push("+" + stack.pop()); break;
			/* 2.5.198.81 */
			case 'PtgPercent': stack.push(stack.pop() + "%"); break;

		/* 2.2.2.1 Binary Value Operator Token */
			/* 2.5.198.26 */
			case 'PtgAdd':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+"+"+e1);
				break;
			/* 2.5.198.90 */
			case 'PtgSub':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+"-"+e1);
				break;
			/* 2.5.198.75 */
			case 'PtgMul':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+"*"+e1);
				break;
			/* 2.5.198.45 */
			case 'PtgDiv':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+"/"+e1);
				break;
			/* 2.5.198.82 */
			case 'PtgPower':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+"^"+e1);
				break;
			/* 2.5.198.43 */
			case 'PtgConcat':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+"&"+e1);
				break;
			/* 2.5.198.69 */
			case 'PtgLt':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+"<"+e1);
				break;
			/* 2.5.198.68 */
			case 'PtgLe':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+"<="+e1);
				break;
			/* 2.5.198.56 */
			case 'PtgEq':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+"="+e1);
				break;
			/* 2.5.198.64 */
			case 'PtgGe':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+">="+e1);
				break;
			/* 2.5.198.65 */
			case 'PtgGt':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+">"+e1);
				break;
			/* 2.5.198.78 */
			case 'PtgNe':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+"<>"+e1);
				break;

		/* 2.2.2.1 Binary Reference Operator Token */
			/* 2.5.198.67 */
			case 'PtgIsect':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+" "+e1);
				break;
			case 'PtgUnion':
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+","+e1);
				break;
			case 'PtgRange': break;

		/* 2.2.2.3 Control Tokens "can be ignored" */
			/* 2.5.198.34 */
			case 'PtgAttrChoose': break;
			/* 2.5.198.35 */
			case 'PtgAttrGoto': break;
			/* 2.5.198.36 */
			case 'PtgAttrIf': break;


			/* 2.5.198.84 */
			case 'PtgRef':
				type = f[1][0]; c = shift_cell_xls(decode_cell(encode_cell(f[1][1])), _range);
				stack.push(encode_cell(c));
				break;
			/* 2.5.198.88 */
			case 'PtgRefN':
				type = f[1][0]; c = shift_cell_xls(decode_cell(encode_cell(f[1][1])), cell);
				stack.push(encode_cell(c));
				break;
			case 'PtgRef3d': // TODO: lots of stuff
				type = f[1][0]; ixti = f[1][1]; c = shift_cell_xls(f[1][2], _range);
				stack.push(supbooks[1][ixti+1]+"!"+encode_cell(c));
				break;

		/* Function Call */
			/* 2.5.198.62 */
			case 'PtgFunc':
			/* 2.5.198.63 */
			case 'PtgFuncVar':
				/* f[1] = [argc, func] */
				var argc = f[1][0], func = f[1][1];
				if(!argc) argc = 0;
				var args = stack.slice(-argc);
				stack.length -= argc;
				if(func === 'User') func = args.shift();
				stack.push(func + "(" + args.join(",") + ")");
				break;

			/* 2.5.198.42 */
			case 'PtgBool': stack.push(f[1] ? "TRUE" : "FALSE"); break;
			/* 2.5.198.66 */
			case 'PtgInt': stack.push(f[1]); break;
			/* 2.5.198.79 TODO: precision? */
			case 'PtgNum': stack.push(String(f[1])); break;
			/* 2.5.198.89 */
			case 'PtgStr': stack.push('"' + f[1] + '"'); break;
			/* 2.5.198.57 */
			case 'PtgErr': stack.push(f[1]); break;
			/* 2.5.198.27 TODO: fixed points */
			case 'PtgArea':
				type = f[1][0]; r = shift_range_xls(f[1][1], _range);
				stack.push(encode_range(r));
				break;
			/* 2.5.198.28 */
			case 'PtgArea3d': // TODO: lots of stuff
				type = f[1][0]; ixti = f[1][1]; r = f[1][2];
				stack.push(supbooks[1][ixti+1]+"!"+encode_range(r));
				break;
			/* 2.5.198.41 */
			case 'PtgAttrSum':
				stack.push("SUM(" + stack.pop() + ")");
				break;

		/* Expression Prefixes */
			/* 2.5.198.37 */
			case 'PtgAttrSemi': break;

			/* 2.5.97.60 TODO: do something different for revisions */
			case 'PtgName':
				/* f[1] = type, 0, nameindex */
				nameidx = f[1][2];
				var lbl = supbooks[0][nameidx];
				var name = lbl.Name;
				if(name in XLSXFutureFunctions) name = XLSXFutureFunctions[name];
				stack.push(name);
				break;

			/* 2.5.97.61 TODO: do something different for revisions */
			case 'PtgNameX':
				/* f[1] = type, ixti, nameindex */
				var bookidx = f[1][1]; nameidx = f[1][2]; var externbook;
				/* TODO: Properly handle missing values */
				if(supbooks[bookidx+1]) externbook = supbooks[bookidx+1][nameidx];
				else if(supbooks[bookidx-1]) externbook = supbooks[bookidx-1][nameidx];
				if(!externbook) externbook = {body: "??NAMEX??"};
				stack.push(externbook.body);
				break;

		/* 2.2.2.4 Display Tokens */
			/* 2.5.198.80 */
			case 'PtgParen': stack.push('(' + stack.pop() + ')'); break;

			/* 2.5.198.86 */
			case 'PtgRefErr': stack.push('#REF!'); break;

		/* */
			/* 2.5.198.58 TODO */
			case 'PtgExp':
				c = {c:f[1][1],r:f[1][0]};
				var q = {c: cell.c, r:cell.r};
				if(supbooks.sharedf[encode_cell(c)]) {
					var parsedf = (supbooks.sharedf[encode_cell(c)]);
					stack.push(stringify_formula(parsedf, _range, q, supbooks, opts));
				}
				else {
					var fnd = false;
					for(e1=0;e1!=supbooks.arrayf.length; ++e1) {
						/* TODO: should be something like range_has */
						e2 = supbooks.arrayf[e1];
						if(c.c < e2[0].s.c || c.c > e2[0].e.c) continue;
						if(c.r < e2[0].s.r || c.r > e2[0].e.r) continue;
						stack.push(stringify_formula(e2[1], _range, q, supbooks, opts));
					}
					if(!fnd) stack.push(f[1]);
				}
				break;

			/* 2.5.198.32 TODO */
			case 'PtgArray':
				stack.push("{" + f[1].map(mapper).join(";") + "}");
				break;

		/* 2.2.2.5 Mem Tokens */
			/* 2.5.198.70 TODO: confirm this is a non-display */
			case 'PtgMemArea':
				//stack.push("(" + f[2].map(encode_range).join(",") + ")");
				break;

			/* 2.5.198.38 TODO */
			case 'PtgAttrSpace': break;

			/* 2.5.198.92 TODO */
			case 'PtgTbl': break;

			/* 2.5.198.71 */
			case 'PtgMemErr': break;

			/* 2.5.198.74 */
			case 'PtgMissArg':
				stack.push("");
				break;

			/* 2.5.198.29 TODO */
			case 'PtgAreaErr': break;

			/* 2.5.198.31 TODO */
			case 'PtgAreaN': stack.push(""); break;

			/* 2.5.198.87 TODO */
			case 'PtgRefErr3d': break;

			/* 2.5.198.72 TODO */
			case 'PtgMemFunc': break;

			default: throw 'Unrecognized Formula Token: ' + f;
		}
		//console.log("::",f, stack)
	}
	//console.log("--",stack);
	return stack[0];
}

