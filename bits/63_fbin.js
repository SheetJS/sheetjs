/* [MS-XLS] 2.5.198.1 TODO */
function parse_ArrayParsedFormula(blob, length, opts/*::, ref*/) {
	var target = blob.l + length, len = opts.biff == 2 ? 1 : 2;
	var rgcb, cce = blob.read_shift(len); // length of rgce
	if(cce == 0xFFFF) return [[],parsenoop(blob, length-2)];
	var rgce = parse_Rgce(blob, cce, opts);
	if(length !== cce + len) rgcb = parse_RgbExtra(blob, length - cce - len, rgce, opts);
	blob.l = target;
	return [rgce, rgcb];
}

/* [MS-XLS] 2.5.198.3 TODO */
function parse_XLSCellParsedFormula(blob, length, opts) {
	var target = blob.l + length, len = opts.biff == 2 ? 1 : 2;
	var rgcb, cce = blob.read_shift(len); // length of rgce
	if(cce == 0xFFFF) return [[],parsenoop(blob, length-2)];
	var rgce = parse_Rgce(blob, cce, opts);
	if(length !== cce + len) rgcb = parse_RgbExtra(blob, length - cce - len, rgce, opts);
	blob.l = target;
	return [rgce, rgcb];
}

/* [MS-XLS] 2.5.198.21 */
function parse_NameParsedFormula(blob, length, opts, cce) {
	var target = blob.l + length;
	var rgce = parse_Rgce(blob, cce, opts);
	var rgcb;
	if(target !== blob.l) rgcb = parse_RgbExtra(blob, target - blob.l, rgce, opts);
	return [rgce, rgcb];
}

/* [MS-XLS] 2.5.198.118 TODO */
function parse_SharedParsedFormula(blob, length, opts) {
	var target = blob.l + length;
	var rgcb, cce = blob.read_shift(2); // length of rgce
	var rgce = parse_Rgce(blob, cce, opts);
	if(cce == 0xFFFF) return [[],parsenoop(blob, length-2)];
	if(length !== cce + 2) rgcb = parse_RgbExtra(blob, target - cce - 2, rgce, opts);
	return [rgce, rgcb];
}

/* [MS-XLS] 2.5.133 TODO: how to emit empty strings? */
function parse_FormulaValue(blob/*::, length*/) {
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
function write_FormulaValue(value) {
	if(value == null) {
		// Blank String Value
		var o = new_buf(8);
		o.write_shift(1, 0x03);
		o.write_shift(1, 0);
		o.write_shift(2, 0);
		o.write_shift(2, 0);
		o.write_shift(2, 0xFFFF);
		return o;
	} else if(typeof value == "number") return write_Xnum(value);
	return write_Xnum(0);
}

/* [MS-XLS] 2.4.127 TODO */
function parse_Formula(blob, length, opts) {
	var end = blob.l + length;
	var cell = parse_XLSCell(blob, 6);
	if(opts.biff == 2) ++blob.l;
	var val = parse_FormulaValue(blob,8);
	var flags = blob.read_shift(1);
	if(opts.biff != 2) {
		blob.read_shift(1);
		if(opts.biff >= 5) {
			/*var chn = */blob.read_shift(4);
		}
	}
	var cbf = parse_XLSCellParsedFormula(blob, end - blob.l, opts);
	return {cell:cell, val:val[0], formula:cbf, shared: (flags >> 3) & 1, tt:val[1]};
}
function write_Formula(cell/*:Cell*/, R/*:number*/, C/*:number*/, opts, os/*:number*/) {
	// Cell
	var o1 = write_XLSCell(R, C, os);

	// FormulaValue
	var o2 = write_FormulaValue(cell.v);

	// flags + cache
	var o3 = new_buf(6);
	var flags = 0x01 | 0x20;
	o3.write_shift(2, flags);
	o3.write_shift(4, 0);

	// CellParsedFormula
	var bf = new_buf(cell.bf.length);
	for(var i = 0; i < cell.bf.length; ++i) bf[i] = cell.bf[i];

	var out = bconcat([o1, o2, o3, bf]);
	return out;
}


/* XLSB Parsed Formula records have the same shape */
function parse_XLSBParsedFormula(data, length, opts) {
	var cce = data.read_shift(4);
	var rgce = parse_Rgce(data, cce, opts);
	var cb = data.read_shift(4);
	var rgcb = cb > 0 ? parse_RgbExtra(data, cb, rgce, opts) : null;
	return [rgce, rgcb];
}

/* [MS-XLSB] 2.5.97.1 ArrayParsedFormula */
var parse_XLSBArrayParsedFormula = parse_XLSBParsedFormula;
/* [MS-XLSB] 2.5.97.4 CellParsedFormula */
var parse_XLSBCellParsedFormula = parse_XLSBParsedFormula;
/* [MS-XLSB] 2.5.97.8 DVParsedFormula */
//var parse_XLSBDVParsedFormula = parse_XLSBParsedFormula;
/* [MS-XLSB] 2.5.97.9 FRTParsedFormula */
//var parse_XLSBFRTParsedFormula = parse_XLSBParsedFormula2;
/* [MS-XLSB] 2.5.97.12 NameParsedFormula */
var parse_XLSBNameParsedFormula = parse_XLSBParsedFormula;
/* [MS-XLSB] 2.5.97.98 SharedParsedFormula */
var parse_XLSBSharedParsedFormula = parse_XLSBParsedFormula;

/* Writes a PtgNum or PtgInt */
function write_XLSBFormulaNum(val/*:number*/) {
	if((val | 0) == val && val < Math.pow(2,16) && val >= 0) {
		var oint = new_buf(11);
		oint.write_shift(4, 3);
		oint.write_shift(1, 0x1e);
		oint.write_shift(2, val);
		oint.write_shift(4, 0);
		return oint;
	}

	var num = new_buf(17);
	num.write_shift(4, 11);
	num.write_shift(1, 0x1f);
	num.write_shift(8, val);
	num.write_shift(4, 0);
	return num;
}
/* Writes a PtgErr */
function write_XLSBFormulaErr(val/*:number*/) {
	var oint = new_buf(10);
	oint.write_shift(4, 2);
	oint.write_shift(1, 0x1C);
	oint.write_shift(1, val);
	oint.write_shift(4, 0);
	return oint;
}
/* Writes a PtgBool */
function write_XLSBFormulaBool(val/*:boolean*/) {
	var oint = new_buf(10);
	oint.write_shift(4, 2);
	oint.write_shift(1, 0x1D);
	oint.write_shift(1, val?1:0);
	oint.write_shift(4, 0);
	return oint;
}

/* Writes a PtgStr */
function write_XLSBFormulaStr(val/*:string*/) {
	var preamble = new_buf(7);
	preamble.write_shift(4, 3 + 2 * val.length);
	preamble.write_shift(1, 0x17);
	preamble.write_shift(2, val.length);

	var body = new_buf(2 * val.length);
	body.write_shift(2 * val.length, val, "utf16le");

	var postamble = new_buf(4);
	postamble.write_shift(4, 0);

	return bconcat([preamble, body, postamble]);
}

/* Writes a PtgRef */
function write_XLSBFormulaRef(str) {
	var cell = decode_cell(str);
	var out = new_buf(15);
	out.write_shift(4, 7);
	out.write_shift(1, 0x04 | ((1)<<5));
	out.write_shift(4, cell.r);
	out.write_shift(2, cell.c | ((str.charAt(0) == "$" ? 0 : 1)<<14) | ((str.match(/\$\d/) ? 0 : 1)<<15)); // <== ColRelShort
	out.write_shift(4, 0);

	return out;
}

/* Writes a PtgRef3d */
function write_XLSBFormulaRef3D(str, wb) {
	var lastbang = str.lastIndexOf("!");
	var sname = str.slice(0, lastbang);
	str = str.slice(lastbang+1);
	var cell = decode_cell(str);
	if(sname.charAt(0) == "'") sname = sname.slice(1, -1).replace(/''/g, "'");

	var out = new_buf(17);
	out.write_shift(4, 9);
	out.write_shift(1, 0x1A | ((1)<<5));
	out.write_shift(2, 2 + wb.SheetNames.map(function(n) { return n.toLowerCase(); }).indexOf(sname.toLowerCase()));
	out.write_shift(4, cell.r);
	out.write_shift(2, cell.c | ((str.charAt(0) == "$" ? 0 : 1)<<14) | ((str.match(/\$\d/) ? 0 : 1)<<15)); // <== ColRelShort
	out.write_shift(4, 0);

	return out;
}

/* Writes a PtgRefErr3d */
function write_XLSBFormulaRefErr3D(str, wb) {
	var lastbang = str.lastIndexOf("!");
	var sname = str.slice(0, lastbang);
	str = str.slice(lastbang+1);
	if(sname.charAt(0) == "'") sname = sname.slice(1, -1).replace(/''/g, "'");

	var out = new_buf(17);
	out.write_shift(4, 9);
	out.write_shift(1, 0x1C | ((1)<<5));
	out.write_shift(2, 2 + wb.SheetNames.map(function(n) { return n.toLowerCase(); }).indexOf(sname.toLowerCase()));
	out.write_shift(4, 0);
	out.write_shift(2, 0); // <== ColRelShort
	out.write_shift(4, 0);

	return out;
}

/* Writes a single sheet range [PtgRef PtgRef PtgRange] */
function write_XLSBFormulaRange(_str) {
	var parts = _str.split(":"), str = parts[0];

	var out = new_buf(23);
	out.write_shift(4, 15);

	/* start cell */
	str = parts[0]; var cell = decode_cell(str);
	out.write_shift(1, 0x04 | ((1)<<5));
	out.write_shift(4, cell.r);
	out.write_shift(2, cell.c | ((str.charAt(0) == "$" ? 0 : 1)<<14) | ((str.match(/\$\d/) ? 0 : 1)<<15)); // <== ColRelShort
	out.write_shift(4, 0);

	/* end cell */
	str = parts[1]; cell = decode_cell(str);
	out.write_shift(1, 0x04 | ((1)<<5));
	out.write_shift(4, cell.r);
	out.write_shift(2, cell.c | ((str.charAt(0) == "$" ? 0 : 1)<<14) | ((str.match(/\$\d/) ? 0 : 1)<<15)); // <== ColRelShort
	out.write_shift(4, 0);

	/* PtgRange */
	out.write_shift(1, 0x11);

	out.write_shift(4, 0);

	return out;
}

/* Writes a range with explicit sheet name [PtgRef3D PtgRef3D PtgRange] */
function write_XLSBFormulaRangeWS(_str, wb) {
	var lastbang = _str.lastIndexOf("!");
	var sname = _str.slice(0, lastbang);
	_str = _str.slice(lastbang+1);
	if(sname.charAt(0) == "'") sname = sname.slice(1, -1).replace(/''/g, "'");
	var parts = _str.split(":"); str = parts[0];

	var out = new_buf(27);
	out.write_shift(4, 19);

	/* start cell */
	var str = parts[0], cell = decode_cell(str);
	out.write_shift(1, 0x1A | ((1)<<5));
	out.write_shift(2, 2 + wb.SheetNames.map(function(n) { return n.toLowerCase(); }).indexOf(sname.toLowerCase()));
	out.write_shift(4, cell.r);
	out.write_shift(2, cell.c | ((str.charAt(0) == "$" ? 0 : 1)<<14) | ((str.match(/\$\d/) ? 0 : 1)<<15)); // <== ColRelShort

	/* end cell */
	str = parts[1]; cell = decode_cell(str);
	out.write_shift(1, 0x1A | ((1)<<5));
	out.write_shift(2, 2 + wb.SheetNames.map(function(n) { return n.toLowerCase(); }).indexOf(sname.toLowerCase()));
	out.write_shift(4, cell.r);
	out.write_shift(2, cell.c | ((str.charAt(0) == "$" ? 0 : 1)<<14) | ((str.match(/\$\d/) ? 0 : 1)<<15)); // <== ColRelShort

	/* PtgRange */
	out.write_shift(1, 0x11);

	out.write_shift(4, 0);

	return out;
}

/* Writes a range with explicit sheet name [PtgArea3d] */
function write_XLSBFormulaArea3D(_str, wb) {
	var lastbang = _str.lastIndexOf("!");
	var sname = _str.slice(0, lastbang);
	_str = _str.slice(lastbang+1);
	if(sname.charAt(0) == "'") sname = sname.slice(1, -1).replace(/''/g, "'");
	var range = decode_range(_str);

	var out = new_buf(23);
	out.write_shift(4, 15);

	out.write_shift(1, 0x1B | ((1)<<5));
	out.write_shift(2, 2 + wb.SheetNames.map(function(n) { return n.toLowerCase(); }).indexOf(sname.toLowerCase()));
	out.write_shift(4, range.s.r);
	out.write_shift(4, range.e.r);
	out.write_shift(2, range.s.c);
	out.write_shift(2, range.e.c);

	out.write_shift(4, 0);

	return out;
}


/* General Formula */
function write_XLSBFormula(val/*:string|number*/, wb) {
	if(typeof val == "number") return write_XLSBFormulaNum(val);
	if(typeof val == "boolean") return write_XLSBFormulaBool(val);
	if(/^#(DIV\/0!|GETTING_DATA|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!)$/.test(val)) return write_XLSBFormulaErr(+RBErr[val]);
	if(val.match(/^\$?(?:[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D]|[A-Z]{1,2})\$?(?:10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6]|[1-9]\d{0,5})$/)) return write_XLSBFormulaRef(val);
	if(val.match(/^\$?(?:[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D]|[A-Z]{1,2})\$?(?:10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6]|[1-9]\d{0,5}):\$?(?:[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D]|[A-Z]{1,2})\$?(?:10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6]|[1-9]\d{0,5})$/)) return write_XLSBFormulaRange(val);
	if(val.match(/^#REF!\$?(?:[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D]|[A-Z]{1,2})\$?(?:10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6]|[1-9]\d{0,5}):\$?(?:[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D]|[A-Z]{1,2})\$?(?:10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6]|[1-9]\d{0,5})$/)) return write_XLSBFormulaArea3D(val, wb);
	if(val.match(/^(?:'[^\\\/?*\[\]:]*'|[^'][^\\\/?*\[\]:'`~!@#$%^()\-=+{}|;,<.>]*)!\$?(?:[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D]|[A-Z]{1,2})\$?(?:10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6]|[1-9]\d{0,5})$/)) return write_XLSBFormulaRef3D(val, wb);
	if(val.match(/^(?:'[^\\\/?*\[\]:]*'|[^'][^\\\/?*\[\]:'`~!@#$%^()\-=+{}|;,<.>]*)!\$?(?:[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D]|[A-Z]{1,2})\$?(?:10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6]|[1-9]\d{0,5}):\$?(?:[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D]|[A-Z]{1,2})\$?(?:10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6]|[1-9]\d{0,5})$/)) return write_XLSBFormulaRangeWS(val, wb);
	if(/^(?:'[^\\\/?*\[\]:]*'|[^'][^\\\/?*\[\]:'`~!@#$%^()\-=+{}|;,<.>]*)!#REF!$/.test(val)) return write_XLSBFormulaRefErr3D(val, wb);
	if(/^".*"$/.test(val)) return write_XLSBFormulaStr(val);
	if(/^[+-]\d+$/.test(val)) return write_XLSBFormulaNum(parseInt(val, 10));
	throw "Formula |" + val + "| not supported for XLSB";
}
var write_XLSBNameParsedFormula = write_XLSBFormula;
