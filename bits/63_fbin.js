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
