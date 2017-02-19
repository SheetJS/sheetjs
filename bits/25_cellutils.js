/* XLS ranges enforced */
function shift_cell_xls(cell, tgt, opts) {
	var out = dup(cell);
	if(tgt.s) {
		if(out.cRel) out.c += tgt.s.c;
		if(out.rRel) out.r += tgt.s.r;
	} else {
		out.c += tgt.c;
		out.r += tgt.r;
	}
	if(!opts || opts.biff < 12) {
		while(out.c >= 0x100) out.c -= 0x100;
		while(out.r >= 0x10000) out.r -= 0x10000;
	}
	return out;
}

function shift_range_xls(cell, range) {
	cell.s = shift_cell_xls(cell.s, range.s);
	cell.e = shift_cell_xls(cell.e, range.s);
	return cell;
}

function encode_cell_xls(c)/*:string*/ {
	var s = encode_cell(c);
	if(c.cRel === 0) s = fix_col(s);
	if(c.rRel === 0) s = fix_row(s);
	return s;
}

function encode_range_xls(r)/*:string*/ {
	return encode_cell_xls(r.s) + ":" + encode_cell_xls(r.e);
}
