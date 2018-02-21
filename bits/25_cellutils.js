/* XLS ranges enforced */
function shift_cell_xls(cell/*:CellAddress*/, tgt/*:any*/, opts/*:?any*/)/*:CellAddress*/ {
	var out = dup(cell);
	if(tgt.s) {
		if(out.cRel) out.c += tgt.s.c;
		if(out.rRel) out.r += tgt.s.r;
	} else {
		if(out.cRel) out.c += tgt.c;
		if(out.rRel) out.r += tgt.r;
	}
	if(!opts || opts.biff < 12) {
		while(out.c >= 0x100) out.c -= 0x100;
		while(out.r >= 0x10000) out.r -= 0x10000;
	}
	return out;
}

function shift_range_xls(cell, range, opts) {
	var out = dup(cell);
	out.s = shift_cell_xls(out.s, range.s, opts);
	out.e = shift_cell_xls(out.e, range.s, opts);
	return out;
}

function encode_cell_xls(c/*:CellAddress*/, biff/*:number*/)/*:string*/ {
	if(c.cRel && c.c < 0) { c = dup(c); c.c += (biff > 8) ? 0x4000 : 0x100; }
	if(c.rRel && c.r < 0) { c = dup(c); c.r += (biff > 8) ? 0x100000 : ((biff > 5) ? 0x10000 : 0x4000); }
	var s = encode_cell(c);
	if(c.cRel === 0) s = fix_col(s);
	if(c.rRel === 0) s = fix_row(s);
	return s;
}

function encode_range_xls(r, opts)/*:string*/ {
	if(r.s.r == 0 && !r.s.rRel) {
		if(r.e.r == (opts.biff >= 12 ? 0xFFFFF : (opts.biff >= 8 ? 0x10000 : 0x4000)) && !r.e.rRel) {
			return (r.s.cRel ? "" : "$") + encode_col(r.s.c) + ":" + (r.e.cRel ? "" : "$") + encode_col(r.e.c);
		}
	}
	if(r.s.c == 0 && !r.s.cRel) {
		if(r.e.c == (opts.biff >= 12 ? 0xFFFF : 0xFF) && !r.e.cRel) {
			return (r.s.rRel ? "" : "$") + encode_row(r.s.r) + ":" + (r.e.rRel ? "" : "$") + encode_row(r.e.r);
		}
	}
	return encode_cell_xls(r.s, opts.biff) + ":" + encode_cell_xls(r.e, opts.biff);
}
