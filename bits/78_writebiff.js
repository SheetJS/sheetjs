/* BIFF2-4 single-sheet workbooks */
function write_biff_rec(ba/*:BufArray*/, t/*:number*/, payload, length/*:?number*/) {
	var len = (length || (payload||[]).length);
	var o = ba.next(4 + len);
	o.write_shift(2, t);
	o.write_shift(2, len);
	if(/*:: len != null &&*/len > 0 && is_buf(payload)) ba.push(payload);
}

function write_BOF(wb/*:Workbook*/, o) {
	if(o.bookType != 'biff2') throw "unsupported BIFF version";
	var out = new_buf(4);
	out.write_shift(2, 0x0002); // "unused"
	out.write_shift(2, 0x0010); // Sheet
	return out;
}

function write_BIFF2Cell(out, r/*:number*/, c/*:number*/) {
	if(!out) out = new_buf(7);
	out.write_shift(2, r);
	out.write_shift(2, c);
	out.write_shift(1, 0);
	out.write_shift(1, 0);
	out.write_shift(1, 0);
	return out;
}

function write_BIFF2INT(r/*:number*/, c/*:number*/, val) {
	var out = new_buf(9);
	write_BIFF2Cell(out, r, c);
	out.write_shift(2, val);
	return out;
}

function write_BIFF2NUMBER(r, c, val) {
	var out = new_buf(15);
	write_BIFF2Cell(out, r, c);
	out.write_shift(8, val, 'f');
	return out;
}

function write_BIFF2BERR(r, c, val, t) {
	var out = new_buf(9);
	write_BIFF2Cell(out, r, c);
	if(t == 'e') { out.write_shift(1, val); out.write_shift(1, 1); }
	else { out.write_shift(1, val?1:0); out.write_shift(1, 0); }
	return out;
}

/* TODO: codepage, large strings */
function write_BIFF2LABEL(r, c, val) {
	var out = new_buf(8 + 2*val.length);
	write_BIFF2Cell(out, r, c);
	out.write_shift(1, val.length);
	out.write_shift(val.length, val, 'sbcs');
	return out.l < out.length ? out.slice(0, out.l) : out;
}

function write_ws_biff_cell(ba/*:BufArray*/, cell/*:Cell*/, R/*:number*/, C/*:number*/, opts) {
	if(cell.v != null) switch(cell.t) {
		case 'd': case 'n':
			var v = cell.t == 'd' ? datenum(cell.v) : cell.v;
			if((v == (v|0)) && (v >= 0) && (v < 65536))
				write_biff_rec(ba, 0x0002, write_BIFF2INT(R, C, v));
			else
				write_biff_rec(ba, 0x0003, write_BIFF2NUMBER(R,C, v));
			return;
		case 'b': case 'e': write_biff_rec(ba, 0x0005, write_BIFF2BERR(R, C, cell.v, cell.t)); return;
		/* TODO: codepage, sst */
		case 's': case 'str':
			write_biff_rec(ba, 0x0004, write_BIFF2LABEL(R, C, cell.v));
			return;
	}
	write_biff_rec(ba, 0x0001, write_BIFF2Cell(null, R, C));
}

function write_biff_ws(ba/*:BufArray*/, ws/*:Worksheet*/, idx/*:number*/, opts, wb/*:Workbook*/) {
	var dense = Array.isArray(ws);
	var range = safe_decode_range(ws['!ref'] || "A1"), ref, rr = "", cols = [];
	for(var R = range.s.r; R <= range.e.r; ++R) {
		rr = encode_row(R);
		for(var C = range.s.c; C <= range.e.c; ++C) {
			if(R === range.s.r) cols[C] = encode_col(C);
			ref = cols[C] + rr;
			var cell = dense ? ws[R][C] : ws[ref];
			if(!cell) continue;
			/* write cell */
			write_ws_biff_cell(ba, cell, R, C, opts);
		}
	}
}

/* Based on test files */
function write_biff_buf(wb/*:Workbook*/, opts/*:WriteOpts*/) {
	var o = opts || {};
	if(DENSE != null && o.dense == null) o.dense = DENSE;
	var ba = buf_array();
	var idx = 0;
	for(var i=0;i<wb.SheetNames.length;++i) if(wb.SheetNames[i] == o.sheet) idx=i;
	if(idx == 0 && !!o.sheet && wb.SheetNames[0] != o.sheet) throw new Error("Sheet not found: " + o.sheet);
	write_biff_rec(ba, 0x0009, write_BOF(wb, o));
	/* ... */
	write_biff_ws(ba, wb.Sheets[wb.SheetNames[idx]], idx, o, wb);
	/* ... */
	write_biff_rec(ba, 0x000a);
	// TODO
	return ba.end();
}
