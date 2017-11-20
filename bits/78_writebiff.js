function write_biff_rec(ba/*:BufArray*/, type/*:number|string*/, payload, length/*:?number*/) {
	var t/*:number*/ = +type || +XLSRE[/*::String(*/type/*::)*/];
	if(isNaN(t)) return;
	var len = length || (payload||[]).length || 0;
	var o = ba.next(4 + len);
	o.write_shift(2, t);
	o.write_shift(2, len);
	if(/*:: len != null &&*/len > 0 && is_buf(payload)) ba.push(payload);
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

function write_BIFF2BERR(r/*:number*/, c/*:number*/, val, t/*:?string*/) {
	var out = new_buf(9);
	write_BIFF2Cell(out, r, c);
	if(t == 'e') { out.write_shift(1, val); out.write_shift(1, 1); }
	else { out.write_shift(1, val?1:0); out.write_shift(1, 0); }
	return out;
}

/* TODO: codepage, large strings */
function write_BIFF2LABEL(r/*:number*/, c/*:number*/, val) {
	var out = new_buf(8 + 2*val.length);
	write_BIFF2Cell(out, r, c);
	out.write_shift(1, val.length);
	out.write_shift(val.length, val, 'sbcs');
	return out.l < out.length ? out.slice(0, out.l) : out;
}

function write_ws_biff2_cell(ba/*:BufArray*/, cell/*:Cell*/, R/*:number*/, C/*:number*/, opts) {
	if(cell.v != null) switch(cell.t) {
		case 'd': case 'n':
			var v = cell.t == 'd' ? datenum(parseDate(cell.v)) : cell.v;
			if((v == (v|0)) && (v >= 0) && (v < 65536))
				write_biff_rec(ba, 0x0002, write_BIFF2INT(R, C, v));
			else
				write_biff_rec(ba, 0x0003, write_BIFF2NUM(R,C, v));
			return;
		case 'b': case 'e': write_biff_rec(ba, 0x0005, write_BIFF2BERR(R, C, cell.v, cell.t)); return;
		/* TODO: codepage, sst */
		case 's': case 'str':
			write_biff_rec(ba, 0x0004, write_BIFF2LABEL(R, C, cell.v));
			return;
	}
	write_biff_rec(ba, 0x0001, write_BIFF2Cell(null, R, C));
}

function write_ws_biff2(ba/*:BufArray*/, ws/*:Worksheet*/, idx/*:number*/, opts, wb/*:Workbook*/) {
	var dense = Array.isArray(ws);
	var range = safe_decode_range(ws['!ref'] || "A1"), ref/*:string*/, rr = "", cols/*:Array<string>*/ = [];
	for(var R = range.s.r; R <= range.e.r; ++R) {
		rr = encode_row(R);
		for(var C = range.s.c; C <= range.e.c; ++C) {
			if(R === range.s.r) cols[C] = encode_col(C);
			ref = cols[C] + rr;
			var cell = dense ? (ws[R]||[])[C] : ws[ref];
			if(!cell) continue;
			/* write cell */
			write_ws_biff2_cell(ba, cell, R, C, opts);
		}
	}
}

/* Based on test files */
function write_biff2_buf(wb/*:Workbook*/, opts/*:WriteOpts*/) {
	var o = opts || {};
	if(DENSE != null && o.dense == null) o.dense = DENSE;
	var ba = buf_array();
	var idx = 0;
	for(var i=0;i<wb.SheetNames.length;++i) if(wb.SheetNames[i] == o.sheet) idx=i;
	if(idx == 0 && !!o.sheet && wb.SheetNames[0] != o.sheet) throw new Error("Sheet not found: " + o.sheet);
	write_biff_rec(ba, 0x0009, write_BOF(wb, 0x10, o));
	/* ... */
	write_ws_biff2(ba, wb.Sheets[wb.SheetNames[idx]], idx, o, wb);
	/* ... */
	write_biff_rec(ba, 0x000A);
	return ba.end();
}

function write_ws_biff8_cell(ba/*:BufArray*/, cell/*:Cell*/, R/*:number*/, C/*:number*/, opts) {
	if(cell.v != null) switch(cell.t) {
		case 'd': case 'n':
			var v = cell.t == 'd' ? datenum(parseDate(cell.v)) : cell.v;
			/* TODO: emit RK as appropriate */
			write_biff_rec(ba, "Number", write_Number(R, C, v, opts));
			return;
		case 'b': case 'e': write_biff_rec(ba, "BoolErr", write_BoolErr(R, C, cell.v, opts, cell.t)); return;
		/* TODO: codepage, sst */
		case 's': case 'str':
			write_biff_rec(ba, "Label", write_Label(R, C, cell.v, opts));
			return;
	}
	write_biff_rec(ba, "Blank", write_XLSCell(R, C));
}

/* [MS-XLS] 2.1.7.20.5 */
function write_ws_biff8(idx/*:number*/, opts, wb/*:Workbook*/) {
	var ba = buf_array();
	var s = wb.SheetNames[idx], ws = wb.Sheets[s] || {};
	var _sheet/*:WBWSProp*/ = ((((wb||{}).Workbook||{}).Sheets||[])[idx]||{}/*:any*/);
	var dense = Array.isArray(ws);
	var ref/*:string*/, rr = "", cols/*:Array<string>*/ = [];
	var range = safe_decode_range(ws['!ref'] || "A1");
	write_biff_rec(ba, 0x0809, write_BOF(wb, 0x10, opts));
	/* ... */
	write_biff_rec(ba, "CalcMode", writeuint16(1));
	write_biff_rec(ba, "CalcCount", writeuint16(100));
	write_biff_rec(ba, "CalcRefMode", writebool(true));
	write_biff_rec(ba, "CalcIter", writebool(false));
	write_biff_rec(ba, "CalcDelta", write_Xnum(0.001));
	write_biff_rec(ba, "CalcSaveRecalc", writebool(true));
	write_biff_rec(ba, "PrintRowCol", writebool(false));
	write_biff_rec(ba, "PrintGrid", writebool(false));
	write_biff_rec(ba, "GridSet", writeuint16(1));
	write_biff_rec(ba, "Guts", write_Guts([0,0]));
	/* ... */
	write_biff_rec(ba, "HCenter", writebool(false));
	write_biff_rec(ba, "VCenter", writebool(false));
	/* ... */
	write_biff_rec(ba, "Dimensions", write_Dimensions(range, opts));
	/* ... */

	for(var R = range.s.r; R <= range.e.r; ++R) {
		rr = encode_row(R);
		for(var C = range.s.c; C <= range.e.c; ++C) {
			if(R === range.s.r) cols[C] = encode_col(C);
			ref = cols[C] + rr;
			var cell = dense ? (ws[R]||[])[C] : ws[ref];
			if(!cell) continue;
			/* write cell */
			write_ws_biff8_cell(ba, cell, R, C, opts);
		}
	}
	var cname/*:string*/ = _sheet.CodeName || _sheet.name || s;
	write_biff_rec(ba, "CodeName", write_XLUnicodeString(cname, opts));
	/* ... */
	write_biff_rec(ba, "EOF");
	return ba.end();
}

/* [MS-XLS] 2.1.7.20.3 */
function write_biff8_global(wb/*:Workbook*/, bufs, opts/*:WriteOpts*/) {
	var A = buf_array();
	var _wb/*:WBWBProps*/ = /*::((*/(wb.Workbook||{}).WBProps||{/*::CodeName:"ThisWorkbook"*/}/*:: ):any)*/;
	var b8 = opts.biff == 8, b5 = opts.biff == 5;
	write_biff_rec(A, 0x0809, write_BOF(wb, 0x05, opts));
	write_biff_rec(A, "InterfaceHdr", b8 ? writeuint16(0x04b0) : null);
	write_biff_rec(A, "Mms", writezeroes(2));
	if(b5) write_biff_rec(A, "ToolbarHdr");
	if(b5) write_biff_rec(A, "ToolbarEnd");
	write_biff_rec(A, "InterfaceEnd");
	write_biff_rec(A, "WriteAccess", write_WriteAccess("SheetJS", opts));
	write_biff_rec(A, "CodePage", writeuint16(b8 ? 0x04b0 : 0x04E4));
	if(b8) write_biff_rec(A, "DSF", writeuint16(0));
	write_biff_rec(A, "RRTabId", write_RRTabId(wb.SheetNames.length));
	if(b8 && wb.vbaraw) {
		write_biff_rec(A, "ObProj");
		// $FlowIgnore
		var cname/*:string*/ = _wb.CodeName || "ThisWorkbook";
		write_biff_rec(A, "CodeName", write_XLUnicodeString(cname, opts));
	}
	write_biff_rec(A, "BuiltInFnGroupCount", writeuint16(0x11));
	write_biff_rec(A, "WinProtect", writebool(false));
	write_biff_rec(A, "Protect", writebool(false));
	write_biff_rec(A, "Password", writeuint16(0));
	if(b8) write_biff_rec(A, "Prot4Rev", writebool(false));
	if(b8) write_biff_rec(A, "Prot4RevPass", writeuint16(0));
	write_biff_rec(A, "Window1", write_Window1(opts));
	write_biff_rec(A, "Backup", writebool(false));
	write_biff_rec(A, "HideObj", writeuint16(0));
	write_biff_rec(A, "Date1904", writebool(safe1904(wb)=="true"));
	write_biff_rec(A, "CalcPrecision", writebool(true));
	if(b8) write_biff_rec(A, "RefreshAll", writebool(false));
	write_biff_rec(A, "BookBool", writeuint16(0));
	/* ... */
	if(b8) write_biff_rec(A, "UsesELFs", writebool(false));
	var a = A.end();

	var C = buf_array();
	if(b8) write_biff_rec(C, "Country", write_Country());
	/* BIFF8: [SST *Continue] ExtSST */
	write_biff_rec(C, "EOF");
	var c = C.end();

	var B = buf_array();
	var blen = 0, j = 0;
	for(j = 0; j < wb.SheetNames.length; ++j) blen += (b8 ? 12 : 11) + (b8 ? 2 : 1) * wb.SheetNames[j].length;
	var start = a.length + blen + c.length;
	for(j = 0; j < wb.SheetNames.length; ++j) {
		write_biff_rec(B, "BoundSheet8", write_BoundSheet8({pos:start, hs:0, dt:0, name:wb.SheetNames[j]}, opts));
		start += bufs[j].length;
	}
	/* 1*BoundSheet8 */
	var b = B.end();
	if(blen != b.length) throw new Error("BS8 " + blen + " != " + b.length);

	var out = [];
	if(a.length) out.push(a);
	if(b.length) out.push(b);
	if(c.length) out.push(c);
	return __toBuffer([out]);
}

/* [MS-XLS] 2.1.7.20 Workbook Stream */
function write_biff8_buf(wb/*:Workbook*/, opts/*:WriteOpts*/) {
	var o = opts || {};
	var bufs = [];
	for(var i = 0; i < wb.SheetNames.length; ++i) bufs[bufs.length] = write_ws_biff8(i, o, wb);
	bufs.unshift(write_biff8_global(wb, bufs, o));
	return __toBuffer([bufs]);
}

function write_biff_buf(wb/*:Workbook*/, opts/*:WriteOpts*/) {
	var o = opts || {};
	switch(o.biff || 2) {
		case 8: case 5: return write_biff8_buf(wb, opts);
		case 4: case 3: case 2: return write_biff2_buf(wb, opts);
	}
	throw new Error("invalid type " + o.bookType + " for BIFF");
}
