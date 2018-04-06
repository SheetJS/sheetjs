
/* [MS-XLSB] 2.4.726 BrtRowHdr */
function parse_BrtRowHdr(data, length) {
	var z = ({}/*:any*/);
	var tgt = data.l + length;
	z.r = data.read_shift(4);
	data.l += 4; // TODO: ixfe
	var miyRw = data.read_shift(2);
	data.l += 1; // TODO: top/bot padding
	var flags = data.read_shift(1);
	data.l = tgt;
	if(flags & 0x07) z.level = flags & 0x07;
	if(flags & 0x10) z.hidden = true;
	if(flags & 0x20) z.hpt = miyRw / 20;
	return z;
}
function write_BrtRowHdr(R/*:number*/, range, ws) {
	var o = new_buf(17+8*16);
	var row = (ws['!rows']||[])[R]||{};
	o.write_shift(4, R);

	o.write_shift(4, 0); /* TODO: ixfe */

	var miyRw = 0x0140;
	if(row.hpx) miyRw = px2pt(row.hpx) * 20;
	else if(row.hpt) miyRw = row.hpt * 20;
	o.write_shift(2, miyRw);

	o.write_shift(1, 0); /* top/bot padding */

	var flags = 0x0;
	if(row.level) flags |= row.level;
	if(row.hidden) flags |= 0x10;
	if(row.hpx || row.hpt) flags |= 0x20;
	o.write_shift(1, flags);

	o.write_shift(1, 0); /* phonetic guide */

	/* [MS-XLSB] 2.5.8 BrtColSpan explains the mechanism */
	var ncolspan = 0, lcs = o.l;
	o.l += 4;

	var caddr = {r:R, c:0};
	for(var i = 0; i < 16; ++i) {
		if((range.s.c > ((i+1) << 10)) || (range.e.c < (i << 10))) continue;
		var first = -1, last = -1;
		for(var j = (i<<10); j < ((i+1)<<10); ++j) {
			caddr.c = j;
			var cell = Array.isArray(ws) ? (ws[caddr.r]||[])[caddr.c] : ws[encode_cell(caddr)];
			if(cell) { if(first < 0) first = j; last = j; }
		}
		if(first < 0) continue;
		++ncolspan;
		o.write_shift(4, first);
		o.write_shift(4, last);
	}

	var l = o.l;
	o.l = lcs;
	o.write_shift(4, ncolspan);
	o.l = l;

	return o.length > o.l ? o.slice(0, o.l) : o;
}
function write_row_header(ba, ws, range, R) {
	var o = write_BrtRowHdr(R, range, ws);
	if((o.length > 17) || (ws['!rows']||[])[R]) write_record(ba, 'BrtRowHdr', o);
}

/* [MS-XLSB] 2.4.820 BrtWsDim */
var parse_BrtWsDim = parse_UncheckedRfX;
var write_BrtWsDim = write_UncheckedRfX;

/* [MS-XLSB] 2.4.821 BrtWsFmtInfo */
function parse_BrtWsFmtInfo(/*::data, length*/) {
}
//function write_BrtWsFmtInfo(ws, o) { }

/* [MS-XLSB] 2.4.823 BrtWsProp */
function parse_BrtWsProp(data, length) {
	var z = {};
	/* TODO: pull flags */
	data.l += 19;
	z.name = parse_XLSBCodeName(data, length - 19);
	return z;
}
function write_BrtWsProp(str, o) {
	if(o == null) o = new_buf(84+4*str.length);
	for(var i = 0; i < 3; ++i) o.write_shift(1,0);
	write_BrtColor({auto:1}, o);
	o.write_shift(-4,-1);
	o.write_shift(-4,-1);
	write_XLSBCodeName(str, o);
	return o.slice(0, o.l);
}

/* [MS-XLSB] 2.4.306 BrtCellBlank */
function parse_BrtCellBlank(data) {
	var cell = parse_XLSBCell(data);
	return [cell];
}
function write_BrtCellBlank(cell, ncell, o) {
	if(o == null) o = new_buf(8);
	return write_XLSBCell(ncell, o);
}


/* [MS-XLSB] 2.4.307 BrtCellBool */
function parse_BrtCellBool(data) {
	var cell = parse_XLSBCell(data);
	var fBool = data.read_shift(1);
	return [cell, fBool, 'b'];
}
function write_BrtCellBool(cell, ncell, o) {
	if(o == null) o = new_buf(9);
	write_XLSBCell(ncell, o);
	o.write_shift(1, cell.v ? 1 : 0);
	return o;
}

/* [MS-XLSB] 2.4.308 BrtCellError */
function parse_BrtCellError(data) {
	var cell = parse_XLSBCell(data);
	var bError = data.read_shift(1);
	return [cell, bError, 'e'];
}

/* [MS-XLSB] 2.4.311 BrtCellIsst */
function parse_BrtCellIsst(data) {
	var cell = parse_XLSBCell(data);
	var isst = data.read_shift(4);
	return [cell, isst, 's'];
}
function write_BrtCellIsst(cell, ncell, o) {
	if(o == null) o = new_buf(12);
	write_XLSBCell(ncell, o);
	o.write_shift(4, ncell.v);
	return o;
}

/* [MS-XLSB] 2.4.313 BrtCellReal */
function parse_BrtCellReal(data) {
	var cell = parse_XLSBCell(data);
	var value = parse_Xnum(data);
	return [cell, value, 'n'];
}
function write_BrtCellReal(cell, ncell, o) {
	if(o == null) o = new_buf(16);
	write_XLSBCell(ncell, o);
	write_Xnum(cell.v, o);
	return o;
}

/* [MS-XLSB] 2.4.314 BrtCellRk */
function parse_BrtCellRk(data) {
	var cell = parse_XLSBCell(data);
	var value = parse_RkNumber(data);
	return [cell, value, 'n'];
}
function write_BrtCellRk(cell, ncell, o) {
	if(o == null) o = new_buf(12);
	write_XLSBCell(ncell, o);
	write_RkNumber(cell.v, o);
	return o;
}


/* [MS-XLSB] 2.4.317 BrtCellSt */
function parse_BrtCellSt(data) {
	var cell = parse_XLSBCell(data);
	var value = parse_XLWideString(data);
	return [cell, value, 'str'];
}
function write_BrtCellSt(cell, ncell, o) {
	if(o == null) o = new_buf(12 + 4 * cell.v.length);
	write_XLSBCell(ncell, o);
	write_XLWideString(cell.v, o);
	return o.length > o.l ? o.slice(0, o.l) : o;
}

/* [MS-XLSB] 2.4.653 BrtFmlaBool */
function parse_BrtFmlaBool(data, length, opts) {
	var end = data.l + length;
	var cell = parse_XLSBCell(data);
	cell.r = opts['!row'];
	var value = data.read_shift(1);
	var o = [cell, value, 'b'];
	if(opts.cellFormula) {
		data.l += 2;
		var formula = parse_XLSBCellParsedFormula(data, end - data.l, opts);
		o[3] = stringify_formula(formula, null/*range*/, cell, opts.supbooks, opts);/* TODO */
	}
	else data.l = end;
	return o;
}

/* [MS-XLSB] 2.4.654 BrtFmlaError */
function parse_BrtFmlaError(data, length, opts) {
	var end = data.l + length;
	var cell = parse_XLSBCell(data);
	cell.r = opts['!row'];
	var value = data.read_shift(1);
	var o = [cell, value, 'e'];
	if(opts.cellFormula) {
		data.l += 2;
		var formula = parse_XLSBCellParsedFormula(data, end - data.l, opts);
		o[3] = stringify_formula(formula, null/*range*/, cell, opts.supbooks, opts);/* TODO */
	}
	else data.l = end;
	return o;
}

/* [MS-XLSB] 2.4.655 BrtFmlaNum */
function parse_BrtFmlaNum(data, length, opts) {
	var end = data.l + length;
	var cell = parse_XLSBCell(data);
	cell.r = opts['!row'];
	var value = parse_Xnum(data);
	var o = [cell, value, 'n'];
	if(opts.cellFormula) {
		data.l += 2;
		var formula = parse_XLSBCellParsedFormula(data, end - data.l, opts);
		o[3] = stringify_formula(formula, null/*range*/, cell, opts.supbooks, opts);/* TODO */
	}
	else data.l = end;
	return o;
}

/* [MS-XLSB] 2.4.656 BrtFmlaString */
function parse_BrtFmlaString(data, length, opts) {
	var end = data.l + length;
	var cell = parse_XLSBCell(data);
	cell.r = opts['!row'];
	var value = parse_XLWideString(data);
	var o = [cell, value, 'str'];
	if(opts.cellFormula) {
		data.l += 2;
		var formula = parse_XLSBCellParsedFormula(data, end - data.l, opts);
		o[3] = stringify_formula(formula, null/*range*/, cell, opts.supbooks, opts);/* TODO */
	}
	else data.l = end;
	return o;
}

/* [MS-XLSB] 2.4.682 BrtMergeCell */
var parse_BrtMergeCell = parse_UncheckedRfX;
var write_BrtMergeCell = write_UncheckedRfX;
/* [MS-XLSB] 2.4.107 BrtBeginMergeCells */
function write_BrtBeginMergeCells(cnt, o) {
	if(o == null) o = new_buf(4);
	o.write_shift(4, cnt);
	return o;
}

/* [MS-XLSB] 2.4.662 BrtHLink */
function parse_BrtHLink(data, length/*::, opts*/) {
	var end = data.l + length;
	var rfx = parse_UncheckedRfX(data, 16);
	var relId = parse_XLNullableWideString(data);
	var loc = parse_XLWideString(data);
	var tooltip = parse_XLWideString(data);
	var display = parse_XLWideString(data);
	data.l = end;
	var o = ({rfx:rfx, relId:relId, loc:loc, display:display}/*:any*/);
	if(tooltip) o.Tooltip = tooltip;
	return o;
}
function write_BrtHLink(l, rId) {
	var o = new_buf(50+4*(l[1].Target.length + (l[1].Tooltip || "").length));
	write_UncheckedRfX({s:decode_cell(l[0]), e:decode_cell(l[0])}, o);
	write_RelID("rId" + rId, o);
	var locidx = l[1].Target.indexOf("#");
	var loc = locidx == -1 ? "" : l[1].Target.slice(locidx+1);
	write_XLWideString(loc || "", o);
	write_XLWideString(l[1].Tooltip || "", o);
	write_XLWideString("", o);
	return o.slice(0, o.l);
}

/* [MS-XLSB] 2.4.6 BrtArrFmla */
function parse_BrtArrFmla(data, length, opts) {
	var end = data.l + length;
	var rfx = parse_RfX(data, 16);
	var fAlwaysCalc = data.read_shift(1);
	var o = [rfx]; o[2] = fAlwaysCalc;
	if(opts.cellFormula) {
		var formula = parse_XLSBArrayParsedFormula(data, end - data.l, opts);
		o[1] = formula;
	} else data.l = end;
	return o;
}

/* [MS-XLSB] 2.4.750 BrtShrFmla */
function parse_BrtShrFmla(data, length, opts) {
	var end = data.l + length;
	var rfx = parse_UncheckedRfX(data, 16);
	var o = [rfx];
	if(opts.cellFormula) {
		var formula = parse_XLSBSharedParsedFormula(data, end - data.l, opts);
		o[1] = formula;
		data.l = end;
	} else data.l = end;
	return o;
}

/* [MS-XLSB] 2.4.323 BrtColInfo */
/* TODO: once XLS ColInfo is set, combine the functions */
function write_BrtColInfo(C/*:number*/, col, o) {
	if(o == null) o = new_buf(18);
	var p = col_obj_w(C, col);
	o.write_shift(-4, C);
	o.write_shift(-4, C);
	o.write_shift(4, (p.width || 10) * 256);
	o.write_shift(4, 0/*ixfe*/); // style
	var flags = 0;
	if(col.hidden) flags |= 0x01;
	if(typeof p.width == 'number') flags |= 0x02;
	o.write_shift(1, flags); // bit flag
	o.write_shift(1, 0); // bit flag
	return o;
}

/* [MS-XLSB] 2.4.678 BrtMargins */
var BrtMarginKeys = ["left","right","top","bottom","header","footer"];
function parse_BrtMargins(data/*::, length, opts*/)/*:Margins*/ {
	var margins = ({}/*:any*/);
	BrtMarginKeys.forEach(function(k) { margins[k] = parse_Xnum(data, 8); });
	return margins;
}
function write_BrtMargins(margins/*:Margins*/, o) {
	if(o == null) o = new_buf(6*8);
	default_margins(margins);
	BrtMarginKeys.forEach(function(k) { write_Xnum((margins/*:any*/)[k], o); });
	return o;
}

/* [MS-XLSB] 2.4.299 BrtBeginWsView */
function parse_BrtBeginWsView(data/*::, length, opts*/) {
	var f = data.read_shift(2);
	data.l += 28;
	return { RTL: f & 0x20 };
}
function write_BrtBeginWsView(ws, Workbook, o) {
	if(o == null) o = new_buf(30);
	var f = 0x39c;
	if((((Workbook||{}).Views||[])[0]||{}).RTL) f |= 0x20;
	o.write_shift(2, f); // bit flag
	o.write_shift(4, 0);
	o.write_shift(4, 0); // view first row
	o.write_shift(4, 0); // view first col
	o.write_shift(1, 0); // gridline color ICV
	o.write_shift(1, 0);
	o.write_shift(2, 0);
	o.write_shift(2, 100); // zoom scale
	o.write_shift(2, 0);
	o.write_shift(2, 0);
	o.write_shift(2, 0);
	o.write_shift(4, 0); // workbook view id
	return o;
}

/* [MS-XLSB] 2.4.309 BrtCellIgnoreEC */
function write_BrtCellIgnoreEC(ref) {
	var o = new_buf(24);
	o.write_shift(4, 4);
	o.write_shift(4, 1);
	write_UncheckedRfX(ref, o);
	return o;
}

/* [MS-XLSB] 2.4.748 BrtSheetProtection */
function write_BrtSheetProtection(sp, o) {
	if(o == null) o = new_buf(16*4+2);
	o.write_shift(2, sp.password ? crypto_CreatePasswordVerifier_Method1(sp.password) : 0);
	o.write_shift(4, 1); // this record should not be written if no protection
	[
		["objects",             false], // fObjects
		["scenarios",           false], // fScenarios
		["formatCells",          true], // fFormatCells
		["formatColumns",        true], // fFormatColumns
		["formatRows",           true], // fFormatRows
		["insertColumns",        true], // fInsertColumns
		["insertRows",           true], // fInsertRows
		["insertHyperlinks",     true], // fInsertHyperlinks
		["deleteColumns",        true], // fDeleteColumns
		["deleteRows",           true], // fDeleteRows
		["selectLockedCells",   false], // fSelLockedCells
		["sort",                 true], // fSort
		["autoFilter",           true], // fAutoFilter
		["pivotTables",          true], // fPivotTables
		["selectUnlockedCells", false]  // fSelUnlockedCells
	].forEach(function(n) {
		/*:: if(o == null) throw "unreachable"; */
		if(n[1]) o.write_shift(4, sp[n[0]] != null && !sp[n[0]] ? 1 : 0);
		else      o.write_shift(4, sp[n[0]] != null && sp[n[0]] ? 0 : 1);
	});
	return o;
}

/* [MS-XLSB] 2.1.7.61 Worksheet */
function parse_ws_bin(data, _opts, idx, rels, wb/*:WBWBProps*/, themes, styles)/*:Worksheet*/ {
	if(!data) return data;
	var opts = _opts || {};
	if(!rels) rels = {'!id':{}};
	if(DENSE != null && opts.dense == null) opts.dense = DENSE;
	var s/*:Worksheet*/ = (opts.dense ? [] : {});

	var ref;
	var refguess = {s: {r:2000000, c:2000000}, e: {r:0, c:0} };

	var pass = false, end = false;
	var row, p, cf, R, C, addr, sstr, rr, cell/*:Cell*/;
	var merges/*:Array<Range>*/ = [];
	opts.biff = 12;
	opts['!row'] = 0;

	var ai = 0, af = false;

	var arrayf/*:Array<[Range, string]>*/ = [];
	var sharedf = {};
	var supbooks = opts.supbooks || ([[]]/*:any*/);
	supbooks.sharedf = sharedf;
	supbooks.arrayf = arrayf;
	supbooks.SheetNames = wb.SheetNames || wb.Sheets.map(function(x) { return x.name; });
	if(!opts.supbooks) {
		opts.supbooks = supbooks;
		if(wb.Names) for(var i = 0; i < wb.Names.length; ++i) supbooks[0][i+1] = wb.Names[i];
	}

	var colinfo/*:Array<ColInfo>*/ = [], rowinfo/*:Array<RowInfo>*/ = [];
	var seencol = false;

	recordhopper(data, function ws_parse(val, R_n, RT) {
		if(end) return;
		switch(RT) {
			case 0x0094: /* 'BrtWsDim' */
				ref = val; break;
			case 0x0000: /* 'BrtRowHdr' */
				row = val;
				if(opts.sheetRows && opts.sheetRows <= row.r) end=true;
				rr = encode_row(R = row.r);
				opts['!row'] = row.r;
				if(val.hidden || val.hpt || val.level != null) {
					if(val.hpt) val.hpx = pt2px(val.hpt);
					rowinfo[val.r] = val;
				}
				break;

			case 0x0002: /* 'BrtCellRk' */
			case 0x0003: /* 'BrtCellError' */
			case 0x0004: /* 'BrtCellBool' */
			case 0x0005: /* 'BrtCellReal' */
			case 0x0006: /* 'BrtCellSt' */
			case 0x0007: /* 'BrtCellIsst' */
			case 0x0008: /* 'BrtFmlaString' */
			case 0x0009: /* 'BrtFmlaNum' */
			case 0x000A: /* 'BrtFmlaBool' */
			case 0x000B: /* 'BrtFmlaError' */
				p = ({t:val[2]}/*:any*/);
				switch(val[2]) {
					case 'n': p.v = val[1]; break;
					case 's': sstr = strs[val[1]]; p.v = sstr.t; p.r = sstr.r; break;
					case 'b': p.v = val[1] ? true : false; break;
					case 'e': p.v = val[1]; if(opts.cellText !== false) p.w = BErr[p.v]; break;
					case 'str': p.t = 's'; p.v = val[1]; break;
				}
				if((cf = styles.CellXf[val[0].iStyleRef])) safe_format(p,cf.numFmtId,null,opts, themes, styles);
				C = val[0].c;
				if(opts.dense) { if(!s[R]) s[R] = []; s[R][C] = p; }
				else s[encode_col(C) + rr] = p;
				if(opts.cellFormula) {
					af = false;
					for(ai = 0; ai < arrayf.length; ++ai) {
						var aii = arrayf[ai];
						if(row.r >= aii[0].s.r && row.r <= aii[0].e.r)
							if(C >= aii[0].s.c && C <= aii[0].e.c) {
								p.F = encode_range(aii[0]); af = true;
							}
					}
					if(!af && val.length > 3) p.f = val[3];
				}
				if(refguess.s.r > row.r) refguess.s.r = row.r;
				if(refguess.s.c > C) refguess.s.c = C;
				if(refguess.e.r < row.r) refguess.e.r = row.r;
				if(refguess.e.c < C) refguess.e.c = C;
				if(opts.cellDates && cf && p.t == 'n' && SSF.is_date(SSF._table[cf.numFmtId])) {
					var _d = SSF.parse_date_code(p.v); if(_d) { p.t = 'd'; p.v = new Date(_d.y, _d.m-1,_d.d,_d.H,_d.M,_d.S,_d.u); }
				}
				break;

			case 0x0001: /* 'BrtCellBlank' */
				if(!opts.sheetStubs || pass) break;
				p = ({t:'z',v:undefined}/*:any*/);
				C = val[0].c;
				if(opts.dense) { if(!s[R]) s[R] = []; s[R][C] = p; }
				else s[encode_col(C) + rr] = p;
				if(refguess.s.r > row.r) refguess.s.r = row.r;
				if(refguess.s.c > C) refguess.s.c = C;
				if(refguess.e.r < row.r) refguess.e.r = row.r;
				if(refguess.e.c < C) refguess.e.c = C;
				break;

			case 0x00B0: /* 'BrtMergeCell' */
				merges.push(val); break;

			case 0x01EE: /* 'BrtHLink' */
				var rel = rels['!id'][val.relId];
				if(rel) {
					val.Target = rel.Target;
					if(val.loc) val.Target += "#"+val.loc;
					val.Rel = rel;
				} else if(val.relId == '') {
					val.Target = "#" + val.loc;
				}
				for(R=val.rfx.s.r;R<=val.rfx.e.r;++R) for(C=val.rfx.s.c;C<=val.rfx.e.c;++C) {
					if(opts.dense) {
						if(!s[R]) s[R] = [];
						if(!s[R][C]) s[R][C] = {t:'z',v:undefined};
						s[R][C].l = val;
					} else {
						addr = encode_cell({c:C,r:R});
						if(!s[addr]) s[addr] = {t:'z',v:undefined};
						s[addr].l = val;
					}
				}
				break;

			case 0x01AA: /* 'BrtArrFmla' */
				if(!opts.cellFormula) break;
				arrayf.push(val);
				cell = ((opts.dense ? s[R][C] : s[encode_col(C) + rr])/*:any*/);
				cell.f = stringify_formula(val[1], refguess, {r:row.r, c:C}, supbooks, opts);
				cell.F = encode_range(val[0]);
				break;
			case 0x01AB: /* 'BrtShrFmla' */
				if(!opts.cellFormula) break;
				sharedf[encode_cell(val[0].s)] = val[1];
				cell = (opts.dense ? s[R][C] : s[encode_col(C) + rr]);
				cell.f = stringify_formula(val[1], refguess, {r:row.r, c:C}, supbooks, opts);
				break;

			/* identical to 'ColInfo' in XLS */
			case 0x003C: /* 'BrtColInfo' */
				if(!opts.cellStyles) break;
				while(val.e >= val.s) {
					colinfo[val.e--] = { width: val.w/256, hidden: !!(val.flags & 0x01) };
					if(!seencol) { seencol = true; find_mdw_colw(val.w/256); }
					process_col(colinfo[val.e+1]);
				}
				break;

			case 0x00A1: /* 'BrtBeginAFilter' */
				s['!autofilter'] = { ref:encode_range(val) };
				break;

			case 0x01DC: /* 'BrtMargins' */
				s['!margins'] = val;
				break;

			case 0x0093: /* 'BrtWsProp' */
				if(!wb.Sheets[idx]) wb.Sheets[idx] = {};
				if(val.name) wb.Sheets[idx].CodeName = val.name;
				break;

			case 0x0089: /* 'BrtBeginWsView' */
				if(!wb.Views) wb.Views = [{}];
				if(!wb.Views[0]) wb.Views[0] = {};
				if(val.RTL) wb.Views[0].RTL = true;
				break;

			case 0x01E5: /* 'BrtWsFmtInfo' */
				break;
			case 0x00AF: /* 'BrtAFilterDateGroupItem' */
			case 0x0284: /* 'BrtActiveX' */
			case 0x0271: /* 'BrtBigName' */
			case 0x0232: /* 'BrtBkHim' */
			case 0x018C: /* 'BrtBrk' */
			case 0x0458: /* 'BrtCFIcon' */
			case 0x047A: /* 'BrtCFRuleExt' */
			case 0x01D7: /* 'BrtCFVO' */
			case 0x041A: /* 'BrtCFVO14' */
			case 0x0289: /* 'BrtCellIgnoreEC' */
			case 0x0451: /* 'BrtCellIgnoreEC14' */
			case 0x0031: /* 'BrtCellMeta' */
			case 0x024D: /* 'BrtCellSmartTagProperty' */
			case 0x025F: /* 'BrtCellWatch' */
			case 0x0234: /* 'BrtColor' */
			case 0x041F: /* 'BrtColor14' */
			case 0x00A8: /* 'BrtColorFilter' */
			case 0x00AE: /* 'BrtCustomFilter' */
			case 0x049C: /* 'BrtCustomFilter14' */
			case 0x01F3: /* 'BrtDRef' */
			case 0x0040: /* 'BrtDVal' */
			case 0x041D: /* 'BrtDVal14' */
			case 0x0226: /* 'BrtDrawing' */
			case 0x00AB: /* 'BrtDynamicFilter' */
			case 0x00A7: /* 'BrtFilter' */
			case 0x0499: /* 'BrtFilter14' */
			case 0x00A9: /* 'BrtIconFilter' */
			case 0x049D: /* 'BrtIconFilter14' */
			case 0x0227: /* 'BrtLegacyDrawing' */
			case 0x0228: /* 'BrtLegacyDrawingHF' */
			case 0x0295: /* 'BrtListPart' */
			case 0x027F: /* 'BrtOleObject' */
			case 0x01DE: /* 'BrtPageSetup' */
			case 0x0097: /* 'BrtPane' */
			case 0x0219: /* 'BrtPhoneticInfo' */
			case 0x01DD: /* 'BrtPrintOptions' */
			case 0x0218: /* 'BrtRangeProtection' */
			case 0x044F: /* 'BrtRangeProtection14' */
			case 0x02A8: /* 'BrtRangeProtectionIso' */
			case 0x0450: /* 'BrtRangeProtectionIso14' */
			case 0x0400: /* 'BrtRwDescent' */
			case 0x0098: /* 'BrtSel' */
			case 0x0297: /* 'BrtSheetCalcProp' */
			case 0x0217: /* 'BrtSheetProtection' */
			case 0x02A6: /* 'BrtSheetProtectionIso' */
			case 0x01F8: /* 'BrtSlc' */
			case 0x0413: /* 'BrtSparkline' */
			case 0x01AC: /* 'BrtTable' */
			case 0x00AA: /* 'BrtTop10Filter' */
			case 0x0C00: /* 'BrtUid' */
			case 0x0032: /* 'BrtValueMeta' */
			case 0x0816: /* 'BrtWebExtension' */
			case 0x0415: /* 'BrtWsFmtInfoEx14' */
				break;

			case 0x0023: /* 'BrtFRTBegin' */
				pass = true; break;
			case 0x0024: /* 'BrtFRTEnd' */
				pass = false; break;
			case 0x0025: /* 'BrtACBegin' */ break;
			case 0x0026: /* 'BrtACEnd' */ break;

			default:
				if((R_n||"").indexOf("Begin") > 0){/* empty */}
				else if((R_n||"").indexOf("End") > 0){/* empty */}
				else if(!pass || opts.WTF) throw new Error("Unexpected record " + RT + " " + R_n);
		}
	}, opts);

	delete opts.supbooks;
	delete opts['!row'];

	if(!s["!ref"] && (refguess.s.r < 2000000 || ref && (ref.e.r > 0 || ref.e.c > 0 || ref.s.r > 0 || ref.s.c > 0))) s["!ref"] = encode_range(ref || refguess);
	if(opts.sheetRows && s["!ref"]) {
		var tmpref = safe_decode_range(s["!ref"]);
		if(opts.sheetRows <= +tmpref.e.r) {
			tmpref.e.r = opts.sheetRows - 1;
			if(tmpref.e.r > refguess.e.r) tmpref.e.r = refguess.e.r;
			if(tmpref.e.r < tmpref.s.r) tmpref.s.r = tmpref.e.r;
			if(tmpref.e.c > refguess.e.c) tmpref.e.c = refguess.e.c;
			if(tmpref.e.c < tmpref.s.c) tmpref.s.c = tmpref.e.c;
			s["!fullref"] = s["!ref"];
			s["!ref"] = encode_range(tmpref);
		}
	}
	if(merges.length > 0) s["!merges"] = merges;
	if(colinfo.length > 0) s["!cols"] = colinfo;
	if(rowinfo.length > 0) s["!rows"] = rowinfo;
	return s;
}

/* TODO: something useful -- this is a stub */
function write_ws_bin_cell(ba/*:BufArray*/, cell/*:Cell*/, R/*:number*/, C/*:number*/, opts, ws/*:Worksheet*/) {
	if(cell.v === undefined) return "";
	var vv = "";
	switch(cell.t) {
		case 'b': vv = cell.v ? "1" : "0"; break;
		case 'd': // no BrtCellDate :(
			cell = dup(cell);
			cell.z = cell.z || SSF._table[14];
			cell.v = datenum(parseDate(cell.v)); cell.t = 'n';
			break;
		/* falls through */
		case 'n': case 'e': vv = ''+cell.v; break;
		default: vv = cell.v; break;
	}
	var o/*:any*/ = ({r:R, c:C}/*:any*/);
	/* TODO: cell style */
	o.s = get_cell_style(opts.cellXfs, cell, opts);
	if(cell.l) ws['!links'].push([encode_cell(o), cell.l]);
	if(cell.c) ws['!comments'].push([encode_cell(o), cell.c]);
	switch(cell.t) {
		case 's': case 'str':
			if(opts.bookSST) {
				vv = get_sst_id(opts.Strings, (cell.v/*:any*/));
				o.t = "s"; o.v = vv;
				write_record(ba, "BrtCellIsst", write_BrtCellIsst(cell, o));
			} else {
				o.t = "str";
				write_record(ba, "BrtCellSt", write_BrtCellSt(cell, o));
			}
			return;
		case 'n':
			/* TODO: determine threshold for Real vs RK */
			if(cell.v == (cell.v | 0) && cell.v > -1000 && cell.v < 1000) write_record(ba, "BrtCellRk", write_BrtCellRk(cell, o));
			else write_record(ba, "BrtCellReal", write_BrtCellReal(cell, o));
			return;
		case 'b':
			o.t = "b";
			write_record(ba, "BrtCellBool", write_BrtCellBool(cell, o));
			return;
		case 'e': /* TODO: error */ o.t = "e"; break;
	}
	write_record(ba, "BrtCellBlank", write_BrtCellBlank(cell, o));
}

function write_CELLTABLE(ba, ws/*:Worksheet*/, idx/*:number*/, opts/*::, wb:Workbook*/) {
	var range = safe_decode_range(ws['!ref'] || "A1"), ref, rr = "", cols/*:Array<string>*/ = [];
	write_record(ba, 'BrtBeginSheetData');
	var dense = Array.isArray(ws);
	var cap = range.e.r;
	if(ws['!rows']) cap = Math.max(range.e.r, ws['!rows'].length - 1);
	for(var R = range.s.r; R <= cap; ++R) {
		rr = encode_row(R);
		/* [ACCELLTABLE] */
		/* BrtRowHdr */
		write_row_header(ba, ws, range, R);
		if(R <= range.e.r) for(var C = range.s.c; C <= range.e.c; ++C) {
			/* *16384CELL */
			if(R === range.s.r) cols[C] = encode_col(C);
			ref = cols[C] + rr;
			var cell = dense ? (ws[R]||[])[C] : ws[ref];
			if(!cell) continue;
			/* write cell */
			write_ws_bin_cell(ba, cell, R, C, opts, ws);
		}
	}
	write_record(ba, 'BrtEndSheetData');
}

function write_MERGECELLS(ba, ws/*:Worksheet*/) {
	if(!ws || !ws['!merges']) return;
	write_record(ba, 'BrtBeginMergeCells', write_BrtBeginMergeCells(ws['!merges'].length));
	ws['!merges'].forEach(function(m) { write_record(ba, 'BrtMergeCell', write_BrtMergeCell(m)); });
	write_record(ba, 'BrtEndMergeCells');
}

function write_COLINFOS(ba, ws/*:Worksheet*//*::, idx:number, opts, wb:Workbook*/) {
	if(!ws || !ws['!cols']) return;
	write_record(ba, 'BrtBeginColInfos');
	ws['!cols'].forEach(function(m, i) { if(m) write_record(ba, 'BrtColInfo', write_BrtColInfo(i, m)); });
	write_record(ba, 'BrtEndColInfos');
}

function write_IGNOREECS(ba, ws/*:Worksheet*/) {
	if(!ws || !ws['!ref']) return;
	write_record(ba, 'BrtBeginCellIgnoreECs');
	write_record(ba, 'BrtCellIgnoreEC', write_BrtCellIgnoreEC(safe_decode_range(ws['!ref'])));
	write_record(ba, 'BrtEndCellIgnoreECs');
}

function write_HLINKS(ba, ws/*:Worksheet*/, rels) {
	/* *BrtHLink */
	ws['!links'].forEach(function(l) {
		if(!l[1].Target) return;
		var rId = add_rels(rels, -1, l[1].Target.replace(/#.*$/, ""), RELS.HLINK);
		write_record(ba, "BrtHLink", write_BrtHLink(l, rId));
	});
	delete ws['!links'];
}
function write_LEGACYDRAWING(ba, ws/*:Worksheet*/, idx/*:number*/, rels) {
	/* [BrtLegacyDrawing] */
	if(ws['!comments'].length > 0) {
		var rId = add_rels(rels, -1, "../drawings/vmlDrawing" + (idx+1) + ".vml", RELS.VML);
		write_record(ba, "BrtLegacyDrawing", write_RelID("rId" + rId));
		ws['!legacy'] = rId;
	}
}

function write_AUTOFILTER(ba, ws) {
	if(!ws['!autofilter']) return;
	write_record(ba, "BrtBeginAFilter", write_UncheckedRfX(safe_decode_range(ws['!autofilter'].ref)));
	/* *FILTERCOLUMN */
	/* [SORTSTATE] */
	/* BrtEndAFilter */
	write_record(ba, "BrtEndAFilter");
}

function write_WSVIEWS2(ba, ws, Workbook) {
	write_record(ba, "BrtBeginWsViews");
	{ /* 1*WSVIEW2 */
		/* [ACUID] */
		write_record(ba, "BrtBeginWsView", write_BrtBeginWsView(ws, Workbook));
		/* [BrtPane] */
		/* *4BrtSel */
		/* *4SXSELECT */
		/* *FRT */
		write_record(ba, "BrtEndWsView");
	}
	/* *FRT */
	write_record(ba, "BrtEndWsViews");
}

function write_WSFMTINFO(/*::ba, ws*/) {
	/* [ACWSFMTINFO] */
	//write_record(ba, "BrtWsFmtInfo", write_BrtWsFmtInfo(ws));
}

function write_SHEETPROTECT(ba, ws) {
	if(!ws['!protect']) return;
	/* [BrtSheetProtectionIso] */
	write_record(ba, "BrtSheetProtection", write_BrtSheetProtection(ws['!protect']));
}

function write_ws_bin(idx/*:number*/, opts, wb/*:Workbook*/, rels) {
	var ba = buf_array();
	var s = wb.SheetNames[idx], ws = wb.Sheets[s] || {};
	var c/*:string*/ = s; try { if(wb && wb.Workbook) c = wb.Workbook.Sheets[idx].CodeName || c; } catch(e) {}
	var r = safe_decode_range(ws['!ref'] || "A1");
	ws['!links'] = [];
	/* passed back to write_zip and removed there */
	ws['!comments'] = [];
	write_record(ba, "BrtBeginSheet");
	if(wb.vbaraw) write_record(ba, "BrtWsProp", write_BrtWsProp(c));
	write_record(ba, "BrtWsDim", write_BrtWsDim(r));
	write_WSVIEWS2(ba, ws, wb.Workbook);
	write_WSFMTINFO(ba, ws);
	write_COLINFOS(ba, ws, idx, opts, wb);
	write_CELLTABLE(ba, ws, idx, opts, wb);
	/* [BrtSheetCalcProp] */
	write_SHEETPROTECT(ba, ws);
	/* *([BrtRangeProtectionIso] BrtRangeProtection) */
	/* [SCENMAN] */
	write_AUTOFILTER(ba, ws);
	/* [SORTSTATE] */
	/* [DCON] */
	/* [USERSHVIEWS] */
	write_MERGECELLS(ba, ws);
	/* [BrtPhoneticInfo] */
	/* *CONDITIONALFORMATTING */
	/* [DVALS] */
	write_HLINKS(ba, ws, rels);
	/* [BrtPrintOptions] */
	if(ws['!margins']) write_record(ba, "BrtMargins", write_BrtMargins(ws['!margins']));
	/* [BrtPageSetup] */
	/* [HEADERFOOTER] */
	/* [RWBRK] */
	/* [COLBRK] */
	/* *BrtBigName */
	/* [CELLWATCHES] */
	write_IGNOREECS(ba, ws);
	/* [SMARTTAGS] */
	/* [BrtDrawing] */
	write_LEGACYDRAWING(ba, ws, idx, rels);
	/* [BrtLegacyDrawingHF] */
	/* [BrtBkHim] */
	/* [OLEOBJECTS] */
	/* [ACTIVEXCONTROLS] */
	/* [WEBPUBITEMS] */
	/* [LISTPARTS] */
	/* FRTWORKSHEET */
	write_record(ba, "BrtEndSheet");
	return ba.end();
}
