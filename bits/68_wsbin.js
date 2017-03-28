
/* [MS-XLSB] 2.4.718 BrtRowHdr */
function parse_BrtRowHdr(data, length) {
	var z = ([]/*:any*/);
	z.r = data.read_shift(4);
	data.l += length-4;
	return z;
}
function write_BrtRowHdr(R/*:number*/, range, ws) {
	var o = new_buf(17+8*16);
	o.write_shift(4, R);

	/* TODO: flags styles */
	o.write_shift(4, 0);
	o.write_shift(2, 0x0140);
	o.write_shift(2, 0);
	o.write_shift(1, 0);

	/* [MS-XLSB] 2.5.8 BrtColSpan explains the mechanism */
	var ncolspan = 0, lcs = o.l;
	o.l += 4;

	var caddr = {r:R, c:0};
	for(var i = 0; i < 16; ++i) {
		if(range.s.c > ((i+1) << 10) || range.e.c < (i << 10)) continue;
		var first = -1, last = -1;
		for(var j = (i<<10); j < ((i+1)<<10); ++j) {
			caddr.c = j;
			if(ws[encode_cell(caddr)]) { if(first < 0) first = j; last = j; }
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
	if(o.length > 17) write_record(ba, 'BrtRowHdr', o);
}

/* [MS-XLSB] 2.4.812 BrtWsDim */
var parse_BrtWsDim = parse_UncheckedRfX;
var write_BrtWsDim = write_UncheckedRfX;

/* [MS-XLSB] 2.4.815 BrtWsProp */
function parse_BrtWsProp(data, length) {
	var z = {};
	/* TODO: pull flags */
	data.l += 19;
	z.name = parse_XLSBCodeName(data, length - 19);
	return z;
}

/* [MS-XLSB] 2.4.303 BrtCellBlank */
function parse_BrtCellBlank(data, length) {
	var cell = parse_XLSBCell(data);
	return [cell];
}
function write_BrtCellBlank(cell, ncell, o) {
	if(o == null) o = new_buf(8);
	return write_XLSBCell(ncell, o);
}


/* [MS-XLSB] 2.4.304 BrtCellBool */
function parse_BrtCellBool(data, length) {
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

/* [MS-XLSB] 2.4.305 BrtCellError */
function parse_BrtCellError(data, length) {
	var cell = parse_XLSBCell(data);
	var bError = data.read_shift(1);
	return [cell, bError, 'e'];
}

/* [MS-XLSB] 2.4.308 BrtCellIsst */
function parse_BrtCellIsst(data, length) {
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

/* [MS-XLSB] 2.4.310 BrtCellReal */
function parse_BrtCellReal(data, length) {
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

/* [MS-XLSB] 2.4.311 BrtCellRk */
function parse_BrtCellRk(data, length) {
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


/* [MS-XLSB] 2.4.314 BrtCellSt */
function parse_BrtCellSt(data, length) {
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

/* [MS-XLSB] 2.4.647 BrtFmlaBool */
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

/* [MS-XLSB] 2.4.648 BrtFmlaError */
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

/* [MS-XLSB] 2.4.649 BrtFmlaNum */
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

/* [MS-XLSB] 2.4.650 BrtFmlaString */
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

/* [MS-XLSB] 2.4.676 BrtMergeCell */
var parse_BrtMergeCell = parse_UncheckedRfX;
var write_BrtMergeCell = write_UncheckedRfX;
/* [MS-XLSB] 2.4.108 BrtBeginMergeCells */
function write_BrtBeginMergeCells(cnt, o) {
	if(o == null) o = new_buf(4);
	o.write_shift(4, cnt);
	return o;
}

/* [MS-XLSB] 2.4.656 BrtHLink */
function parse_BrtHLink(data, length, opts) {
	var end = data.l + length;
	var rfx = parse_UncheckedRfX(data, 16);
	var relId = parse_XLNullableWideString(data);
	var loc = parse_XLWideString(data);
	var tooltip = parse_XLWideString(data);
	var display = parse_XLWideString(data);
	data.l = end;
	return {rfx:rfx, relId:relId, loc:loc, Tooltip:tooltip, display:display};
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

/* [MS-XLSB] 2.4.742 BrtShrFmla */
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
	o.write_shift(4, p.width * 256);
	o.write_shift(4, 0/*ixfe*/); // style
	o.write_shift(1, 2); // bit flag
	o.write_shift(1, 0); // bit flag
	return o;
}

/* [MS-XLSB] 2.1.7.61 Worksheet */
function parse_ws_bin(data, opts, rels, wb, themes, styles)/*:Worksheet*/ {
	if(!data) return data;
	if(!rels) rels = {'!id':{}};
	var s = {};

	var ref;
	var refguess = {s: {r:2000000, c:2000000}, e: {r:0, c:0} };

	var pass = false, end = false;
	var row, p, cf, R, C, addr, sstr, rr;
	var mergecells = [];
	if(!opts) opts = {};
	opts.biff = 12;
	opts['!row'] = 0;

	var ai = 0, af = false;

	var array_formulae = [];
	var shared_formulae = {};
	var supbooks = ([[]]/*:any*/);
	supbooks.sharedf = shared_formulae;
	supbooks.arrayf = array_formulae;
	opts.supbooks = supbooks;

	for(var i = 0; i < wb.Names['!names'].length; ++i) supbooks[0][i+1] = wb.Names[wb.Names['!names'][i]];

	var colinfo = [], rowinfo = [];
	var defwidth = 0, defheight = 0; // twips / MDW respectively
	var seencol = false;

	recordhopper(data, function ws_parse(val, Record) {
		if(end) return;
		switch(Record.n) {
			case 'BrtWsDim': ref = val; break;
			case 'BrtRowHdr':
				row = val;
				if(opts.sheetRows && opts.sheetRows <= row.r) end=true;
				rr = encode_row(row.r);
				opts['!row'] = row.r;
				break;

			case 'BrtFmlaBool':
			case 'BrtFmlaError':
			case 'BrtFmlaNum':
			case 'BrtFmlaString':
			case 'BrtCellBool':
			case 'BrtCellError':
			case 'BrtCellIsst':
			case 'BrtCellReal':
			case 'BrtCellRk':
			case 'BrtCellSt':
				p = ({t:val[2]}/*:any*/);
				switch(val[2]) {
					case 'n': p.v = val[1]; break;
					case 's': sstr = strs[val[1]]; p.v = sstr.t; p.r = sstr.r; break;
					case 'b': p.v = val[1] ? true : false; break;
					case 'e': p.v = val[1]; p.w = BErr[p.v]; break;
					case 'str': p.t = 's'; p.v = utf8read(val[1]); break;
				}
				if((cf = styles.CellXf[val[0].iStyleRef])) safe_format(p,cf.ifmt,null,opts, themes, styles);
				s[encode_col(C=val[0].c) + rr] = p;
				if(opts.cellFormula) {
					af = false;
					for(ai = 0; ai < array_formulae.length; ++ai) {
						var aii = array_formulae[ai];
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
				if(opts.cellDates && cf && p.t == 'n' && SSF.is_date(SSF._table[cf.ifmt])) {
					var _d = SSF.parse_date_code(p.v); if(_d) { p.t = 'd'; p.v = new Date(Date.UTC(_d.y, _d.m-1,_d.d,_d.H,_d.M,_d.S,_d.u)); }
				}
				break;

			case 'BrtCellBlank': if(!opts.sheetStubs) break;
				p = ({t:'z',v:undefined}/*:any*/);
				s[encode_col(C=val[0].c) + rr] = p;
				if(refguess.s.r > row.r) refguess.s.r = row.r;
				if(refguess.s.c > C) refguess.s.c = C;
				if(refguess.e.r < row.r) refguess.e.r = row.r;
				if(refguess.e.c < C) refguess.e.c = C;
				break;

			/* Merge Cells */
			case 'BrtBeginMergeCells': break;
			case 'BrtEndMergeCells': break;
			case 'BrtMergeCell': mergecells.push(val); break;

			case 'BrtHLink':
				var rel = rels['!id'][val.relId];
				if(rel) {
					val.Target = rel.Target;
					if(val.loc) val.Target += "#"+val.loc;
					val.Rel = rel;
				}
				for(R=val.rfx.s.r;R<=val.rfx.e.r;++R) for(C=val.rfx.s.c;C<=val.rfx.e.c;++C) {
					addr = encode_cell({c:C,r:R});
					if(!s[addr]) s[addr] = {t:'s',v:undefined};
					s[addr].l = val;
				}
				break;

			case 'BrtArrFmla': if(!opts.cellFormula) break;
				array_formulae.push(val);
				s[encode_col(C) + rr].f = stringify_formula(val[1], refguess, {r:row.r, c:C}, supbooks, opts);
				s[encode_col(C) + rr].F = encode_range(val[0]);
				break;
			case 'BrtShrFmla': if(!opts.cellFormula) break;
				// TODO
				shared_formulae[encode_cell(val[0].s)] = val[1];
				s[encode_col(C) + rr].f = stringify_formula(val[1], refguess, {r:row.r, c:C}, supbooks, opts);
				break;

			/* identical to 'ColInfo' in XLS */
			case 'BrtColInfo': {
				if(!opts.cellStyles) break;
				while(val.e >= val.s) {
					colinfo[val.e--] = { width: val.w/256 };
					if(!seencol) { seencol = true; find_mdw_colw(val.w/256); }
					process_col(colinfo[val.e+1]);
				}
			} break;

			case 'BrtBeginSheet': break;
			case 'BrtWsProp': break; // TODO
			case 'BrtSheetCalcProp': break; // TODO
			case 'BrtBeginWsViews': break; // TODO
			case 'BrtBeginWsView': break; // TODO
			case 'BrtPane': break; // TODO
			case 'BrtSel': break; // TODO
			case 'BrtEndWsView': break; // TODO
			case 'BrtEndWsViews': break; // TODO
			case 'BrtACBegin': break; // TODO
			case 'BrtRwDescent': break; // TODO
			case 'BrtACEnd': break; // TODO
			case 'BrtWsFmtInfoEx14': break; // TODO
			case 'BrtWsFmtInfo': break; // TODO
			case 'BrtBeginColInfos': break; // TODO
			case 'BrtEndColInfos': break; // TODO
			case 'BrtBeginSheetData': break; // TODO
			case 'BrtEndSheetData': break; // TODO
			case 'BrtSheetProtection': break; // TODO
			case 'BrtPrintOptions': break; // TODO
			case 'BrtMargins': break; // TODO
			case 'BrtPageSetup': break; // TODO
			case 'BrtFRTBegin': pass = true; break;
			case 'BrtFRTEnd': pass = false; break;
			case 'BrtEndSheet': break; // TODO
			case 'BrtDrawing': break; // TODO
			case 'BrtLegacyDrawing': break; // TODO
			case 'BrtLegacyDrawingHF': break; // TODO
			case 'BrtPhoneticInfo': break; // TODO
			case 'BrtBeginHeaderFooter': break; // TODO
			case 'BrtEndHeaderFooter': break; // TODO
			case 'BrtBrk': break; // TODO
			case 'BrtBeginRwBrk': break; // TODO
			case 'BrtEndRwBrk': break; // TODO
			case 'BrtBeginColBrk': break; // TODO
			case 'BrtEndColBrk': break; // TODO
			case 'BrtBeginUserShViews': break; // TODO
			case 'BrtBeginUserShView': break; // TODO
			case 'BrtEndUserShView': break; // TODO
			case 'BrtEndUserShViews': break; // TODO
			case 'BrtBkHim': break; // TODO
			case 'BrtBeginOleObjects': break; // TODO
			case 'BrtOleObject': break; // TODO
			case 'BrtEndOleObjects': break; // TODO
			case 'BrtBeginListParts': break; // TODO
			case 'BrtListPart': break; // TODO
			case 'BrtEndListParts': break; // TODO
			case 'BrtBeginSortState': break; // TODO
			case 'BrtBeginSortCond': break; // TODO
			case 'BrtEndSortCond': break; // TODO
			case 'BrtEndSortState': break; // TODO
			case 'BrtBeginConditionalFormatting': break; // TODO
			case 'BrtEndConditionalFormatting': break; // TODO
			case 'BrtBeginCFRule': break; // TODO
			case 'BrtEndCFRule': break; // TODO
			case 'BrtBeginDVals': break; // TODO
			case 'BrtDVal': break; // TODO
			case 'BrtEndDVals': break; // TODO
			case 'BrtRangeProtection': break; // TODO
			case 'BrtBeginDCon': break; // TODO
			case 'BrtEndDCon': break; // TODO
			case 'BrtBeginDRefs': break;
			case 'BrtDRef': break;
			case 'BrtEndDRefs': break;

			/* ActiveX */
			case 'BrtBeginActiveXControls': break;
			case 'BrtActiveX': break;
			case 'BrtEndActiveXControls': break;

			/* AutoFilter */
			case 'BrtBeginAFilter': break;
			case 'BrtEndAFilter': break;
			case 'BrtBeginFilterColumn': break;
			case 'BrtBeginFilters': break;
			case 'BrtFilter': break;
			case 'BrtEndFilters': break;
			case 'BrtEndFilterColumn': break;
			case 'BrtDynamicFilter': break;
			case 'BrtTop10Filter': break;
			case 'BrtBeginCustomFilters': break;
			case 'BrtCustomFilter': break;
			case 'BrtEndCustomFilters': break;

			/* Smart Tags */
			case 'BrtBeginSmartTags': break;
			case 'BrtBeginCellSmartTags': break;
			case 'BrtBeginCellSmartTag': break;
			case 'BrtCellSmartTagProperty': break;
			case 'BrtEndCellSmartTag': break;
			case 'BrtEndCellSmartTags': break;
			case 'BrtEndSmartTags': break;

			/* Cell Watch */
			case 'BrtBeginCellWatches': break;
			case 'BrtCellWatch': break;
			case 'BrtEndCellWatches': break;

			/* Table */
			case 'BrtTable': break;

			/* Ignore Cell Errors */
			case 'BrtBeginCellIgnoreECs': break;
			case 'BrtCellIgnoreEC': break;
			case 'BrtEndCellIgnoreECs': break;

			default: if(!pass || opts.WTF) throw new Error("Unexpected record " + Record.n);
		}
	}, opts);

	delete opts.supbooks;
	delete opts['!row'];

	if(!s["!ref"] && (refguess.s.r < 2000000 || ref && (ref.e.r > 0 || ref.e.c > 0 || ref.s.r > 0 || ref.s.c > 0))) s["!ref"] = encode_range(ref || refguess);
	if(opts.sheetRows && s["!ref"]) {
		var tmpref = safe_decode_range(s["!ref"]);
		if(opts.sheetRows < +tmpref.e.r) {
			tmpref.e.r = opts.sheetRows - 1;
			if(tmpref.e.r > refguess.e.r) tmpref.e.r = refguess.e.r;
			if(tmpref.e.r < tmpref.s.r) tmpref.s.r = tmpref.e.r;
			if(tmpref.e.c > refguess.e.c) tmpref.e.c = refguess.e.c;
			if(tmpref.e.c < tmpref.s.c) tmpref.s.c = tmpref.e.c;
			s["!fullref"] = s["!ref"];
			s["!ref"] = encode_range(tmpref);
		}
	}
	if(mergecells.length > 0) s["!merges"] = mergecells;
	if(colinfo.length > 0) s["!cols"] = colinfo;
	if(rowinfo.length > 0) s["!rows"] = rowinfo;
	return s;
}

/* TODO: something useful -- this is a stub */
function write_ws_bin_cell(ba/*:BufArray*/, cell/*:Cell*/, R/*:number*/, C/*:number*/, opts) {
	if(cell.v === undefined) return "";
	var vv = ""; var olddate = null;
	switch(cell.t) {
		case 'b': vv = cell.v ? "1" : "0"; break;
		case 'd': // no BrtCellDate :(
			cell.z = cell.z || SSF._table[14];
			olddate = cell.v;
			cell.v = datenum((cell.v/*:any*/)); cell.t = 'n';
			break;
		/* falls through */
		case 'n': case 'e': vv = ''+cell.v; break;
		default: vv = cell.v; break;
	}
	var o/*:any*/ = ({r:R, c:C}/*:any*/);
	/* TODO: cell style */
	//o.s = get_cell_style(opts.cellXfs, cell, opts);
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
			if(olddate) { cell.t = 'd'; cell.v = olddate; }
			return;
		case 'b':
			o.t = "b";
			write_record(ba, "BrtCellBool", write_BrtCellBool(cell, o));
			return;
		case 'e': /* TODO: error */ o.t = "e"; break;
	}
	write_record(ba, "BrtCellBlank", write_BrtCellBlank(cell, o));
}

function write_CELLTABLE(ba, ws/*:Worksheet*/, idx/*:number*/, opts, wb/*:Workbook*/) {
	var range = safe_decode_range(ws['!ref'] || "A1"), ref, rr = "", cols = [];
	write_record(ba, 'BrtBeginSheetData');
	for(var R = range.s.r; R <= range.e.r; ++R) {
		rr = encode_row(R);
		/* [ACCELLTABLE] */
		/* BrtRowHdr */
		write_row_header(ba, ws, range, R);
		for(var C = range.s.c; C <= range.e.c; ++C) {
			/* *16384CELL */
			if(R === range.s.r) cols[C] = encode_col(C);
			ref = cols[C] + rr;
			if(!ws[ref]) continue;
			/* write cell */
			write_ws_bin_cell(ba, ws[ref], R, C, opts);
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

function write_COLINFOS(ba, ws/*:Worksheet*/, idx/*:number*/, opts, wb/*:Workbook*/) {
	if(!ws || !ws['!cols']) return;
	write_record(ba, 'BrtBeginColInfos');
	ws['!cols'].forEach(function(m, i) { if(m) write_record(ba, 'BrtColInfo', write_BrtColInfo(i, m)); });
	write_record(ba, 'BrtEndColInfos');
}

function write_ws_bin(idx/*:number*/, opts, wb/*:Workbook*/) {
	var ba = buf_array();
	var s = wb.SheetNames[idx], ws = wb.Sheets[s] || {};
	var r = safe_decode_range(ws['!ref'] || "A1");
	write_record(ba, "BrtBeginSheet");
	/* [BrtWsProp] */
	write_record(ba, "BrtWsDim", write_BrtWsDim(r));
	/* [WSVIEWS2] */
	/* [WSFMTINFO] */
	write_COLINFOS(ba, ws, idx, opts, wb);
	write_CELLTABLE(ba, ws, idx, opts, wb);
	/* [BrtSheetCalcProp] */
	/* [[BrtSheetProtectionIso] BrtSheetProtection] */
	/* *([BrtRangeProtectionIso] BrtRangeProtection) */
	/* [SCENMAN] */
	/* [AUTOFILTER] */
	/* [SORTSTATE] */
	/* [DCON] */
	/* [USERSHVIEWS] */
	write_MERGECELLS(ba, ws);
	/* [BrtPhoneticInfo] */
	/* *CONDITIONALFORMATTING */
	/* [DVALS] */
	/* *BrtHLink */
	/* [BrtPrintOptions] */
	/* [BrtMargins] */
	/* [BrtPageSetup] */
	/* [HEADERFOOTER] */
	/* [RWBRK] */
	/* [COLBRK] */
	/* *BrtBigName */
	/* [CELLWATCHES] */
	/* [IGNOREECS] */
	/* [SMARTTAGS] */
	/* [BrtDrawing] */
	/* [BrtLegacyDrawing] */
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
