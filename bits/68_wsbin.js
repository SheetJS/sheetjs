
/* [MS-XLSB] 2.4.718 BrtRowHdr */
function parse_BrtRowHdr(data, length) {
	var z = [];
	z.r = data.read_shift(4);
	data.l += length-4;
	return z;
}

/* [MS-XLSB] 2.4.812 BrtWsDim */
var parse_BrtWsDim = parse_UncheckedRfX;
var write_BrtWsDim = write_UncheckedRfX;

/* [MS-XLSB] 2.4.815 BrtWsProp */
function parse_BrtWsProp(data, length) {
	var z = {};
	/* TODO: pull flags */
	data.l += 19;
	z.name = parse_CodeName(data, length - 19);
	return z;
}

/* [MS-XLSB] 2.4.303 BrtCellBlank */
function parse_BrtCellBlank(data, length) {
	var cell = parse_Cell(data);
	return [cell];
}
function write_BrtCellBlank(cell, val, o) {
	if(o == null) o = new_buf(8);
	return write_Cell(val, o);
}


/* [MS-XLSB] 2.4.304 BrtCellBool */
function parse_BrtCellBool(data, length) {
	var cell = parse_Cell(data);
	var fBool = data.read_shift(1);
	return [cell, fBool, 'b'];
}

/* [MS-XLSB] 2.4.305 BrtCellError */
function parse_BrtCellError(data, length) {
	var cell = parse_Cell(data);
	var fBool = data.read_shift(1);
	return [cell, fBool, 'e'];
}

/* [MS-XLSB] 2.4.308 BrtCellIsst */
function parse_BrtCellIsst(data, length) {
	var cell = parse_Cell(data);
	var isst = data.read_shift(4);
	return [cell, isst, 's'];
}

/* [MS-XLSB] 2.4.310 BrtCellReal */
function parse_BrtCellReal(data, length) {
	var cell = parse_Cell(data);
	var value = parse_Xnum(data);
	return [cell, value, 'n'];
}

/* [MS-XLSB] 2.4.311 BrtCellRk */
function parse_BrtCellRk(data, length) {
	var cell = parse_Cell(data);
	var value = parse_RkNumber(data);
	return [cell, value, 'n'];
}

/* [MS-XLSB] 2.4.314 BrtCellSt */
function parse_BrtCellSt(data, length) {
	var cell = parse_Cell(data);
	var value = parse_XLWideString(data);
	return [cell, value, 'str'];
}

/* [MS-XLSB] 2.4.647 BrtFmlaBool */
function parse_BrtFmlaBool(data, length, opts) {
	var cell = parse_Cell(data);
	var value = data.read_shift(1);
	var o = [cell, value, 'b'];
	if(opts.cellFormula) {
		var formula = parse_CellParsedFormula(data, length-9);
		o[3] = ""; /* TODO */
	}
	else data.l += length-9;
	return o;
}

/* [MS-XLSB] 2.4.648 BrtFmlaError */
function parse_BrtFmlaError(data, length, opts) {
	var cell = parse_Cell(data);
	var value = data.read_shift(1);
	var o = [cell, value, 'e'];
	if(opts.cellFormula) {
		var formula = parse_CellParsedFormula(data, length-9);
		o[3] = ""; /* TODO */
	}
	else data.l += length-9;
	return o;
}

/* [MS-XLSB] 2.4.649 BrtFmlaNum */
function parse_BrtFmlaNum(data, length, opts) {
	var cell = parse_Cell(data);
	var value = parse_Xnum(data);
	var o = [cell, value, 'n'];
	if(opts.cellFormula) {
		var formula = parse_CellParsedFormula(data, length - 16);
		o[3] = ""; /* TODO */
	}
	else data.l += length-16;
	return o;
}

/* [MS-XLSB] 2.4.650 BrtFmlaString */
function parse_BrtFmlaString(data, length, opts) {
	var start = data.l;
	var cell = parse_Cell(data);
	var value = parse_XLWideString(data);
	var o = [cell, value, 'str'];
	if(opts.cellFormula) {
		var formula = parse_CellParsedFormula(data, start + length - data.l);
	}
	else data.l = start + length;
	return o;
}

/* [MS-XLSB] 2.4.676 BrtMergeCell */
var parse_BrtMergeCell = parse_UncheckedRfX;

/* [MS-XLSB] 2.4.656 BrtHLink */
function parse_BrtHLink(data, length, opts) {
	var end = data.l + length;
	var rfx = parse_UncheckedRfX(data, 16);
	var relId = parse_XLNullableWideString(data);
	var loc = parse_XLWideString(data);
	var tooltip = parse_XLWideString(data);
	var display = parse_XLWideString(data);
	data.l = end;
	return {rfx:rfx, relId:relId, loc:loc, tooltip:tooltip, display:display};
}

/* [MS-XLSB] 2.1.7.61 Worksheet */
function parse_ws_bin(data, opts, rels) {
	if(!data) return data;
	if(!rels) rels = {'!id':{}};
	var s = {};

	var ref;
	var refguess = {s: {r:1000000, c:1000000}, e: {r:0, c:0} };

	var pass = false, end = false;
	var row, p, cf, R, C, addr, sstr, rr;
	var mergecells = [];
	recordhopper(data, function ws_parse(val, R) {
		if(end) return;
		switch(R.n) {
			case 'BrtWsDim': ref = val; break;
			case 'BrtRowHdr':
				row = val;
				if(opts.sheetRows && opts.sheetRows <= row.r) end=true;
				rr = encode_row(row.r);
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
				p = {t:val[2]};
				switch(val[2]) {
					case 'n': p.v = val[1]; break;
					case 's': sstr = strs[val[1]]; p.v = sstr.t; p.r = sstr.r; break;
					case 'b': p.v = val[1] ? true : false; break;
					case 'e': p.v = val[1]; p.w = BErr[p.v]; break;
					case 'str': p.t = 's'; p.v = utf8read(val[1]); break;
				}
				if(opts.cellFormula && val.length > 3) p.f = val[3];
				if((cf = styles.CellXf[val[0].iStyleRef])) safe_format(p,cf.ifmt,null,opts);
				s[encode_col(C=val[0].c) + rr] = p;
				if(refguess.s.r > row.r) refguess.s.r = row.r;
				if(refguess.s.c > C) refguess.s.c = C;
				if(refguess.e.r < row.r) refguess.e.r = row.r;
				if(refguess.e.c < C) refguess.e.c = C;
				break;

			case 'BrtCellBlank': if(!opts.sheetStubs) break;
				p = {t:'s',v:undefined};
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

			case 'BrtArrFmla': break; // TODO
			case 'BrtShrFmla': break; // TODO
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
			case 'BrtColInfo': break; // TODO
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

			default: if(!pass || opts.WTF) throw new Error("Unexpected record " + R.n);
		}
	}, opts);
	if(!s["!ref"] && (refguess.s.r < 1000000 || ref.e.r > 0 || ref.e.c > 0 || ref.s.r > 0 || ref.s.c > 0)) s["!ref"] = encode_range(ref);
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
	return s;
}

/* TODO: something useful -- this is a stub */
function write_ws_bin_cell(ba, cell, R, C, opts) {
	if(cell.v === undefined) return "";
	var vv = "";
	switch(cell.t) {
		case 'b': vv = cell.v ? "1" : "0"; break;
		case 'n': case 'e': vv = ''+cell.v; break;
		default: vv = cell.v; break;
	}
	var o = {r:R, c:C};
	/* TODO: cell style */
	o.s = get_cell_style(opts.cellXfs, cell, opts);
	switch(cell.t) {
		case 's': case 'str':
			if(opts.bookSST) {
				vv = get_sst_id(opts.Strings, cell.v);
				o.t = "s"; break;
			}
			o.t = "str"; break;
		case 'n': break;
		case 'b': o.t = "b"; break;
		case 'e': o.t = "e"; break;
	}
	write_record(ba, "BrtCellBlank", write_BrtCellBlank(cell, o));
}

function write_CELLTABLE(ba, ws, idx, opts, wb) {
	var range = safe_decode_range(ws['!ref'] || "A1"), ref, rr = "", cols = [];
	write_record(ba, 'BrtBeginSheetData');
	for(var R = range.s.r; R <= range.e.r; ++R) {
		rr = encode_row(R);
		/* [ACCELLTABLE] */
		/* BrtRowHdr */
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

function write_ws_bin(idx, opts, wb) {
	var ba = buf_array();
	var s = wb.SheetNames[idx], ws = wb.Sheets[s] || {};
	var r = safe_decode_range(ws['!ref'] || "A1");
	write_record(ba, "BrtBeginSheet");
	/* [BrtWsProp] */
	write_record(ba, "BrtWsDim", write_BrtWsDim(r));
	/* [WSVIEWS2] */
	/* [WSFMTINFO] */
	/* *COLINFOS */
	write_CELLTABLE(ba, ws, idx, opts, wb);
	/* [BrtSheetCalcProp] */
	/* [[BrtSheetProtectionIso] BrtSheetProtection] */
	/* *([BrtRangeProtectionIso] BrtRangeProtection) */
	/* [SCENMAN] */
	/* [AUTOFILTER] */
	/* [SORTSTATE] */
	/* [DCON] */
	/* [USERSHVIEWS] */
	/* [MERGECELLS] */
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
