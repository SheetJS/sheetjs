
/* [MS-XLSB] 2.4.718 BrtRowHdr */
var parse_BrtRowHdr = function(data, length) {
	var z = {};
	z.r = data.read_shift(4);
	data.l += length-4;
	return z;
};

/* [MS-XLSB] 2.4.812 BrtWsDim */
var parse_BrtWsDim = parse_UncheckedRfX;

/* [MS-XLSB] 2.4.815 BrtWsProp */
var parse_BrtWsProp = function(data, length) {
	var z = {};
	/* TODO: pull flags */
	data.l += 19;
	z.name = parse_CodeName(data, length - 19);
	return z;
};

/* [MS-XLSB] 2.4.303 BrtCellBlank */
var parse_BrtCellBlank = parsenoop;

/* [MS-XLSB] 2.4.304 BrtCellBool */
var parse_BrtCellBool = function(data, length) {
	var cell = parse_Cell(data);
	var fBool = data.read_shift(1);
	return [cell, fBool, 'b'];
};

/* [MS-XLSB] 2.4.305 BrtCellError */
var parse_BrtCellError = function(data, length) {
	var cell = parse_Cell(data);
	var fBool = data.read_shift(1);
	return [cell, fBool, 'e'];
};

/* [MS-XLSB] 2.4.308 BrtCellIsst */
var parse_BrtCellIsst = function(data, length) {
	var cell = parse_Cell(data);
	var isst = data.read_shift(4);
	return [cell, isst, 's'];
};

/* [MS-XLSB] 2.4.310 BrtCellReal */
var parse_BrtCellReal = function(data, length) {
	var cell = parse_Cell(data);
	var value = parse_Xnum(data);
	return [cell, value, 'n'];
};

/* [MS-XLSB] 2.4.311 BrtCellRk */
var parse_BrtCellRk = function(data, length) {
	var cell = parse_Cell(data);
	var value = parse_RkNumber(data);
	return [cell, value, 'n'];
};

/* [MS-XLSB] 2.4.314 BrtCellSt */
var parse_BrtCellSt = function(data, length) {
	var cell = parse_Cell(data);
	var value = parse_XLWideString(data);
	return [cell, value, 'str'];
};

/* [MS-XLSB] 2.4.647 BrtFmlaBool */
var parse_BrtFmlaBool = function(data, length, opts) {
	var cell = parse_Cell(data);
	var value = data.read_shift(1);
	var o = [cell, value, 'b'];
	if(opts.cellFormula) {
		var formula = parse_CellParsedFormula(data, length-9);
		o[3] = ""; /* TODO */
	}
	else data.l += length-9;
	return o;
};

/* [MS-XLSB] 2.4.648 BrtFmlaError */
var parse_BrtFmlaError = function(data, length, opts) {
	var cell = parse_Cell(data);
	var value = data.read_shift(1);
	var o = [cell, value, 'e'];
	if(opts.cellFormula) {
		var formula = parse_CellParsedFormula(data, length-9);
		o[3] = ""; /* TODO */
	}
	else data.l += length-9;
	return o;
};

/* [MS-XLSB] 2.4.649 BrtFmlaNum */
var parse_BrtFmlaNum = function(data, length, opts) {
	var cell = parse_Cell(data);
	var value = parse_Xnum(data);
	var o = [cell, value, 'n'];
	if(opts.cellFormula) {
		var formula = parse_CellParsedFormula(data, length - 16);
		o[3] = ""; /* TODO */
	}
	else data.l += length-16;
	return o;
};

/* [MS-XLSB] 2.4.650 BrtFmlaString */
var parse_BrtFmlaString = function(data, length, opts) {
	var start = data.l;
	var cell = parse_Cell(data);
	var value = parse_XLWideString(data);
	var o = [cell, value, 'str'];
	if(opts.cellFormula) {
		var formula = parse_CellParsedFormula(data, start + length - data.l);
		o[3] = ""; /* TODO */
	}
	else data.l = start + length;
	return o;
};

/* [MS-XLSB] 2.1.7.61 Worksheet */
var parse_ws_bin = function(data, opts) {
	if(!data) return data;
	var s = {};

	var ref;
	var refguess = {s: {r:1000000, c:1000000}, e: {r:0, c:0} };

	var pass = false, end = false;
	var row, p, cf;
	recordhopper(data, function(val, R) {
		if(end) return;
		switch(R.n) {
			case 'BrtWsDim': ref = val; break;
			case 'BrtRowHdr':
				row = val;
				if(opts.sheetRows && opts.sheetRows <= row.r) end=true;
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
					case 's': p.v = strs[val[1]].t; p.r = strs[val[1]].r; break;
					case 'b': p.v = val[1] ? true : false; break;
					case 'e': p.raw = val[1]; p.v = BErr[p.raw]; break;
					case 'str': p.v = utf8read(val[1]); break;
				}
				if(opts.cellFormula && val.length > 3) p.f = val[3];
				if((cf = styles.CellXf[val[0].iStyleRef])) try {
					p.w = SSF.format(cf.ifmt,p.v,_ssfopts);
					if(opts.cellNF) p.z = SSF._table[cf.ifmt];
				} catch(e) { if(opts.WTF) throw e; }
				s[encode_cell({c:val[0].c,r:row.r})] = p;
				if(refguess.s.r > row.r) refguess.s.r = row.r;
				if(refguess.s.c > val[0].c) refguess.s.c = val[0].c;
				if(refguess.e.r < row.r) refguess.e.r = row.r;
				if(refguess.e.c < val[0].c) refguess.e.c = val[0].c;
				break;

			case 'BrtCellBlank': break; // (blank cell)

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
			case 'BrtBeginMergeCells': break; // TODO
			case 'BrtMergeCell': break; // TODO
			case 'BrtEndMergeCells': break; // TODO
			case 'BrtHLink': break; // TODO
			case 'BrtDrawing': break; // TODO
			case 'BrtLegacyDrawing': break; // TODO
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

			default: if(!pass) throw new Error("Unexpected record " + R.n);
		}
	}, opts);
	s["!ref"] = encode_range(ref);
	if(opts.sheetRows) {
		var tmpref = decode_range(s["!ref"]);
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
	return s;
};

