
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

	var pass = false;
	var row, p, cf;
	recordhopper(data, function(val, R) {
		switch(R.n) {
			case 'BrtWsDim': ref = val; break;
			case 'BrtRowHdr': row = val; break;

			case 'BrtFmlaBool':
			case 'BrtFmlaError':
			case 'BrtFmlaNum':
			case 'BrtFmlaString':
			case 'BrtCellBool':
			case 'BrtCellError':
			case 'BrtCellIsst':
			case 'BrtCellReal':
			case 'BrtCellRk':
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
				break; // TODO

			case 'BrtCellBlank': break; // (blank cell)

			case 'BrtArrFmla': break; // TODO
			case 'BrtShrFmla': break; // TODO
			case 'BrtBeginSheet': break;
			case 'BrtWsProp': break; // TODO
			case 'BrtSheetCalcProp': break; // TODO
			case 'BrtBeginWsViews': break; // TODO
			case 'BrtBeginWsView': break; // TODO
			case 'BrtEndWsView': break; // TODO
			case 'BrtEndWsViews': break; // TODO
			case 'BrtSel': break; // TODO
			case 'BrtACBegin': break; // TODO
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
			case 'BrtLegacyDrawing': break; // TODO
			//default: if(!pass) throw new Error("Unexpected record " + R.n);
		}
	}, opts);
	s["!ref"] = encode_range(ref);
	return s;
};

