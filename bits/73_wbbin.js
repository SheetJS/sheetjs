/* [MS-XLSB] 2.4.301 BrtBundleSh */
function parse_BrtBundleSh(data, length/*:number*/) {
	var z = {};
	z.Hidden = data.read_shift(4); //hsState ST_SheetState
	z.iTabID = data.read_shift(4);
	z.strRelID = parse_RelID(data,length-8);
	z.name = parse_XLWideString(data);
	return z;
}
function write_BrtBundleSh(data, o) {
	if(!o) o = new_buf(127);
	o.write_shift(4, data.Hidden);
	o.write_shift(4, data.iTabID);
	write_RelID(data.strRelID, o);
	write_XLWideString(data.name.substr(0,31), o);
	return o.length > o.l ? o.slice(0, o.l) : o;
}

/* [MS-XLSB] 2.4.807 BrtWbProp */
function parse_BrtWbProp(data, length) {
	data.read_shift(4);
	var dwThemeVersion = data.read_shift(4);
	var strName = (length > 8) ? parse_XLWideString(data) : "";
	return [dwThemeVersion, strName];
}
function write_BrtWbProp(data, o) {
	if(!o) o = new_buf(68);
	o.write_shift(4, 0);
	o.write_shift(4, 0);
	write_XLSBCodeName("ThisWorkbook", o);
	return o.slice(0, o.l);
}

function parse_BrtFRTArchID$(data, length) {
	var o = {};
	data.read_shift(4);
	o.ArchID = data.read_shift(4);
	data.l += length - 8;
	return o;
}

/* [MS-XLSB] 2.4.680 BrtName */
function parse_BrtName(data, length, opts) {
	var end = data.l + length;
	var flags = data.read_shift(4);
	var chKey = data.read_shift(1);
	var itab = data.read_shift(4);
	var name = parse_XLNameWideString(data);
	var formula = parse_XLSBNameParsedFormula(data, 0, opts);
	var comment = parse_XLNullableWideString(data);
	//if(0 /* fProc */) {
		// unusedstring1: XLNullableWideString
		// description: XLNullableWideString
		// helpTopic: XLNullableWideString
		// unusedstring2: XLNullableWideString
	//}
	data.l = end;
	return {Name:name, Ptg:formula, Comment:comment};
}

/* [MS-XLSB] 2.1.7.60 Workbook */
function parse_wb_bin(data, opts)/*:WorkbookFile*/ {
	var wb = { AppVersion:{}, WBProps:{}, WBView:[], Sheets:[], CalcPr:{}, xmlns: "" };
	var pass = false, z;

	if(!opts) opts = {};
	opts.biff = 12;

	var Names = {}, NameList = [];

	recordhopper(data, function hopper_wb(val, R) {
		switch(R.n) {
			case 'BrtBundleSh': wb.Sheets.push(val); break;

			case 'BrtName':
				Names[val.Name] = val; NameList.push(val.Name);
				break;

			case 'BrtBeginBook': break;
			case 'BrtFileVersion': break;
			case 'BrtWbProp14': case 'BrtWbProp': break;
			case 'BrtACBegin': break;
			case 'BrtAbsPath15': break;
			case 'BrtACEnd': break;
			case 'BrtWbFactoid': break;
			/*case 'BrtBookProtectionIso': break;*/
			case 'BrtBookProtection': break;
			case 'BrtBeginBookViews': break;
			case 'BrtBookView': break;
			case 'BrtEndBookViews': break;
			case 'BrtBeginBundleShs': break;
			case 'BrtEndBundleShs': break;
			case 'BrtBeginFnGroup': break;
			case 'BrtEndFnGroup': break;
			case 'BrtBeginExternals': break;
			case 'BrtSupSelf': break;
			case 'BrtSupBookSrc': break;
			case 'BrtExternSheet': break;
			case 'BrtEndExternals': break;
			case 'BrtCalcProp': break;
			case 'BrtUserBookView': break;
			case 'BrtBeginPivotCacheIDs': break;
			case 'BrtBeginPivotCacheID': break;
			case 'BrtEndPivotCacheID': break;
			case 'BrtEndPivotCacheIDs': break;
			case 'BrtWebOpt': break;
			case 'BrtFileRecover': break;
			case 'BrtFileSharing': break;
			/*case 'BrtBeginWebPubItems': break;
			case 'BrtBeginWebPubItem': break;
			case 'BrtEndWebPubItem': break;
			case 'BrtEndWebPubItems': break;*/

			/* Smart Tags */
			case 'BrtBeginSmartTagTypes': break;
			case 'BrtSmartTagType': break;
			case 'BrtEndSmartTagTypes': break;

			case 'BrtFRTBegin': pass = true; break;
			case 'BrtFRTArchID$': break;
			case 'BrtWorkBookPr15': break;
			case 'BrtFRTEnd': pass = false; break;
			case 'BrtEndBook': break;
			default: if(!pass || opts.WTF) throw new Error("Unexpected record " + R.n);
		}
	}, opts);

	parse_wb_defaults(wb);

	Names['!names'] = NameList;
	// $FlowIgnore
	wb.Names = Names;

	return wb;
}

/* [MS-XLSB] 2.1.7.60 Workbook */
function write_BUNDLESHS(ba, wb, opts) {
	write_record(ba, "BrtBeginBundleShs");
	for(var idx = 0; idx != wb.SheetNames.length; ++idx) {
		var viz = wb.Workbook && wb.Workbook.Sheets && wb.Workbook.Sheets[idx] && wb.Workbook.Sheets[idx].Hidden || 0;
		var d = { Hidden: viz, iTabID: idx+1, strRelID: 'rId' + (idx+1), name: wb.SheetNames[idx] };
		write_record(ba, "BrtBundleSh", write_BrtBundleSh(d));
	}
	write_record(ba, "BrtEndBundleShs");
}

/* [MS-XLSB] 2.4.643 BrtFileVersion */
function write_BrtFileVersion(data, o) {
	if(!o) o = new_buf(127);
	for(var i = 0; i != 4; ++i) o.write_shift(4, 0);
	write_XLWideString("SheetJS", o);
	write_XLWideString(XLSX.version, o);
	write_XLWideString(XLSX.version, o);
	write_XLWideString("7262", o);
	o.length = o.l;
	return o.length > o.l ? o.slice(0, o.l) : o;
}

/* [MS-XLSB] 2.4.298 BrtBookView */
function write_BrtBookView(idx, o) {
	if(!o) o = new_buf(29);
	o.write_shift(-4, 0);
	o.write_shift(-4, 460);
	o.write_shift(4,  28800);
	o.write_shift(4,  17600);
	o.write_shift(4,  500);
	o.write_shift(4,  idx);
	o.write_shift(4,  idx);
	var flags = 0x78;
	o.write_shift(1,  flags);
	return o.length > o.l ? o.slice(0, o.l) : o;
}

/* [MS-XLSB] 2.1.7.60 Workbook */
function write_BOOKVIEWS(ba, wb, opts) {
	/* required if hidden tab appears before visible tab */
	if(!wb.Workbook || !wb.Workbook.Sheets) return;
	var sheets = wb.Workbook.Sheets;
	var i = 0, vistab = -1, hidden = -1;
	for(; i < sheets.length; ++i) {
		if(!sheets[i] || !sheets[i].Hidden && vistab == -1) vistab = i;
		else if(sheets[i].Hidden == 1 && hidden == -1) hidden = i;
	}
	if(hidden > vistab) return;
	write_record(ba, "BrtBeginBookViews");
	write_record(ba, "BrtBookView", write_BrtBookView(vistab));
	/* 1*(BrtBookView *FRT) */
	write_record(ba, "BrtEndBookViews");
}

/* [MS-XLSB] 2.4.302 BrtCalcProp */
function write_BrtCalcProp(data, o) {
	if(!o) o = new_buf(26);
	o.write_shift(4,0); /* force recalc */
	o.write_shift(4,1);
	o.write_shift(4,0);
	write_Xnum(0, o);
	o.write_shift(-4, 1023);
	o.write_shift(1, 0x33);
	o.write_shift(1, 0x00);
	return o;
}

/* [MS-XLSB] 2.4.640 BrtFileRecover */
function write_BrtFileRecover(data, o) {
	if(!o) o = new_buf(1);
	o.write_shift(1,0);
	return o;
}

/* [MS-XLSB] 2.1.7.60 Workbook */
function write_wb_bin(wb, opts) {
	var ba = buf_array();
	write_record(ba, "BrtBeginBook");
	write_record(ba, "BrtFileVersion", write_BrtFileVersion());
	/* [[BrtFileSharingIso] BrtFileSharing] */
	write_record(ba, "BrtWbProp", write_BrtWbProp());
	/* [ACABSPATH] */
	/* [[BrtBookProtectionIso] BrtBookProtection] */
	write_BOOKVIEWS(ba, wb, opts);
	write_BUNDLESHS(ba, wb, opts);
	/* [FNGROUP] */
	/* [EXTERNALS] */
	/* *BrtName */
	/* write_record(ba, "BrtCalcProp", write_BrtCalcProp()); */
	/* [BrtOleSize] */
	/* *(BrtUserBookView *FRT) */
	/* [PIVOTCACHEIDS] */
	/* [BrtWbFactoid] */
	/* [SMARTTAGTYPES] */
	/* [BrtWebOpt] */
	/* write_record(ba, "BrtFileRecover", write_BrtFileRecover()); */
	/* [WEBPUBITEMS] */
	/* [CRERRS] */
	/* FRTWORKBOOK */
	write_record(ba, "BrtEndBook");

	return ba.end();
}
