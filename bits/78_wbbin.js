/* [MS-XLSB] 2.4.301 BrtBundleSh */
function parse_BrtBundleSh(data, length) {
	var z = {};
	z.hsState = data.read_shift(4); //ST_SheetState
	z.iTabID = data.read_shift(4);
	z.strRelID = parse_RelID(data,length-8);
	z.name = parse_XLWideString(data);
	return z;
}
function write_BrtBundleSh(data, o) {
	if(!o) o = new_buf(127);
	o.write_shift(4, data.hsState);
	o.write_shift(4, data.iTabID);
	write_RelID(data.strRelID, o);
	write_XLWideString(data.name.substr(0,31), o);
	return o;
}

/* [MS-XLSB] 2.4.807 BrtWbProp */
function parse_BrtWbProp(data, length) {
	data.read_shift(4);
	var dwThemeVersion = data.read_shift(4);
	var strName = (length > 8) ? parse_XLWideString(data) : "";
	return [dwThemeVersion, strName];
}
function write_BrtWbProp(data, o) {
	if(!o) o = new_buf(8);
	o.write_shift(4, 0);
	o.write_shift(4, 0);
	return o;
}

function parse_BrtFRTArchID$(data, length) {
	var o = {};
	data.read_shift(4);
	o.ArchID = data.read_shift(4);
	data.l += length - 8;
	return o;
}

/* [MS-XLSB] 2.1.7.60 Workbook */
function parse_wb_bin(data, opts) {
	var wb = { AppVersion:{}, WBProps:{}, WBView:[], Sheets:[], CalcPr:{}, xmlns: "" };
	var pass = false, z;

	recordhopper(data, function hopper_wb(val, R) {
		switch(R.n) {
			case 'BrtBundleSh': wb.Sheets.push(val); break;

			case 'BrtBeginBook': break;
			case 'BrtFileVersion': break;
			case 'BrtWbProp': break;
			case 'BrtACBegin': break;
			case 'BrtAbsPath15': break;
			case 'BrtACEnd': break;
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
			case 'BrtName': break;
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
			case 'BrtFRTBegin': pass = true; break;
			case 'BrtFRTArchID$': break;
			case 'BrtFRTEnd': pass = false; break;
			case 'BrtEndBook': break;
			default: if(!pass) throw new Error("Unexpected record " + R.n);
		}
	});

	parse_wb_defaults(wb);

	return wb;
}

/* [MS-XLSB] 2.1.7.60 Workbook */
function write_BUNDLESHS(ba, wb, opts) {
	write_record(ba, "BrtBeginBundleShs");
	for(var idx = 0; idx != wb.SheetNames.length; ++idx) {
		var d = { hsState: 0, iTabID: idx+1, strRelID: 'rId' + (idx+1), name: wb.SheetNames[idx] };
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
	return o;
}

/* [MS-XLSB] 2.1.7.60 Workbook */
function write_BOOKVIEWS(ba, wb, opts) {
	write_record(ba, "BrtBeginBookViews");
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
	write_record(ba, "BrtCalcProp", write_BrtCalcProp());
	/* [BrtOleSize] */
	/* *(BrtUserBookView *FRT) */
	/* [PIVOTCACHEIDS] */
	/* [BrtWbFactoid] */
	/* [SMARTTAGTYPES] */
	/* [BrtWebOpt] */
	write_record(ba, "BrtFileRecover", write_BrtFileRecover());
	/* [WEBPUBITEMS] */
	/* [CRERRS] */
	/* FRTWORKBOOK */
	write_record(ba, "BrtEndBook");

	return ba.end();
}
