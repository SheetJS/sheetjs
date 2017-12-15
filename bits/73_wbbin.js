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
function parse_BrtWbProp(data, length)/*:WBProps*/ {
	var o/*:WBProps*/ = ({}/*:any*/);
	var flags = data.read_shift(4);
	o.defaultThemeVersion = data.read_shift(4);
	var strName = (length > 8) ? parse_XLWideString(data) : "";
	if(strName.length > 0) o.CodeName = strName;
	o.autoCompressPictures = !!(flags & 0x10000);
	o.backupFile = !!(flags & 0x40);
	o.checkCompatibility = !!(flags & 0x1000);
	o.date1904 = !!(flags & 0x01);
	o.filterPrivacy = !!(flags & 0x08);
	o.hidePivotFieldList = !!(flags & 0x400);
	o.promptedSolutions = !!(flags & 0x10);
	o.publishItems = !!(flags & 0x800);
	o.refreshAllConnections = !!(flags & 0x40000);
	o.saveExternalLinkValues = !!(flags & 0x80);
	o.showBorderUnselectedTables = !!(flags & 0x04);
	o.showInkAnnotation = !!(flags & 0x20);
	o.showObjects = ["all", "placeholders", "none"][(flags >> 13) & 0x03];
	o.showPivotChartFilter = !!(flags & 0x8000);
	o.updateLinks = ["userSet", "never", "always"][(flags >> 8) & 0x03];
	return o;
}
function write_BrtWbProp(data/*:?WBProps*/, o) {
	if(!o) o = new_buf(72);
	var flags = 0;
	if(data) {
		/* TODO: mirror parse_BrtWbProp fields */
		if(data.filterPrivacy) flags |= 0x08;
	}
	o.write_shift(4, flags);
	o.write_shift(4, 0);
	write_XLSBCodeName(data && data.CodeName || "ThisWorkbook", o);
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
	var out = ({Name:name, Ptg:formula}/*:any*/);
	if(itab < 0xFFFFFFF) out.Sheet = itab;
	if(comment) out.Comment = comment;
	return out;
}

/* [MS-XLSB] 2.1.7.60 Workbook */
function parse_wb_bin(data, opts)/*:WorkbookFile*/ {
	var wb = { AppVersion:{}, WBProps:{}, WBView:[], Sheets:[], CalcPr:{}, xmlns: "" };
	var pass = false, z;

	if(!opts) opts = {};
	opts.biff = 12;

	var Names = [];
	var supbooks = ([[]]/*:any*/);
	supbooks.SheetNames = [];
	supbooks.XTI = [];

	recordhopper(data, function hopper_wb(val, R_n, RT) {
		switch(RT) {
			case 0x009C: /* 'BrtBundleSh' */
				supbooks.SheetNames.push(val.name);
				wb.Sheets.push(val); break;

			case 0x0099: /* 'BrtWbProp' */
				wb.WBProps = val; break;

			case 0x0027: /* 'BrtName' */
				if(val.Sheet != null) opts.SID = val.Sheet;
				val.Ref = stringify_formula(val.Ptg, null, null, supbooks, opts);
				delete opts.SID;
				delete val.Ptg;
				Names.push(val);
				break;
			case 0x040C: /* 'BrtNameExt' */ break;

			case 0x0165: /* 'BrtSupSelf' */
			case 0x0166: /* 'BrtSupSame' */
			case 0x0163: /* 'BrtSupBookSrc' */
			case 0x029B: /* 'BrtSupAddin' */
				if(!supbooks[0].length) supbooks[0] = [RT, val];
				else supbooks.push([RT, val]);
				supbooks[supbooks.length - 1].XTI = [];
				break;
			case 0x016A: /* 'BrtExternSheet' */
				if(supbooks.length === 0) { supbooks[0] = []; supbooks[0].XTI = []; }
				supbooks[supbooks.length - 1].XTI = supbooks[supbooks.length - 1].XTI.concat(val);
				supbooks.XTI = supbooks.XTI.concat(val);
				break;
			case 0x0169: /* 'BrtPlaceholderName' */
				break;

			/* case 'BrtModelTimeGroupingCalcCol' */
			/* case 'BrtRevisionPtr' */
			/* case 'BrtUid' */
			case 0x0817: /* 'BrtAbsPath15' */
			case 0x0216: /* 'BrtBookProtection' */
			case 0x02A5: /* 'BrtBookProtectionIso' */
			case 0x009E: /* 'BrtBookView' */
			case 0x009D: /* 'BrtCalcProp' */
			case 0x0262: /* 'BrtCrashRecErr' */
			case 0x0802: /* 'BrtDecoupledPivotCacheID' */
			case 0x009B: /* 'BrtFileRecover' */
			case 0x0224: /* 'BrtFileSharing' */
			case 0x02A4: /* 'BrtFileSharingIso' */
			case 0x0080: /* 'BrtFileVersion' */
			case 0x0299: /* 'BrtFnGroup' */
			case 0x0850: /* 'BrtModelRelationship' */
			case 0x084D: /* 'BrtModelTable' */
			case 0x0225: /* 'BrtOleSize' */
			case 0x0805: /* 'BrtPivotTableRef' */
			case 0x0254: /* 'BrtSmartTagType' */
			case 0x081C: /* 'BrtTableSlicerCacheID' */
			case 0x081B: /* 'BrtTableSlicerCacheIDs' */
			case 0x0822: /* 'BrtTimelineCachePivotCacheID' */
			case 0x018D: /* 'BrtUserBookView' */
			case 0x009A: /* 'BrtWbFactoid' */
			case 0x045D: /* 'BrtWbProp14' */
			case 0x0229: /* 'BrtWebOpt' */
			case 0x082B: /* 'BrtWorkBookPr15' */
				break;

			case 0x0023: /* 'BrtFRTBegin' */
				pass = true; break;
			case 0x0024: /* 'BrtFRTEnd' */
				pass = false; break;
			case 0x0025: /* 'BrtACBegin' */ break;
			case 0x0026: /* 'BrtACEnd' */ break;

			case 0x0010: /* 'BrtFRTArchID$' */ break;

			default:
				if((R_n||"").indexOf("Begin") > 0){/* empty */}
				else if((R_n||"").indexOf("End") > 0){/* empty */}
				else if(!pass || opts.WTF) throw new Error("Unexpected record " + RT + " " + R_n);
		}
	}, opts);

	parse_wb_defaults(wb);

	// $FlowIgnore
	wb.Names = Names;

	(wb/*:any*/).supbooks = supbooks;
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
	write_record(ba, "BrtWbProp", write_BrtWbProp(wb.Workbook && wb.Workbook.WBProps || null));
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
