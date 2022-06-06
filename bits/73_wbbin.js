/* [MS-XLSB] 2.4.304 BrtBundleSh */
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
	write_XLWideString(data.name.slice(0,31), o);
	return o.length > o.l ? o.slice(0, o.l) : o;
}

/* [MS-XLSB] 2.4.815 BrtWbProp */
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
		if(data.date1904) flags |= 0x01;
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

/* [MS-XLSB] 2.4.687 BrtName */
function parse_BrtName(data, length, opts) {
	var end = data.l + length;
	var flags = data.read_shift(4);
	data.l += 1; //var chKey = data.read_shift(1);
	var itab = data.read_shift(4);
	var name = parse_XLNameWideString(data);
	var formula = parse_XLSBNameParsedFormula(data, 0, opts);
	var comment = parse_XLNullableWideString(data);
	if(flags & 0x20) name = "_xlnm." + name;
	//if(0 /* fProc */) {
		// unusedstring1: XLNullableWideString
		// description: XLNullableWideString
		// helpTopic: XLNullableWideString
		// unusedstring2: XLNullableWideString
	//}
	data.l = end;
	var out = ({Name:name, Ptg:formula, Flags: flags}/*:any*/);
	if(itab < 0xFFFFFFF) out.Sheet = itab;
	if(comment) out.Comment = comment;
	return out;
}
function write_BrtName(name, wb) {
	var o = new_buf(9);
	var flags = 0;
	var dname = name.Name;
	if(XLSLblBuiltIn.indexOf(dname) > -1) { flags |= 0x20; dname = dname.slice(6); }
	o.write_shift(4, flags); // flags
	o.write_shift(1, 0); // chKey
	o.write_shift(4, name.Sheet == null ? 0xFFFFFFFF : name.Sheet);

	var arr = [
		o,
		write_XLWideString(dname),
		write_XLSBNameParsedFormula(name.Ref, wb)
	];
	if(name.Comment) arr.push(write_XLNullableWideString(name.Comment));
	else {
		var x = new_buf(4);
		x.write_shift(4, 0xFFFFFFFF);
		arr.push(x);
	}

	// if macro (flags & 0x0F):
	// write_shift(4, 0xFFFFFFFF);
	// write_XLNullableWideString(description)
	// write_XLNullableWideString(helpTopic)
	// write_shift(4, 0xFFFFFFFF);

	return bconcat(arr);
}

/* [MS-XLSB] 2.1.7.61 Workbook */
function parse_wb_bin(data, opts)/*:WorkbookFile*/ {
	var wb = { AppVersion:{}, WBProps:{}, WBView:[], Sheets:[], CalcPr:{}, xmlns: "" };
	var state/*:Array<string>*/ = [];
	var pass = false;

	if(!opts) opts = {};
	opts.biff = 12;

	var Names = [];
	var supbooks = ([[]]/*:any*/);
	supbooks.SheetNames = [];
	supbooks.XTI = [];

	XLSBRecordEnum[0x0010] = { n:"BrtFRTArchID$", f:parse_BrtFRTArchID$ };

	recordhopper(data, function hopper_wb(val, R, RT) {
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

			case 0x0817: /* 'BrtAbsPath15' */
			case 0x009E: /* 'BrtBookView' */
			case 0x008F: /* 'BrtBeginBundleShs' */
			case 0x0298: /* 'BrtBeginFnGroup' */
			case 0x0161: /* 'BrtBeginExternals' */
				break;

			/* case 'BrtModelTimeGroupingCalcCol' */
			case 0x0C00: /* 'BrtUid' */
			case 0x0C01: /* 'BrtRevisionPtr' */
			case 0x0216: /* 'BrtBookProtection' */
			case 0x02A5: /* 'BrtBookProtectionIso' */
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
				state.push(RT); pass = true; break;
			case 0x0024: /* 'BrtFRTEnd' */
				state.pop(); pass = false; break;
			case 0x0025: /* 'BrtACBegin' */
				state.push(RT); pass = true; break;
			case 0x0026: /* 'BrtACEnd' */
				state.pop(); pass = false; break;

			case 0x0010: /* 'BrtFRTArchID$' */ break;

			default:
				if(R.T){/* empty */}
				else if(!pass || (opts.WTF && state[state.length-1] != 0x0025 /* BrtACBegin */ && state[state.length-1] != 0x0023 /* BrtFRTBegin */)) throw new Error("Unexpected record 0x" + RT.toString(16));
		}
	}, opts);

	parse_wb_defaults(wb);

	// $FlowIgnore
	wb.Names = Names;

	(wb/*:any*/).supbooks = supbooks;
	return wb;
}

function write_BUNDLESHS(ba, wb/*::, opts*/) {
	write_record(ba, 0x008F /* BrtBeginBundleShs */);
	for(var idx = 0; idx != wb.SheetNames.length; ++idx) {
		var viz = wb.Workbook && wb.Workbook.Sheets && wb.Workbook.Sheets[idx] && wb.Workbook.Sheets[idx].Hidden || 0;
		var d = { Hidden: viz, iTabID: idx+1, strRelID: 'rId' + (idx+1), name: wb.SheetNames[idx] };
		write_record(ba, 0x009C /* BrtBundleSh */, write_BrtBundleSh(d));
	}
	write_record(ba, 0x0090 /* BrtEndBundleShs */);
}

/* [MS-XLSB] 2.4.649 BrtFileVersion */
function write_BrtFileVersion(data, o) {
	if(!o) o = new_buf(127);
	for(var i = 0; i != 4; ++i) o.write_shift(4, 0);
	write_XLWideString("SheetJS", o);
	write_XLWideString(XLSX.version, o);
	write_XLWideString(XLSX.version, o);
	write_XLWideString("7262", o);
	return o.length > o.l ? o.slice(0, o.l) : o;
}

/* [MS-XLSB] 2.4.301 BrtBookView */
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

function write_BOOKVIEWS(ba, wb/*::, opts*/) {
	/* required if hidden tab appears before visible tab */
	if(!wb.Workbook || !wb.Workbook.Sheets) return;
	var sheets = wb.Workbook.Sheets;
	var i = 0, vistab = -1, hidden = -1;
	for(; i < sheets.length; ++i) {
		if(!sheets[i] || !sheets[i].Hidden && vistab == -1) vistab = i;
		else if(sheets[i].Hidden == 1 && hidden == -1) hidden = i;
	}
	if(hidden > vistab) return;
	write_record(ba, 0x0087 /* BrtBeginBookViews */);
	write_record(ba, 0x009E /* BrtBookView */, write_BrtBookView(vistab));
	/* 1*(BrtBookView *FRT) */
	write_record(ba, 0x0088 /* BrtEndBookViews */);
}

function write_BRTNAMES(ba, wb) {
	if(!wb.Workbook || !wb.Workbook.Names) return;
	wb.Workbook.Names.forEach(function(name) { try {
		if(name.Flags & 0x0e) return; // TODO: macro name write
		write_record(ba, 0x0027 /* BrtName */, write_BrtName(name, wb));
	} catch(e) {
		console.error("Could not serialize defined name " + JSON.stringify(name));
	} });
}

function write_SELF_EXTERNS_xlsb(wb) {
	var L = wb.SheetNames.length;
	var o = new_buf(12 * L + 28);
	o.write_shift(4, L + 2);
	o.write_shift(4, 0); o.write_shift(4, -2); o.write_shift(4, -2); // workbook-level reference
	o.write_shift(4, 0); o.write_shift(4, -1); o.write_shift(4, -1); // #REF!...
	for(var i = 0; i < L; ++i) {
		o.write_shift(4, 0); o.write_shift(4, i); o.write_shift(4, i);
	}
	return o;
}
function write_EXTERNALS_xlsb(ba, wb) {
	write_record(ba, 0x0161 /* BrtBeginExternals */);
	write_record(ba, 0x0165 /* BrtSupSelf */);
	write_record(ba, 0x016A /* BrtExternSheet */, write_SELF_EXTERNS_xlsb(wb, 0));
	write_record(ba, 0x0162 /* BrtEndExternals */);
}

/* [MS-XLSB] 2.4.305 BrtCalcProp */
/*function write_BrtCalcProp(data, o) {
	if(!o) o = new_buf(26);
	o.write_shift(4,0); // force recalc
	o.write_shift(4,1);
	o.write_shift(4,0);
	write_Xnum(0, o);
	o.write_shift(-4, 1023);
	o.write_shift(1, 0x33);
	o.write_shift(1, 0x00);
	return o;
}*/

/* [MS-XLSB] 2.4.646 BrtFileRecover */
/*function write_BrtFileRecover(data, o) {
	if(!o) o = new_buf(1);
	o.write_shift(1,0);
	return o;
}*/

/* [MS-XLSB] 2.1.7.61 Workbook */
function write_wb_bin(wb, opts) {
	var ba = buf_array();
	write_record(ba, 0x0083 /* BrtBeginBook */);
	write_record(ba, 0x0080 /* BrtFileVersion */, write_BrtFileVersion());
	/* [[BrtFileSharingIso] BrtFileSharing] */
	write_record(ba, 0x0099 /* BrtWbProp */, write_BrtWbProp(wb.Workbook && wb.Workbook.WBProps || null));
	/* [ACABSPATH] */
	/* [[BrtBookProtectionIso] BrtBookProtection] */
	write_BOOKVIEWS(ba, wb, opts);
	write_BUNDLESHS(ba, wb, opts);
	/* [FNGROUP] */
	write_EXTERNALS_xlsb(ba, wb);
	if((wb.Workbook||{}).Names) write_BRTNAMES(ba, wb);
	/* write_record(ba, 0x009D BrtCalcProp, write_BrtCalcProp()); */
	/* [BrtOleSize] */
	/* *(BrtUserBookView *FRT) */
	/* [PIVOTCACHEIDS] */
	/* [BrtWbFactoid] */
	/* [SMARTTAGTYPES] */
	/* [BrtWebOpt] */
	/* write_record(ba, 0x009B BrtFileRecover, write_BrtFileRecover()); */
	/* [WEBPUBITEMS] */
	/* [CRERRS] */
	/* FRTWORKBOOK */
	write_record(ba, 0x0084 /* BrtEndBook */);

	return ba.end();
}
