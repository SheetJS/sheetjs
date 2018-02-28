function get_sheet_type(n/*:string*/)/*:string*/ {
	if(RELS.WS.indexOf(n) > -1) return "sheet";
	if(RELS.CS && n == RELS.CS) return "chart";
	if(RELS.DS && n == RELS.DS) return "dialog";
	if(RELS.MS && n == RELS.MS) return "macro";
	return (n && n.length) ? n : "sheet";
}
function safe_parse_wbrels(wbrels, sheets) {
	if(!wbrels) return 0;
	try {
		wbrels = sheets.map(function pwbr(w) { if(!w.id) w.id = w.strRelID; return [w.name, wbrels['!id'][w.id].Target, get_sheet_type(wbrels['!id'][w.id].Type)]; });
	} catch(e) { return null; }
	return !wbrels || wbrels.length === 0 ? null : wbrels;
}

function safe_parse_sheet(zip, path/*:string*/, relsPath/*:string*/, sheet, idx/*:number*/, sheetRels, sheets, stype/*:string*/, opts, wb, themes, styles) {
	try {
		sheetRels[sheet]=parse_rels(getzipstr(zip, relsPath, true), path);
		var data = getzipdata(zip, path);
		switch(stype) {
			case 'sheet': sheets[sheet]=parse_ws(data, path, idx, opts, sheetRels[sheet], wb, themes, styles); break;
			case 'chart':
				var cs = parse_cs(data, path, idx, opts, sheetRels[sheet], wb, themes, styles);
				sheets[sheet] = cs;
				if(!cs || !cs['!chart']) break;
				var dfile = resolve_path(cs['!chart'].Target, path);
				var drelsp = get_rels_path(dfile);
				var draw = parse_drawing(getzipstr(zip, dfile, true), parse_rels(getzipstr(zip, drelsp, true), dfile));
				var chartp = resolve_path(draw, dfile);
				var crelsp = get_rels_path(chartp);
				cs = parse_chart(getzipstr(zip, chartp, true), chartp, opts, parse_rels(getzipstr(zip, crelsp, true), chartp), wb, cs);
				break;
			case 'macro': sheets[sheet]=parse_ms(data, path, idx, opts, sheetRels[sheet], wb, themes, styles); break;
			case 'dialog': sheets[sheet]=parse_ds(data, path, idx, opts, sheetRels[sheet], wb, themes, styles); break;
		}
	} catch(e) { if(opts.WTF) throw e; }
}

function strip_front_slash(x/*:string*/)/*:string*/ { return x.charAt(0) == '/' ? x.slice(1) : x; }

function parse_zip(zip/*:ZIP*/, opts/*:?ParseOpts*/)/*:Workbook*/ {
	make_ssf(SSF);
	opts = opts || {};
	fix_read_opts(opts);

	/* OpenDocument Part 3 Section 2.2.1 OpenDocument Package */
	if(safegetzipfile(zip, 'META-INF/manifest.xml')) return parse_ods(zip, opts);
	/* UOC */
	if(safegetzipfile(zip, 'objectdata.xml')) return parse_ods(zip, opts);
	/* Numbers */
	if(safegetzipfile(zip, 'Index/Document.iwa')) throw new Error('Unsupported NUMBERS file');

	var entries = zipentries(zip);
	var dir = parse_ct((getzipstr(zip, '[Content_Types].xml')/*:?any*/));
	var xlsb = false;
	var sheets, binname;
	if(dir.workbooks.length === 0) {
		binname = "xl/workbook.xml";
		if(getzipdata(zip,binname, true)) dir.workbooks.push(binname);
	}
	if(dir.workbooks.length === 0) {
		binname = "xl/workbook.bin";
		if(!getzipdata(zip,binname,true)) throw new Error("Could not find workbook");
		dir.workbooks.push(binname);
		xlsb = true;
	}
	if(dir.workbooks[0].slice(-3) == "bin") xlsb = true;

	var themes = ({}/*:any*/);
	var styles = ({}/*:any*/);
	if(!opts.bookSheets && !opts.bookProps) {
		strs = [];
		if(dir.sst) strs=parse_sst(getzipdata(zip, strip_front_slash(dir.sst)), dir.sst, opts);

		if(opts.cellStyles && dir.themes.length) themes = parse_theme(getzipstr(zip, dir.themes[0].replace(/^\//,''), true)||"",dir.themes[0], opts);

		if(dir.style) styles = parse_sty(getzipdata(zip, strip_front_slash(dir.style)), dir.style, themes, opts);
	}

	/*var externbooks = */dir.links.map(function(link) {
		return parse_xlink(getzipdata(zip, strip_front_slash(link)), link, opts);
	});

	var wb = parse_wb(getzipdata(zip, strip_front_slash(dir.workbooks[0])), dir.workbooks[0], opts);

	var props = {}, propdata = "";

	if(dir.coreprops.length) {
		propdata = getzipdata(zip, strip_front_slash(dir.coreprops[0]), true);
		if(propdata) props = parse_core_props(propdata);
		if(dir.extprops.length !== 0) {
			propdata = getzipdata(zip, strip_front_slash(dir.extprops[0]), true);
			if(propdata) parse_ext_props(propdata, props, opts);
		}
	}

	var custprops = {};
	if(!opts.bookSheets || opts.bookProps) {
		if (dir.custprops.length !== 0) {
			propdata = getzipstr(zip, strip_front_slash(dir.custprops[0]), true);
			if(propdata) custprops = parse_cust_props(propdata, opts);
		}
	}

	var out = ({}/*:any*/);
	if(opts.bookSheets || opts.bookProps) {
		if(wb.Sheets) sheets = wb.Sheets.map(function pluck(x){ return x.name; });
		else if(props.Worksheets && props.SheetNames.length > 0) sheets=props.SheetNames;
		if(opts.bookProps) { out.Props = props; out.Custprops = custprops; }
		if(opts.bookSheets && typeof sheets !== 'undefined') out.SheetNames = sheets;
		if(opts.bookSheets ? out.SheetNames : opts.bookProps) return out;
	}
	sheets = {};

	var deps = {};
	if(opts.bookDeps && dir.calcchain) deps=parse_cc(getzipdata(zip, strip_front_slash(dir.calcchain)),dir.calcchain,opts);

	var i=0;
	var sheetRels = ({}/*:any*/);
	var path, relsPath;

	{
		var wbsheets = wb.Sheets;
		props.Worksheets = wbsheets.length;
		props.SheetNames = [];
		for(var j = 0; j != wbsheets.length; ++j) {
			props.SheetNames[j] = wbsheets[j].name;
		}
	}

	var wbext = xlsb ? "bin" : "xml";
	var wbrelsfile = 'xl/_rels/workbook.' + wbext + '.rels';
	var wbrels = parse_rels(getzipstr(zip, wbrelsfile, true), wbrelsfile);
	if(wbrels) wbrels = safe_parse_wbrels(wbrels, wb.Sheets);
	/* Numbers iOS hack */
	var nmode = (getzipdata(zip,"xl/worksheets/sheet.xml",true))?1:0;
	for(i = 0; i != props.Worksheets; ++i) {
		var stype = "sheet";
		if(wbrels && wbrels[i]) {
			path = 'xl/' + (wbrels[i][1]).replace(/[\/]?xl\//, "");
			stype = wbrels[i][2];
		} else {
			path = 'xl/worksheets/sheet'+(i+1-nmode)+"." + wbext;
			path = path.replace(/sheet0\./,"sheet.");
		}
		relsPath = path.replace(/^(.*)(\/)([^\/]*)$/, "$1/_rels/$3.rels");
		safe_parse_sheet(zip, path, relsPath, props.SheetNames[i], i, sheetRels, sheets, stype, opts, wb, themes, styles);
	}

	if(dir.comments) parse_comments(zip, dir.comments, sheets, sheetRels, opts);

	out = ({
		Directory: dir,
		Workbook: wb,
		Props: props,
		Custprops: custprops,
		Deps: deps,
		Sheets: sheets,
		SheetNames: props.SheetNames,
		Strings: strs,
		Styles: styles,
		Themes: themes,
		SSF: SSF.get_table()
	}/*:any*/);
	if(opts.bookFiles) {
		out.keys = entries;
		out.files = zip.files;
	}
	if(opts.bookVBA) {
		if(dir.vba.length > 0) out.vbaraw = getzipdata(zip,strip_front_slash(dir.vba[0]),true);
		else if(dir.defaults && dir.defaults.bin === CT_VBA) out.vbaraw = getzipdata(zip, 'xl/vbaProject.bin',true);
	}
	return out;
}

/* [MS-OFFCRYPTO] 2.1.1 */
function parse_xlsxcfb(cfb, _opts/*:?ParseOpts*/)/*:Workbook*/ {
	var opts = _opts || {};
	var f = '/!DataSpaces/Version';
	var data = CFB.find(cfb, f); if(!data || !data.content) throw new Error("ECMA-376 Encrypted file missing " + f);
	/*var version = */parse_DataSpaceVersionInfo(data.content);

	/* 2.3.4.1 */
	f = '/!DataSpaces/DataSpaceMap';
	data = CFB.find(cfb, f); if(!data || !data.content) throw new Error("ECMA-376 Encrypted file missing " + f);
	var dsm = parse_DataSpaceMap(data.content);
	if(dsm.length !== 1 || dsm[0].comps.length !== 1 || dsm[0].comps[0].t !== 0 || dsm[0].name !== "StrongEncryptionDataSpace" || dsm[0].comps[0].v !== "EncryptedPackage")
		throw new Error("ECMA-376 Encrypted file bad " + f);

	/* 2.3.4.2 */
	f = '/!DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace';
	data = CFB.find(cfb, f); if(!data || !data.content) throw new Error("ECMA-376 Encrypted file missing " + f);
	var seds = parse_DataSpaceDefinition(data.content);
	if(seds.length != 1 || seds[0] != "StrongEncryptionTransform")
		throw new Error("ECMA-376 Encrypted file bad " + f);

	/* 2.3.4.3 */
	f = '/!DataSpaces/TransformInfo/StrongEncryptionTransform/!Primary';
	data = CFB.find(cfb, f); if(!data || !data.content) throw new Error("ECMA-376 Encrypted file missing " + f);
	/*var hdr = */parse_Primary(data.content);

	f = '/EncryptionInfo';
	data = CFB.find(cfb, f); if(!data || !data.content) throw new Error("ECMA-376 Encrypted file missing " + f);
	var einfo = parse_EncryptionInfo(data.content);

	/* 2.3.4.4 */
	f = '/EncryptedPackage';
	data = CFB.find(cfb, f); if(!data || !data.content) throw new Error("ECMA-376 Encrypted file missing " + f);

/*global decrypt_agile */
/*:: declare var decrypt_agile:any; */
	if(einfo[0] == 0x04 && typeof decrypt_agile !== 'undefined') return decrypt_agile(einfo[1], data.content, opts.password || "", opts);
/*global decrypt_std76 */
/*:: declare var decrypt_std76:any; */
	if(einfo[0] == 0x02 && typeof decrypt_std76 !== 'undefined') return decrypt_std76(einfo[1], data.content, opts.password || "", opts);
	throw new Error("File is password-protected");
}

