function get_sheet_type(n) {
	if(RELS.WS.indexOf(n) > -1) return "sheet";
	if(RELS.CS && n == RELS.CS) return "chart";
	if(RELS.DS && n == RELS.DS) return "dialog";
	if(RELS.MS && n == RELS.MS) return "macro";
	if(!n || !n.length) return "sheet";
	return n;
}
function safe_parse_wbrels(wbrels, sheets) {
	if(!wbrels) return 0;
	try {
		wbrels = sheets.map(function pwbr(w) { if(!w.id) w.id = w.strRelID; return [w.name, wbrels['!id'][w.id].Target, get_sheet_type(wbrels['!id'][w.id].Type)]; });
	} catch(e) { return null; }
	return !wbrels || wbrels.length === 0 ? null : wbrels;
}

function safe_parse_sheet(zip, path/*:string*/, relsPath/*:string*/, sheet, sheetRels, sheets, stype/*:string*/, opts, wb, themes, styles) {
	try {
		sheetRels[sheet]=parse_rels(getzipstr(zip, relsPath, true), path);
		var data = getzipdata(zip, path);
		switch(stype) {
			case 'sheet': sheets[sheet]=parse_ws(data, path, opts,sheetRels[sheet], wb, themes, styles); break;
			case 'chart':
				var cs = parse_cs(data, path, opts,sheetRels[sheet], wb, themes, styles);
				sheets[sheet] = cs;
				if(!cs || !cs['!chart']) break;
				var dfile = resolve_path(cs['!chart'].Target, path);
				var drelsp = get_rels_path(dfile);
				var draw = parse_drawing(getzipstr(zip, dfile, true), parse_rels(getzipstr(zip,drelsp,true), dfile));
				var chartp = resolve_path(draw, dfile);
				var crelsp = get_rels_path(chartp);
				cs = parse_chart(getzipstr(zip, chartp, true), chartp, opts, parse_rels(getzipstr(zip, crelsp,true), chartp), wb, cs);
				break;
			case 'macro': sheets[sheet]=parse_ms(data, path, opts,sheetRels[sheet], wb, themes, styles); break;
			case 'dialog': sheets[sheet]=parse_ds(data, path, opts,sheetRels[sheet], wb, themes, styles); break;
		}
	} catch(e) { if(opts.WTF) throw e; }
}

var nodirs = function nodirs(x/*:string*/)/*:boolean*/{return x.slice(-1) != '/';};
function parse_zip(zip/*:ZIP*/, opts/*:?ParseOpts*/)/*:Workbook*/ {
	make_ssf(SSF);
	opts = opts || {};
	fix_read_opts(opts);
	reset_cp();

	/* OpenDocument Part 3 Section 2.2.1 OpenDocument Package */
	if(safegetzipfile(zip, 'META-INF/manifest.xml')) return parse_ods(zip, opts);
	/* UOC */
	if(safegetzipfile(zip, 'objectdata.xml')) return parse_ods(zip, opts);

	var entries = keys(zip.files).filter(nodirs).sort();
	var dir = parse_ct((getzipstr(zip, '[Content_Types].xml')/*:?any*/), opts);
	var xlsb = false;
	var sheets, binname;
	if(dir.workbooks.length === 0) {
		binname = "xl/workbook.xml";
		if(getzipdata(zip,binname, true)) dir.workbooks.push(binname);
	}
	if(dir.workbooks.length === 0) {
		binname = "xl/workbook.bin";
		if(!getzipfile(zip,binname,true)) throw new Error("Could not find workbook");
		dir.workbooks.push(binname);
		xlsb = true;
	}
	if(dir.workbooks[0].slice(-3) == "bin") xlsb = true;
	if(xlsb) set_cp(1200);

	var themes = ({}/*:any*/);
	var styles = ({}/*:any*/);
	if(!opts.bookSheets && !opts.bookProps) {
		strs = [];
		if(dir.sst) strs=parse_sst(getzipdata(zip, dir.sst.replace(/^\//,'')), dir.sst, opts);

		if(opts.cellStyles && dir.themes.length) themes = parse_theme(getzipstr(zip, dir.themes[0].replace(/^\//,''), true)||"",dir.themes[0], opts);

		if(dir.style) styles = parse_sty(getzipdata(zip, dir.style.replace(/^\//,'')),dir.style, themes, opts);
	}

	var wb = parse_wb(getzipdata(zip, dir.workbooks[0].replace(/^\//,'')), dir.workbooks[0], opts);

	var props = {}, propdata = "";

	if(dir.coreprops.length !== 0) {
		propdata = getzipstr(zip, dir.coreprops[0].replace(/^\//,''), true);
		if(propdata) props = parse_core_props(propdata);
		if(dir.extprops.length !== 0) {
			propdata = getzipstr(zip, dir.extprops[0].replace(/^\//,''), true);
			if(propdata) parse_ext_props(propdata, props);
		}
	}

	var custprops = {};
	if(!opts.bookSheets || opts.bookProps) {
		if (dir.custprops.length !== 0) {
			propdata = getzipstr(zip, dir.custprops[0].replace(/^\//,''), true);
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
	if(opts.bookDeps && dir.calcchain) deps=parse_cc(getzipdata(zip, dir.calcchain.replace(/^\//,'')),dir.calcchain,opts);

	var i=0;
	var sheetRels = ({}/*:any*/);
	var path, relsPath;
	if(!props.Worksheets) {
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
		safe_parse_sheet(zip, path, relsPath, props.SheetNames[i], sheetRels, sheets, stype, opts, wb, themes, styles);
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
		if(dir.vba.length > 0) out.vbaraw = getzipdata(zip,dir.vba[0].replace(/^\//,''),true);
		else if(dir.defaults && dir.defaults.bin === 'application/vnd.ms-office.vbaProject') out.vbaraw = getzipdata(zip,'xl/vbaProject.bin',true);
	}
	return out;
}
