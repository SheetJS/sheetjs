function write_zip(wb/*:Workbook*/, opts/*:WriteOpts*/)/*:ZIP*/ {
	_shapeid = 1024;
	if(opts.bookType == "ods") return write_ods(wb, opts);
	if(wb && !wb.SSF) {
		wb.SSF = SSF.get_table();
	}
	if(wb && wb.SSF) {
		// $FlowIgnore
		make_ssf(SSF); SSF.load_table(wb.SSF);
		// $FlowIgnore
		opts.revssf = evert_num(wb.SSF); opts.revssf[wb.SSF[65535]] = 0;
		opts.ssf = wb.SSF;
	}
	opts.rels = {}; opts.wbrels = {};
	opts.Strings = /*::((*/[]/*:: :any):SST)*/; opts.Strings.Count = 0; opts.Strings.Unique = 0;
	var wbext = opts.bookType == "xlsb" ? "bin" : "xml";
	var vbafmt = opts.bookType == "xlsb" || opts.bookType == "xlsm";
	var ct = ({
		workbooks:[], sheets:[], charts:[], dialogs:[], macros:[],
		rels:[], strs:[], comments:[],
		coreprops:[], extprops:[], custprops:[], themes:[], styles:[],
		calcchains:[], vba: [], drawings: [],
		TODO:[], xmlns: "" }/*:any*/);
	fix_write_opts(opts = opts || {});
	/*:: if(!jszip) throw new Error("JSZip is not available"); */
	var zip = new jszip();
	var f = "", rId = 0;

	opts.cellXfs = [];
	get_cell_style(opts.cellXfs, {}, {revssf:{"General":0}});

	if(!wb.Props) wb.Props = {};

	f = "docProps/core.xml";
	zip.file(f, write_core_props(wb.Props, opts));
	ct.coreprops.push(f);
	add_rels(opts.rels, 2, f, RELS.CORE_PROPS);

	/*::if(!wb.Props) throw "unreachable"; */
	f = "docProps/app.xml";
	if(wb.Props && wb.Props.SheetNames){/* empty */}
	else if(!wb.Workbook || !wb.Workbook.Sheets) wb.Props.SheetNames = wb.SheetNames;
	// $FlowIgnore
	else wb.Props.SheetNames = wb.SheetNames.map(function(x,i) { return [(wb.Workbook.Sheets[i]||{}).Hidden != 2, x];}).filter(function(x) { return x[0]; }).map(function(x) { return x[1]; });
	wb.Props.Worksheets = wb.Props.SheetNames.length;
	zip.file(f, write_ext_props(wb.Props, opts));
	ct.extprops.push(f);
	add_rels(opts.rels, 3, f, RELS.EXT_PROPS);

	if(wb.Custprops !== wb.Props && keys(wb.Custprops||{}).length > 0) {
		f = "docProps/custom.xml";
		zip.file(f, write_cust_props(wb.Custprops, opts));
		ct.custprops.push(f);
		add_rels(opts.rels, 4, f, RELS.CUST_PROPS);
	}

	f = "xl/workbook." + wbext;
	zip.file(f, write_wb(wb, f, opts));
	ct.workbooks.push(f);
	add_rels(opts.rels, 1, f, RELS.WB);

	for(rId=1;rId <= wb.SheetNames.length; ++rId) {
		var wsrels = {'!id':{}};
		var ws = wb.Sheets[wb.SheetNames[rId-1]];
		var _type = (ws || {})["!type"] || "sheet";
		switch(_type) {
		case "chart": /*
			f = "xl/chartsheets/sheet" + rId + "." + wbext;
			zip.file(f, write_cs(rId-1, f, opts, wb, wsrels));
			ct.charts.push(f);
			add_rels(wsrels, -1, "chartsheets/sheet" + rId + "." + wbext, RELS.CS);
			break; */
			/* falls through */
		default:
			f = "xl/worksheets/sheet" + rId + "." + wbext;
			zip.file(f, write_ws(rId-1, f, opts, wb, wsrels));
			ct.sheets.push(f);
			add_rels(opts.wbrels, -1, "worksheets/sheet" + rId + "." + wbext, RELS.WS[0]);
		}

		if(ws) {
			var comments = ws['!comments'];
			if(comments && comments.length > 0) {
				var cf = "xl/comments" + rId + "." + wbext;
				zip.file(cf, write_cmnt(comments, cf, opts));
				ct.comments.push(cf);
				add_rels(wsrels, -1, "../comments" + rId + "." + wbext, RELS.CMNT);
			}
			if(ws['!legacy']) {
				zip.file("xl/drawings/vmlDrawing" + (rId) + ".vml", write_comments_vml(rId, ws['!comments']));
			}
			delete ws['!comments'];
			delete ws['!legacy'];
		}

		if(wsrels['!id'].rId1) zip.file(get_rels_path(f), write_rels(wsrels));
	}

	if(opts.Strings != null && opts.Strings.length > 0) {
		f = "xl/sharedStrings." + wbext;
		zip.file(f, write_sst(opts.Strings, f, opts));
		ct.strs.push(f);
		add_rels(opts.wbrels, -1, "sharedStrings." + wbext, RELS.SST);
	}

	/* TODO: something more intelligent with themes */

	f = "xl/theme/theme1.xml";
	zip.file(f, write_theme(wb.Themes, opts));
	ct.themes.push(f);
	add_rels(opts.wbrels, -1, "theme/theme1.xml", RELS.THEME);

	/* TODO: something more intelligent with styles */

	f = "xl/styles." + wbext;
	zip.file(f, write_sty(wb, f, opts));
	ct.styles.push(f);
	add_rels(opts.wbrels, -1, "styles." + wbext, RELS.STY);

	if(wb.vbaraw && vbafmt) {
		f = "xl/vbaProject.bin";
		zip.file(f, wb.vbaraw);
		ct.vba.push(f);
		add_rels(opts.wbrels, -1, "vbaProject.bin", RELS.VBA);
	}

	zip.file("[Content_Types].xml", write_ct(ct, opts));
	zip.file('_rels/.rels', write_rels(opts.rels));
	zip.file('xl/_rels/workbook.' + wbext + '.rels', write_rels(opts.wbrels));

	delete opts.revssf; delete opts.ssf;
	return zip;
}
