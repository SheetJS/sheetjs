function write_zip_xlsb(wb/*:Workbook*/, opts/*:WriteOpts*/)/*:ZIP*/ {
	if(wb && !wb.SSF) {
		wb.SSF = dup(table_fmt);
	}
	if(wb && wb.SSF) {
		make_ssf(); SSF_load_table(wb.SSF);
		// $FlowIgnore
		opts.revssf = evert_num(wb.SSF); opts.revssf[wb.SSF[65535]] = 0;
		opts.ssf = wb.SSF;
	}
	opts.rels = {}; opts.wbrels = {};
	opts.Strings = /*::((*/[]/*:: :any):SST)*/; opts.Strings.Count = 0; opts.Strings.Unique = 0;
	if(browser_has_Map) opts.revStrings = new Map();
	else { opts.revStrings = {}; opts.revStrings.foo = []; delete opts.revStrings.foo; }
	var wbext = "bin";
	var vbafmt = true;
	var ct = new_ct();
	fix_write_opts(opts = opts || {});
	var zip = zip_new();
	var f = "", rId = 0;

	opts.cellXfs = [];
	get_cell_style(opts.cellXfs, {}, {revssf:{"General":0}});

	if(!wb.Props) wb.Props = {};

	f = "docProps/core.xml";
	zip_add_file(zip, f, write_core_props(wb.Props, opts));
	ct.coreprops.push(f);
	add_rels(opts.rels, 2, f, RELS.CORE_PROPS);

	/*::if(!wb.Props) throw "unreachable"; */
	f = "docProps/app.xml";
	if(wb.Props && wb.Props.SheetNames){/* empty */}
	else if(!wb.Workbook || !wb.Workbook.Sheets) wb.Props.SheetNames = wb.SheetNames;
	else {
		var _sn = [];
		for(var _i = 0; _i < wb.SheetNames.length; ++_i)
			if((wb.Workbook.Sheets[_i]||{}).Hidden != 2) _sn.push(wb.SheetNames[_i]);
		wb.Props.SheetNames = _sn;
	}
	wb.Props.Worksheets = wb.Props.SheetNames.length;
	zip_add_file(zip, f, write_ext_props(wb.Props, opts));
	ct.extprops.push(f);
	add_rels(opts.rels, 3, f, RELS.EXT_PROPS);

	if(wb.Custprops !== wb.Props && keys(wb.Custprops||{}).length > 0) {
		f = "docProps/custom.xml";
		zip_add_file(zip, f, write_cust_props(wb.Custprops, opts));
		ct.custprops.push(f);
		add_rels(opts.rels, 4, f, RELS.CUST_PROPS);
	}

	for(rId=1;rId <= wb.SheetNames.length; ++rId) {
		var wsrels = {'!id':{}};
		var ws = wb.Sheets[wb.SheetNames[rId-1]];
		var _type = (ws || {})["!type"] || "sheet";
		switch(_type) {
		case "chart":
			/* falls through */
		default:
			f = "xl/worksheets/sheet" + rId + "." + wbext;
			zip_add_file(zip, f, write_ws_bin(rId-1, opts, wb, wsrels));
			ct.sheets.push(f);
			add_rels(opts.wbrels, -1, "worksheets/sheet" + rId + "." + wbext, RELS.WS[0]);
		}

		if(ws) {
			var comments = ws['!comments'];
			var need_vml = false;
			var cf = "";
			if(comments && comments.length > 0) {
				cf = "xl/comments" + rId + "." + wbext;
				zip_add_file(zip, cf, write_comments_bin(comments, opts));
				ct.comments.push(cf);
				add_rels(wsrels, -1, "../comments" + rId + "." + wbext, RELS.CMNT);
				need_vml = true;
			}
			if(ws['!legacy']) {
				if(need_vml) zip_add_file(zip, "xl/drawings/vmlDrawing" + (rId) + ".vml", write_vml(rId, ws['!comments']));
			}
			delete ws['!comments'];
			delete ws['!legacy'];
		}

		if(wsrels['!id'].rId1) zip_add_file(zip, get_rels_path(f), write_rels(wsrels));
	}

	if(opts.Strings != null && opts.Strings.length > 0) {
		f = "xl/sharedStrings." + wbext;
		zip_add_file(zip, f, write_sst_bin(opts.Strings, opts));
		ct.strs.push(f);
		add_rels(opts.wbrels, -1, "sharedStrings." + wbext, RELS.SST);
	}

	f = "xl/workbook." + wbext;
	zip_add_file(zip, f, write_wb_bin(wb, opts));
	ct.workbooks.push(f);
	add_rels(opts.rels, 1, f, RELS.WB);

	/* TODO: something more intelligent with themes */

	f = "xl/theme/theme1.xml";
	var ww = write_theme(wb.Themes, opts);
	zip_add_file(zip, f, ww);
	ct.themes.push(f);
	add_rels(opts.wbrels, -1, "theme/theme1.xml", RELS.THEME);

	/* TODO: something more intelligent with styles */

	f = "xl/styles." + wbext;
	zip_add_file(zip, f, write_sty_bin(wb, opts));
	ct.styles.push(f);
	add_rels(opts.wbrels, -1, "styles." + wbext, RELS.STY);

	if(wb.vbaraw && vbafmt) {
		f = "xl/vbaProject.bin";
		zip_add_file(zip, f, wb.vbaraw);
		ct.vba.push(f);
		add_rels(opts.wbrels, -1, "vbaProject.bin", RELS.VBA);
	}

	f = "xl/metadata." + wbext;
	zip_add_file(zip, f, write_xlmeta_bin());
	ct.metadata.push(f);
	add_rels(opts.wbrels, -1, "metadata." + wbext, RELS.XLMETA);

	zip_add_file(zip, "[Content_Types].xml", write_ct(ct, opts));
	zip_add_file(zip, '_rels/.rels', write_rels(opts.rels));
	zip_add_file(zip, 'xl/_rels/workbook.' + wbext + '.rels', write_rels(opts.wbrels));

	delete opts.revssf; delete opts.ssf;
	return zip;
}

function write_zip_xlsx(wb/*:Workbook*/, opts/*:WriteOpts*/)/*:ZIP*/ {
	if(wb && !wb.SSF) {
		wb.SSF = dup(table_fmt);
	}
	if(wb && wb.SSF) {
		make_ssf(); SSF_load_table(wb.SSF);
		// $FlowIgnore
		opts.revssf = evert_num(wb.SSF); opts.revssf[wb.SSF[65535]] = 0;
		opts.ssf = wb.SSF;
	}
	opts.rels = {}; opts.wbrels = {};
	opts.Strings = /*::((*/[]/*:: :any):SST)*/; opts.Strings.Count = 0; opts.Strings.Unique = 0;
	if(browser_has_Map) opts.revStrings = new Map();
	else { opts.revStrings = {}; opts.revStrings.foo = []; delete opts.revStrings.foo; }
	var wbext = "xml";
	var vbafmt = VBAFMTS.indexOf(opts.bookType) > -1;
	var ct = new_ct();
	fix_write_opts(opts = opts || {});
	var zip = zip_new();
	var f = "", rId = 0;

	opts.cellXfs = [];
	get_cell_style(opts.cellXfs, {}, {revssf:{"General":0}});

	if(!wb.Props) wb.Props = {};

	f = "docProps/core.xml";
	zip_add_file(zip, f, write_core_props(wb.Props, opts));
	ct.coreprops.push(f);
	add_rels(opts.rels, 2, f, RELS.CORE_PROPS);

	/*::if(!wb.Props) throw "unreachable"; */
	f = "docProps/app.xml";
	if(wb.Props && wb.Props.SheetNames){/* empty */}
	else if(!wb.Workbook || !wb.Workbook.Sheets) wb.Props.SheetNames = wb.SheetNames;
	else {
		var _sn = [];
		for(var _i = 0; _i < wb.SheetNames.length; ++_i)
			if((wb.Workbook.Sheets[_i]||{}).Hidden != 2) _sn.push(wb.SheetNames[_i]);
		wb.Props.SheetNames = _sn;
	}
	wb.Props.Worksheets = wb.Props.SheetNames.length;
	zip_add_file(zip, f, write_ext_props(wb.Props, opts));
	ct.extprops.push(f);
	add_rels(opts.rels, 3, f, RELS.EXT_PROPS);

	if(wb.Custprops !== wb.Props && keys(wb.Custprops||{}).length > 0) {
		f = "docProps/custom.xml";
		zip_add_file(zip, f, write_cust_props(wb.Custprops, opts));
		ct.custprops.push(f);
		add_rels(opts.rels, 4, f, RELS.CUST_PROPS);
	}

	var people = ["SheetJ5"];
	opts.tcid = 0;

	for(rId=1;rId <= wb.SheetNames.length; ++rId) {
		var wsrels = {'!id':{}};
		var ws = wb.Sheets[wb.SheetNames[rId-1]];
		var _type = (ws || {})["!type"] || "sheet";
		switch(_type) {
		case "chart":
			/* falls through */
		default:
			f = "xl/worksheets/sheet" + rId + "." + wbext;
			zip_add_file(zip, f, write_ws_xml(rId-1, opts, wb, wsrels));
			ct.sheets.push(f);
			add_rels(opts.wbrels, -1, "worksheets/sheet" + rId + "." + wbext, RELS.WS[0]);
		}

		if(ws) {
			var comments = ws['!comments'];
			var need_vml = false;
			var cf = "";
			if(comments && comments.length > 0) {
				var needtc = false;
				comments.forEach(function(carr) {
					carr[1].forEach(function(c) { if(c.T == true) needtc = true; });
				});
				if(needtc) {
					cf = "xl/threadedComments/threadedComment" + rId + ".xml";
					zip_add_file(zip, cf, write_tcmnt_xml(comments, people, opts));
					ct.threadedcomments.push(cf);
					add_rels(wsrels, -1, "../threadedComments/threadedComment" + rId + ".xml", RELS.TCMNT);
				}

				cf = "xl/comments" + rId + "." + wbext;
				zip_add_file(zip, cf, write_comments_xml(comments, opts));
				ct.comments.push(cf);
				add_rels(wsrels, -1, "../comments" + rId + "." + wbext, RELS.CMNT);
				need_vml = true;
			}
			if(ws['!legacy']) {
				if(need_vml) zip_add_file(zip, "xl/drawings/vmlDrawing" + (rId) + ".vml", write_vml(rId, ws['!comments']));
			}
			delete ws['!comments'];
			delete ws['!legacy'];
		}

		if(wsrels['!id'].rId1) zip_add_file(zip, get_rels_path(f), write_rels(wsrels));
	}

	if(opts.Strings != null && opts.Strings.length > 0) {
		f = "xl/sharedStrings." + wbext;
		zip_add_file(zip, f, write_sst_xml(opts.Strings, opts));
		ct.strs.push(f);
		add_rels(opts.wbrels, -1, "sharedStrings." + wbext, RELS.SST);
	}

	f = "xl/workbook." + wbext;
	zip_add_file(zip, f, write_wb_xml(wb, opts));
	ct.workbooks.push(f);
	add_rels(opts.rels, 1, f, RELS.WB);

	/* TODO: something more intelligent with themes */

	f = "xl/theme/theme1.xml";
	zip_add_file(zip, f, write_theme(wb.Themes, opts));
	ct.themes.push(f);
	add_rels(opts.wbrels, -1, "theme/theme1.xml", RELS.THEME);

	/* TODO: something more intelligent with styles */

	f = "xl/styles." + wbext;
	zip_add_file(zip, f, write_sty_xml(wb, opts));
	ct.styles.push(f);
	add_rels(opts.wbrels, -1, "styles." + wbext, RELS.STY);

	if(wb.vbaraw && vbafmt) {
		f = "xl/vbaProject.bin";
		zip_add_file(zip, f, wb.vbaraw);
		ct.vba.push(f);
		add_rels(opts.wbrels, -1, "vbaProject.bin", RELS.VBA);
	}

	f = "xl/metadata." + wbext;
	zip_add_file(zip, f, write_xlmeta_xml());
	ct.metadata.push(f);
	add_rels(opts.wbrels, -1, "metadata." + wbext, RELS.XLMETA);

	if(people.length > 1) {
		f = "xl/persons/person.xml";
		zip_add_file(zip, f, write_people_xml(people, opts));
		ct.people.push(f);
		add_rels(opts.wbrels, -1, "persons/person.xml", RELS.PEOPLE);
	}

	zip_add_file(zip, "[Content_Types].xml", write_ct(ct, opts));
	zip_add_file(zip, '_rels/.rels', write_rels(opts.rels));
	zip_add_file(zip, 'xl/_rels/workbook.' + wbext + '.rels', write_rels(opts.wbrels));

	delete opts.revssf; delete opts.ssf;
	return zip;
}

