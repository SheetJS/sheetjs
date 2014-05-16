function add_rels(rels, rId, f, type, relobj) {
	if(!relobj) relobj = {};
	if(!rels['!id']) rels['!id'] = {};
	relobj.Id = 'rId' + rId;
	relobj.Type = type; 
	relobj.Target = f;
	if(rels['!id'][relobj.Id]) throw new Error("Cannot rewrite rId " + rId);
	rels['!id'][relobj.Id] = relobj;
	rels[('/' + relobj.Target).replace("//","/")] = relobj;
}

function write_zip(wb, opts) {
	if(wb && wb.SSF) {
		make_ssf(SSF); SSF.load_table(wb.SSF);
		opts.revssf = evert(wb.SSF); opts.revssf[wb.SSF[65535]] = 0;
	}
	opts.rels = {}; opts.wbrels = {};
	opts.Strings = []; opts.Strings.Count = 0; opts.Strings.Unique = 0;
	var wbext = opts.bookType == "xlsb" ? "bin" : "xml";
	var ct = { workbooks: [], sheets: [], calcchains: [], themes: [], styles: [],
		coreprops: [], extprops: [], custprops: [], strs:[], comments: [], vba: [],
		TODO:[], rels:[], xmlns: "" };
	fix_write_opts(opts = opts || {});
	var zip = new jszip();
	var f = "", rId = 0;

	opts.cellXfs = [];

	f = "docProps/core.xml";
	zip.file(f, write_core_props(wb.Props, opts));
	ct.coreprops.push(f);
	add_rels(opts.rels, 3, f, RELS.CORE_PROPS);

	f = "docProps/app.xml";
	wb.Props.SheetNames = wb.SheetNames;
	wb.Props.Worksheets = wb.SheetNames.length;
	zip.file(f, write_ext_props(wb.Props, opts));
	ct.extprops.push(f);
	add_rels(opts.rels, 4, f, RELS.EXT_PROPS);

	if(wb.Custprops !== wb.Props) { /* TODO: fix xlsjs */
		f = "docProps/custom.xml";
		zip.file(f, write_cust_props(wb.Custprops, opts));
		ct.custprops.push(f);
		add_rels(opts.rels, 5, f, RELS.CUST_PROPS);
	}

	f = "xl/workbook." + wbext;
	zip.file(f, write_wb(wb, f, opts));
	ct.workbooks.push(f);
	add_rels(opts.rels, 1, f, RELS.WB);

	wb.SheetNames.forEach(function(s, i) {
		rId = i+1; f = "xl/worksheets/sheet" + rId + "." + wbext;
		zip.file(f, write_ws(i, f, opts, wb));
		ct.sheets.push(f);
		add_rels(opts.wbrels, rId, "worksheets/sheet" + rId + "." + wbext, RELS.WS);
	});

	if((opts.Strings||[]).length > 0) {
		f = "xl/sharedStrings." + wbext;
		zip.file(f, write_sst(opts.Strings, f, opts));
		ct.strs.push(f);
		add_rels(opts.wbrels, ++rId, "sharedStrings." + wbext, RELS.SST);
	}

	/* TODO: something more intelligent with themes */
	
/*	f = "xl/theme/theme1.xml"
	zip.file(f, write_theme());
	ct.themes.push(f);
	add_rels(opts.wbrels, ++rId, "theme/theme1.xml", RELS.THEME);*/

	/* TODO: something more intelligent with styles */

	f = "xl/styles.xml";
	zip.file(f, write_sty(wb, f, opts));
	ct.styles.push(f);
	add_rels(opts.wbrels, ++rId, "styles." + wbext, RELS.STY);

	zip.file("[Content_Types].xml", write_ct(ct, opts));
	zip.file('_rels/.rels', write_rels(opts.rels));
	zip.file('xl/_rels/workbook.xml.rels', write_rels(opts.wbrels));
	return zip;
}
