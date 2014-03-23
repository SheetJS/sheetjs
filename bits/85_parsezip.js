function parseZip(zip, opts) {
	opts = opts || {};
	fixopts(opts);
	reset_cp();
	var entries = Object.keys(zip.files);
	var keys = entries.filter(function(x){return x.substr(-1) != '/';}).sort();
	var dir = parseCT(getzipdata(zip, '[Content_Types].xml'));
	var xlsb = false;
	var sheets;
	if(dir.workbooks.length === 0) {
		var binname = "xl/workbook.bin";
		if(!getzipfile(zip,binname)) throw new Error("Could not find workbook entry");
		dir.workbooks.push(binname);
		xlsb = true;
	}

	if(!opts.bookSheets && !opts.bookProps) {
		strs = {};
		if(dir.sst) strs=parse_sst(getzipdata(zip, dir.sst.replace(/^\//,'')), dir.sst, opts);

		styles = {};
		if(dir.style) styles = parse_sty(getzipdata(zip, dir.style.replace(/^\//,'')),dir.style, opts);
	}

	var wb = parse_wb(getzipdata(zip, dir.workbooks[0].replace(/^\//,'')), dir.workbooks[0], opts);

	var props = {}, propdata = "";
	try {
		propdata = dir.coreprops.length !== 0 ? getzipdata(zip, dir.coreprops[0].replace(/^\//,'')) : "";
		propdata += dir.extprops.length !== 0 ? getzipdata(zip, dir.extprops[0].replace(/^\//,'')) : "";
		props = propdata !== "" ? parseProps(propdata) : {};
	} catch(e) { }

	var custprops = {};
	if(!opts.bookSheets || opts.bookProps) {
		if (dir.custprops.length !== 0) {
			propdata = getzipdata(zip, dir.custprops[0].replace(/^\//,''), true);
			if(propdata) custprops = parseCustomProps(propdata);
		}
	}

	var out = {};
	if(opts.bookSheets || opts.bookProps) {
		if(props.Worksheets && props.SheetNames.length > 0) sheets=props.SheetNames;
		else if(wb.Sheets) sheets = wb.Sheets.map(function(x){ return x.name; });
		if(opts.bookProps) { out.Props = props; out.Custprops = custprops; }
		if(typeof sheets !== 'undefined') out.SheetNames = sheets;
		if(opts.bookSheets ? out.SheetNames : opts.bookProps) return out;
	}
	sheets = {};

	var deps = {};
	if(opts.bookDeps && dir.calcchain) deps=parseDeps(getzipdata(zip, dir.calcchain.replace(/^\//,'')));

	var i=0;
	var sheetRels = {};
	var path, relsPath;
	if(!props.Worksheets) {
		/* Google Docs doesn't generate the appropriate metadata, so we impute: */
		var wbsheets = wb.Sheets;
		props.Worksheets = wbsheets.length;
		props.SheetNames = [];
		for(var j = 0; j != wbsheets.length; ++j) {
			props.SheetNames[j] = wbsheets[j].name;
		}
	}
	/* Numbers iOS hack TODO: parse workbook rels to get names */
	var nmode = (getzipdata(zip,"xl/worksheets/sheet.xml",true))?1:0;
	for(i = 0; i != props.Worksheets; ++i) {
		try {
			//path = dir.sheets[i].replace(/^\//,'');
			path = 'xl/worksheets/sheet'+(i+1-nmode)+(xlsb?'.bin':'.xml');
			path = path.replace(/sheet0\./,"sheet.");
			relsPath = path.replace(/^(.*)(\/)([^\/]*)$/, "$1/_rels/$3.rels");
			sheets[props.SheetNames[i]]=parse_ws(getzipdata(zip, path),path,opts);
			sheetRels[props.SheetNames[i]]=parseRels(getzipdata(zip, relsPath, true), path);
		} catch(e) { if(opts.WTF) throw e; }
	}

	if(dir.comments) parse_comments(zip, dir.comments, sheets, sheetRels, opts);

	out = {
		Directory: dir,
		Workbook: wb,
		Props: props,
		Custprops: custprops,
		Deps: deps,
		Sheets: sheets,
		SheetNames: props.SheetNames,
		Strings: strs,
		Styles: styles,
	};
	if(opts.bookFiles) {
		out.keys = keys;
		out.files = zip.files;
	}
	return out;
}
