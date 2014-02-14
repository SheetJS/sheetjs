function parseZip(zip, opts) {
	opts = opts || {};
	fixopts(opts);
	reset_cp();
	var entries = Object.keys(zip.files);
	var keys = entries.filter(function(x){return x.substr(-1) != '/';}).sort();
	var dir = parseCT(getdata(getzipfile(zip, '[Content_Types].xml')));
	var xlsb = false;
	if(dir.workbooks.length === 0) {
		var binname = "xl/workbook.bin";
		if(!getzipfile(zip,binname)) throw new Error("Could not find workbook entry");
		dir.workbooks.push(binname);
		xlsb = true;
	}

	if(!opts.bookSheets && !opts.bookProps) {
		strs = {};
		if(dir.sst) strs=parse_sst(getdata(getzipfile(zip, dir.sst.replace(/^\//,''))), dir.sst, opts);

		styles = {};
		if(dir.style) styles = parse_sty(getdata(getzipfile(zip, dir.style.replace(/^\//,''))),dir.style);
	}

	var wb = parse_wb(getdata(getzipfile(zip, dir.workbooks[0].replace(/^\//,''))), dir.workbooks[0], opts);

	var props = {}, propdata = "";
	try {
		propdata = dir.coreprops.length !== 0 ? getdata(getzipfile(zip, dir.coreprops[0].replace(/^\//,''))) : "";
	propdata += dir.extprops.length !== 0 ? getdata(getzipfile(zip, dir.extprops[0].replace(/^\//,''))) : "";
		props = propdata !== "" ? parseProps(propdata) : {};
	} catch(e) { }

	var custprops = {};
	if(!opts.bookSheets || opts.bookProps) {
		if (dir.custprops.length !== 0) try {
			propdata = getdata(getzipfile(zip, dir.custprops[0].replace(/^\//,'')));
			custprops = parseCustomProps(propdata);
		} catch(e) {/*console.error(e);*/}
	}

	var out = {};
	if(opts.bookSheets || opts.bookProps) {
		var sheets;
		if(props.Worksheets && props.SheetNames.length > 0) sheets=props.SheetNames;
		else if(wb.Sheets) sheets = wb.Sheets.map(function(x){ return x.name; });
		if(opts.bookProps) { out.Props = props; out.Custprops = custprops; }
		if(typeof sheets !== 'undefined') out.SheetNames = sheets;
		if(opts.bookSheets ? out.SheetNames : opts.bookProps) return out;
	}

	var deps = {};
	if(dir.calcchain) deps=parseDeps(getdata(getzipfile(zip, dir.calcchain.replace(/^\//,''))));
	var sheets = {}, i=0;
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
		for(i = 0; i != props.Worksheets; ++i) {
			try { /* TODO: remove these guards */
				path = 'xl/worksheets/sheet' + (i+1) + (xlsb?'.bin':'.xml');
				relsPath = path.replace(/^(.*)(\/)([^\/]*)$/, "$1/_rels/$3.rels");
				sheets[props.SheetNames[i]]=parse_ws(getdata(getzipfile(zip, path)),path,opts);
				sheetRels[props.SheetNames[i]]=parseRels(getdata(getzipfile(zip, relsPath)), path);
			} catch(e) {}
		}
	} else {
		for(i = 0; i != props.Worksheets; ++i) {
			try {
				//var path = dir.sheets[i].replace(/^\//,'');
				path = 'xl/worksheets/sheet' + (i+1) + (xlsb?'.bin':'.xml');
				relsPath = path.replace(/^(.*)(\/)([^\/]*)$/, "$1/_rels/$3.rels");
				sheets[props.SheetNames[i]]=parse_ws(getdata(getzipfile(zip, path)),path,opts);
				sheetRels[props.SheetNames[i]]=parseRels(getdata(getzipfile(zip, relsPath)), path);
			} catch(e) {/*console.error(e);*/}
		}
	}

	if(dir.comments) parse_comments(zip, dir.comments, sheets, sheetRels, opts);

	return {
		Directory: dir,
		Workbook: wb,
		Props: props,
		Custprops: custprops,
		Deps: deps,
		Sheets: sheets,
		SheetNames: props.SheetNames,
		Strings: strs,
		Styles: styles,
		keys: keys,
		files: zip.files
	};
}
