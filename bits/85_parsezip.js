function parseZip(zip) {
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
	strs = {};
	if(dir.sst) strs=parse_sst(getdata(getzipfile(zip, dir.sst.replace(/^\//,''))), dir.sst);

	styles = {};
	if(dir.style) styles = parse_sty(getdata(getzipfile(zip, dir.style.replace(/^\//,''))),dir.style);

	var wb = parse_wb(getdata(getzipfile(zip, dir.workbooks[0].replace(/^\//,''))), dir.workbooks[0]);
	var props = {}, propdata = "";
	try {
		propdata = dir.coreprops.length !== 0 ? getdata(getzipfile(zip, dir.coreprops[0].replace(/^\//,''))) : "";
	propdata += dir.extprops.length !== 0 ? getdata(getzipfile(zip, dir.extprops[0].replace(/^\//,''))) : "";
		props = propdata !== "" ? parseProps(propdata) : {};
	} catch(e) { }
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
				sheets[props.SheetNames[i]]=parse_ws(getdata(getzipfile(zip, path)),path);
				sheetRels[props.SheetNames[i]]=parseRels(getdata(getzipfile(zip, relsPath)), path);
			} catch(e) {}
		}
	} else {
		for(i = 0; i != props.Worksheets; ++i) {
			try {
				//var path = dir.sheets[i].replace(/^\//,'');
				path = 'xl/worksheets/sheet' + (i+1) + (xlsb?'.bin':'.xml');
				relsPath = path.replace(/^(.*)(\/)([^\/]*)$/, "$1/_rels/$3.rels");
				sheets[props.SheetNames[i]]=parse_ws(getdata(getzipfile(zip, path)),path);
				sheetRels[props.SheetNames[i]]=parseRels(getdata(getzipfile(zip, relsPath)), path);
			} catch(e) {/*console.error(e);*/}
		}
	}

	if(dir.comments) parseCommentsAddToSheets(zip, dir.comments, sheets, sheetRels);

	return {
		Directory: dir,
		Workbook: wb,
		Props: props,
		Deps: deps,
		Sheets: sheets,
		SheetNames: props.SheetNames,
		Strings: strs,
		Styles: styles,
		keys: keys,
		files: zip.files
	};
}
