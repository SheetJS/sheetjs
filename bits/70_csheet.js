RELS.CS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";

/* 18.3 Worksheets also covers Chartsheets */
function parse_cs_xml(data/*:?string*/, opts, rels, wb, themes, styles)/*:Worksheet*/ {
	if(!data) return data;
	/* 18.3.1.12 chartsheet CT_ChartSheet */
	if(!rels) rels = {'!id':{}};
	var s = {'!type':"chart", '!chart':null, '!rel':""};
	var m;

	/* 18.3.1.36 drawing CT_Drawing */
	if((m = data.match(/drawing r:id="(.*?)"/))) s['!rel'] = m[1];

	if(rels['!id'][s['!rel']]) s['!chart'] = rels['!id'][s['!rel']];
	return s;
}

/* [MS-XLSB] 2.1.7.7 Chart Sheet */
function parse_cs_bin(data, opts, rels, wb, themes, styles)/*:Worksheet*/ {
	if(!data) return data;
	if(!rels) rels = {'!id':{}};
	var s = {'!type':"chart", '!chart':null, '!rel':""};
	var pass = false;
	recordhopper(data, function cs_parse(val, Record) {
		switch(Record.n) {

			case 'BrtDrawing': s['!rel'] = val; break;

			case 'BrtBeginSheet': break;
			case 'BrtCsProp': break; // TODO
			case 'BrtBeginCsViews': break; // TODO
			case 'BrtBeginCsView': break; // TODO
			case 'BrtEndCsView': break; // TODO
			case 'BrtEndCsViews': break; // TODO
			case 'BrtCsProtection': break; // TODO
			case 'BrtMargins': break; // TODO
			case 'BrtCsPageSetup': break; // TODO
			case 'BrtEndSheet': break; // TODO
			case 'BrtBeginHeaderFooter': break; // TODO
			case 'BrtEndHeaderFooter': break; // TODO
			default: if(!pass || opts.WTF) throw new Error("Unexpected record " + Record.n);
		}
	}, opts);

	if(rels['!id'][s['!rel']]) s['!chart'] = rels['!id'][s['!rel']];
	return s;
}
