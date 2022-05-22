/* 18.3 Worksheets also covers Chartsheets */
function parse_cs_xml(data/*:?string*/, opts, idx/*:number*/, rels, wb/*::, themes, styles*/)/*:Worksheet*/ {
	if(!data) return data;
	/* 18.3.1.12 chartsheet CT_ChartSheet */
	if(!rels) rels = {'!id':{}};
	var s = ({'!type':"chart", '!drawel':null, '!rel':""}/*:any*/);
	var m;

	/* 18.3.1.83 sheetPr CT_ChartsheetPr */
	var sheetPr = data.match(sheetprregex);
	if(sheetPr) parse_ws_xml_sheetpr(sheetPr[0], s, wb, idx);

	/* 18.3.1.36 drawing CT_Drawing */
	if((m = data.match(/drawing r:id="(.*?)"/))) s['!rel'] = m[1];

	if(rels['!id'][s['!rel']]) s['!drawel'] = rels['!id'][s['!rel']];
	return s;
}
//function write_cs_xml(idx/*:number*/, opts, wb/*:Workbook*/, rels)/*:string*/ {
//	var o = [XML_HEADER, writextag('chartsheet', null, {
//		'xmlns': XMLNS_main[0],
//		'xmlns:r': XMLNS.r
//	})];
//	o[o.length] = writextag("drawing", null, {"r:id": "rId1"});
//	add_rels(rels, -1, "../drawings/drawing" + (idx+1) + ".xml", RELS.DRAW);
//	if(o.length>2) { o[o.length] = ('</chartsheet>'); o[1]=o[1].replace("/>",">"); }
//	return o.join("");
//}
