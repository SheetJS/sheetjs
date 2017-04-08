RELS.CS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";

var CS_XML_ROOT = writextag('chartsheet', null, {
	'xmlns': XMLNS.main[0],
	'xmlns:r': XMLNS.r
});

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
function write_cs_xml(idx/*:number*/, opts, wb/*:Workbook*/, rels)/*:string*/ {
	var o = [XML_HEADER, CS_XML_ROOT];
	o[o.length] = writextag("drawing", null, {"r:id": "rId1"});
	add_rels(rels, -1, "../drawings/drawing" + (idx+1) + ".xml", RELS.DRAW);
	if(o.length>2) { o[o.length] = ('</chartsheet>'); o[1]=o[1].replace("/>",">"); }
	return o.join("");
}

/* [MS-XLSB] 2.1.7.7 Chart Sheet */
function parse_cs_bin(data, opts, rels, wb, themes, styles)/*:Worksheet*/ {
	if(!data) return data;
	if(!rels) rels = {'!id':{}};
	var s = {'!type':"chart", '!chart':null, '!rel':""};
	var state = [];
	var pass = false;
	recordhopper(data, function cs_parse(val, Record, RT) {
		switch(Record.n) {

			case 'BrtDrawing': s['!rel'] = val; break;

			case 'BrtUid': break;
			case 'BrtMargins': break; // TODO
			case 'BrtLegacyDrawing': break; // TODO
			case 'BrtLegacyDrawingHF': break; // TODO
			case 'BrtBkHim': break; // TODO
			case 'BrtCsProp': break; // TODO
			case 'BrtCsProtection': break; // TODO
			case 'BrtCsProtectionIso': break; // TODO
			case 'BrtCsPageSetup': break; // TODO

			case 'BrtFRTBegin': pass = true; break;
			case 'BrtFRTEnd': pass = false; break;
			case 'BrtACBegin': state.push(R.n); break;
			case 'BrtACEnd': state.pop(); break;

			default:
				if((Record.n||"").indexOf("Begin") > 0) state.push(Record.n);
				else if((Record.n||"").indexOf("End") > 0) state.pop();
				else if(!pass || opts.WTF) throw new Error("Unexpected record " + RT + " " + Record.n);
		}
	}, opts);

	if(rels['!id'][s['!rel']]) s['!chart'] = rels['!id'][s['!rel']];
	return s;
}
function write_cs_bin(idx/*:number*/, opts, wb/*:Workbook*/, rels) {
	var ba = buf_array();
	write_record(ba, "BrtBeginSheet");
	/* [BrtCsProp] */
	/* CSVIEWS */
	/* [[BrtCsProtectionIso] BrtCsProtection] */
	/* [USERCSVIEWS] */
	/* [BrtMargins] */
	/* [BrtCsPageSetup] */
	/* [HEADERFOOTER] */
	/* BrtDrawing */
	/* [BrtLegacyDrawing] */
	/* [BrtLegacyDrawingHF] */
	/* [BrtBkHim] */
	/* [WEBPUBITEMS] */
	/* FRTCHARTSHEET */
	write_record(ba, "BrtEndSheet");
	return ba.end();
}
