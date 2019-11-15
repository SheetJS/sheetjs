RELS.CS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";

var CS_XML_ROOT = writextag('chartsheet', null, {
	'xmlns': XMLNS.main[0],
	'xmlns:r': XMLNS.r
});

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
function write_cs_xml(idx/*:number*/, opts, wb/*:Workbook*/, rels)/*:string*/ {
	var o = [XML_HEADER, CS_XML_ROOT];
	o[o.length] = writextag("drawing", null, {"r:id": "rId1"});
	add_rels(rels, -1, "../drawings/drawing" + (idx+1) + ".xml", RELS.DRAW);
	if(o.length>2) { o[o.length] = ('</chartsheet>'); o[1]=o[1].replace("/>",">"); }
	return o.join("");
}

/* [MS-XLSB] 2.4.331 BrtCsProp */
function parse_BrtCsProp(data, length/*:number*/) {
	data.l += 10;
	var name = parse_XLWideString(data, length - 10);
	return { name: name };
}

/* [MS-XLSB] 2.1.7.7 Chart Sheet */
function parse_cs_bin(data, opts, idx/*:number*/, rels, wb/*::, themes, styles*/)/*:Worksheet*/ {
	if(!data) return data;
	if(!rels) rels = {'!id':{}};
	var s = {'!type':"chart", '!drawel':null, '!rel':""};
	var state/*:Array<string>*/ = [];
	var pass = false;
	recordhopper(data, function cs_parse(val, R_n, RT) {
		switch(RT) {

			case 0x0226: /* 'BrtDrawing' */
				s['!rel'] = val; break;

			case 0x028B: /* 'BrtCsProp' */
				if(!wb.Sheets[idx]) wb.Sheets[idx] = {};
				if(val.name) wb.Sheets[idx].CodeName = val.name;
				break;

			case 0x0232: /* 'BrtBkHim' */
			case 0x028C: /* 'BrtCsPageSetup' */
			case 0x029D: /* 'BrtCsProtection' */
			case 0x02A7: /* 'BrtCsProtectionIso' */
			case 0x0227: /* 'BrtLegacyDrawing' */
			case 0x0228: /* 'BrtLegacyDrawingHF' */
			case 0x01DC: /* 'BrtMargins' */
			case 0x0C00: /* 'BrtUid' */
				break;

			case 0x0023: /* 'BrtFRTBegin' */
				pass = true; break;
			case 0x0024: /* 'BrtFRTEnd' */
				pass = false; break;
			case 0x0025: /* 'BrtACBegin' */
				state.push(R_n); break;
			case 0x0026: /* 'BrtACEnd' */
				state.pop(); break;

			default:
				if((R_n||"").indexOf("Begin") > 0) state.push(R_n);
				else if((R_n||"").indexOf("End") > 0) state.pop();
				else if(!pass || opts.WTF) throw new Error("Unexpected record " + RT + " " + R_n);
		}
	}, opts);

	if(rels['!id'][s['!rel']]) s['!drawel'] = rels['!id'][s['!rel']];
	return s;
}
function write_cs_bin(/*::idx:number, opts, wb:Workbook, rels*/) {
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
