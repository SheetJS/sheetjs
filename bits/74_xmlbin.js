function parse_wb(data, name/*:string*/, opts)/*:WorkbookFile*/ {
	if(name.slice(-4)===".bin") return parse_wb_bin((data/*:any*/), opts);
	return parse_wb_xml((data/*:any*/), opts);
}

function parse_ws(data, name/*:string*/, idx/*:number*/, opts, rels, wb, themes, styles)/*:Worksheet*/ {
	if(name.slice(-4)===".bin") return parse_ws_bin((data/*:any*/), opts, idx, rels, wb, themes, styles);
	return parse_ws_xml((data/*:any*/), opts, idx, rels, wb, themes, styles);
}

function parse_cs(data, name/*:string*/, idx/*:number*/, opts, rels, wb, themes, styles)/*:Worksheet*/ {
	if(name.slice(-4)===".bin") return parse_cs_bin((data/*:any*/), opts, idx, rels, wb, themes, styles);
	return parse_cs_xml((data/*:any*/), opts, idx, rels, wb, themes, styles);
}

function parse_ms(data, name/*:string*/, idx/*:number*/, opts, rels, wb, themes, styles)/*:Worksheet*/ {
	if(name.slice(-4)===".bin") return parse_ms_bin((data/*:any*/), opts, idx, rels, wb, themes, styles);
	return parse_ms_xml((data/*:any*/), opts, idx, rels, wb, themes, styles);
}

function parse_ds(data, name/*:string*/, idx/*:number*/, opts, rels, wb, themes, styles)/*:Worksheet*/ {
	if(name.slice(-4)===".bin") return parse_ds_bin((data/*:any*/), opts, idx, rels, wb, themes, styles);
	return parse_ds_xml((data/*:any*/), opts, idx, rels, wb, themes, styles);
}

function parse_sty(data, name/*:string*/, themes, opts) {
	if(name.slice(-4)===".bin") return parse_sty_bin((data/*:any*/), themes, opts);
	return parse_sty_xml((data/*:any*/), themes, opts);
}

function parse_sst(data, name/*:string*/, opts)/*:SST*/ {
	if(name.slice(-4)===".bin") return parse_sst_bin((data/*:any*/), opts);
	return parse_sst_xml((data/*:any*/), opts);
}

function parse_cmnt(data, name/*:string*/, opts)/*:Array<RawComment>*/ {
	if(name.slice(-4)===".bin") return parse_comments_bin((data/*:any*/), opts);
	return parse_comments_xml((data/*:any*/), opts);
}

function parse_cc(data, name/*:string*/, opts) {
	if(name.slice(-4)===".bin") return parse_cc_bin((data/*:any*/), name, opts);
	return parse_cc_xml((data/*:any*/), name, opts);
}

function parse_xlink(data, rel, name/*:string*/, opts) {
	if(name.slice(-4)===".bin") return parse_xlink_bin((data/*:any*/), rel, name, opts);
	return parse_xlink_xml((data/*:any*/), rel, name, opts);
}

function parse_xlmeta(data, name/*:string*/, opts) {
	if(name.slice(-4)===".bin") return parse_xlmeta_bin((data/*:any*/), name, opts);
	return parse_xlmeta_xml((data/*:any*/), name, opts);
}
