function parse_wb(data, name, opts) {
	return name.substr(-4)===".bin" ? parse_wb_bin(data, opts) : parse_wb_xml(data, opts);
}

function parse_ws(data, name, opts, rels) {
	return name.substr(-4)===".bin" ? parse_ws_bin(data, opts, rels) : parse_ws_xml(data, opts, rels);
}

function parse_sty(data, name, opts) {
	return name.substr(-4)===".bin" ? parse_sty_bin(data, opts) : parse_sty_xml(data, opts);
}

function parse_sst(data, name, opts) {
	return name.substr(-4)===".bin" ? parse_sst_bin(data, opts) : parse_sst_xml(data, opts);
}

function parse_cmnt(data, name, opts) {
	return name.substr(-4)===".bin" ? parse_comments_bin(data, opts) : parse_comments_xml(data, opts);
}

function parse_cc(data, name, opts) {
	return name.substr(-4)===".bin" ? parse_cc_bin(data, opts) : parse_cc_xml(data, opts);
}
