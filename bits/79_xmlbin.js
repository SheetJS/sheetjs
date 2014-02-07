function parse_wb(data, name, opts) {
	return name.substr(-4)===".bin" ? parse_wb_bin(data, opts) : parse_wb_xml(data, opts);
}

function parse_ws(data, name, opts) {
	return name.substr(-4)===".bin" ? parse_ws_bin(data, opts) : parse_ws_xml(data, opts);
}

function parse_sty(data, name, opts) {
	return name.substr(-4)===".bin" ? parse_sty_bin(data, opts) : parse_sty_xml(data, opts);
}
