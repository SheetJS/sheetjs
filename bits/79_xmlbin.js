function parse_wb(data, name) {
	return name.substr(-4)===".bin" ? parse_wb_bin(data) : parse_workbook(data);
}

function parse_ws(data, name) {
	return name.substr(-4)===".bin" ? parse_ws_bin(data) : parse_worksheet(data);
}

function parse_sty(data, name) {
	return name.substr(-4)===".bin" ? parse_sty_bin(data) : parse_styles(data);
}
