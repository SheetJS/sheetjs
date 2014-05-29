function parse_wb(data, name, opts) {
	return (name.substr(-4)===".bin" ? parse_wb_bin : parse_wb_xml)(data, opts);
}

function parse_ws(data, name, opts, rels) {
	return (name.substr(-4)===".bin" ? parse_ws_bin : parse_ws_xml)(data, opts, rels);
}

function parse_sty(data, name, opts) {
	return (name.substr(-4)===".bin" ? parse_sty_bin : parse_sty_xml)(data, opts);
}

function parse_theme(data, name, opts) {
	return parse_theme_xml(data, opts);
}

function parse_sst(data, name, opts) {
	return (name.substr(-4)===".bin" ? parse_sst_bin : parse_sst_xml)(data, opts);
}

function parse_cmnt(data, name, opts) {
	return (name.substr(-4)===".bin" ? parse_comments_bin : parse_comments_xml)(data, opts);
}

function parse_cc(data, name, opts) {
	return (name.substr(-4)===".bin" ? parse_cc_bin : parse_cc_xml)(data, opts);
}

function write_wb(wb, name, opts) {
	return (name.substr(-4)===".bin" ? write_wb_bin : write_wb_xml)(wb, opts);
}

function write_ws(data, name, opts, wb) {
	return (name.substr(-4)===".bin" ? write_ws_bin : write_ws_xml)(data, opts, wb);
}

function write_sty(data, name, opts) {
	return (name.substr(-4)===".bin" ? write_sty_bin : write_sty_xml)(data, opts);
}

function write_sst(data, name, opts) {
	return (name.substr(-4)===".bin" ? write_sst_bin : write_sst_xml)(data, opts);
}
/*
function write_cmnt(data, name, opts) {
	return (name.substr(-4)===".bin" ? write_comments_bin : write_comments_xml)(data, opts);
}

function write_cc(data, name, opts) {
	return (name.substr(-4)===".bin" ? write_cc_bin : write_cc_xml)(data, opts);
}
*/
