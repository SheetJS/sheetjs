
var parse_sst = function(data, name) {
	return name.substr(-4)===".bin" ? parse_sst_bin(data) : parse_sst_xml(data);
};
