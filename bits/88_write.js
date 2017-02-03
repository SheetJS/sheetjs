function write_zip_type(wb, opts) {
	var o = opts||{};
	var z = write_zip(wb, o);
	var oopts = {};
	if(opts.compression) oopts.compression = 'DEFLATE';
	switch(o.type) {
		case "base64": oopts.type = "base64"; break;
		case "binary": oopts.type = "string"; break;
		case "buffer":
		case "file": oopts.type = "nodebuffer"; break;
		default: throw new Error("Unrecognized type " + o.type);
	}
	if(o.type === "file") return _fs.writeFileSync(o.file, z.generate(oopts));
	return z.generate(oopts);
}

function writeSync(wb, opts) {
	var o = opts||{};
	switch(o.bookType) {
		case 'xml': return write_xlml(wb, o);
		default: return write_zip_type(wb, o);
	}
}

function writeFileSync(wb, filename, opts) {
	var o = opts||{}; o.type = 'file';
	o.file = filename;
	if(!o.bookType) switch(o.file.substr(-5).toLowerCase()) {
		case '.xlsx': o.bookType = 'xlsx'; break;
		case '.xlsm': o.bookType = 'xlsm'; break;
		case '.xlsb': o.bookType = 'xlsb'; break;
	default: switch(o.file.substr(-4).toLowerCase()) {
		case '.xls': o.bookType = 'xls'; break;
		case '.xml': o.bookType = 'xml'; break;
		case '.ods': o.bookType = 'ods'; break;
	}}
	return writeSync(wb, o);
}

