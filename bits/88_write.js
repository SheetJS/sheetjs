function write_zip_type(wb, opts) {
	var o = {}, z;
	opts = opts||{};
	z = write_zip(wb, opts);
	
	switch(opts.type) {
		case "base64": o.type = "base64"; if (typeof opts.compression === 'string') {o.compression = opts.compression;} break;
		case "binary": o.type = "string"; if (typeof opts.compression === 'string') {o.compression = opts.compression;} break;
		case "buffer": o.type = "nodebuffer"; if (typeof opts.compression === 'string') {o.compression = opts.compression;} break;
		case "file": o.type = "nodebuffer"; if (typeof opts.compression === 'string') {o.compression = opts.compression;} break;
		default: throw new Error("Unrecognized type " + opts.type);
	}
	
	if(opts.type === "file") {
		return _fs.writeFileSync(opts.file, z.generate(o));
	}
	else {
		return z.generate(o);
	}
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
	switch(o.file.substr(-5).toLowerCase()) {
		case '.xlsx': o.bookType = 'xlsx'; break;
		case '.xlsm': o.bookType = 'xlsm'; break;
		case '.xlsb': o.bookType = 'xlsb'; break;
	default: switch(o.file.substr(-4).toLowerCase()) {
		case '.xls': o.bookType = 'xls'; break;
		case '.xml': o.bookType = 'xml'; break;
	}}
	return writeSync(wb, o);
}

