function write_zip_type(wb, opts) {
	var o = opts||{};
  style_builder  = new StyleBuilder(opts);

  var z = write_zip(wb, o);
	switch(o.type) {
		case "base64": return z.generate({type:"base64"});
		case "binary": return z.generate({type:"string"});
		case "buffer": return z.generate({type:"nodebuffer"});
		case "file": return _fs.writeFileSync(o.file, z.generate({type:"nodebuffer"}));
		default: throw new Error("Unrecognized type " + o.type);
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

