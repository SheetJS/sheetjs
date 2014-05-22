function writeSync(wb, opts) {
	var o = opts||{};
	var z = write_zip(wb, o);
	switch(o.type) {
		case "base64": return z.generate({type:"base64"});
		case "binary": return z.generate({type:"string"});
		case "buffer": return z.generate({type:"nodebuffer"});
		case "file": return _fs.writeFileSync(o.file, z.generate({type:"nodebuffer"}));
		default: throw new Error("Unrecognized type " + o.type);
	}
}

function writeFileSync(wb, filename, opts) {
	var o = opts||{}; o.type = 'file';
	o.file = filename;
	switch(o.file.substr(-5).toLowerCase()) {
		case '.xlsm': o.bookType = 'xlsm'; break;
		case '.xlsb': o.bookType = 'xlsb'; break;
	}
	return writeSync(wb, o);
}

