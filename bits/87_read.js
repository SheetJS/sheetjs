function readSync(data, opts) {
	var zip, d = data;
	var o = opts||{};
	if(!o.type) o.type = (typeof Buffer !== 'undefined' && data instanceof Buffer) ? "buffer" : "base64";
	switch(o.type) {
		case "base64": zip = new jszip(d, { base64:true }); break;
		case "binary": zip = new jszip(d, { base64:false }); break;
		case "buffer": zip = new jszip(d); break;
		case "file": zip=new jszip(d=_fs.readFileSync(data)); break;
		default: throw new Error("Unrecognized type " + o.type);
	}
	return parse_zip(zip, o);
}

function readFileSync(data, opts) {
	var o = opts||{}; o.type = 'file';
	return readSync(data, o);
}

