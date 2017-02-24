function firstbyte(f/*:RawData*/,o/*:?TypeOpts*/)/*:number*/ {
	switch((o||{}).type || "base64") {
		case 'buffer': return f[0];
		case 'base64': return Base64.decode(f.substr(0,12)).charCodeAt(0);
		case 'binary': return f.charCodeAt(0);
		case 'array': return f[0];
		default: throw new Error("Unrecognized type " + (o ? o.type : "undefined"));
	}
}

function read_zip(data/*:RawData*/, opts/*:?ParseOpts*/)/*:Workbook*/ {
	/*:: if(!jszip) throw new Error("JSZip is not available"); */
	var zip, d = data;
	var o = opts||{};
	if(!o.type) o.type = (has_buf && Buffer.isBuffer(data)) ? "buffer" : "base64";
	switch(o.type) {
		case "base64": zip = new jszip(d, { base64:true }); break;
		case "binary": case "array": zip = new jszip(d, { base64:false }); break;
		case "buffer": zip = new jszip(d); break;
		default: throw new Error("Unrecognized type " + o.type);
	}
	return parse_zip(zip, o);
}

function readSync(data/*:RawData*/, opts/*:?ParseOpts*/)/*:Workbook*/ {
	var zip, d = data, n=0;
	var o = opts||{};
	if(!o.type) o.type = (has_buf && Buffer.isBuffer(data)) ? "buffer" : "base64";
	if(o.type == "file") { o.type = "buffer"; d = _fs.readFileSync(data); }
	switch((n = firstbyte(d, o))) {
		case 0xD0: return parse_xlscfb(CFB.read(d, o), o);
		case 0x09: return parse_xlscfb(s2a(o.type === 'base64' ? Base64.decode(d) : d), o);
		case 0x3C: return parse_xlml(d, o);
		case 0x50: return read_zip(d, o);
		case 0xEF: return parse_xlml(d, o);
		default: throw new Error("Unsupported file " + n);
	}
}

function readFileSync(filename/*:string*/, opts/*:?ParseOpts*/)/*:Workbook*/ {
	var o = opts||{}; o.type = 'file';
	return readSync(filename, o);
}
