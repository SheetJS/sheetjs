function firstbyte(f/*:RawData*/,o/*:?TypeOpts*/)/*:Array<number>*/ {
	var x = "";
	switch((o||{}).type || "base64") {
		case 'buffer': return [f[0], f[1], f[2], f[3]];
		case 'base64': x = Base64.decode(f.substr(0,24)); break;
		case 'binary': x = f; break;
		case 'array':  return [f[0], f[1], f[2], f[3]];
		default: throw new Error("Unrecognized type " + (o ? o.type : "undefined"));
	}
	return [x.charCodeAt(0), x.charCodeAt(1), x.charCodeAt(2), x.charCodeAt(3)];
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
	var zip, d = data, n=[0];
	var o = opts||{};
	if(!o.type) o.type = (has_buf && Buffer.isBuffer(data)) ? "buffer" : "base64";
	if(o.type == "file") { o.type = "buffer"; d = _fs.readFileSync(data); }
	switch((n = firstbyte(d, o))[0]) {
		case 0xD0: return parse_xlscfb(CFB.read(d, o), o);
		case 0x09: return parse_xlscfb(s2a(o.type === 'base64' ? Base64.decode(d) : d), o);
		case 0x3C: return parse_xlml(d, o);
		case 0x50: if(n[1] == 0x4B && n[2] < 0x20 && n[3] < 0x20) return read_zip(d, o); break;
		case 0xEF: return parse_xlml(d, o);
		case 0x03: case 0x83: case 0x8B: return DBF.to_workbook(d, o);
		case 0x30: case 0x31: if(n[2] <= 12 && n[3] <= 31) return DBF.to_workbook(d, o); break;
		default: throw new Error("Unsupported file " + n.join("|"));
	}
	throw new Error("Unsupported file format " + n.join("|"));
}

function readFileSync(filename/*:string*/, opts/*:?ParseOpts*/)/*:Workbook*/ {
	var o = opts||{}; o.type = 'file';
	return readSync(filename, o);
}
