function firstbyte(f/*:RawData*/,o/*:?TypeOpts*/)/*:Array<number>*/ {
	var x = "";
	switch((o||{}).type || "base64") {
		case 'buffer': return [f[0], f[1], f[2], f[3]];
		case 'base64': x = Base64.decode(f.substr(0,24)); break;
		case 'binary': x = f; break;
		case 'array':  return [f[0], f[1], f[2], f[3]];
		default: throw new Error("Unrecognized type " + (o && o.type || "undefined"));
	}
	return [x.charCodeAt(0), x.charCodeAt(1), x.charCodeAt(2), x.charCodeAt(3)];
}

function read_cfb(cfb/*:CFBContainer*/, opts/*:?ParseOpts*/)/*:Workbook*/ {
	if(CFB.find(cfb, "EncryptedPackage")) return parse_xlsxcfb(cfb, opts);
	return parse_xlscfb(cfb, opts);
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

function read_plaintext(data/*:string*/, o/*:ParseOpts*/)/*:Workbook*/ {
	var i = 0;
	main: while(i < data.length) switch(data.charCodeAt(i)) {
		case 0x0A: case 0x0D: case 0x20: ++i; break;
		case 0x3C: return parse_xlml(data.slice(i),o);
		default: break main;
	}
	return PRN.to_workbook(data, o);
}

function read_plaintext_raw(data/*:RawData*/, o/*:ParseOpts*/)/*:Workbook*/ {
	var str = "", bytes = firstbyte(data, o);
	switch(o.type) {
		case 'base64': str = Base64.decode(data); break;
		case 'binary': str = data; break;
		case 'buffer': str = data.toString('binary'); break;
		case 'array': str = cc2str(data); break;
		default: throw new Error("Unrecognized type " + o.type);
	}
	if(bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) str = utf8read(str);
	return read_plaintext(str, o);
}

function read_utf16(data/*:RawData*/, o/*:ParseOpts*/)/*:Workbook*/ {
	var d = data;
	if(o.type == 'base64') d = Base64.decode(d);
	d = cptable.utils.decode(1200, d.slice(2), 'str');
	o.type = "binary";
	return read_plaintext(d, o);
}

function readSync(data/*:RawData*/, opts/*:?ParseOpts*/)/*:Workbook*/ {
	var zip, d = data, n=[0];
	var o = opts||{};
	_ssfopts = {};
	if(o.dateNF) _ssfopts.dateNF = o.dateNF;
	if(!o.type) o.type = (has_buf && Buffer.isBuffer(data)) ? "buffer" : "base64";
	if(o.type == "file") { o.type = "buffer"; d = _fs.readFileSync(data); }
	switch((n = firstbyte(d, o))[0]) {
		case 0xD0: return read_cfb(CFB.read(d, o), o);
		case 0x09: return parse_xlscfb(d, o);
		case 0x3C: return parse_xlml(d, o);
		case 0x49: if(n[1] == 0x44) return read_wb_ID(d, o); break;
		case 0x54: if(n[1] == 0x41 && n[2] == 0x42 && n[3] == 0x4C) return DIF.to_workbook(d, o); break;
		case 0x50: if(n[1] == 0x4B && n[2] < 0x20 && n[3] < 0x20) return read_zip(d, o); break;
		case 0xEF: return n[3] == 0x3C ? parse_xlml(d, o) : PRN.to_workbook(d,o);
		case 0xFF: if(n[1] == 0xFE){ return read_utf16(d, o); } break;
		case 0x00: if(n[1] == 0x00 && n[2] >= 0x02 && n[3] == 0x00) return WK_.to_workbook(d, o); break;
		case 0x03: case 0x83: case 0x8B: return DBF.to_workbook(d, o);
		case 0x7B: if(n[1] == 0x5C && n[2] == 0x72 && n[3] == 0x74) return RTF.to_workbook(d, o); break;
		case 0x0A: case 0x0D: case 0x20: return read_plaintext_raw(d, o);
	}
	if(n[2] <= 12 && n[3] <= 31) return DBF.to_workbook(d, o);
	if(0x20>n[0]||n[0]>0x7F) throw new Error("Unsupported file " + n.join("|"));
	return PRN.to_workbook(d, o);
}

function readFileSync(filename/*:string*/, opts/*:?ParseOpts*/)/*:Workbook*/ {
	var o = opts||{}; o.type = 'file';
	return readSync(filename, o);
}
