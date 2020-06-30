function firstbyte(f/*:RawData*/,o/*:?TypeOpts*/)/*:Array<number>*/ {
	var x = "";
	switch((o||{}).type || "base64") {
		case 'buffer': return [f[0], f[1], f[2], f[3], f[4], f[5], f[6], f[7]];
		case 'base64': x = Base64.decode(f.slice(0,12)); break;
		case 'binary': x = f; break;
		case 'array':  return [f[0], f[1], f[2], f[3], f[4], f[5], f[6], f[7]];
		default: throw new Error("Unrecognized type " + (o && o.type || "undefined"));
	}
	return [x.charCodeAt(0), x.charCodeAt(1), x.charCodeAt(2), x.charCodeAt(3), x.charCodeAt(4), x.charCodeAt(5), x.charCodeAt(6), x.charCodeAt(7)];
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
	zip = zip_read(d, o);
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

function bstrify(data/*:string*/)/*:string*/ {
	return !data.match(/[^\x00-\x7F]/) ? data : utf8write(data);
}

function read_prn(data, d, o, str) {
	if(str) { o.type = "string"; return PRN.to_workbook(data, o); }
	return PRN.to_workbook(d, o);
}

function readSync(data/*:RawData*/, opts/*:?ParseOpts*/)/*:Workbook*/ {
	reset_cp();
	if(typeof ArrayBuffer !== 'undefined' && data instanceof ArrayBuffer) return readSync(new Uint8Array(data), opts);
	var d = data, n = [0,0,0,0], str = false;
	var o = opts||{};
	if(o.cellStyles) { o.cellNF = true; o.sheetStubs = true; }
	_ssfopts = {};
	if(o.dateNF) _ssfopts.dateNF = o.dateNF;
	if(!o.type) o.type = (has_buf && Buffer.isBuffer(data)) ? "buffer" : "base64";
	if(o.type == "file") { o.type = has_buf ? "buffer" : "binary"; d = read_binary(data); }
	if(o.type == "string") { str = true; o.type = "binary"; o.codepage = 65001; d = bstrify(data); }
	if(o.type == 'array' && typeof Uint8Array !== 'undefined' && data instanceof Uint8Array && typeof ArrayBuffer !== 'undefined') {
		// $FlowIgnore
		var ab=new ArrayBuffer(3), vu=new Uint8Array(ab); vu.foo="bar";
		// $FlowIgnore
		if(!vu.foo) {o=dup(o); o.type='array'; return readSync(ab2a(d), o);}
	}
	switch((n = firstbyte(d, o))[0]) {
		case 0xD0: if(n[1] === 0xCF && n[2] === 0x11 && n[3] === 0xE0 && n[4] === 0xA1 && n[5] === 0xB1 && n[6] === 0x1A && n[7] === 0xE1) return read_cfb(CFB.read(d, o), o); break;
		case 0x09: if(n[1] <= 0x04) return parse_xlscfb(d, o); break;
		case 0x3C: return parse_xlml(d, o);
		case 0x49: if(n[1] === 0x44) return read_wb_ID(d, o); break;
		case 0x54: if(n[1] === 0x41 && n[2] === 0x42 && n[3] === 0x4C) return DIF.to_workbook(d, o); break;
		case 0x50: return (n[1] === 0x4B && n[2] < 0x09 && n[3] < 0x09) ? read_zip(d, o) : read_prn(data, d, o, str);
		case 0xEF: return n[3] === 0x3C ? parse_xlml(d, o) : read_prn(data, d, o, str);
		case 0xFF: if(n[1] === 0xFE) { return read_utf16(d, o); } break;
		case 0x00: if(n[1] === 0x00 && n[2] >= 0x02 && n[3] === 0x00) return WK_.to_workbook(d, o); break;
		case 0x03: case 0x83: case 0x8B: case 0x8C: return DBF.to_workbook(d, o);
		case 0x7B: if(n[1] === 0x5C && n[2] === 0x72 && n[3] === 0x74) return RTF.to_workbook(d, o); break;
		case 0x0A: case 0x0D: case 0x20: return read_plaintext_raw(d, o);
	}
	if(DBF.versions.indexOf(n[0]) > -1 && n[2] <= 12 && n[3] <= 31) return DBF.to_workbook(d, o);
	return read_prn(data, d, o, str);
}

function readFileSync(filename/*:string*/, opts/*:?ParseOpts*/)/*:Workbook*/ {
	var o = opts||{}; o.type = 'file';
	return readSync(filename, o);
}
