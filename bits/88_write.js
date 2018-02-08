function write_zip_type(wb/*:Workbook*/, opts/*:?WriteOpts*/)/*:any*/ {
	var o = opts||{};
	var z = write_zip(wb, o);
	var oopts = {};
	if(o.compression) oopts.compression = 'DEFLATE';
	switch(o.type) {
		case "base64": oopts.type = "base64"; break;
		case "binary": oopts.type = "string"; break;
		case "string": throw new Error("'string' output type invalid for '" + o.bookType + "' files");
		case "buffer":
		case "file": oopts.type = has_buf ? "nodebuffer" : "string"; break;
		default: throw new Error("Unrecognized type " + o.type);
	}
	if(o.type === "file") return write_dl(o.file, z.generate(oopts));
	var out = z.generate(oopts);
	// $FlowIgnore
	return o.type == "string" ? utf8read(out) : out;
}

function write_cfb_type(wb/*:Workbook*/, opts/*:?WriteOpts*/)/*:any*/ {
	var o = opts||{};
	var cfb/*:CFBContainer*/ = write_xlscfb(wb, o);
	switch(o.type) {
		case "base64": case "binary": break;
		case "buffer": case "array": o.type = ""; break;
		case "file": return write_dl(o.file, CFB.write(cfb, {type:has_buf ? 'buffer' : ""}));
		case "string": throw new Error("'string' output type invalid for '" + o.bookType + "' files");
		default: throw new Error("Unrecognized type " + o.type);
	}
	return CFB.write(cfb, o);
}

function write_string_type(out/*:string*/, opts/*:WriteOpts*/, bom/*:?string*/)/*:any*/ {
	if(!bom) bom = "";
	var o = bom + out;
	switch(opts.type) {
		case "base64": return Base64.encode(utf8write(o));
		case "binary": return utf8write(o);
		case "string": return out;
		case "file": return write_dl(opts.file, o, 'utf8');
		case "buffer": {
			if(has_buf) return new Buffer(o, 'utf8');
			else return write_string_type(o, {type:'binary'}).split("").map(function(c) { return c.charCodeAt(0); });
		}
	}
	throw new Error("Unrecognized type " + opts.type);
}

function write_stxt_type(out/*:string*/, opts/*:WriteOpts*/)/*:any*/ {
	switch(opts.type) {
		case "base64": return Base64.encode(out);
		case "binary": return out;
		case "string": return out; /* override in sheet_to_txt */
		case "file": return write_dl(opts.file, out, 'binary');
		case "buffer": {
			if(has_buf) return new Buffer(out, 'binary');
			else return out.split("").map(function(c) { return c.charCodeAt(0); });
		}
	}
	throw new Error("Unrecognized type " + opts.type);
}

/* TODO: test consistency */
function write_binary_type(out, opts/*:WriteOpts*/)/*:any*/ {
	switch(opts.type) {
		case "string":
		case "base64":
		case "binary":
			var bstr = "";
			// $FlowIgnore
			for(var i = 0; i < out.length; ++i) bstr += String.fromCharCode(out[i]);
			return opts.type == 'base64' ? Base64.encode(bstr) : opts.type == 'string' ? utf8read(bstr) : bstr;
		case "file": return write_dl(opts.file, out);
		case "buffer": return out;
		default: throw new Error("Unrecognized type " + opts.type);
	}
}

function writeSync(wb/*:Workbook*/, opts/*:?WriteOpts*/) {
	check_wb(wb);
	var o = opts||{};
	if(o.type == "array") { o.type = "binary"; var out/*:string*/ = (writeSync(wb, o)/*:any*/); o.type = "array"; return s2ab(out); }
	switch(o.bookType || 'xlsb') {
		case 'xml':
		case 'xlml': return write_string_type(write_xlml(wb, o), o);
		case 'slk':
		case 'sylk': return write_string_type(write_slk_str(wb, o), o);
		case 'htm':
		case 'html': return write_string_type(write_htm_str(wb, o), o);
		case 'txt': return write_stxt_type(write_txt_str(wb, o), o);
		case 'csv': return write_string_type(write_csv_str(wb, o), o, "\ufeff");
		case 'dif': return write_string_type(write_dif_str(wb, o), o);
		// $FlowIgnore
		case 'dbf': return write_binary_type(write_dbf_buf(wb, o), o);
		case 'prn': return write_string_type(write_prn_str(wb, o), o);
		case 'rtf': return write_string_type(write_rtf_str(wb, o), o);
		case 'eth': return write_string_type(write_eth_str(wb, o), o);
		case 'fods': return write_string_type(write_ods(wb, o), o);
		case 'biff2': if(!o.biff) o.biff = 2; /* falls through */
		case 'biff3': if(!o.biff) o.biff = 3; /* falls through */
		case 'biff4': if(!o.biff) o.biff = 4; return write_binary_type(write_biff_buf(wb, o), o);
		case 'biff5': if(!o.biff) o.biff = 5; /* falls through */
		case 'biff8':
		case 'xla':
		case 'xls': if(!o.biff) o.biff = 8; return write_cfb_type(wb, o);
		case 'xlsx':
		case 'xlsm':
		case 'xlam':
		case 'xlsb':
		case 'ods': return write_zip_type(wb, o);
		default: throw new Error ("Unrecognized bookType |" + o.bookType + "|");
	}
}

function resolve_book_type(o/*:WriteFileOpts*/) {
	if(o.bookType) return;
	var _BT = {
		"xls": "biff8",
		"htm": "html",
		"slk": "sylk",
		"socialcalc": "eth",
		"Sh33tJS": "WTF"
	};
	var ext = o.file.slice(o.file.lastIndexOf(".")).toLowerCase();
	if(ext.match(/^\.[a-z]+$/)) o.bookType = ext.slice(1);
	o.bookType = _BT[o.bookType] || o.bookType;
}

function writeFileSync(wb/*:Workbook*/, filename/*:string*/, opts/*:?WriteFileOpts*/) {
	var o = opts||{}; o.type = 'file';
	o.file = filename;
	resolve_book_type(o);
	return writeSync(wb, o);
}

function writeFileAsync(filename/*:string*/, wb/*:Workbook*/, opts/*:?WriteFileOpts*/, cb/*:?(e?:ErrnoError)=>void*/) {
	var o = opts||{}; o.type = 'file';
	o.file = filename;
	resolve_book_type(o);
	o.type = 'buffer';
	var _cb = cb; if(!(_cb instanceof Function)) _cb = (opts/*:any*/);
	return _fs.writeFile(filename, writeSync(wb, o), _cb);
}
