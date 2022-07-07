function write_cfb_ctr(cfb/*:CFBContainer*/, o/*:WriteOpts*/)/*:any*/ {
	switch(o.type) {
		case "base64": case "binary": break;
		case "buffer": case "array": o.type = ""; break;
		case "file": return write_dl(o.file, CFB.write(cfb, {type:has_buf ? 'buffer' : ""}));
		case "string": throw new Error("'string' output type invalid for '" + o.bookType + "' files");
		default: throw new Error("Unrecognized type " + o.type);
	}
	return CFB.write(cfb, o);
}

function write_zip(wb/*:Workbook*/, opts/*:WriteOpts*/)/*:ZIP*/ {
	switch(opts.bookType) {
		case "ods": return write_ods(wb, opts);
		case "numbers": return write_numbers_iwa(wb, opts);
		case "xlsb": return write_zip_xlsb(wb, opts);
		default: return write_zip_xlsx(wb, opts);
	}
}

/*:: declare var encrypt_agile:any; */
function write_zip_type(wb/*:Workbook*/, opts/*:?WriteOpts*/)/*:any*/ {
	var o = dup(opts||{});
	var z = write_zip(wb, o);
	return write_zip_denouement(z, o);
}
function write_zip_typeXLSX(wb/*:Workbook*/, opts/*:?WriteOpts*/)/*:any*/ {
	var o = dup(opts||{});
	var z = write_zip_xlsx(wb, o);
	return write_zip_denouement(z, o);
}
function write_zip_denouement(z/*:any*/, o/*:?WriteOpts*/)/*:any*/ {
	var oopts = {};
	var ftype = has_buf ? "nodebuffer" : (typeof Uint8Array !== "undefined" ? "array" : "string");
	if(o.compression) oopts.compression = 'DEFLATE';
	if(o.password) oopts.type = ftype;
	else switch(o.type) {
		case "base64": oopts.type = "base64"; break;
		case "binary": oopts.type = "string"; break;
		case "string": throw new Error("'string' output type invalid for '" + o.bookType + "' files");
		case "buffer":
		case "file": oopts.type = ftype; break;
		default: throw new Error("Unrecognized type " + o.type);
	}
	var out = z.FullPaths ? CFB.write(z, {fileType:"zip", type: /*::(*/{"nodebuffer": "buffer", "string": "binary"}/*:: :any)*/[oopts.type] || oopts.type, compression: !!o.compression}) : z.generate(oopts);
	if(typeof Deno !== "undefined") {
		if(typeof out == "string") {
			if(o.type == "binary" || o.type == "base64") return out;
			out = new Uint8Array(s2ab(out));
		}
	}
/*jshint -W083 */
	if(o.password && typeof encrypt_agile !== 'undefined') return write_cfb_ctr(encrypt_agile(out, o.password), o); // eslint-disable-line no-undef
/*jshint +W083 */
	if(o.type === "file") return write_dl(o.file, out);
	return o.type == "string" ? utf8read(/*::(*/out/*:: :any)*/) : out;
}

function write_cfb_type(wb/*:Workbook*/, opts/*:?WriteOpts*/)/*:any*/ {
	var o = opts||{};
	var cfb/*:CFBContainer*/ = write_xlscfb(wb, o);
	return write_cfb_ctr(cfb, o);
}

function write_string_type(out/*:string*/, opts/*:WriteOpts*/, bom/*:?string*/)/*:any*/ {
	if(!bom) bom = "";
	var o = bom + out;
	switch(opts.type) {
		case "base64": return Base64_encode(utf8write(o));
		case "binary": return utf8write(o);
		case "string": return out;
		case "file": return write_dl(opts.file, o, 'utf8');
		case "buffer": {
			if(has_buf) return Buffer_from(o, 'utf8');
			else if(typeof TextEncoder !== "undefined") return new TextEncoder().encode(o);
			else return write_string_type(o, {type:'binary'}).split("").map(function(c) { return c.charCodeAt(0); });
		}
	}
	throw new Error("Unrecognized type " + opts.type);
}

function write_stxt_type(out/*:string*/, opts/*:WriteOpts*/)/*:any*/ {
	switch(opts.type) {
		case "base64": return Base64_encode_pass(out);
		case "binary": return out;
		case "string": return out; /* override in sheet_to_txt */
		case "file": return write_dl(opts.file, out, 'binary');
		case "buffer": {
			if(has_buf) return Buffer_from(out, 'binary');
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
			return opts.type == 'base64' ? Base64_encode(bstr) : opts.type == 'string' ? utf8read(bstr) : bstr;
		case "file": return write_dl(opts.file, out);
		case "buffer": return out;
		default: throw new Error("Unrecognized type " + opts.type);
	}
}

function writeSyncXLSX(wb/*:Workbook*/, opts/*:?WriteOpts*/) {
	reset_cp();
	check_wb(wb);
	var o = dup(opts||{});
	if(o.cellStyles) { o.cellNF = true; o.sheetStubs = true; }
	if(o.type == "array") { o.type = "binary"; var out/*:string*/ = (writeSyncXLSX(wb, o)/*:any*/); o.type = "array"; return s2ab(out); }
	return write_zip_typeXLSX(wb, o);
}

function writeSync(wb/*:Workbook*/, opts/*:?WriteOpts*/) {
	reset_cp();
	check_wb(wb);
	var o = dup(opts||{});
	if(o.cellStyles) { o.cellNF = true; o.sheetStubs = true; }
	if(o.type == "array") { o.type = "binary"; var out/*:string*/ = (writeSync(wb, o)/*:any*/); o.type = "array"; return s2ab(out); }
	var idx = 0;
	if(o.sheet) {
		if(typeof o.sheet == "number") idx = o.sheet;
		else idx = wb.SheetNames.indexOf(o.sheet);
		if(!wb.SheetNames[idx]) throw new Error("Sheet not found: " + o.sheet + " : " + (typeof o.sheet));
	}
	switch(o.bookType || 'xlsb') {
		case 'xml':
		case 'xlml': return write_string_type(write_xlml(wb, o), o);
		case 'slk':
		case 'sylk': return write_string_type(SYLK.from_sheet(wb.Sheets[wb.SheetNames[idx]], o, wb), o);
		case 'htm':
		case 'html': return write_string_type(sheet_to_html(wb.Sheets[wb.SheetNames[idx]], o), o);
		case 'txt': return write_stxt_type(sheet_to_txt(wb.Sheets[wb.SheetNames[idx]], o), o);
		case 'csv': return write_string_type(sheet_to_csv(wb.Sheets[wb.SheetNames[idx]], o), o, "\ufeff");
		case 'dif': return write_string_type(DIF.from_sheet(wb.Sheets[wb.SheetNames[idx]], o), o);
		case 'dbf': return write_binary_type(DBF.from_sheet(wb.Sheets[wb.SheetNames[idx]], o), o);
		case 'prn': return write_string_type(PRN.from_sheet(wb.Sheets[wb.SheetNames[idx]], o), o);
		case 'rtf': return write_string_type(sheet_to_rtf(wb.Sheets[wb.SheetNames[idx]], o), o);
		case 'eth': return write_string_type(ETH.from_sheet(wb.Sheets[wb.SheetNames[idx]], o), o);
		case 'fods': return write_string_type(write_ods(wb, o), o);
		case 'wk1': return write_binary_type(WK_.sheet_to_wk1(wb.Sheets[wb.SheetNames[idx]], o), o);
		case 'wk3': return write_binary_type(WK_.book_to_wk3(wb, o), o);
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
		case 'numbers':
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

function writeFileSyncXLSX(wb/*:Workbook*/, filename/*:string*/, opts/*:?WriteFileOpts*/) {
	var o = opts||{}; o.type = 'file';
	o.file = filename;
	resolve_book_type(o);
	return writeSyncXLSX(wb, o);
}


function writeFileAsync(filename/*:string*/, wb/*:Workbook*/, opts/*:?WriteFileOpts*/, cb/*:?(e?:ErrnoError)=>void*/) {
	var o = opts||{}; o.type = 'file';
	o.file = filename;
	resolve_book_type(o);
	o.type = 'buffer';
	var _cb = cb; if(!(_cb instanceof Function)) _cb = (opts/*:any*/);
	return _fs.writeFile(filename, writeSync(wb, o), _cb);
}
