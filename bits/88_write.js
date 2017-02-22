function write_zip_type(wb/*:Workbook*/, opts/*:?WriteOpts*/) {
	var o = opts||{};
	var z = write_zip(wb, o);
	var oopts = {};
	if(o.compression) oopts.compression = 'DEFLATE';
	switch(o.type) {
		case "base64": oopts.type = "base64"; break;
		case "binary": oopts.type = "string"; break;
		case "buffer":
		case "file": oopts.type = "nodebuffer"; break;
		default: throw new Error("Unrecognized type " + o.type);
	}
	if(o.type === "file") return _fs.writeFileSync(o.file, z.generate(oopts));
	return z.generate(oopts);
}

function write_string_type(out/*:string*/, opts/*:WriteOpts*/) {
	switch(opts.type) {
		case "base64": break; // TODO
		case "binary": break; // TODO
		case "file": return _fs.writeFileSync(opts.file, out, {encoding:'utf8'});
		case "buffer": break; // TODO
		default: return out;
	}
}

function write_binary_type(out, opts/*:WriteOpts*/) {
	switch(opts.type) {
		case "base64": break; // TODO
		case "binary":
			var bstr = "";
			for(var i = 0; i < out.length; ++i) bstr += String.fromCharCode(out[i]);
			return bstr;
		case "file": return _fs.writeFileSync(opts.file, out);
		case "buffer": return out;
		default: throw new Error("Unrecognized type " + opts.type);
	}
}

function writeSync(wb/*:Workbook*/, opts/*:?WriteOpts*/) {
	var o = opts||{};
	switch(o.bookType) {
		case 'xml': return write_string_type(write_xlml(wb, o), o);
		case 'csv': return write_string_type(write_csv_str(wb, o), o);
		case 'fods': return write_string_type(write_ods(wb, o), o);
		case 'biff2': return write_binary_type(write_biff_buf(wb, o), o);
		default: return write_zip_type(wb, o);
	}
}

function writeFileSync(wb/*:Workbook*/, filename/*:string*/, opts/*:?WriteFileOpts*/) {
	var o = opts||{}; o.type = 'file';
	o.file = filename;
	if(!o.bookType) switch(o.file.slice(-5).toLowerCase()) {
		case '.xlsx': o.bookType = 'xlsx'; break;
		case '.xlsm': o.bookType = 'xlsm'; break;
		case '.xlsb': o.bookType = 'xlsb'; break;
		case '.fods': o.bookType = 'fods'; break;
	default: switch(o.file.slice(-4).toLowerCase()) {
		case '.xls': o.bookType = 'biff2'; break;
		case '.xml': o.bookType = 'xml'; break;
		case '.ods': o.bookType = 'ods'; break;
		case '.csv': o.bookType = 'csv'; break;
	}}
	return writeSync(wb, o);
}

