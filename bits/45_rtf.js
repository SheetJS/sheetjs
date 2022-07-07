function rtf_to_workbook(d/*:RawData*/, opts)/*:Workbook*/ {
	switch(opts.type) {
		case 'base64': return rtf_to_book_str(Base64_decode(d), opts);
		case 'binary': return rtf_to_book_str(d, opts);
		case 'buffer': return rtf_to_book_str(has_buf && Buffer.isBuffer(d) ? d.toString('binary') : a2s(d), opts);
		case 'array':  return rtf_to_book_str(cc2str(d), opts);
	}
	throw new Error("Unrecognized type " + opts.type);
}

/* TODO: RTF technically can store multiple tables, even if Excel does not */
function rtf_to_book_str(str/*:string*/, opts)/*:Workbook*/ {
	var o = opts || {};
	var sname = o.sheet || "Sheet1";
	var ws/*:Worksheet*/ = o.dense ? ([]/*:any*/) : ({}/*:any*/);
	var wb/*:Workbook*/ = { SheetNames: [ sname ], Sheets: {} };
	wb.Sheets[sname] = ws;

	var rows = str.match(/\\trowd[\s\S]*?\\row\b/g);
	if(!rows.length) throw new Error("RTF missing table");
	var range/*:Range*/ = ({s: {c:0, r:0}, e: {c:0, r:rows.length - 1}}/*:any*/);
	rows.forEach(function(rowtf, R) {
		if(Array.isArray(ws)) ws[R] = [];
		var rtfre = /\\[\w\-]+\b/g;
		var last_index = 0;
		var res;
		var C = -1;
		var payload = [];
		while((res = rtfre.exec(rowtf))) {
			var data = rowtf.slice(last_index, rtfre.lastIndex - res[0].length);
			if(data.charCodeAt(0) == 0x20) data = data.slice(1);
			if(data.length) payload.push(data);
			switch(res[0]) {
				case "\\cell":
					++C;
					if(payload.length) {
						// TODO: value parsing, including codepage adjustments
						var cell = {v: payload.join(""), t:"s"};
						if(cell.v == "TRUE" || cell.v == "FALSE") { cell.v = cell.v == "TRUE"; cell.t = "b"; }
						else if(!isNaN(fuzzynum(cell.v))) { cell.t = 'n'; if(o.cellText !== false) cell.w = cell.v; cell.v = fuzzynum(cell.v); }

						if(Array.isArray(ws)) ws[R][C] = cell;
						else ws[encode_cell({r:R, c:C})] = cell;
					}
					payload = [];
					break;
				case "\\par": // NOTE: Excel serializes both "\r" and "\n" as "\\par"
					payload.push("\n");
					break;
			}
			last_index = rtfre.lastIndex;
		}
		if(C > range.e.c) range.e.c = C;
	});
	ws['!ref'] = encode_range(range);
	return wb;
}

/* TODO: standardize sheet names as titles for tables */
function sheet_to_rtf(ws/*:Worksheet*//*::, opts*/)/*:string*/ {
	var o = ["{\\rtf1\\ansi"];
	var r = safe_decode_range(ws['!ref']), cell/*:Cell*/;
	var dense = Array.isArray(ws);
	for(var R = r.s.r; R <= r.e.r; ++R) {
		o.push("\\trowd\\trautofit1");
		for(var C = r.s.c; C <= r.e.c; ++C) o.push("\\cellx" + (C+1));
		o.push("\\pard\\intbl");
		for(C = r.s.c; C <= r.e.c; ++C) {
			var coord = encode_cell({r:R,c:C});
			cell = dense ? (ws[R]||[])[C]: ws[coord];
			if(!cell || cell.v == null && (!cell.f || cell.F)) continue;
			o.push(" " + (cell.w || (format_cell(cell), cell.w)).replace(/[\r\n]/g, "\\par "));
			o.push("\\cell");
		}
		o.push("\\pard\\intbl\\row");
	}
	return o.join("") + "}";
}
