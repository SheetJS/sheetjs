var RTF = (function() {
	function rtf_to_sheet(d/*:RawData*/, opts)/*:Worksheet*/ {
		switch(opts.type) {
			case 'base64': return rtf_to_sheet_str(Base64.decode(d), opts);
			case 'binary': return rtf_to_sheet_str(d, opts);
			case 'buffer': return rtf_to_sheet_str(has_buf && Buffer.isBuffer(d) ? d.toString('binary') : a2s(d), opts);
			case 'array':  return rtf_to_sheet_str(cc2str(d), opts);
		}
		throw new Error("Unrecognized type " + opts.type);
	}

	/* TODO: this is a stub */
	function rtf_to_sheet_str(str/*:string*/, opts)/*:Worksheet*/ {
		var o = opts || {};
		var ws/*:Worksheet*/ = o.dense ? ([]/*:any*/) : ({}/*:any*/);

		var rows = str.match(/\\trowd.*?\\row\b/g);
		if(!rows.length) throw new Error("RTF missing table");
		var range/*:Range*/ = ({s: {c:0, r:0}, e: {c:0, r:rows.length - 1}}/*:any*/);
		rows.forEach(function(rowtf, R) {
			if(Array.isArray(ws)) ws[R] = [];
			var rtfre = /\\\w+\b/g;
			var last_index = 0;
			var res;
			var C = -1;
			while((res = rtfre.exec(rowtf))) {
				switch(res[0]) {
					case "\\cell":
						var data = rowtf.slice(last_index, rtfre.lastIndex - res[0].length);
						if(data[0] == " ") data = data.slice(1);
						++C;
						if(data.length) {
							// TODO: value parsing, including codepage adjustments
							var cell = {v: data, t:"s"};
							if(Array.isArray(ws)) ws[R][C] = cell;
							else ws[encode_cell({r:R, c:C})] = cell;
						}
						break;
				}
				last_index = rtfre.lastIndex;
			}
			if(C > range.e.c) range.e.c = C;
		});
		ws['!ref'] = encode_range(range);
		return ws;
	}

	function rtf_to_workbook(d/*:RawData*/, opts)/*:Workbook*/ { return sheet_to_workbook(rtf_to_sheet(d, opts), opts); }

	/* TODO: this is a stub */
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
				o.push(" " + (cell.w || (format_cell(cell), cell.w)));
				o.push("\\cell");
			}
			o.push("\\pard\\intbl\\row");
		}
		return o.join("") + "}";
	}

	return {
		to_workbook: rtf_to_workbook,
		to_sheet: rtf_to_sheet,
		from_sheet: sheet_to_rtf
	};
})();
