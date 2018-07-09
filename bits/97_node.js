if(has_buf && typeof require != 'undefined') (function() {
	var Readable = require('stream').Readable;

	var write_csv_stream = function(sheet/*:Worksheet*/, opts/*:?Sheet2CSVOpts*/) {
		var stream = Readable();
		var o = opts == null ? {} : opts;
		if(sheet == null || sheet["!ref"] == null) { stream.push(null); return stream; }
		var r = safe_decode_range(sheet["!ref"]);
		var FS = o.FS !== undefined ? o.FS : ",", fs = FS.charCodeAt(0);
		var RS = o.RS !== undefined ? o.RS : "\n", rs = RS.charCodeAt(0);
		var endregex = new RegExp((FS=="|" ? "\\|" : FS)+"+$");
		var row/*:?string*/ = "", cols/*:Array<string>*/ = [];
		o.dense = Array.isArray(sheet);
		var colinfo/*:Array<ColInfo>*/ = o.skipHidden && sheet["!cols"] || [];
		var rowinfo/*:Array<RowInfo>*/ = o.skipHidden && sheet["!rows"] || [];
		for(var C = r.s.c; C <= r.e.c; ++C) if (!((colinfo[C]||{}).hidden)) cols[C] = encode_col(C);
		var R = r.s.r;
		var BOM = false;
		stream._read = function() {
			if(!BOM) { BOM = true; return stream.push("\uFEFF"); }
			if(R > r.e.r) return stream.push(null);
			while(R <= r.e.r) {
				++R;
				if ((rowinfo[R-1]||{}).hidden) continue;
				row = make_csv_row(sheet, r, R-1, cols, fs, rs, FS, o);
				if(row != null) {
					if(o.strip) row = row.replace(endregex,"");
					stream.push(row + RS);
					break;
				}
			}
		};
		return stream;
	};

	var write_html_stream = function(ws/*:Worksheet*/, opts/*:?Sheet2HTMLOpts*/) {
		var stream = Readable();

		var o = opts || {};
		var header = o.header != null ? o.header : HTML_.BEGIN;
		var footer = o.footer != null ? o.footer : HTML_.END;
		stream.push(header);
		var r = decode_range(ws['!ref']);
		o.dense = Array.isArray(ws);
		stream.push(HTML_._preamble(ws, r, o));
		var R = r.s.r;
		var end = false;
		stream._read = function() {
			if(R > r.e.r) {
				if(!end) { end = true; stream.push("</table>" + footer); }
				return stream.push(null);
			}
			while(R <= r.e.r) {
				stream.push(HTML_._row(ws, r, R, o));
				++R;
				break;
			}
		};
		return stream;
	};

	var write_json_stream = function(sheet/*:Worksheet*/, opts/*:?Sheet2CSVOpts*/) {
		var stream = Readable({objectMode:true});

		if(sheet == null || sheet["!ref"] == null) { stream.push(null); return stream; }
		var val = {t:'n',v:0}, header = 0, offset = 1, hdr/*:Array<any>*/ = [], v=0, vv="";
		var r = {s:{r:0,c:0},e:{r:0,c:0}};
		var o = opts || {};
		var range = o.range != null ? o.range : sheet["!ref"];
		if(o.header === 1) header = 1;
		else if(o.header === "A") header = 2;
		else if(Array.isArray(o.header)) header = 3;
		switch(typeof range) {
			case 'string': r = safe_decode_range(range); break;
			case 'number': r = safe_decode_range(sheet["!ref"]); r.s.r = range; break;
			default: r = range;
		}
		if(header > 0) offset = 0;
		var rr = encode_row(r.s.r);
		var cols/*:Array<string>*/ = [];
		var counter = 0;
		var dense = Array.isArray(sheet);
		var R = r.s.r, C = 0, CC = 0;
		if(dense && !sheet[R]) sheet[R] = [];
		for(C = r.s.c; C <= r.e.c; ++C) {
			cols[C] = encode_col(C);
			val = dense ? sheet[R][C] : sheet[cols[C] + rr];
			switch(header) {
				case 1: hdr[C] = C - r.s.c; break;
				case 2: hdr[C] = cols[C]; break;
				case 3: hdr[C] = o.header[C - r.s.c]; break;
				default:
					if(val == null) val = {w: "__EMPTY", t: "s"};
					vv = v = format_cell(val, null, o);
					counter = 0;
					for(CC = 0; CC < hdr.length; ++CC) if(hdr[CC] == vv) vv = v + "_" + (++counter);
					hdr[C] = vv;
			}
		}
		R = r.s.r + offset;
		stream._read = function() {
			if(R > r.e.r) return stream.push(null);
			while(R <= r.e.r) {
				++R;
				//if ((rowinfo[R-1]||{}).hidden) continue;
				var row = make_json_row(sheet, r, R, cols, header, hdr, dense, o);
				if((row.isempty === false) || (header === 1 ? o.blankrows !== false : !!o.blankrows)) {
					stream.push(row.row);
					break;
				}
			}
		};
		return stream;
	};

	XLSX.stream = {
		to_json: write_json_stream,
		to_html: write_html_stream,
		to_csv: write_csv_stream
	};
})();

