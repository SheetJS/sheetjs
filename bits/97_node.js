if(has_buf && typeof require != 'undefined') (function() {
	var Readable = require('stream').Readable;

	var write_csv_stream = function(sheet/*:Worksheet*/, opts/*:?Sheet2CSVOpts*/) {
		var stream = Readable();
		var out = "";
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
		stream._read = function() {
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

	XLSX.stream = {
		to_html: write_html_stream,
		to_csv: write_csv_stream
	};
})();

