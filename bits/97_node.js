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
		var row/*:?string*/ = "", cols = [];
		o.dense = Array.isArray(sheet);
		for(var C = r.s.c; C <= r.e.c; ++C) cols[C] = encode_col(C);
		var R = r.s.r;
		stream._read = function() {
			if(R > r.e.r) return stream.push(null);
			while(R <= r.e.r) {
				row = make_csv_row(sheet, r, R, cols, fs, rs, FS, o);
				++R;
				if(row != null) {
					if(o.strip) row = row.replace(endregex,"");
					stream.push(row + RS);
					break;
				}
			}
		};
		return stream;
	};

	var write_html_stream = function(sheet/*:Worksheet*/, opts/*:?Sheet2HTMLOpts*/) {
		var stream = Readable();

		var o = opts == null ? {} : opts;
		var r = decode_range(sheet['!ref']), cell/*:Cell*/;
		o.dense = Array.isArray(sheet);
		stream.push(HTML_.BEGIN);

		var R = r.s.r;
		var end = false;
		stream._read = function() {
			if(R > r.e.r) {
				if(!end) { end = true; stream.push(HTML_.END); }
				return stream.push(null);
			}
			while(R <= r.e.r) {
				stream.push(HTML_._row(sheet, r, R, o));
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

