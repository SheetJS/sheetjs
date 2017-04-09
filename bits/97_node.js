if(has_buf && typeof require != 'undefined') (function() {
	var Readable = require('stream').Readable;

	var write_csv_stream = function(sheet/*:Worksheet*/, opts/*:?Sheet2CSVOpts*/) {
		var stream = Readable();
		var out = "", txt = "", qreg = /"/g;
		var o = opts == null ? {} : opts;
		if(sheet == null || sheet["!ref"] == null) { stream.push(null); return stream; }
		var r = safe_decode_range(sheet["!ref"]);
		var FS = o.FS !== undefined ? o.FS : ",", fs = FS.charCodeAt(0);
		var RS = o.RS !== undefined ? o.RS : "\n", rs = RS.charCodeAt(0);
		var endregex = new RegExp((FS=="|" ? "\\|" : FS)+"+$");
		var row = "", rr = "", cols = [];
		var i = 0, cc = 0, val;
		var R = 0, C = 0;
		var dense = Array.isArray(sheet);
		for(C = r.s.c; C <= r.e.c; ++C) cols[C] = encode_col(C);
		R = r.s.r;
		stream._read = function() {
			if(R > r.e.r) return stream.push(null);
			while(true) {
			var isempty = true;
			row = "";
			rr = encode_row(R);
			for(C = r.s.c; C <= r.e.c; ++C) {
				val = dense ? (sheet[R]||[])[C]: sheet[cols[C] + rr];
				if(val == null) txt = "";
				else if(val.v != null) {
					isempty = false;
					txt = ''+format_cell(val, null, o);
					for(i = 0, cc = 0; i !== txt.length; ++i) if((cc = txt.charCodeAt(i)) === fs || cc === rs || cc === 34) {
						txt = "\"" + txt.replace(qreg, '""') + "\""; break; }
				} else if(val.f != null && !val.F) {
					isempty = false;
					txt = '=' + val.f; if(txt.indexOf(",") >= 0) txt = '"' + txt.replace(qreg, '""') + '"';
				} else txt = "";
				/* NOTE: Excel CSV does not support array formulae */
				row += (C === r.s.c ? "" : FS) + txt;
			}
			if(o.blankrows === false && isempty) { ++R; continue; }
			if(o.strip) row = row.replace(endregex,"");
			stream.push(row + RS);
			++R;
			break;
			}
		};
		return stream;
	};


	XLSX.stream = {
		to_csv: write_csv_stream
	};
})();
