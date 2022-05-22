var _Readable;
function set_readable(R) { _Readable = R; }

function write_csv_stream(sheet/*:Worksheet*/, opts/*:?Sheet2CSVOpts*/) {
	var stream = _Readable();
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
	var BOM = false, w = 0;
	stream._read = function() {
		if(!BOM) { BOM = true; return stream.push("\uFEFF"); }
		while(R <= r.e.r) {
			++R;
			if ((rowinfo[R-1]||{}).hidden) continue;
			row = make_csv_row(sheet, r, R-1, cols, fs, rs, FS, o);
			if(row != null) {
				if(o.strip) row = row.replace(endregex,"");
				if(row || (o.blankrows !== false)) return stream.push((w++ ? RS : "") + row);
			}
		}
		return stream.push(null);
	};
	return stream;
}

function write_html_stream(ws/*:Worksheet*/, opts/*:?Sheet2HTMLOpts*/) {
	var stream = _Readable();

	var o = opts || {};
	var header = o.header != null ? o.header : HTML_BEGIN;
	var footer = o.footer != null ? o.footer : HTML_END;
	stream.push(header);
	var r = decode_range(ws['!ref']);
	o.dense = Array.isArray(ws);
	stream.push(make_html_preamble(ws, r, o));
	var R = r.s.r;
	var end = false;
	stream._read = function() {
		if(R > r.e.r) {
			if(!end) { end = true; stream.push("</table>" + footer); }
			return stream.push(null);
		}
		while(R <= r.e.r) {
			stream.push(make_html_row(ws, r, R, o));
			++R;
			break;
		}
	};
	return stream;
}

function write_json_stream(sheet/*:Worksheet*/, opts/*:?Sheet2CSVOpts*/) {
	var stream = _Readable({objectMode:true});

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
	var R = r.s.r, C = 0;
	var header_cnt = {};
	if(dense && !sheet[R]) sheet[R] = [];
	var colinfo/*:Array<ColInfo>*/ = o.skipHidden && sheet["!cols"] || [];
	var rowinfo/*:Array<RowInfo>*/ = o.skipHidden && sheet["!rows"] || [];
	for(C = r.s.c; C <= r.e.c; ++C) {
		if(((colinfo[C]||{}).hidden)) continue;
		cols[C] = encode_col(C);
		val = dense ? sheet[R][C] : sheet[cols[C] + rr];
		switch(header) {
			case 1: hdr[C] = C - r.s.c; break;
			case 2: hdr[C] = cols[C]; break;
			case 3: hdr[C] = o.header[C - r.s.c]; break;
			default:
				if(val == null) val = {w: "__EMPTY", t: "s"};
				vv = v = format_cell(val, null, o);
				counter = header_cnt[v] || 0;
				if(!counter) header_cnt[v] = 1;
				else {
					do { vv = v + "_" + (counter++); } while(header_cnt[vv]); header_cnt[v] = counter;
					header_cnt[vv] = 1;
				}
				hdr[C] = vv;
		}
	}
	R = r.s.r + offset;
	stream._read = function() {
		while(R <= r.e.r) {
			if ((rowinfo[R-1]||{}).hidden) continue;
			var row = make_json_row(sheet, r, R, cols, header, hdr, dense, o);
			++R;
			if((row.isempty === false) || (header === 1 ? o.blankrows !== false : !!o.blankrows)) {
				stream.push(row.row);
				return;
			}
		}
		return stream.push(null);
	};
	return stream;
}

var __stream = {
	to_json: write_json_stream,
	to_html: write_html_stream,
	to_csv: write_csv_stream,
	set_readable: set_readable
};
