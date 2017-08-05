function sheet_to_json(sheet/*:Worksheet*/, opts/*:?Sheet2JSONOpts*/){
	if(sheet == null || sheet["!ref"] == null) return [];
	var val = {t:'n',v:0}, header = 0, offset = 1, hdr/*:Array<any>*/ = [], isempty = true, v=0, vv="";
	var r = {s:{r:0,c:0},e:{r:0,c:0}};
	var o = opts != null ? opts : {};
	var raw = o.raw;
	var defval = o.defval;
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
	var cols = new Array(r.e.c-r.s.c+1);
	var out = new Array(r.e.r-r.s.r-offset+1);
	var outi = 0, counter = 0;
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
				if(val == null) continue;
				vv = v = format_cell(val, null, o);
				counter = 0;
				for(CC = 0; CC < hdr.length; ++CC) if(hdr[CC] == vv) vv = v + "_" + (++counter);
				hdr[C] = vv;
		}
	}
	var row = (header === 1) ? [] : {};
	for (R = r.s.r + offset; R <= r.e.r; ++R) {
		rr = encode_row(R);
		isempty = true;
		if(header === 1) row = [];
		else {
			row = {};
			if(Object.defineProperty) try { Object.defineProperty(row, '__rowNum__', {value:R, enumerable:false}); } catch(e) { row.__rowNum__ = R; }
			else row.__rowNum__ = R;
		}
		if(!dense || sheet[R]) for (C = r.s.c; C <= r.e.c; ++C) {
			val = dense ? sheet[R][C] : sheet[cols[C] + rr];
			if(val === undefined || val.t === undefined) {
				if(defval === undefined) continue;
				if(hdr[C] != null) { row[hdr[C]] = defval; isempty = false; }
				continue;
			}
			v = val.v;
			switch(val.t){
				case 'z': if(v == null) break; continue;
				case 'e': continue;
				case 's': case 'd': case 'b': case 'n': break;
				default: throw new Error('unrecognized type ' + val.t);
			}
			if(hdr[C] != null) {
				if(v == null) {
					if(defval !== undefined) row[hdr[C]] = defval;
					else if(raw && v === null) row[hdr[C]] = null;
					else continue;
				} else {
					row[hdr[C]] = raw ? v : format_cell(val,v,o);
				}
				isempty = false;
			}
		}
		if((isempty === false) || (header === 1 ? o.blankrows !== false : !!o.blankrows)) out[outi++] = row;
	}
	out.length = outi;
	return out;
}

var qreg = /"/g;
function make_csv_row(sheet/*:Worksheet*/, r/*:Range*/, R/*:number*/, cols/*:Array<string>*/, fs/*:number*/, rs/*:number*/, FS/*:string*/, o/*:Sheet2CSVOpts*/)/*:?string*/ {
	var isempty = true;
	var row = [], txt = "", rr = encode_row(R);
	for(var C = r.s.c; C <= r.e.c; ++C) {
		if (!cols[C]) continue;
		var val = o.dense ? (sheet[R]||[])[C]: sheet[cols[C] + rr];
		if(val == null) txt = "";
		else if(val.v != null) {
			isempty = false;
			txt = ''+format_cell(val, null, o);
			for(var i = 0, cc = 0; i !== txt.length; ++i) if((cc = txt.charCodeAt(i)) === fs || cc === rs || cc === 34) {txt = "\"" + txt.replace(qreg, '""') + "\""; break; }
			if(txt == "ID") txt = '"ID"';
		} else if(val.f != null && !val.F) {
			isempty = false;
			txt = '=' + val.f; if(txt.indexOf(",") >= 0) txt = '"' + txt.replace(qreg, '""') + '"';
		} else txt = "";
		/* NOTE: Excel CSV does not support array formulae */
		row.push(txt);
	}
	if(o.blankrows === false && isempty) return null;
	return row.join(FS);
}

function sheet_to_csv(sheet/*:Worksheet*/, opts/*:?Sheet2CSVOpts*/)/*:string*/ {
	var out = [];
	var o = opts == null ? {} : opts;
	if(sheet == null || sheet["!ref"] == null) return "";
	var r = safe_decode_range(sheet["!ref"]);
	var FS = o.FS !== undefined ? o.FS : ",", fs = FS.charCodeAt(0);
	var RS = o.RS !== undefined ? o.RS : "\n", rs = RS.charCodeAt(0);
	var endregex = new RegExp((FS=="|" ? "\\|" : FS)+"+$");
	var row = "", cols = [];
	o.dense = Array.isArray(sheet);
	var colInfos = o.skipHidden && sheet["!cols"] || [];
	var rowInfos = o.skipHidden && sheet["!rows"] || [];
	for(var C = r.s.c; C <= r.e.c; ++C) if (!((colInfos[C]||{}).hidden)) cols[C] = encode_col(C);
	for(var R = r.s.r; R <= r.e.r; ++R) {
		if ((rowInfos[R]||{}).hidden) continue;
		row = make_csv_row(sheet, r, R, cols, fs, rs, FS, o);
		if(row == null) { continue; }
		if(o.strip) row = row.replace(endregex,"");
		out.push(row + RS);
	}
	delete o.dense;
	return out.join("");
}

function sheet_to_txt(sheet/*:Worksheet*/, opts/*:?Sheet2CSVOpts*/) {
	if(!opts) opts = {}; opts.FS = "\t"; opts.RS = "\n";
	var s = sheet_to_csv(sheet, opts);
	if(typeof cptable == 'undefined') return s;
	var o = cptable.utils.encode(1200, s);
	return "\xff\xfe" + o;
}

function sheet_to_formulae(sheet/*:Worksheet*/)/*:Array<string>*/ {
	var y = "", x, val="";
	if(sheet == null || sheet["!ref"] == null) return [];
	var r = safe_decode_range(sheet['!ref']), rr = "", cols = [], C;
	var cmds/*:Array<string>*/ = [];
	var dense = Array.isArray(sheet);
	for(C = r.s.c; C <= r.e.c; ++C) cols[C] = encode_col(C);
	for(var R = r.s.r; R <= r.e.r; ++R) {
		rr = encode_row(R);
		for(C = r.s.c; C <= r.e.c; ++C) {
			y = cols[C] + rr;
			x = dense ? (sheet[R]||[])[C] : sheet[y];
			val = "";
			if(x === undefined) continue;
			else if(x.F != null) {
				y = x.F;
				if(!x.f) continue;
				val = x.f;
				if(y.indexOf(":") == -1) y = y + ":" + y;
			}
			if(x.f != null) val = x.f;
			else if(x.t == 'z') continue;
			else if(x.t == 'n' && x.v != null) val = "" + x.v;
			else if(x.t == 'b') val = x.v ? "TRUE" : "FALSE";
			else if(x.w !== undefined) val = "'" + x.w;
			else if(x.v === undefined) continue;
			else if(x.t == 's') val = "'" + x.v;
			else val = ""+x.v;
			cmds[cmds.length] = y + "=" + val;
		}
	}
	return cmds;
}

function json_to_sheet(js/*:Array<any>*/, opts)/*:Worksheet*/ {
	var o = opts || {};
	var ws = ({}/*:any*/);
	var cell/*:Cell*/;
	var range/*:Range*/ = ({s: {c:0, r:0}, e: {c:0, r:js.length}}/*:any*/);
	var hdr = o.header || [], C = 0;

	for(var R = 0; R != js.length; ++R) {
		Object.keys(js[R]).filter(function(x) { return js[R].hasOwnProperty(x); }).forEach(function(k) {
			if((C=hdr.indexOf(k)) == -1) hdr[C=hdr.length] = k;
			var v = js[R][k];
			var t = 'z';
			var z = "";
			if(typeof v == 'number') t = 'n';
			else if(typeof v == 'boolean') t = 'b';
			else if(typeof v == 'string') t = 's';
			else if(v instanceof Date) {
				t = 'd';
				if(!o.cellDates) { t = 'n'; v = datenum(v); }
				z = o.dateNF || SSF._table[14];
			}
			ws[encode_cell({c:C,r:R+1})] = cell = ({t:t, v:v}/*:any*/);
			if(z) cell.z = z;
		});
	}
	range.e.c = hdr.length - 1;
	for(C = 0; C < hdr.length; ++C) ws[encode_col(C) + "1"] = {t:'s', v:hdr[C]};
	ws['!ref'] = encode_range(range);
	return ws;
}

var utils/*:any*/ = {
	encode_col: encode_col,
	encode_row: encode_row,
	encode_cell: encode_cell,
	encode_range: encode_range,
	decode_col: decode_col,
	decode_row: decode_row,
	split_cell: split_cell,
	decode_cell: decode_cell,
	decode_range: decode_range,
	format_cell: format_cell,
	get_formulae: sheet_to_formulae,
	make_csv: sheet_to_csv,
	make_json: sheet_to_json,
	make_formulae: sheet_to_formulae,
	aoa_to_sheet: aoa_to_sheet,
	json_to_sheet: json_to_sheet,
	table_to_sheet: parse_dom_table,
	table_to_book: table_to_book,
	sheet_to_csv: sheet_to_csv,
	sheet_to_json: sheet_to_json,
	sheet_to_html: HTML_.from_sheet,
	sheet_to_formulae: sheet_to_formulae,
	sheet_to_row_object_array: sheet_to_json
};

