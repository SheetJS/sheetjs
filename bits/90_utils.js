var _chr = function(c) { return String.fromCharCode(c); };

function encode_col(col) { var s=""; for(++col; col; col=Math.floor((col-1)/26)) s = _chr(((col-1)%26) + 65) + s; return s; }
function encode_row(row) { return "" + (row + 1); }
function encode_cell(cell) { return encode_col(cell.c) + encode_row(cell.r); }

function decode_col(c) { var d = 0, i = 0; for(; i !== c.length; ++i) d = 26*d + c.charCodeAt(i) - 64; return d - 1; }
function decode_row(rowstr) { return Number(rowstr) - 1; }
function split_cell(cstr) { return cstr.replace(/(\$?[A-Z]*)(\$?[0-9]*)/,"$1,$2").split(","); }
function decode_cell(cstr) { var splt = split_cell(cstr); return { c:decode_col(splt[0]), r:decode_row(splt[1]) }; }
function decode_range(range) { var x =range.split(":").map(decode_cell); return {s:x[0],e:x[x.length-1]}; }
function encode_range(range) { return encode_cell(range.s) + ":" + encode_cell(range.e); }

function sheet_to_row_object_array(sheet, opts){
	var val, row, r, hdr = {}, isempty, R, C, v;
	var out = [];
	opts = opts || {};
	if(!sheet || !sheet["!ref"]) return out;
	r = XLSX.utils.decode_range(sheet["!ref"]);
	for(R=r.s.r, C = r.s.c; C <= r.e.c; ++C) {
		val = sheet[encode_cell({c:C,r:R})];
		if(!val) continue;
		if(val.w) hdr[C] = val.w;
		else switch(val.t) {
			case 's': case 'str': hdr[C] = val.v; break;
			case 'n': hdr[C] = val.v; break;
		}
	}

	for (R = r.s.r + 1; R <= r.e.r; ++R) {
		isempty = true;
		/* row index available as __rowNum__ */
		row = Object.create({ __rowNum__ : R });
		for (C = r.s.c; C <= r.e.c; ++C) {
			val = sheet[encode_cell({c: C,r: R})];
			if(!val || !val.t) continue;
			if(typeof val.w !== 'undefined' && !opts.raw) { row[hdr[C]] = val.w; isempty = false; }
			else switch(val.t){
				case 's': case 'str': case 'b': case 'n':
					if(val.v !== undefined) {
						row[hdr[C]] = val.v;
						isempty = false;
					}
					break;
				case 'e': break; /* throw */
				default: throw 'unrecognized type ' + val.t;
			}
		}
		if(!isempty) out.push(row);
	}
	return out;
}

function sheet_to_csv(sheet, opts) {
	var stringify = function stringify(val) {
		if(!val.t) return "";
		if(typeof val.w !== 'undefined') return '"' + val.w.replace(/"/,'""') + '"';
		switch(val.t){
			case 'n': return String(val.v);
			case 's': case 'str':
				if(typeof val.v === 'undefined') return "";
				return '"' + val.v.replace(/"/,'""') + '"';
			case 'b': return val.v ? "TRUE" : "FALSE";
			case 'e': return val.v; /* throw out value in case of error */
			default: throw 'unrecognized type ' + val.t;
		}
	};
	var out = "", txt = "";
	opts = opts || {};
	if(!sheet || !sheet["!ref"]) return out;
	var r = XLSX.utils.decode_range(sheet["!ref"]);
	for(var R = r.s.r; R <= r.e.r; ++R) {
		var row = [];
		for(var C = r.s.c; C <= r.e.c; ++C) {
			var val = sheet[XLSX.utils.encode_cell({c:C,r:R})];
			if(!val) { row.push(""); continue; }
			txt = stringify(val);
			row.push(String(txt).replace(/\\r\\n/g,"\n").replace(/\\t/g,"\t").replace(/\\\\/g,"\\").replace("\\\"","\"\""));
		}
		out += row.join(opts.FS||",") + (opts.RS||"\n");
	}
	return out;
}
var make_csv = sheet_to_csv;

function get_formulae(ws) {
	var cmds = [];
	for(var y in ws) if(y[0] !=='!' && ws.hasOwnProperty(y)) {
		var x = ws[y];
		var val = "";
		if(x.f) val = x.f;
		else if(typeof x.v === 'undefined') continue;
		else if(typeof x.v === 'number') val = x.v;
		else val = x.v;
		cmds.push(y + "=" + val);
	}
	return cmds;
}

XLSX.utils = {
	encode_col: encode_col,
	encode_row: encode_row,
	encode_cell: encode_cell,
	encode_range: encode_range,
	decode_col: decode_col,
	decode_row: decode_row,
	split_cell: split_cell,
	decode_cell: decode_cell,
	decode_range: decode_range,
	sheet_to_csv: sheet_to_csv,
	make_csv: sheet_to_csv,
	get_formulae: get_formulae,
	sheet_to_row_object_array: sheet_to_row_object_array
};
