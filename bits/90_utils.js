function decode_row(rowstr) { return Number(unfix_row(rowstr)) - 1; }
function encode_row(row) { return "" + (row + 1); }
function fix_row(cstr) { return cstr.replace(/([A-Z]|^)([0-9]+)$/,"$1$$$2"); }
function unfix_row(cstr) { return cstr.replace(/\$([0-9]+)$/,"$1"); }

function decode_col(colstr) { var c = unfix_col(colstr), d = 0, i = 0; for(; i !== c.length; ++i) d = 26*d + c.charCodeAt(i) - 64; return d - 1; }
function encode_col(col) { var s=""; for(++col; col; col=Math.floor((col-1)/26)) s = _chr(((col-1)%26) + 65) + s; return s; }
function fix_col(cstr) { return cstr.replace(/^([A-Z])/,"$$$1"); }
function unfix_col(cstr) { return cstr.replace(/^\$([A-Z])/,"$1"); }

function split_cell(cstr) { return cstr.replace(/(\$?[A-Z]*)(\$?[0-9]*)/,"$1,$2").split(","); }
function decode_cell(cstr) { var splt = split_cell(cstr); return { c:decode_col(splt[0]), r:decode_row(splt[1]) }; }
function encode_cell(cell) { return encode_col(cell.c) + encode_row(cell.r); }
function fix_cell(cstr) { return fix_col(fix_row(cstr)); }
function unfix_cell(cstr) { return unfix_col(unfix_row(cstr)); }
function decode_range(range) { var x =range.split(":").map(decode_cell); return {s:x[0],e:x[x.length-1]}; }
function encode_range(cs,ce) {
	if(typeof ce === 'undefined' || typeof ce === 'number') return encode_range(cs.s, cs.e);
	if(typeof cs !== 'string') cs = encode_cell(cs); if(typeof ce !== 'string') ce = encode_cell(ce);
	return cs == ce ? cs : cs + ":" + ce;
}

function format_cell(cell, v) {
	if(!cell || !cell.t) return "";
	if(typeof cell.w !== 'undefined') return cell.w;
	if(typeof v === 'undefined') v = cell.v;
	if(typeof cell.z !== 'undefined') try { return (cell.w = SSF.format(cell.z, v)); } catch(e) { }
	if(!cell.XF) return v;
	try { return (cell.w = SSF.format(cell.XF.ifmt||0, v)); } catch(e) { return v; }
}

function sheet_to_json(sheet, opts){
	var val, row, range, header, offset = 1, r, hdr = {}, isempty, R, C, v;
	var out = [];
	opts = opts || {};
	if(!sheet || !sheet["!ref"]) return out;
	range = opts.range || sheet["!ref"];
	header = opts.header || "";
	switch(typeof range) {
		case 'string': r = decode_range(range); break;
		case 'number': r = decode_range(sheet["!ref"]); r.s.r = range; break;
		default: r = range;
	}
	if(header) offset = 0;
	for(R=r.s.r, C = r.s.c; C <= r.e.c; ++C) {
		val = sheet[encode_cell({c:C,r:R})];
		if(header === "A") hdr[C] = encode_col(C);
		else if(header === 1) hdr[C] = C;
		else if(Array.isArray(header)) hdr[C] = header[C - r.s.c];
		else if(!val) continue;
		else hdr[C] = format_cell(val);
	}

	for (R = r.s.r + offset; R <= r.e.r; ++R) {
		isempty = true;
		row = header === 1 ? [] : Object.create({ __rowNum__ : R });
		for (C = r.s.c; C <= r.e.c; ++C) {
			val = sheet[encode_cell({c: C,r: R})];
			if(!val || !val.t) continue;
			v = (val || {}).v;
			switch(val.t){
				case 'e': continue;
				case 's': case 'str': break;
				case 'b': case 'n': break;
				default: throw 'unrecognized type ' + val.t;
			}
			if(typeof v !== 'undefined') {
				row[hdr[C]] = opts.raw ? v||val.v : format_cell(val,v);
				isempty = false;
			}
		}
		if(!isempty) out.push(row);
	}
	return out;
}

function sheet_to_row_object_array(sheet, opts) { if(!opts) opts = {}; delete opts.range; return sheet_to_json(sheet, opts); }

function sheet_to_csv(sheet, opts) {
	var out = [], txt = "";
	opts = opts || {};
	if(!sheet || !sheet["!ref"]) return "";
	var r = decode_range(sheet["!ref"]);
	var fs = opts.FS||",", rs = opts.RS||"\n";

	for(var R = r.s.r; R <= r.e.r; ++R) {
		var row = [];
		for(var C = r.s.c; C <= r.e.c; ++C) {
			var val = sheet[encode_cell({c:C,r:R})];
			if(!val) { row.push(""); continue; }
			txt = String(format_cell(val));
			if(txt.indexOf(fs)!==-1 || txt.indexOf(rs)!==-1 || txt.indexOf('"')!==-1)
				txt = "\"" + txt.replace(/"/g, '""') + "\"";
			row.push(txt);
		}
		out.push(row.join(fs));
	}
	return out.join(rs) + (out.length ? rs : "");
}
var make_csv = sheet_to_csv;

function get_formulae(ws) {
	var cmds = [];
	for(var y in ws) if(y[0] !=='!' && ws.hasOwnProperty(y)) {
		var x = ws[y];
		var val = "";
		if(x.f) val = x.f;
		else if(typeof x.w !== 'undefined') val = "'" + x.w;
		else if(typeof x.v === 'undefined') continue;
		else val = x.v;
		cmds.push(y + "=" + val);
	}
	return cmds;
}

var utils = {
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
	make_json: sheet_to_json,
	get_formulae: get_formulae,
	format_cell: format_cell,
	sheet_to_json: sheet_to_json,
	sheet_to_row_object_array: sheet_to_row_object_array
};
