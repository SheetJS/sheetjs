/* from js-harb (C) 2014-present  SheetJS */
var DBF = (function() {
var dbf_codepage_map = {
	/* Code Pages Supported by Visual FoxPro */
	/*::[*/0x01/*::]*/:   437,           /*::[*/0x02/*::]*/:   850,
	/*::[*/0x03/*::]*/:  1252,           /*::[*/0x04/*::]*/: 10000,
	/*::[*/0x64/*::]*/:   852,           /*::[*/0x65/*::]*/:   866,
	/*::[*/0x66/*::]*/:   865,           /*::[*/0x67/*::]*/:   861,
	/*::[*/0x68/*::]*/:   895,           /*::[*/0x69/*::]*/:   620,
	/*::[*/0x6A/*::]*/:   737,           /*::[*/0x6B/*::]*/:   857,
	/*::[*/0x78/*::]*/:   950,           /*::[*/0x79/*::]*/:   949,
	/*::[*/0x7A/*::]*/:   936,           /*::[*/0x7B/*::]*/:   932,
	/*::[*/0x7C/*::]*/:   874,           /*::[*/0x7D/*::]*/:  1255,
	/*::[*/0x7E/*::]*/:  1256,           /*::[*/0x96/*::]*/: 10007,
	/*::[*/0x97/*::]*/: 10029,           /*::[*/0x98/*::]*/: 10006,
	/*::[*/0xC8/*::]*/:  1250,           /*::[*/0xC9/*::]*/:  1251,
	/*::[*/0xCA/*::]*/:  1254,           /*::[*/0xCB/*::]*/:  1253,

	/* shapefile DBF extension */
	/*::[*/0x00/*::]*/: 20127,           /*::[*/0x08/*::]*/:   865,
	/*::[*/0x09/*::]*/:   437,           /*::[*/0x0A/*::]*/:   850,
	/*::[*/0x0B/*::]*/:   437,           /*::[*/0x0D/*::]*/:   437,
	/*::[*/0x0E/*::]*/:   850,           /*::[*/0x0F/*::]*/:   437,
	/*::[*/0x10/*::]*/:   850,           /*::[*/0x11/*::]*/:   437,
	/*::[*/0x12/*::]*/:   850,           /*::[*/0x13/*::]*/:   932,
	/*::[*/0x14/*::]*/:   850,           /*::[*/0x15/*::]*/:   437,
	/*::[*/0x16/*::]*/:   850,           /*::[*/0x17/*::]*/:   865,
	/*::[*/0x18/*::]*/:   437,           /*::[*/0x19/*::]*/:   437,
	/*::[*/0x1A/*::]*/:   850,           /*::[*/0x1B/*::]*/:   437,
	/*::[*/0x1C/*::]*/:   863,           /*::[*/0x1D/*::]*/:   850,
	/*::[*/0x1F/*::]*/:   852,           /*::[*/0x22/*::]*/:   852,
	/*::[*/0x23/*::]*/:   852,           /*::[*/0x24/*::]*/:   860,
	/*::[*/0x25/*::]*/:   850,           /*::[*/0x26/*::]*/:   866,
	/*::[*/0x37/*::]*/:   850,           /*::[*/0x40/*::]*/:   852,
	/*::[*/0x4D/*::]*/:   936,           /*::[*/0x4E/*::]*/:   949,
	/*::[*/0x4F/*::]*/:   950,           /*::[*/0x50/*::]*/:   874,
	/*::[*/0x57/*::]*/:  1252,           /*::[*/0x58/*::]*/:  1252,
	/*::[*/0x59/*::]*/:  1252,

	/*::[*/0xFF/*::]*/: 16969
};

/* TODO: find an actual specification */
function dbf_to_aoa(buf, opts)/*:AOA*/ {
	var out/*:AOA*/ = [];
	/* TODO: browser based */
	var d/*:Block*/ = (new_raw_buf(1)/*:any*/);
	switch(opts.type) {
		case 'base64': d = s2a(Base64.decode(buf)); break;
		case 'binary': d = s2a(buf); break;
		case 'buffer':
		case 'array': d = buf; break;
	}
	prep_blob(d, 0);
	/* header */
	var ft = d.read_shift(1);
	var memo = false;
	var vfp = false;
	switch(ft) {
		case 0x02: case 0x03: break;
		case 0x30: vfp = true; memo = true; break;
		case 0x31: vfp = true; break;
		case 0x83: memo = true; break;
		case 0x8B: memo = true; break;
		case 0xF5: memo = true; break;
		default: throw new Error("DBF Unsupported Version: " + ft.toString(16));
	}
	var filedate = new Date(), nrow = 0, fpos = 0;
	if(ft == 0x02) nrow = d.read_shift(2);
	filedate = new Date(d.read_shift(1) + 1900, d.read_shift(1) - 1, d.read_shift(1));
	if(ft != 0x02) nrow = d.read_shift(4);
	if(ft != 0x02) fpos = d.read_shift(2);
	var rlen = d.read_shift(2);

	var flags = 0, current_cp = 1252;
	if(ft != 0x02) {
	d.l+=16;
	flags = d.read_shift(1);
	//if(memo && ((flags & 0x02) === 0)) throw new Error("DBF Flags " + flags.toString(16) + " ft " + ft.toString(16));

	/* codepage present in FoxPro */
	if(d[d.l] !== 0) current_cp = dbf_codepage_map[d[d.l]];
	d.l+=1;

	d.l+=2;
	}
	var fields = [], field = {};
	var hend = fpos - 10 - (vfp ? 264 : 0);
	while(ft == 0x02 ? d.l < d.length && d[d.l] != 0x0d: d.l < hend) {
		field = {};
		field.name = cptable.utils.decode(current_cp, d.slice(d.l, d.l+10)).replace(/[\u0000\r\n].*$/g,"");
		d.l += 11;
		field.type = String.fromCharCode(d.read_shift(1));
		if(ft != 0x02) field.offset = d.read_shift(4);
		field.len = d.read_shift(1);
		if(ft == 0x02) field.offset = d.read_shift(2);
		field.dec = d.read_shift(1);
		if(field.name.length) fields.push(field);
		if(ft != 0x02) d.l += 14;
		switch(field.type) {
			// case 'B': break; // Binary
			case 'C': break; // character
			case 'D': break; // date
			case 'F': break; // floating point
			// case 'G': break; // General
			case 'I': break; // long
			case 'L': break; // boolean
			case 'M': break; // memo
			case 'N': break; // number
			// case 'O': break; // double
			// case 'P': break; // Picture
			case 'T': break; // datetime
			case 'Y': break; // currency
			case '0': break; // null ?
			case '+': break; // autoincrement
			case '@': break; // timestamp
			default: throw new Error('Unknown Field Type: ' + field.type);
		}
	}
	if(d[d.l] !== 0x0D) d.l = fpos-1;
	else if(ft == 0x02) d.l = 0x209;
	if(ft != 0x02) {
		if(d.read_shift(1) !== 0x0D) throw new Error("DBF Terminator not found " + d.l + " " + d[d.l]);
		d.l = fpos;
	}
	/* data */
	var R = 0, C = 0;
	out[0] = [];
	for(C = 0; C != fields.length; ++C) out[0][C] = fields[C].name;
	while(nrow-- > 0) {
		if(d[d.l] === 0x2A) { d.l+=rlen; continue; }
		++d.l;
		out[++R] = []; C = 0;
		for(C = 0; C != fields.length; ++C) {
			var dd = d.slice(d.l, d.l+fields[C].len); d.l+=fields[C].len;
			prep_blob(dd, 0);
			var s = cptable.utils.decode(current_cp, dd);
			switch(fields[C].type) {
				case 'C':
					out[R][C] = cptable.utils.decode(current_cp, dd);
					out[R][C] = out[R][C].trim();
					break;
				case 'D':
					if(s.length === 8) out[R][C] = new Date(+s.substr(0,4), +s.substr(4,2)-1, +s.substr(6,2));
					else out[R][C] = s;
					break;
				case 'F': out[R][C] = parseFloat(s.trim()); break;
				case 'I': out[R][C] = dd.read_shift(4, 'i'); break;
				case 'L': switch(s.toUpperCase()) {
					case 'Y': case 'T': out[R][C] = true; break;
					case 'N': case 'F': out[R][C] = false; break;
					case ' ': case '?': out[R][C] = false; break; /* NOTE: technically unitialized */
					default: throw new Error("DBF Unrecognized L:|" + s + "|");
					} break;
				case 'M': /* TODO: handle memo files */
					if(!memo) throw new Error("DBF Unexpected MEMO for type " + ft.toString(16));
					out[R][C] = "##MEMO##" + dd.read_shift(4);
					break;
				case 'N': out[R][C] = +s.replace(/\u0000/g,"").trim(); break;
				case 'T':
					var day = dd.read_shift(4), ms = dd.read_shift(4);
					throw new Error(day + " | " + ms);
					//out[R][C] = new Date(); // TODO
					//break;
				case 'Y': out[R][C] = dd.read(4,'i')/1e4; break;
				case '0':
					if(fields[C].name === '_NullFlags') break;
					/* falls through */
				default: throw new Error("DBF Unsupported data type " + fields[C].type);
			}
		}
	}
	if(ft != 0x02) if(d.l < d.length && d[d.l++] != 0x1A) throw new Error("DBF EOF Marker missing " + (d.l-1) + " of " + d.length + " " + d[d.l-1].toString(16));
	return out;
}

function dbf_to_sheet(buf, opts)/*:Worksheet*/ {
	var o = opts || {};
	if(!o.dateNF) o.dateNF = "yyyymmdd";
	return aoa_to_sheet(dbf_to_aoa(buf, o), o);
}

function dbf_to_workbook(buf, opts)/*:Workbook*/ {
	try { return sheet_to_workbook(dbf_to_sheet(buf, opts), opts); }
	catch(e) { if(opts && opts.WTF) throw e; }
	return ({SheetNames:[],Sheets:{}});
}
	return {
		to_workbook: dbf_to_workbook,
		to_sheet: dbf_to_sheet
	};
})();

var SYLK = (function() {
	/* TODO: find an actual specification */
	function sylk_to_aoa(d/*:RawData*/, opts)/*:[AOA, Worksheet]*/ {
		switch(opts.type) {
			case 'base64': return sylk_to_aoa_str(Base64.decode(d), opts);
			case 'binary': return sylk_to_aoa_str(d, opts);
			case 'buffer': return sylk_to_aoa_str(d.toString('binary'), opts);
			case 'array': return sylk_to_aoa_str(cc2str(d), opts);
		}
		throw new Error("Unrecognized type " + opts.type);
	}
	function sylk_to_aoa_str(str/*:string*/, opts)/*:[AOA, Worksheet]*/ {
		var records = str.split(/[\n\r]+/), R = -1, C = -1, ri = 0, rj = 0, arr = [];
		var formats = [];
		var next_cell_format = null;
		var sht = {}, rowinfo = [], colinfo = [], cw = [];
		var Mval = 0, j;
		for (; ri !== records.length; ++ri) {
			Mval = 0;
			var rstr=records[ri].trim();
			var record=rstr.replace(/;;/g, "\u0001").split(";").map(function(x) { return x.replace(/\u0001/g, ";"); });
			var RT=record[0], val;
			if(rstr.length > 0) switch(RT) {
			case 'ID': break; /* header */
			case 'E': break; /* EOF */
			case 'B': break; /* dimensions */
			case 'O': break; /* options? */
			case 'P':
				if(record[1].charAt(0) == 'P')
					formats.push(rstr.substr(3).replace(/;;/g, ";"));
				break;
			case 'C':
			for(rj=1; rj<record.length; ++rj) switch(record[rj].charAt(0)) {
				case 'X': C = parseInt(record[rj].substr(1))-1; break;
				case 'Y':
					R = parseInt(record[rj].substr(1))-1; C = 0;
					for(j = arr.length; j <= R; ++j) arr[j] = [];
					break;
				case 'K':
					val = record[rj].substr(1);
					if(val.charAt(0) === '"') val = val.substr(1,val.length - 2);
					else if(val === 'TRUE') val = true;
					else if(val === 'FALSE') val = false;
					else if(+val === +val) {
						val = +val;
						if(next_cell_format !== null && SSF.is_date(next_cell_format)) val = numdate(val);
					} else if(!isNaN(fuzzydate(val).getDate())) {
						val = parseDate(val);
					}
					arr[R][C] = val;
					next_cell_format = null;
					break;
				case 'E':
					var formula = rc_to_a1(record[rj].substr(1), {r:R,c:C});
					arr[R][C] = [arr[R][C], formula];
					break;
				default: if(opts && opts.WTF) throw new Error("SYLK bad record " + rstr);
			} break;
			case 'F':
			var F_seen = 0;
			for(rj=1; rj<record.length; ++rj) switch(record[rj].charAt(0)) {
				case 'X': C = parseInt(record[rj].substr(1))-1; ++F_seen; break;
				case 'Y':
					R = parseInt(record[rj].substr(1))-1; /*C = 0;*/
					for(j = arr.length; j <= R; ++j) arr[j] = [];
					break;
				case 'M': Mval = parseInt(record[rj].substr(1)) / 20; break;
				case 'F': break; /* ??? */
				case 'P':
					next_cell_format = formats[parseInt(record[rj].substr(1))];
					break;
				case 'S': break; /* cell style */
				case 'D': break; /* column */
				case 'N': break; /* font */
				case 'W':
					cw = record[rj].substr(1).split(" ");
					for(j = parseInt(cw[0], 10); j <= parseInt(cw[1], 10); ++j) {
						Mval = parseInt(cw[2], 10);
						colinfo[j-1] = Mval == 0 ? {hidden:true}: {wch:Mval}; process_col(colinfo[j-1]);
					} break;
				case 'C': /* default column format */
					C = parseInt(record[rj].substr(1))-1;
					if(!colinfo[C]) colinfo[C] = {};
					break;
				case 'R': /* row properties */
					R = parseInt(record[rj].substr(1))-1;
					if(!rowinfo[R]) rowinfo[R] = {};
					if(Mval > 0) { rowinfo[R].hpt = Mval; rowinfo[R].hpx = pt2px(Mval); }
					else if(Mval == 0) rowinfo[R].hidden = true;
					break;
				default: if(opts && opts.WTF) throw new Error("SYLK bad record " + rstr);
			}
			if(F_seen < 1) next_cell_format = null; break;
			default: if(opts && opts.WTF) throw new Error("SYLK bad record " + rstr);
			}
		}
		if(rowinfo.length > 0) sht['!rows'] = rowinfo;
		if(colinfo.length > 0) sht['!cols'] = colinfo;
		return [arr, sht];
	}

	function sylk_to_sheet(d/*:RawData*/, opts)/*:Worksheet*/ {
		var aoasht = sylk_to_aoa(d, opts);
		var aoa = aoasht[0], ws = aoasht[1];
		var o = aoa_to_sheet(aoa, opts);
		keys(ws).forEach(function(k) { o[k] = ws[k]; });
		return o;
	}

	function sylk_to_workbook(d/*:RawData*/, opts)/*:Workbook*/ { return sheet_to_workbook(sylk_to_sheet(d, opts), opts); }

	function write_ws_cell_sylk(cell/*:Cell*/, ws/*:Worksheet*/, R/*:number*/, C/*:number*/, opts)/*:string*/ {
		var o = "C;Y" + (R+1) + ";X" + (C+1) + ";K";
		switch(cell.t) {
			case 'n':
				o += (cell.v||0);
				if(cell.f && !cell.F) o += ";E" + a1_to_rc(cell.f, {r:R, c:C}); break;
			case 'b': o += cell.v ? "TRUE" : "FALSE"; break;
			case 'e': o += cell.w || cell.v; break;
			case 'd': o += '"' + (cell.w || cell.v) + '"'; break;
			case 's': o += '"' + cell.v.replace(/"/g,"") + '"'; break;
		}
		return o;
	}

	function write_ws_cols_sylk(out, cols) {
		cols.forEach(function(col, i) {
			var rec = "F;W" + (i+1) + " " + (i+1) + " ";
			if(col.hidden) rec += "0";
			else {
				if(typeof col.width == 'number') col.wpx = width2px(col.width);
				if(typeof col.wpx == 'number') col.wch = px2char(col.wpx);
				if(typeof col.wch == 'number') rec += Math.round(col.wch);
			}
			if(rec.charAt(rec.length - 1) != " ") out.push(rec);
		});
	}

	function write_ws_rows_sylk(out/*:Array<string>*/, rows/*:Array<RowInfo>*/) {
		rows.forEach(function(row, i) {
			var rec = "F;";
			if(row.hidden) rec += "M0;";
			else if(row.hpt) rec += "M" + 20 * row.hpt + ";";
			else if(row.hpx) rec += "M" + 20 * px2pt(row.hpx) + ";";
			if(rec.length > 2) out.push(rec + "R" + (i+1));
		});
	}

	function sheet_to_sylk(ws/*:Worksheet*/, opts/*:?any*/)/*:string*/ {
		var preamble/*:Array<string>*/ = ["ID;PWXL;N;E"], o/*:Array<string>*/ = [];
		var r = decode_range(ws['!ref']), cell/*:Cell*/;
		var dense = Array.isArray(ws);
		var RS = "\r\n";

		preamble.push("P;PGeneral");
		preamble.push("F;P0;DG0G8;M255");
		if(ws['!cols']) write_ws_cols_sylk(preamble, ws['!cols']);
		if(ws['!rows']) write_ws_rows_sylk(preamble, ws['!rows']);

		preamble.push("B;Y" + (r.e.r - r.s.r + 1) + ";X" + (r.e.c - r.s.c + 1) + ";D" + [r.s.c,r.s.r,r.e.c,r.e.r].join(" "));
		for(var R = r.s.r; R <= r.e.r; ++R) {
			for(var C = r.s.c; C <= r.e.c; ++C) {
				var coord = encode_cell({r:R,c:C});
				cell = dense ? (ws[R]||[])[C]: ws[coord];
				if(!cell || cell.v == null && (!cell.f || cell.F)) continue;
				o.push(write_ws_cell_sylk(cell, ws, R, C, opts));
			}
		}
		return preamble.join(RS) + RS + o.join(RS) + RS + "E" + RS;
	}

	return {
		to_workbook: sylk_to_workbook,
		to_sheet: sylk_to_sheet,
		from_sheet: sheet_to_sylk
	};
})();

var DIF = (function() {
	function dif_to_aoa(d/*:RawData*/, opts)/*:AOA*/ {
		switch(opts.type) {
			case 'base64': return dif_to_aoa_str(Base64.decode(d), opts);
			case 'binary': return dif_to_aoa_str(d, opts);
			case 'buffer': return dif_to_aoa_str(d.toString('binary'), opts);
			case 'array': return dif_to_aoa_str(cc2str(d), opts);
		}
		throw new Error("Unrecognized type " + opts.type);
	}
	function dif_to_aoa_str(str/*:string*/, opts)/*:AOA*/ {
		var records = str.split('\n'), R = -1, C = -1, ri = 0, arr = [];
		for (; ri !== records.length; ++ri) {
			if (records[ri].trim() === 'BOT') { arr[++R] = []; C = 0; continue; }
			if (R < 0) continue;
			var metadata = records[ri].trim().split(",");
			var type = metadata[0], value = metadata[1];
			++ri;
			var data = records[ri].trim();
			switch (+type) {
				case -1:
					if (data === 'BOT') { arr[++R] = []; C = 0; continue; }
					else if (data !== 'EOD') throw new Error("Unrecognized DIF special command " + data);
					break;
				case 0:
					if(data === 'TRUE') arr[R][C] = true;
					else if(data === 'FALSE') arr[R][C] = false;
					else if(+value == +value) arr[R][C] = +value;
					else if(!isNaN(fuzzydate(value).getDate())) arr[R][C] = parseDate(value);
					else arr[R][C] = value;
					++C; break;
				case 1:
					data = data.substr(1,data.length-2);
					arr[R][C++] = data !== '' ? data : null;
					break;
			}
			if (data === 'EOD') break;
		}
		return arr;
	}

	function dif_to_sheet(str/*:string*/, opts)/*:Worksheet*/ { return aoa_to_sheet(dif_to_aoa(str, opts), opts); }
	function dif_to_workbook(str/*:string*/, opts)/*:Workbook*/ { return sheet_to_workbook(dif_to_sheet(str, opts), opts); }

	var sheet_to_dif = (function() {
		var push_field = function pf(o/*:Array<string>*/, topic/*:string*/, v/*:number*/, n/*:number*/, s/*:string*/) {
			o.push(topic);
			o.push(v + "," + n);
			o.push('"' + s.replace(/"/g,'""') + '"');
		};
		var push_value = function po(o/*:Array<string>*/, type/*:number*/, v/*:any*/, s/*:string*/) {
			o.push(type + "," + v);
			o.push(type == 1 ? '"' + s.replace(/"/g,'""') + '"' : s);
		};
		return function sheet_to_dif(ws/*:Worksheet*/, opts/*:?any*/)/*:string*/ {
			var o/*:Array<string>*/ = [];
			var r = decode_range(ws['!ref']), cell/*:Cell*/;
			var dense = Array.isArray(ws);
			push_field(o, "TABLE", 0, 1, "sheetjs");
			push_field(o, "VECTORS", 0, r.e.r - r.s.r + 1,"");
			push_field(o, "TUPLES", 0, r.e.c - r.s.c + 1,"");
			push_field(o, "DATA", 0, 0,"");
			for(var R = r.s.r; R <= r.e.r; ++R) {
				push_value(o, -1, 0, "BOT");
				for(var C = r.s.c; C <= r.e.c; ++C) {
					var coord = encode_cell({r:R,c:C});
					cell = dense ? (ws[R]||[])[C] : ws[coord];
					if(!cell) { push_value(o, 1, 0, ""); continue;}
					switch(cell.t) {
						case 'n':
							var val = DIF_XL ? cell.w : cell.v;
							if(!val && cell.v != null) val = cell.v;
							if(val == null) {
								if(DIF_XL && cell.f && !cell.F) push_value(o, 1, 0, "=" + cell.f);
								else push_value(o, 1, 0, "");
							}
							else push_value(o, 0, val, "V");
							break;
						case 'b':
							push_value(o, 0, cell.v ? 1 : 0, cell.v ? "TRUE" : "FALSE");
							break;
						case 's':
							push_value(o, 1, 0, (!DIF_XL || isNaN(cell.v)) ? cell.v : '="' + cell.v + '"');
							break;
						case 'd':
							if(!cell.w) cell.w = SSF.format(cell.z || SSF._table[14], datenum(parseDate(cell.v)));
							if(DIF_XL) push_value(o, 0, cell.w, "V");
							else push_value(o, 1, 0, cell.w);
							break;
						default: push_value(o, 1, 0, "");
					}
				}
			}
			push_value(o, -1, 0, "EOD");
			var RS = "\r\n";
			var oo = o.join(RS);
			//while((oo.length & 0x7F) != 0) oo += "\0";
			return oo;
		};
	})();
	return {
		to_workbook: dif_to_workbook,
		to_sheet: dif_to_sheet,
		from_sheet: sheet_to_dif
	};
})();

var PRN = (function() {
	function set_text_arr(data/*:string*/, arr/*:AOA*/, R/*:number*/, C/*:number*/) {
		if(data === 'TRUE') arr[R][C] = true;
		else if(data === 'FALSE') arr[R][C] = false;
		else if(data === ""){/* empty */}
		else if(+data == +data) arr[R][C] = +data;
		else if(!isNaN(fuzzydate(data).getDate())) arr[R][C] = parseDate(data);
		else arr[R][C] = data;
	}

	function prn_to_aoa_str(f/*:string*/, opts)/*:AOA*/ {
		var arr/*:AOA*/ = ([]/*:any*/);
		if(!f || f.length === 0) return arr;
		var lines = f.split(/[\r\n]/);
		var L = lines.length - 1;
		while(L >= 0 && lines[L].length === 0) --L;
		var start = 10, idx = 0;
		var R = 0;
		for(; R <= L; ++R) {
			idx = lines[R].indexOf(" ");
			if(idx == -1) idx = lines[R].length; else idx++;
			start = Math.max(start, idx);
		}
		for(R = 0; R <= L; ++R) {
			arr[R] = [];
			/* TODO: confirm that widths are always 10 */
			var C = 0;
			set_text_arr(lines[R].slice(0, start).trim(), arr, R, C);
			for(C = 1; C <= (lines[R].length - start)/10 + 1; ++C)
				set_text_arr(lines[R].slice(start+(C-1)*10,start+C*10).trim(),arr,R,C);
		}
		return arr;
	}

	function dsv_to_sheet_str(str/*:string*/, opts)/*:Worksheet*/ {
		var o = opts || {};
		var sep = "";
		if(DENSE != null && o.dense == null) o.dense = DENSE;
		var ws/*:Worksheet*/ = o.dense ? ([]/*:any*/) : ({}/*:any*/);
		var range/*:Range*/ = ({s: {c:0, r:0}, e: {c:0, r:0}}/*:any*/);

		/* known sep */
		if(str.substr(0,4) == "sep=" && str.charCodeAt(5) == 10) { sep = str.charAt(4); str = str.substr(6); }
		else if(str.substr(0,1024).indexOf("\t") == -1) sep = ","; else sep = "\t";
		var R = 0, C = 0, v = 0;
		var start = 0, end = 0, sepcc = sep.charCodeAt(0), instr = false, cc=0;
		str = str.replace(/\r\n/mg, "\n");
		var _re/*:?RegExp*/ = o.dateNF != null ? dateNF_regex(o.dateNF) : null;
		function finish_cell() {
			var s = str.slice(start, end);
			var cell = ({}/*:any*/);
			if(o.raw) { cell.t = 's'; cell.v = s; }
			else if(s.charCodeAt(0) == 0x3D) { cell.t = 'n'; cell.f = s.substr(1); }
			else if(s == "TRUE") { cell.t = 'b'; cell.v = true; }
			else if(s == "FALSE") { cell.t = 'b'; cell.v = false; }
			else if(!isNaN(v = +s)) { cell.t = 'n'; cell.w = s; cell.v = v; }
			else if(!isNaN(fuzzydate(s).getDate()) || _re && s.match(_re)) {
				cell.z = o.dateNF || SSF._table[14];
				var k = 0;
				if(_re && s.match(_re)){ s=dateNF_fix(s, o.dateNF, (s.match(_re)||[])); k=1; }
				if(o.cellDates) { cell.t = 'd'; cell.v = parseDate(s, k); }
				else { cell.t = 'n'; cell.v = datenum(parseDate(s, k)); }
				cell.w = SSF.format(cell.z, cell.v instanceof Date ? datenum(cell.v):cell.v);
			} else {
				cell.t = 's';
				if(s.charAt(0) == '"' && s.charAt(s.length - 1) == '"') s = s.slice(1,-1).replace(/""/g,'"');
				cell.v = s;
			}
			if(o.dense) { if(!ws[R]) ws[R] = []; ws[R][C] = cell; }
			else ws[encode_cell({c:C,r:R})] = cell;
			start = end+1;
			if(range.e.c < C) range.e.c = C;
			if(range.e.r < R) range.e.r = R;
			if(cc == sepcc) ++C; else { C = 0; ++R; }
		}
		for(;end < str.length;++end) switch((cc=str.charCodeAt(end))) {
			case 0x22: instr = !instr; break;
			case sepcc: case 0x0a: case 0x0d: if(!instr) finish_cell(); break;
			default: break;
		}
		if(end - start > 0) finish_cell();

		ws['!ref'] = encode_range(range);
		return ws;
	}

	function prn_to_sheet_str(str/*:string*/, opts)/*:Worksheet*/ {
		if(str.substr(0,4) == "sep=") return dsv_to_sheet_str(str, opts);
		if(str.indexOf("\t") >= 0 || str.indexOf(",") >= 0) return dsv_to_sheet_str(str, opts);
		return aoa_to_sheet(prn_to_aoa_str(str, opts), opts);
	}

	function prn_to_sheet(d/*:RawData*/, opts)/*:Worksheet*/ {
		var str = "", bytes = firstbyte(d, opts);
		switch(opts.type) {
			case 'base64': str = Base64.decode(d); break;
			case 'binary': str = d; break;
			case 'buffer': str = d.toString('binary'); break;
			case 'array': str = cc2str(d); break;
			default: throw new Error("Unrecognized type " + opts.type);
		}
		if(bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) str = utf8read(str);
		return prn_to_sheet_str(str, opts);
	}

	function prn_to_workbook(d/*:RawData*/, opts)/*:Workbook*/ { return sheet_to_workbook(prn_to_sheet(d, opts), opts); }

	function sheet_to_prn(ws/*:Worksheet*/, opts/*:?any*/)/*:string*/ {
		var o/*:Array<string>*/ = [];
		var r = decode_range(ws['!ref']), cell/*:Cell*/;
		var dense = Array.isArray(ws);
		for(var R = r.s.r; R <= r.e.r; ++R) {
			var oo = [];
			for(var C = r.s.c; C <= r.e.c; ++C) {
				var coord = encode_cell({r:R,c:C});
				cell = dense ? (ws[R]||[])[C] : ws[coord];
				if(!cell || cell.v == null) { oo.push("          "); continue; }
				var w = (cell.w || (format_cell(cell), cell.w) || "").substr(0,10);
				while(w.length < 10) w += " ";
				oo.push(w + (C == 0 ? " " : ""));
			}
			o.push(oo.join(""));
		}
		return o.join("\n");
	}

	return {
		to_workbook: prn_to_workbook,
		to_sheet: prn_to_sheet,
		from_sheet: sheet_to_prn
	};
})();

/* Excel defaults to SYLK but warns if data is not valid */
function read_wb_ID(d, opts) {
	var o = opts || {}, OLD_WTF = !!o.WTF; o.WTF = true;
	try {
		var out = SYLK.to_workbook(d, o);
		o.WTF = OLD_WTF;
		return out;
	} catch(e) {
		o.WTF = OLD_WTF;
		if(!e.message.match(/SYLK bad record ID/) && OLD_WTF) throw e;
		return PRN.to_workbook(d, opts);
	}
}
