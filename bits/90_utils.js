/*::
type MJRObject = {
	row: any;
	isempty: boolean;
};
*/
function make_json_row(sheet/*:Worksheet*/, r/*:Range*/, R/*:number*/, cols/*:Array<string>*/, header/*:number*/, hdr/*:Array<any>*/, dense/*:boolean*/, o/*:Sheet2JSONOpts*/)/*:MJRObject*/ {
	var rr = encode_row(R);
	var defval = o.defval, raw = o.raw || !Object.prototype.hasOwnProperty.call(o, "raw");
	var isempty = true;
	var row/*:any*/ = (header === 1) ? [] : {};
	if(header !== 1) {
		if(Object.defineProperty) try { Object.defineProperty(row, '__rowNum__', {value:R, enumerable:false}); } catch(e) { row.__rowNum__ = R; }
		else row.__rowNum__ = R;
	}
	if(!dense || sheet[R]) for (var C = r.s.c; C <= r.e.c; ++C) {
		var val = dense ? sheet[R][C] : sheet[cols[C] + rr];
		if(val === undefined || val.t === undefined) {
			if(defval === undefined) continue;
			if(hdr[C] != null) { row[hdr[C]] = defval; }
			continue;
		}
		var v = val.v;
		switch(val.t){
			case 'z': if(v == null) break; continue;
			case 'e': v = (v == 0 ? null : void 0); break;
			case 's': case 'd': case 'b': case 'n': break;
			default: throw new Error('unrecognized type ' + val.t);
		}
		if(hdr[C] != null) {
			if(v == null) {
				if(val.t == "e" && v === null) row[hdr[C]] = null;
				else if(defval !== undefined) row[hdr[C]] = defval;
				else if(raw && v === null) row[hdr[C]] = null;
				else continue;
			} else {
				row[hdr[C]] = raw && (val.t !== "n" || (val.t === "n" && o.rawNumbers !== false)) ? v : format_cell(val,v,o);
			}
			if(v != null) isempty = false;
		}
	}
	return { row: row, isempty: isempty };
}


function sheet_to_json(sheet/*:Worksheet*/, opts/*:?Sheet2JSONOpts*/) {
	if(sheet == null || sheet["!ref"] == null) return [];
	var val = {t:'n',v:0}, header = 0, offset = 1, hdr/*:Array<any>*/ = [], v=0, vv="";
	var r = {s:{r:0,c:0},e:{r:0,c:0}};
	var o = opts || {};
	var range = o.range != null ? o.range : sheet["!ref"];
	if(o.header === 1) header = 1;
	else if(o.header === "A") header = 2;
	else if(Array.isArray(o.header)) header = 3;
	else if(o.header == null) header = 0;
	switch(typeof range) {
		case 'string': r = safe_decode_range(range); break;
		case 'number': r = safe_decode_range(sheet["!ref"]); r.s.r = range; break;
		default: r = range;
	}
	if(header > 0) offset = 0;
	var rr = encode_row(r.s.r);
	var cols/*:Array<string>*/ = [];
	var out/*:Array<any>*/ = [];
	var outi = 0, counter = 0;
	var dense = Array.isArray(sheet);
	var R = r.s.r, C = 0;
	var header_cnt = {};
	if(dense && !sheet[R]) sheet[R] = [];
	var colinfo/*:Array<ColInfo>*/ = o.skipHidden && sheet["!cols"] || [];
	var rowinfo/*:Array<ColInfo>*/ = o.skipHidden && sheet["!rows"] || [];
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
	for (R = r.s.r + offset; R <= r.e.r; ++R) {
		if ((rowinfo[R]||{}).hidden) continue;
		var row = make_json_row(sheet, r, R, cols, header, hdr, dense, o);
		if((row.isempty === false) || (header === 1 ? o.blankrows !== false : !!o.blankrows)) out[outi++] = row.row;
	}
	out.length = outi;
	return out;
}

var qreg = /"/g;
function make_csv_row(sheet/*:Worksheet*/, r/*:Range*/, R/*:number*/, cols/*:Array<string>*/, fs/*:number*/, rs/*:number*/, FS/*:string*/, o/*:Sheet2CSVOpts*/)/*:?string*/ {
	var isempty = true;
	var row/*:Array<string>*/ = [], txt = "", rr = encode_row(R);
	for(var C = r.s.c; C <= r.e.c; ++C) {
		if (!cols[C]) continue;
		var val = o.dense ? (sheet[R]||[])[C]: sheet[cols[C] + rr];
		if(val == null) txt = "";
		else if(val.v != null) {
			isempty = false;
			txt = ''+(o.rawNumbers && val.t == "n" ? val.v : format_cell(val, null, o));
			for(var i = 0, cc = 0; i !== txt.length; ++i) if((cc = txt.charCodeAt(i)) === fs || cc === rs || cc === 34 || o.forceQuotes) {txt = "\"" + txt.replace(qreg, '""') + "\""; break; }
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
	var out/*:Array<string>*/ = [];
	var o = opts == null ? {} : opts;
	if(sheet == null || sheet["!ref"] == null) return "";
	var r = safe_decode_range(sheet["!ref"]);
	var FS = o.FS !== undefined ? o.FS : ",", fs = FS.charCodeAt(0);
	var RS = o.RS !== undefined ? o.RS : "\n", rs = RS.charCodeAt(0);
	var endregex = new RegExp((FS=="|" ? "\\|" : FS)+"+$");
	var row = "", cols/*:Array<string>*/ = [];
	o.dense = Array.isArray(sheet);
	var colinfo/*:Array<ColInfo>*/ = o.skipHidden && sheet["!cols"] || [];
	var rowinfo/*:Array<ColInfo>*/ = o.skipHidden && sheet["!rows"] || [];
	for(var C = r.s.c; C <= r.e.c; ++C) if (!((colinfo[C]||{}).hidden)) cols[C] = encode_col(C);
	var w = 0;
	for(var R = r.s.r; R <= r.e.r; ++R) {
		if ((rowinfo[R]||{}).hidden) continue;
		row = make_csv_row(sheet, r, R, cols, fs, rs, FS, o);
		if(row == null) { continue; }
		if(o.strip) row = row.replace(endregex,"");
		if(row || (o.blankrows !== false)) out.push((w++ ? RS : "") + row);
	}
	delete o.dense;
	return out.join("");
}

function sheet_to_txt(sheet/*:Worksheet*/, opts/*:?Sheet2CSVOpts*/) {
	if(!opts) opts = {}; opts.FS = "\t"; opts.RS = "\n";
	var s = sheet_to_csv(sheet, opts);
	if(typeof $cptable == 'undefined' || opts.type == 'string') return s;
	var o = $cptable.utils.encode(1200, s, 'str');
	return String.fromCharCode(255) + String.fromCharCode(254) + o;
}

function sheet_to_formulae(sheet/*:Worksheet*/)/*:Array<string>*/ {
	var y = "", x, val="";
	if(sheet == null || sheet["!ref"] == null) return [];
	var r = safe_decode_range(sheet['!ref']), rr = "", cols/*:Array<string>*/ = [], C;
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

function sheet_add_json(_ws/*:?Worksheet*/, js/*:Array<any>*/, opts)/*:Worksheet*/ {
	var o = opts || {};
	var dense = _ws ? Array.isArray(_ws) : o.dense;
	if(DENSE != null && dense == null) dense = DENSE;
	var offset = +!o.skipHeader;
	var ws/*:Worksheet*/ = _ws || (dense ? ([]/*:any*/) : ({}/*:any*/));
	var _R = 0, _C = 0;
	if(ws && o.origin != null) {
		if(typeof o.origin == 'number') _R = o.origin;
		else {
			var _origin/*:CellAddress*/ = typeof o.origin == "string" ? decode_cell(o.origin) : o.origin;
			_R = _origin.r; _C = _origin.c;
		}
	}
	var range/*:Range*/ = ({s: {c:0, r:0}, e: {c:_C, r:_R + js.length - 1 + offset}}/*:any*/);
	if(ws['!ref']) {
		var _range = safe_decode_range(ws['!ref']);
		range.e.c = Math.max(range.e.c, _range.e.c);
		range.e.r = Math.max(range.e.r, _range.e.r);
		if(_R == -1) { _R = _range.e.r + 1; range.e.r = _R + js.length - 1 + offset; }
	} else {
		if(_R == -1) { _R = 0; range.e.r = js.length - 1 + offset; }
	}
	var hdr/*:Array<string>*/ = o.header || [], C = 0;
	var ROW = [];
	js.forEach(function (JS, R/*:number*/) {
		if(dense && !ws[_R + R + offset]) ws[_R + R + offset] = [];
		if(dense) ROW = ws[_R + R + offset];
		keys(JS).forEach(function(k) {
			if((C=hdr.indexOf(k)) == -1) hdr[C=hdr.length] = k;
			var v = JS[k];
			var t = 'z';
			var z = "";
			var ref = dense ? "" : encode_cell({c:_C + C,r:_R + R + offset});
			var cell/*:Cell*/ = dense ? ROW[_C + C] : ws[ref];
			if(v && typeof v === 'object' && !(v instanceof Date)){
				ws[ref] = v;
			} else {
				if(typeof v == 'number') t = 'n';
				else if(typeof v == 'boolean') t = 'b';
				else if(typeof v == 'string') t = 's';
				else if(v instanceof Date) {
					t = 'd';
					if(!o.cellDates) { t = 'n'; v = datenum(v); }
					z = (cell != null && cell.z && fmt_is_date(cell.z)) ? cell.z : (o.dateNF || table_fmt[14]);
				}
				else if(v === null && o.nullError) { t = 'e'; v = 0; }
				if(!cell) {
					if(!dense) ws[ref] = cell = ({t:t, v:v}/*:any*/);
					else ROW[_C + C] = cell = ({t:t, v:v}/*:any*/);
				} else {
					cell.t = t; cell.v = v;
					delete cell.w; delete cell.R;
					if(z) cell.z = z;
				}
				if(z) cell.z = z;
			}
		});
	});
	range.e.c = Math.max(range.e.c, _C + hdr.length - 1);
	var __R = encode_row(_R);
	if(dense && !ws[_R]) ws[_R] = [];
	if(offset) for(C = 0; C < hdr.length; ++C) {
		if(dense) ws[_R][C + _C] = {t:'s', v:hdr[C]};
		else ws[encode_col(C + _C) + __R] = {t:'s', v:hdr[C]};
	}
	ws['!ref'] = encode_range(range);
	return ws;
}
function json_to_sheet(js/*:Array<any>*/, opts)/*:Worksheet*/ { return sheet_add_json(null, js, opts); }

/* get cell, creating a stub if necessary */
function ws_get_cell_stub(ws/*:Worksheet*/, R, C/*:?number*/)/*:Cell*/ {
	/* A1 cell address */
	if(typeof R == "string") {
		/* dense */
		if(Array.isArray(ws)) {
			var RC = decode_cell(R);
			if(!ws[RC.r]) ws[RC.r] = [];
			return ws[RC.r][RC.c] || (ws[RC.r][RC.c] = {t:'z'});
		}
		return ws[R] || (ws[R] = {t:'z'});
	}
	/* cell address object */
	if(typeof R != "number") return ws_get_cell_stub(ws, encode_cell(R));
	/* R and C are 0-based indices */
	return ws_get_cell_stub(ws, encode_cell({r:R,c:C||0}));
}

/* find sheet index for given name / validate index */
function wb_sheet_idx(wb/*:Workbook*/, sh/*:number|string*/) {
	if(typeof sh == "number") {
		if(sh >= 0 && wb.SheetNames.length > sh) return sh;
		throw new Error("Cannot find sheet # " + sh);
	} else if(typeof sh == "string") {
		var idx = wb.SheetNames.indexOf(sh);
		if(idx > -1) return idx;
		throw new Error("Cannot find sheet name |" + sh + "|");
	} else throw new Error("Cannot find sheet |" + sh + "|");
}

/* simple blank workbook object */
function book_new()/*:Workbook*/ {
	return { SheetNames: [], Sheets: {} };
}

/* add a worksheet to the end of a given workbook */
function book_append_sheet(wb/*:Workbook*/, ws/*:Worksheet*/, name/*:?string*/, roll/*:?boolean*/)/*:string*/ {
	var i = 1;
	if(!name) for(; i <= 0xFFFF; ++i, name = undefined) if(wb.SheetNames.indexOf(name = "Sheet" + i) == -1) break;
	if(!name || wb.SheetNames.length >= 0xFFFF) throw new Error("Too many worksheets");
	if(roll && wb.SheetNames.indexOf(name) >= 0) {
		var m = name.match(/(^.*?)(\d+)$/);
		i = m && +m[2] || 0;
		var root = m && m[1] || name;
		for(++i; i <= 0xFFFF; ++i) if(wb.SheetNames.indexOf(name = root + i) == -1) break;
	}
	check_ws_name(name);
	if(wb.SheetNames.indexOf(name) >= 0) throw new Error("Worksheet with name |" + name + "| already exists!");

	wb.SheetNames.push(name);
	wb.Sheets[name] = ws;
	return name;
}

/* set sheet visibility (visible/hidden/very hidden) */
function book_set_sheet_visibility(wb/*:Workbook*/, sh/*:number|string*/, vis/*:number*/) {
	if(!wb.Workbook) wb.Workbook = {};
	if(!wb.Workbook.Sheets) wb.Workbook.Sheets = [];

	var idx = wb_sheet_idx(wb, sh);
	// $FlowIgnore
	if(!wb.Workbook.Sheets[idx]) wb.Workbook.Sheets[idx] = {};

	switch(vis) {
		case 0: case 1: case 2: break;
		default: throw new Error("Bad sheet visibility setting " + vis);
	}
	// $FlowIgnore
	wb.Workbook.Sheets[idx].Hidden = vis;
}

/* set number format */
function cell_set_number_format(cell/*:Cell*/, fmt/*:string|number*/) {
	cell.z = fmt;
	return cell;
}

/* set cell hyperlink */
function cell_set_hyperlink(cell/*:Cell*/, target/*:string*/, tooltip/*:?string*/) {
	if(!target) {
		delete cell.l;
	} else {
		cell.l = ({ Target: target }/*:Hyperlink*/);
		if(tooltip) cell.l.Tooltip = tooltip;
	}
	return cell;
}
function cell_set_internal_link(cell/*:Cell*/, range/*:string*/, tooltip/*:?string*/) { return cell_set_hyperlink(cell, "#" + range, tooltip); }

/* add to cell comments */
function cell_add_comment(cell/*:Cell*/, text/*:string*/, author/*:?string*/) {
	if(!cell.c) cell.c = [];
	cell.c.push({t:text, a:author||"SheetJS"});
}

/* set array formula and flush related cells */
function sheet_set_array_formula(ws/*:Worksheet*/, range, formula/*:string*/, dynamic/*:boolean*/) {
	var rng = typeof range != "string" ? range : safe_decode_range(range);
	var rngstr = typeof range == "string" ? range : encode_range(range);
	for(var R = rng.s.r; R <= rng.e.r; ++R) for(var C = rng.s.c; C <= rng.e.c; ++C) {
		var cell = ws_get_cell_stub(ws, R, C);
		cell.t = 'n';
		cell.F = rngstr;
		delete cell.v;
		if(R == rng.s.r && C == rng.s.c) {
			cell.f = formula;
			if(dynamic) cell.D = true;
		}
	}
	var wsr = decode_range(ws["!ref"]);
	if(wsr.s.r > rng.s.r) wsr.s.r = rng.s.r;
	if(wsr.s.c > rng.s.c) wsr.s.c = rng.s.c;
	if(wsr.e.r < rng.e.r) wsr.e.r = rng.e.r;
	if(wsr.e.c < rng.e.c) wsr.e.c = rng.e.c;
	ws["!ref"] = encode_range(wsr);
	return ws;
}

