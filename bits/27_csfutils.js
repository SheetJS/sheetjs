function decode_row(rowstr/*:string*/)/*:number*/ { return parseInt(unfix_row(rowstr),10) - 1; }
function encode_row(row/*:number*/)/*:string*/ { return "" + (row + 1); }
function fix_row(cstr/*:string*/)/*:string*/ { return cstr.replace(/([A-Z]|^)(\d+)$/,"$1$$$2"); }
function unfix_row(cstr/*:string*/)/*:string*/ { return cstr.replace(/\$(\d+)$/,"$1"); }

function decode_col(colstr/*:string*/)/*:number*/ { var c = unfix_col(colstr), d = 0, i = 0; for(; i !== c.length; ++i) d = 26*d + c.charCodeAt(i) - 64; return d - 1; }
function encode_col(col/*:number*/)/*:string*/ { var s=""; for(++col; col; col=Math.floor((col-1)/26)) s = String.fromCharCode(((col-1)%26) + 65) + s; return s; }
function fix_col(cstr/*:string*/)/*:string*/ { return cstr.replace(/^([A-Z])/,"$$$1"); }
function unfix_col(cstr/*:string*/)/*:string*/ { return cstr.replace(/^\$([A-Z])/,"$1"); }

function split_cell(cstr/*:string*/)/*:Array<string>*/ { return cstr.replace(/(\$?[A-Z]*)(\$?\d*)/,"$1,$2").split(","); }
function decode_cell(cstr/*:string*/)/*:CellAddress*/ { var splt = split_cell(cstr); return { c:decode_col(splt[0]), r:decode_row(splt[1]) }; }
function encode_cell(cell/*:CellAddress*/)/*:string*/ { return encode_col(cell.c) + encode_row(cell.r); }
function fix_cell(cstr/*:string*/)/*:string*/ { return fix_col(fix_row(cstr)); }
function unfix_cell(cstr/*:string*/)/*:string*/ { return unfix_col(unfix_row(cstr)); }
function decode_range(range/*:string*/)/*:Range*/ { var x =range.split(":").map(decode_cell); return {s:x[0],e:x[x.length-1]}; }
/*# if only one arg, it is assumed to be a Range.  If 2 args, both are cell addresses */
function encode_range(cs/*:CellAddrSpec|Range*/,ce/*:?CellAddrSpec*/)/*:string*/ {
	if(typeof ce === 'undefined' || typeof ce === 'number') {
/*:: if(!(cs instanceof Range)) throw "unreachable"; */
		return encode_range(cs.s, cs.e);
	}
/*:: if((cs instanceof Range)) throw "unreachable"; */
	if(typeof cs !== 'string') cs = encode_cell((cs/*:any*/));
	if(typeof ce !== 'string') ce = encode_cell((ce/*:any*/));
/*:: if(typeof cs !== 'string') throw "unreachable"; */
/*:: if(typeof ce !== 'string') throw "unreachable"; */
	return cs == ce ? cs : cs + ":" + ce;
}

function safe_decode_range(range/*:string*/)/*:Range*/ {
	var o = {s:{c:0,r:0},e:{c:0,r:0}};
	var idx = 0, i = 0, cc = 0;
	var len = range.length;
	for(idx = 0; i < len; ++i) {
		if((cc=range.charCodeAt(i)-64) < 1 || cc > 26) break;
		idx = 26*idx + cc;
	}
	o.s.c = --idx;

	for(idx = 0; i < len; ++i) {
		if((cc=range.charCodeAt(i)-48) < 0 || cc > 9) break;
		idx = 10*idx + cc;
	}
	o.s.r = --idx;

	if(i === len || range.charCodeAt(++i) === 58) { o.e.c=o.s.c; o.e.r=o.s.r; return o; }

	for(idx = 0; i != len; ++i) {
		if((cc=range.charCodeAt(i)-64) < 1 || cc > 26) break;
		idx = 26*idx + cc;
	}
	o.e.c = --idx;

	for(idx = 0; i != len; ++i) {
		if((cc=range.charCodeAt(i)-48) < 0 || cc > 9) break;
		idx = 10*idx + cc;
	}
	o.e.r = --idx;
	return o;
}

function safe_format_cell(cell/*:Cell*/, v/*:any*/) {
	var q = (cell.t == 'd' && v instanceof Date);
	if(cell.z != null) try { return (cell.w = SSF.format(cell.z, q ? datenum(v) : v)); } catch(e) { }
	try { return (cell.w = SSF.format((cell.XF||{}).numFmtId||(q ? 14 : 0),  q ? datenum(v) : v)); } catch(e) { return ''+v; }
}

function format_cell(cell/*:Cell*/, v/*:any*/, o/*:any*/) {
	if(cell == null || cell.t == null || cell.t == 'z') return "";
	if(cell.w !== undefined) return cell.w;
	if(cell.t == 'd' && !cell.z && o && o.dateNF) cell.z = o.dateNF;
	if(v == undefined) return safe_format_cell(cell, cell.v);
	return safe_format_cell(cell, v);
}

function sheet_to_workbook(sheet/*:Worksheet*/, opts)/*:Workbook*/ {
	var n = opts && opts.sheet ? opts.sheet : "Sheet1";
	var sheets = {}; sheets[n] = sheet;
	return { SheetNames: [n], Sheets: sheets };
}

function aoa_to_sheet(data/*:AOA*/, opts/*:?any*/)/*:Worksheet*/ {
	var o = opts || {};
	if(DENSE != null && o.dense == null) o.dense = DENSE;
	var ws/*:Worksheet*/ = o.dense ? ([]/*:any*/) : ({}/*:any*/);
	var range/*:Range*/ = ({s: {c:10000000, r:10000000}, e: {c:0, r:0}}/*:any*/);
	for(var R = 0; R != data.length; ++R) {
		for(var C = 0; C != data[R].length; ++C) {
			if(typeof data[R][C] === 'undefined') continue;
			var cell/*:Cell*/ = ({v: data[R][C] }/*:any*/);
			if(Array.isArray(cell.v)) { cell.f = data[R][C][1]; cell.v = cell.v[0]; }
			if(range.s.r > R) range.s.r = R;
			if(range.s.c > C) range.s.c = C;
			if(range.e.r < R) range.e.r = R;
			if(range.e.c < C) range.e.c = C;
			if(cell.v === null) { if(cell.f) cell.t = 'n'; else if(!o.cellStubs) continue; else cell.t = 'z'; }
			else if(typeof cell.v === 'number') cell.t = 'n';
			else if(typeof cell.v === 'boolean') cell.t = 'b';
			else if(cell.v instanceof Date) {
				cell.z = o.dateNF || SSF._table[14];
				if(o.cellDates) { cell.t = 'd'; cell.w = SSF.format(cell.z, datenum(cell.v)); }
				else { cell.t = 'n'; cell.v = datenum(cell.v); cell.w = SSF.format(cell.z, cell.v); }
			}
			else cell.t = 's';
			if(o.dense) {
				if(!ws[R]) ws[R] = [];
				ws[R][C] = cell;
			} else {
				var cell_ref = encode_cell(({c:C,r:R}/*:any*/));
				ws[cell_ref] = cell;
			}
		}
	}
	if(range.s.c < 10000000) ws['!ref'] = encode_range(range);
	return ws;
}

