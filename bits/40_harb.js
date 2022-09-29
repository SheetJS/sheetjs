var DBF_SUPPORTED_VERSIONS = [0x02, 0x03, 0x30, 0x31, 0x83, 0x8B, 0x8C, 0xF5];
var DBF = /*#__PURE__*/(function() {
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
	/*::[*/0x59/*::]*/:  1252,           /*::[*/0x6C/*::]*/:   863,
	/*::[*/0x86/*::]*/:   737,           /*::[*/0x87/*::]*/:   852,
	/*::[*/0x88/*::]*/:   857,           /*::[*/0xCC/*::]*/:  1257,

	/*::[*/0xFF/*::]*/: 16969
};
var dbf_reverse_map = evert({
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
	/*::[*/0x00/*::]*/: 20127
});
/* TODO: find an actual specification */
function dbf_to_aoa(buf, opts)/*:AOA*/ {
	var out/*:AOA*/ = [];
	var d/*:Block*/ = (new_raw_buf(1)/*:any*/);
	switch(opts.type) {
		case 'base64': d = s2a(Base64_decode(buf)); break;
		case 'binary': d = s2a(buf); break;
		case 'buffer':
		case 'array': d = buf; break;
	}
	prep_blob(d, 0);

	/* header */
	var ft = d.read_shift(1);
	var memo = !!(ft & 0x88);
	var vfp = false, l7 = false;
	switch(ft) {
		case 0x02: break; // dBASE II
		case 0x03: break; // dBASE III
		case 0x30: vfp = true; memo = true; break; // VFP
		case 0x31: vfp = true; memo = true; break; // VFP with autoincrement
		// 0x43 dBASE IV SQL table files
		// 0x63 dBASE IV SQL system files
		case 0x83: break; // dBASE III with memo
		case 0x8B: break; // dBASE IV with memo
		case 0x8C: l7 = true; break; // dBASE Level 7 with memo
		// case 0xCB dBASE IV SQL table files with memo
		case 0xF5: break; // FoxPro 2.x with memo
		// case 0xFB FoxBASE
		default: throw new Error("DBF Unsupported Version: " + ft.toString(16));
	}

	var nrow = 0, fpos = 0x0209;
	if(ft == 0x02) nrow = d.read_shift(2);
	d.l += 3; // dBASE II stores DDMMYY date, others use YYMMDD
	if(ft != 0x02) nrow = d.read_shift(4);
	if(nrow > 1048576) nrow = 1e6;

	if(ft != 0x02) fpos = d.read_shift(2); // header length
	var rlen = d.read_shift(2); // record length

	var /*flags = 0,*/ current_cp = opts.codepage || 1252;
	if(ft != 0x02) { // 20 reserved bytes
		d.l+=16;
		/*flags = */d.read_shift(1);
		//if(memo && ((flags & 0x02) === 0)) throw new Error("DBF Flags " + flags.toString(16) + " ft " + ft.toString(16));

		/* codepage present in FoxPro and dBASE Level 7 */
		if(d[d.l] !== 0) current_cp = dbf_codepage_map[d[d.l]];
		d.l+=1;

		d.l+=2;
	}
	if(l7) d.l += 36; // Level 7: 32 byte "Language driver name", 4 byte reserved

/*:: type DBFField = { name:string; len:number; type:string; } */
	var fields/*:Array<DBFField>*/ = [], field/*:DBFField*/ = ({}/*:any*/);
	var hend = Math.min(d.length, (ft == 0x02 ? 0x209 : (fpos - 10 - (vfp ? 264 : 0))));
	var ww = l7 ? 32 : 11;
	while(d.l < hend && d[d.l] != 0x0d) {
		field = ({}/*:any*/);
		field.name = (typeof $cptable !== "undefined" ? $cptable.utils.decode(current_cp, d.slice(d.l, d.l+ww)) : a2s(d.slice(d.l, d.l + ww))).replace(/[\u0000\r\n].*$/g,"");
		d.l += ww;
		field.type = String.fromCharCode(d.read_shift(1));
		if(ft != 0x02 && !l7) field.offset = d.read_shift(4);
		field.len = d.read_shift(1);
		if(ft == 0x02) field.offset = d.read_shift(2);
		field.dec = d.read_shift(1);
		if(field.name.length) fields.push(field);
		if(ft != 0x02) d.l += l7 ? 13 : 14;
		switch(field.type) {
			case 'B': // Double (VFP) / Binary (dBASE L7)
				if((!vfp || field.len != 8) && opts.WTF) console.log('Skipping ' + field.name + ':' + field.type);
				break;
			case 'G': // General (FoxPro and dBASE L7)
			case 'P': // Picture (FoxPro and dBASE L7)
				if(opts.WTF) console.log('Skipping ' + field.name + ':' + field.type);
				break;
			case '+': // Autoincrement (dBASE L7 only)
			case '0': // _NullFlags (VFP only)
			case '@': // Timestamp (dBASE L7 only)
			case 'C': // Character (dBASE II)
			case 'D': // Date (dBASE III)
			case 'F': // Float (dBASE IV)
			case 'I': // Long (VFP and dBASE L7)
			case 'L': // Logical (dBASE II)
			case 'M': // Memo (dBASE III)
			case 'N': // Number (dBASE II)
			case 'O': // Double (dBASE L7 only)
			case 'T': // Datetime (VFP only)
			case 'Y': // Currency (VFP only)
				break;
			default: throw new Error('Unknown Field Type: ' + field.type);
		}
	}

	if(d[d.l] !== 0x0D) d.l = fpos-1;
	if(d.read_shift(1) !== 0x0D) throw new Error("DBF Terminator not found " + d.l + " " + d[d.l]);
	d.l = fpos;

	/* data */
	var R = 0, C = 0;
	out[0] = [];
	for(C = 0; C != fields.length; ++C) out[0][C] = fields[C].name;
	while(nrow-- > 0) {
		if(d[d.l] === 0x2A) {
			// TODO: record marked as deleted -- create a hidden row?
			d.l+=rlen;
			continue;
		}
		++d.l;
		out[++R] = []; C = 0;
		for(C = 0; C != fields.length; ++C) {
			var dd = d.slice(d.l, d.l+fields[C].len); d.l+=fields[C].len;
			prep_blob(dd, 0);
			var s = typeof $cptable !== "undefined" ? $cptable.utils.decode(current_cp, dd) : a2s(dd);
			switch(fields[C].type) {
				case 'C':
					// NOTE: it is conventional to write '  /  /  ' for empty dates
					if(s.trim().length) out[R][C] = s.replace(/\s+$/,"");
					break;
				case 'D':
					if(s.length === 8) out[R][C] = new Date(+s.slice(0,4), +s.slice(4,6)-1, +s.slice(6,8));
					else out[R][C] = s;
					break;
				case 'F': out[R][C] = parseFloat(s.trim()); break;
				case '+': case 'I': out[R][C] = l7 ? dd.read_shift(-4, 'i') ^ 0x80000000 : dd.read_shift(4, 'i'); break;
				case 'L': switch(s.trim().toUpperCase()) {
					case 'Y': case 'T': out[R][C] = true; break;
					case 'N': case 'F': out[R][C] = false; break;
					case '': case '?': break;
					default: throw new Error("DBF Unrecognized L:|" + s + "|");
					} break;
				case 'M': /* TODO: handle memo files */
					if(!memo) throw new Error("DBF Unexpected MEMO for type " + ft.toString(16));
					out[R][C] = "##MEMO##" + (l7 ? parseInt(s.trim(), 10): dd.read_shift(4));
					break;
				case 'N':
					s = s.replace(/\u0000/g,"").trim();
					// NOTE: dBASE II interprets "  .  " as 0
					if(s && s != ".") out[R][C] = +s || 0; break;
				case '@':
					// NOTE: dBASE specs appear to be incorrect
					out[R][C] = new Date(dd.read_shift(-8, 'f') - 0x388317533400);
					break;
				case 'T': out[R][C] = new Date((dd.read_shift(4) - 0x253D8C) * 0x5265C00 + dd.read_shift(4)); break;
				case 'Y': out[R][C] = dd.read_shift(4,'i')/1e4 + (dd.read_shift(4, 'i')/1e4)*Math.pow(2,32); break;
				case 'O': out[R][C] = -dd.read_shift(-8, 'f'); break;
				case 'B': if(vfp && fields[C].len == 8) { out[R][C] = dd.read_shift(8,'f'); break; }
					/* falls through */
				case 'G': case 'P': dd.l += fields[C].len; break;
				case '0':
					if(fields[C].name === '_NullFlags') break;
					/* falls through */
				default: throw new Error("DBF Unsupported data type " + fields[C].type);
			}
		}
	}
	if(ft != 0x02) if(d.l < d.length && d[d.l++] != 0x1A) throw new Error("DBF EOF Marker missing " + (d.l-1) + " of " + d.length + " " + d[d.l-1].toString(16));
	if(opts && opts.sheetRows) out = out.slice(0, opts.sheetRows);
	opts.DBF = fields;
	return out;
}

function dbf_to_sheet(buf, opts)/*:Worksheet*/ {
	var o = opts || {};
	if(!o.dateNF) o.dateNF = "yyyymmdd";
	var ws = aoa_to_sheet(dbf_to_aoa(buf, o), o);
	ws["!cols"] = o.DBF.map(function(field) { return {
		wch: field.len,
		DBF: field
	};});
	delete o.DBF;
	return ws;
}

function dbf_to_workbook(buf, opts)/*:Workbook*/ {
	try {
		var o = sheet_to_workbook(dbf_to_sheet(buf, opts), opts);
		o.bookType = "dbf";
		return o;
	} catch(e) { if(opts && opts.WTF) throw e; }
	return ({SheetNames:[],Sheets:{}});
}

var _RLEN = { 'B': 8, 'C': 250, 'L': 1, 'D': 8, '?': 0, '': 0 };
function sheet_to_dbf(ws/*:Worksheet*/, opts/*:WriteOpts*/) {
	var o = opts || {};
	var old_cp = current_codepage;
	if(+o.codepage >= 0) set_cp(+o.codepage);
	if(o.type == "string") throw new Error("Cannot write DBF to JS string");
	var ba = buf_array();
	var aoa/*:AOA*/ = sheet_to_json(ws, {header:1, raw:true, cellDates:true});
	var headers = aoa[0], data = aoa.slice(1), cols = ws["!cols"] || [];
	var i = 0, j = 0, hcnt = 0, rlen = 1;
	for(i = 0; i < headers.length; ++i) {
		if(((cols[i]||{}).DBF||{}).name) { headers[i] = cols[i].DBF.name; ++hcnt; continue; }
		if(headers[i] == null) continue;
		++hcnt;
		if(typeof headers[i] === 'number') headers[i] = headers[i].toString(10);
		if(typeof headers[i] !== 'string') throw new Error("DBF Invalid column name " + headers[i] + " |" + (typeof headers[i]) + "|");
		if(headers.indexOf(headers[i]) !== i) for(j=0; j<1024;++j)
			if(headers.indexOf(headers[i] + "_" + j) == -1) { headers[i] += "_" + j; break; }
	}
	var range = safe_decode_range(ws['!ref']);
	var coltypes/*:Array<string>*/ = [];
	var colwidths/*:Array<number>*/ = [];
	var coldecimals/*:Array<number>*/ = [];
	for(i = 0; i <= range.e.c - range.s.c; ++i) {
		var guess = '', _guess = '', maxlen = 0;
		var col/*:Array<any>*/ = [];
		for(j=0; j < data.length; ++j) {
			if(data[j][i] != null) col.push(data[j][i]);
		}
		if(col.length == 0 || headers[i] == null) { coltypes[i] = '?'; continue; }
		for(j = 0; j < col.length; ++j) {
			switch(typeof col[j]) {
				/* TODO: check if L2 compat is desired */
				case 'number': _guess = 'B'; break;
				case 'string': _guess = 'C'; break;
				case 'boolean': _guess = 'L'; break;
				case 'object': _guess = col[j] instanceof Date ? 'D' : 'C'; break;
				default: _guess = 'C';
			}
			/* TODO: cache the values instead of encoding twice */
			maxlen = Math.max(maxlen, (typeof $cptable !== "undefined" && typeof col[j] == "string" ? $cptable.utils.encode(current_ansi, col[j]): String(col[j])).length);
			guess = guess && guess != _guess ? 'C' : _guess;
			//if(guess == 'C') break;
		}
		if(maxlen > 250) maxlen = 250;
		_guess = ((cols[i]||{}).DBF||{}).type;
		/* TODO: more fine grained control over DBF type resolution */
		if(_guess == 'C') {
			if(cols[i].DBF.len > maxlen) maxlen = cols[i].DBF.len;
		}
		if(guess == 'B' && _guess == 'N') {
			guess = 'N';
			coldecimals[i] = cols[i].DBF.dec;
			maxlen = cols[i].DBF.len;
		}
		colwidths[i] = guess == 'C' || _guess == 'N' ? maxlen : (_RLEN[guess] || 0);
		rlen += colwidths[i];
		coltypes[i] = guess;
	}

	var h = ba.next(32);
	h.write_shift(4, 0x13021130);
	h.write_shift(4, data.length);
	h.write_shift(2, 296 + 32 * hcnt);
	h.write_shift(2, rlen);
	for(i=0; i < 4; ++i) h.write_shift(4, 0);
	var cp = +dbf_reverse_map[/*::String(*/current_codepage/*::)*/] || 0x03;
	h.write_shift(4, 0x00000000 | (cp<<8));
	if(dbf_codepage_map[cp] != +o.codepage) {
		if(o.codepage) console.error("DBF Unsupported codepage " + current_codepage + ", using 1252");
		current_codepage = 1252;
	}

	for(i = 0, j = 0; i < headers.length; ++i) {
		if(headers[i] == null) continue;
		var hf = ba.next(32);
		/* TODO: test how applications handle non-ASCII field names */
		var _f = (headers[i].slice(-10) + "\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00").slice(0, 11);
		hf.write_shift(1, _f, "sbcs");
		hf.write_shift(1, coltypes[i] == '?' ? 'C' : coltypes[i], "sbcs");
		hf.write_shift(4, j);
		hf.write_shift(1, colwidths[i] || _RLEN[coltypes[i]] || 0);
		hf.write_shift(1, coldecimals[i] || 0);
		hf.write_shift(1, 0x02);
		hf.write_shift(4, 0);
		hf.write_shift(1, 0);
		hf.write_shift(4, 0);
		hf.write_shift(4, 0);
		j += (colwidths[i] || _RLEN[coltypes[i]] || 0);
	}

	var hb = ba.next(264);
	hb.write_shift(4, 0x0000000D);
	for(i=0; i < 65;++i) hb.write_shift(4, 0x00000000);
	for(i=0; i < data.length; ++i) {
		var rout = ba.next(rlen);
		rout.write_shift(1, 0);
		for(j=0; j<headers.length; ++j) {
			if(headers[j] == null) continue;
			switch(coltypes[j]) {
				case 'L': rout.write_shift(1, data[i][j] == null ? 0x3F : data[i][j] ? 0x54 : 0x46); break;
				case 'B': rout.write_shift(8, data[i][j]||0, 'f'); break;
				case 'N':
					var _n = "0";
					if(typeof data[i][j] == "number") _n = data[i][j].toFixed(coldecimals[j]||0);
					for(hcnt=0; hcnt < colwidths[j]-_n.length; ++hcnt) rout.write_shift(1, 0x20);
					rout.write_shift(1, _n, "sbcs");
					break;
				case 'D':
					if(!data[i][j]) rout.write_shift(8, "00000000", "sbcs");
					else {
						rout.write_shift(4, ("0000"+data[i][j].getFullYear()).slice(-4), "sbcs");
						rout.write_shift(2, ("00"+(data[i][j].getMonth()+1)).slice(-2), "sbcs");
						rout.write_shift(2, ("00"+data[i][j].getDate()).slice(-2), "sbcs");
					} break;
				case 'C':
					var _l = rout.l;
					var _s = String(data[i][j] != null ? data[i][j] : "").slice(0, colwidths[j]);
					rout.write_shift(1, _s, "cpstr");
					_l += colwidths[j] - rout.l;
					for(hcnt=0; hcnt < _l; ++hcnt) rout.write_shift(1, 0x20); break;
			}
		}
		// data
	}
	current_codepage = old_cp;
	ba.next(1).write_shift(1, 0x1A);
	return ba.end();
}
	return {
		to_workbook: dbf_to_workbook,
		to_sheet: dbf_to_sheet,
		from_sheet: sheet_to_dbf
	};
})();

var SYLK = /*#__PURE__*/(function() {
	/* TODO: stress test sequences */
	var sylk_escapes = ({
		AA:'À', BA:'Á', CA:'Â', DA:195, HA:'Ä', JA:197,
		AE:'È', BE:'É', CE:'Ê',         HE:'Ë',
		AI:'Ì', BI:'Í', CI:'Î',         HI:'Ï',
		AO:'Ò', BO:'Ó', CO:'Ô', DO:213, HO:'Ö',
		AU:'Ù', BU:'Ú', CU:'Û',         HU:'Ü',
		Aa:'à', Ba:'á', Ca:'â', Da:227, Ha:'ä', Ja:229,
		Ae:'è', Be:'é', Ce:'ê',         He:'ë',
		Ai:'ì', Bi:'í', Ci:'î',         Hi:'ï',
		Ao:'ò', Bo:'ó', Co:'ô', Do:245, Ho:'ö',
		Au:'ù', Bu:'ú', Cu:'û',         Hu:'ü',
		KC:'Ç', Kc:'ç', q:'æ',  z:'œ',  a:'Æ',  j:'Œ',
		DN:209, Dn:241, Hy:255,
		S:169,  c:170,  R:174,  "B ":180,
		/*::[*/0/*::]*/:176,    /*::[*/1/*::]*/:177,  /*::[*/2/*::]*/:178,
		/*::[*/3/*::]*/:179,    /*::[*/5/*::]*/:181,  /*::[*/6/*::]*/:182,
		/*::[*/7/*::]*/:183,    Q:185,  k:186,  b:208,  i:216,  l:222,  s:240,  y:248,
		"!":161, '"':162, "#":163, "(":164, "%":165, "'":167, "H ":168,
		"+":171, ";":187, "<":188, "=":189, ">":190, "?":191, "{":223
	}/*:any*/);
	var sylk_char_regex = new RegExp("\u001BN(" + keys(sylk_escapes).join("|").replace(/\|\|\|/, "|\\||").replace(/([?()+])/g,"\\$1") + "|\\|)", "gm");
	var sylk_char_fn = function(_, $1){ var o = sylk_escapes[$1]; return typeof o == "number" ? _getansi(o) : o; };
	var decode_sylk_char = function($$, $1, $2) { var newcc = (($1.charCodeAt(0) - 0x20)<<4) | ($2.charCodeAt(0) - 0x30); return newcc == 59 ? $$ : _getansi(newcc); };
	sylk_escapes["|"] = 254;
	/* https://oss.sheetjs.com/notes/sylk/ for more details */
	function sylk_to_aoa(d/*:RawData*/, opts)/*:[AOA, Worksheet]*/ {
		switch(opts.type) {
			case 'base64': return sylk_to_aoa_str(Base64_decode(d), opts);
			case 'binary': return sylk_to_aoa_str(d, opts);
			case 'buffer': return sylk_to_aoa_str(has_buf && Buffer.isBuffer(d) ? d.toString('binary') : a2s(d), opts);
			case 'array': return sylk_to_aoa_str(cc2str(d), opts);
		}
		throw new Error("Unrecognized type " + opts.type);
	}
	function sylk_to_aoa_str(str/*:string*/, opts)/*:[AOA, Worksheet]*/ {
		var records = str.split(/[\n\r]+/), R = -1, C = -1, ri = 0, rj = 0, arr/*:AOA*/ = [];
		var formats/*:Array<string>*/ = [];
		var next_cell_format/*:string|null*/ = null;
		var sht = {}, rowinfo/*:Array<RowInfo>*/ = [], colinfo/*:Array<ColInfo>*/ = [], cw/*:Array<string>*/ = [];
		var Mval = 0, j;
		var wb = { Workbook: { WBProps: {}, Names: [] } };
		if(+opts.codepage >= 0) set_cp(+opts.codepage);
		for (; ri !== records.length; ++ri) {
			Mval = 0;
			var rstr=records[ri].trim().replace(/\x1B([\x20-\x2F])([\x30-\x3F])/g, decode_sylk_char).replace(sylk_char_regex, sylk_char_fn);
			var record=rstr.replace(/;;/g, "\u0000").split(";").map(function(x) { return x.replace(/\u0000/g, ";"); });
			var RT=record[0], val;
			if(rstr.length > 0) switch(RT) {
			case 'ID': break; /* header */
			case 'E': break; /* EOF */
			case 'B': break; /* dimensions */
			case 'O': /* workbook options */
			for(rj=1; rj<record.length; ++rj) switch(record[rj].charAt(0)) {
				case 'V': {
					var d1904 = parseInt(record[rj].slice(1), 10);
					// NOTE: it is technically an error if d1904 >= 5 or < 0
					if(d1904 >= 1 && d1904 <= 4) wb.Workbook.WBProps.date1904 = true;
				} break;
			} break;
			case 'W': break; /* window */
			case 'P':
				switch(record[1].charAt(0)){
					case 'P': formats.push(rstr.slice(3).replace(/;;/g, ";")); break;
				} break;
			case 'NN': { /* defined name */
				var nn = {Sheet: 0};
				for(rj=1; rj<record.length; ++rj) switch(record[rj].charAt(0)) {
					case 'N': nn.Name = record[rj].slice(1); break;
					case 'E': nn.Ref = (opts && opts.sheet || "Sheet1") + "!" + rc_to_a1(record[rj].slice(1)); break;
				}
				wb.Workbook.Names.push(nn);
			} break;
			// case 'NE': // ??
			// case 'NU': // ??
			case 'C': /* cell */
			var C_seen_K = false, C_seen_X = false, C_seen_S = false, C_seen_E = false, _R = -1, _C = -1, formula = "", cell_t = "z";
			for(rj=1; rj<record.length; ++rj) switch(record[rj].charAt(0)) {
				case 'A': break; // TODO: comment
				case 'X': C = parseInt(record[rj].slice(1), 10)-1; C_seen_X = true; break;
				case 'Y':
					R = parseInt(record[rj].slice(1), 10)-1; if(!C_seen_X) C = 0;
					for(j = arr.length; j <= R; ++j) arr[j] = [];
					break;
				case 'K':
					val = record[rj].slice(1);
					if(val.charAt(0) === '"') { val = val.slice(1,val.length - 1); cell_t = "s"; }
					else if(val === 'TRUE' || val === 'FALSE') { val = val === 'TRUE'; cell_t = "b"; }
					else if(!isNaN(fuzzynum(val))) {
						val = fuzzynum(val); cell_t = "n";
						if(next_cell_format !== null && fmt_is_date(next_cell_format) && opts.cellDates) { val = numdate(wb.Workbook.WBProps.date1904 ? val + 1462 : val); cell_t = "d"; }
					} else if(!isNaN(fuzzydate(val).getDate())) {
						val = parseDate(val); cell_t = "d";
						if(!opts.cellDates) { cell_t = "n"; val = datenum(val, wb.Workbook.WBProps.date1904); }
					}
					if(typeof $cptable !== 'undefined' && typeof val == "string" && ((opts||{}).type != "string") && (opts||{}).codepage) val = $cptable.utils.decode(opts.codepage, val);
					C_seen_K = true;
					break;
				case 'E':
					C_seen_E = true;
					formula = rc_to_a1(record[rj].slice(1), {r:R,c:C});
					break;
				case 'S':
					C_seen_S = true;
					break;
				case 'G': break; // unknown
				case 'R': _R = parseInt(record[rj].slice(1), 10)-1; break;
				case 'C': _C = parseInt(record[rj].slice(1), 10)-1; break;
				// case 'P': // ??
				// case 'D': // ??
				default: if(opts && opts.WTF) throw new Error("SYLK bad record " + rstr);
			}
			if(C_seen_K) {
				if(!arr[R][C]) arr[R][C] = { t: cell_t, v: val };
				else { arr[R][C].t = cell_t; arr[R][C].v = val; }
				if(next_cell_format) arr[R][C].z = next_cell_format;
				if(opts.cellText !== false && next_cell_format) arr[R][C].w = SSF_format(arr[R][C].z, arr[R][C].v, { date1904: wb.Workbook.WBProps.date1904 });
				next_cell_format = null;
			}
			if(C_seen_S) {
				if(C_seen_E) throw new Error("SYLK shared formula cannot have own formula");
				var shrbase = _R > -1 && arr[_R][_C];
				if(!shrbase || !shrbase[1]) throw new Error("SYLK shared formula cannot find base");
				formula = shift_formula_str(shrbase[1], {r: R - _R, c: C - _C});
			}
			if(formula) {
				if(!arr[R][C]) arr[R][C] = { t: 'n', f: formula };
				else arr[R][C].f = formula;
			}
			break;
			case 'F': /* Format */
			var F_seen = 0;
			for(rj=1; rj<record.length; ++rj) switch(record[rj].charAt(0)) {
				case 'X': C = parseInt(record[rj].slice(1), 10)-1; ++F_seen; break;
				case 'Y':
					R = parseInt(record[rj].slice(1), 10)-1; /*C = 0;*/
					for(j = arr.length; j <= R; ++j) arr[j] = [];
					break;
				case 'M': Mval = parseInt(record[rj].slice(1), 10) / 20; break;
				case 'F': break; /* ??? */
				case 'G': break; /* hide grid */
				case 'P':
					next_cell_format = formats[parseInt(record[rj].slice(1), 10)];
					break;
				case 'S': break; /* cell style */
				case 'D': break; /* column */
				case 'N': break; /* font */
				case 'W':
					cw = record[rj].slice(1).split(" ");
					for(j = parseInt(cw[0], 10); j <= parseInt(cw[1], 10); ++j) {
						Mval = parseInt(cw[2], 10);
						colinfo[j-1] = Mval === 0 ? {hidden:true}: {wch:Mval};
					} break;
				case 'C': /* default column format */
					C = parseInt(record[rj].slice(1), 10)-1;
					if(!colinfo[C]) colinfo[C] = {};
					break;
				case 'R': /* row properties */
					R = parseInt(record[rj].slice(1), 10)-1;
					if(!rowinfo[R]) rowinfo[R] = {};
					if(Mval > 0) { rowinfo[R].hpt = Mval; rowinfo[R].hpx = pt2px(Mval); }
					else if(Mval === 0) rowinfo[R].hidden = true;
					break;
				// case 'K': // ??
				// case 'E': // ??
				default: if(opts && opts.WTF) throw new Error("SYLK bad record " + rstr);
			}
			if(F_seen < 1) next_cell_format = null; break;
			default: if(opts && opts.WTF) throw new Error("SYLK bad record " + rstr);
			}
		}
		if(rowinfo.length > 0) sht['!rows'] = rowinfo;
		if(colinfo.length > 0) sht['!cols'] = colinfo;
		colinfo.forEach(function(col) { process_col(col); });
		if(opts && opts.sheetRows) arr = arr.slice(0, opts.sheetRows);
		return [arr, sht, wb];
	}

	function sylk_to_workbook(d/*:RawData*/, opts)/*:Workbook*/ {
		var aoasht = sylk_to_aoa(d, opts);
		var aoa = aoasht[0], ws = aoasht[1], wb = aoasht[2];
		var _opts = dup(opts); _opts.date1904 = (((wb||{}).Workbook || {}).WBProps || {}).date1904;
		var o = aoa_to_sheet(aoa, _opts);
		keys(ws).forEach(function(k) { o[k] = ws[k]; });
		var outwb = sheet_to_workbook(o, opts);
		keys(wb).forEach(function(k) { outwb[k] = wb[k]; });
		outwb.bookType = "sylk";
		return outwb;
	}

	function write_ws_cell_sylk(cell/*:Cell*/, ws/*:Worksheet*/, R/*:number*/, C/*:number*//*::, opts*/)/*:string*/ {
		var o = "C;Y" + (R+1) + ";X" + (C+1) + ";K";
		switch(cell.t) {
			case 'n':
				o += (cell.v||0);
				if(cell.f && !cell.F) o += ";E" + a1_to_rc(cell.f, {r:R, c:C}); break;
			case 'b': o += cell.v ? "TRUE" : "FALSE"; break;
			case 'e': o += cell.w || cell.v; break;
			case 'd': o += '"' + (cell.w || cell.v) + '"'; break;
			case 's': o += '"' + (cell.v == null ? "" : String(cell.v)).replace(/"/g,"").replace(/;/g, ";;") + '"'; break;
		}
		return o;
	}

	function write_ws_cols_sylk(out, cols) {
		cols.forEach(function(col, i) {
			var rec = "F;W" + (i+1) + " " + (i+1) + " ";
			if(col.hidden) rec += "0";
			else {
				if(typeof col.width == 'number' && !col.wpx) col.wpx = width2px(col.width);
				if(typeof col.wpx == 'number' && !col.wch) col.wch = px2char(col.wpx);
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

	function sheet_to_sylk(ws/*:Worksheet*/, opts/*:?any*/, wb/*:?WorkBook*/)/*:string*/ {
		/* TODO: codepage */
		var preamble/*:Array<string>*/ = ["ID;PSheetJS;N;E"], o/*:Array<string>*/ = [];
		var r = safe_decode_range(ws['!ref']), cell/*:Cell*/;
		var dense = Array.isArray(ws);
		var RS = "\r\n";
		var d1904 = (((wb||{}).Workbook||{}).WBProps||{}).date1904;

		preamble.push("P;PGeneral");
		preamble.push("F;P0;DG0G8;M255");
		if(ws['!cols']) write_ws_cols_sylk(preamble, ws['!cols']);
		if(ws['!rows']) write_ws_rows_sylk(preamble, ws['!rows']);

		preamble.push("B;Y" + (r.e.r - r.s.r + 1) + ";X" + (r.e.c - r.s.c + 1) + ";D" + [r.s.c,r.s.r,r.e.c,r.e.r].join(" "));
		preamble.push("O;L;D;B" + (d1904 ? ";V4" : "") + ";K47;G100 0.001");
		for(var R = r.s.r; R <= r.e.r; ++R) {
			var p = [];
			for(var C = r.s.c; C <= r.e.c; ++C) {
				var coord = encode_cell({r:R,c:C});
				cell = dense ? (ws[R]||[])[C]: ws[coord];
				if(!cell || (cell.v == null && (!cell.f || cell.F))) continue;
				p.push(write_ws_cell_sylk(cell, ws, R, C, opts)); // TODO: pass date1904 info
			}
			o.push(p.join(RS));
		}
		return preamble.join(RS) + RS + o.join(RS) + RS + "E" + RS;
	}

	return {
		to_workbook: sylk_to_workbook,
		from_sheet: sheet_to_sylk
	};
})();

var DIF = /*#__PURE__*/(function() {
	function dif_to_aoa(d/*:RawData*/, opts)/*:AOA*/ {
		switch(opts.type) {
			case 'base64': return dif_to_aoa_str(Base64_decode(d), opts);
			case 'binary': return dif_to_aoa_str(d, opts);
			case 'buffer': return dif_to_aoa_str(has_buf && Buffer.isBuffer(d) ? d.toString('binary') : a2s(d), opts);
			case 'array': return dif_to_aoa_str(cc2str(d), opts);
		}
		throw new Error("Unrecognized type " + opts.type);
	}
	function dif_to_aoa_str(str/*:string*/, opts)/*:AOA*/ {
		var records = str.split('\n'), R = -1, C = -1, ri = 0, arr/*:AOA*/ = [];
		for (; ri !== records.length; ++ri) {
			if (records[ri].trim() === 'BOT') { arr[++R] = []; C = 0; continue; }
			if (R < 0) continue;
			var metadata = records[ri].trim().split(",");
			var type = metadata[0], value = metadata[1];
			++ri;
			var data = records[ri] || "";
			while(((data.match(/["]/g)||[]).length & 1) && ri < records.length - 1) data += "\n" + records[++ri];
			data = data.trim();
			switch (+type) {
				case -1:
					if (data === 'BOT') { arr[++R] = []; C = 0; continue; }
					else if (data !== 'EOD') throw new Error("Unrecognized DIF special command " + data);
					break;
				case 0:
					if(data === 'TRUE') arr[R][C] = true;
					else if(data === 'FALSE') arr[R][C] = false;
					else if(!isNaN(fuzzynum(value))) arr[R][C] = fuzzynum(value);
					else if(!isNaN(fuzzydate(value).getDate())) arr[R][C] = parseDate(value);
					else arr[R][C] = value;
					++C; break;
				case 1:
					data = data.slice(1,data.length-1);
					data = data.replace(/""/g, '"');
					if(DIF_XL && data && data.match(/^=".*"$/)) data = data.slice(2, -1);
					arr[R][C++] = data !== '' ? data : null;
					break;
			}
			if (data === 'EOD') break;
		}
		if(opts && opts.sheetRows) arr = arr.slice(0, opts.sheetRows);
		return arr;
	}

	function dif_to_sheet(str/*:string*/, opts)/*:Worksheet*/ { return aoa_to_sheet(dif_to_aoa(str, opts), opts); }
	function dif_to_workbook(str/*:string*/, opts)/*:Workbook*/ {
		var o = sheet_to_workbook(dif_to_sheet(str, opts), opts);
		o.bookType = "dif";
		return o;
	}

	function make_value(v/*:number*/, s/*:string*/)/*:string*/ { return "0," + String(v) + "\r\n" + s; }
	function make_value_str(s/*:string*/)/*:string*/ { return "1,0\r\n\"" + s.replace(/"/g,'""') + '"'; }
	function sheet_to_dif(ws/*:Worksheet*//*::, opts:?any*/)/*:string*/ {
		var _DIF_XL = DIF_XL;
		var r = safe_decode_range(ws['!ref']);
		var dense = Array.isArray(ws);
		var o/*:Array<string>*/ = [
			"TABLE\r\n0,1\r\n\"sheetjs\"\r\n",
			"VECTORS\r\n0," + (r.e.r - r.s.r + 1) + "\r\n\"\"\r\n",
			"TUPLES\r\n0," + (r.e.c - r.s.c + 1) + "\r\n\"\"\r\n",
			"DATA\r\n0,0\r\n\"\"\r\n"
		];
		for(var R = r.s.r; R <= r.e.r; ++R) {
			var p = "-1,0\r\nBOT\r\n";
			for(var C = r.s.c; C <= r.e.c; ++C) {
				var cell/*:Cell*/ = dense ? (ws[R] && ws[R][C]) : ws[encode_cell({r:R,c:C})];
				if(cell == null) { p +=("1,0\r\n\"\"\r\n"); continue;}
				switch(cell.t) {
					case 'n':
						if(_DIF_XL) {
							if(cell.w != null) p +=("0," + cell.w + "\r\nV");
							else if(cell.v != null) p +=(make_value(cell.v, "V")); // TODO: should this call SSF_format?
							else if(cell.f != null && !cell.F) p +=(make_value_str("=" + cell.f));
							else p +=("1,0\r\n\"\"");
						} else {
							if(cell.v == null) p +=("1,0\r\n\"\"");
							else p +=(make_value(cell.v, "V"));
						}
						break;
					case 'b':
						p +=(cell.v ? make_value(1, "TRUE") : make_value(0, "FALSE"));
						break;
					case 's':
						p +=(make_value_str((!_DIF_XL || isNaN(+cell.v)) ? cell.v : '="' + cell.v + '"'));
						break;
					case 'd':
						if(!cell.w) cell.w = SSF_format(cell.z || table_fmt[14], datenum(parseDate(cell.v)));
						if(_DIF_XL) p +=(make_value(cell.w, "V"));
						else p +=(make_value_str(cell.w));
						break;
					default: p +=("1,0\r\n\"\"");
				}
				p += "\r\n";
			}
			o.push(p);
		}
		return o.join("") + "-1,0\r\nEOD";
	}
	return {
		to_workbook: dif_to_workbook,
		to_sheet: dif_to_sheet,
		from_sheet: sheet_to_dif
	};
})();

var ETH = /*#__PURE__*/(function() {
	function decode(s/*:string*/)/*:string*/ { return s.replace(/\\b/g,"\\").replace(/\\c/g,":").replace(/\\n/g,"\n"); }
	function encode(s/*:string*/)/*:string*/ { return s.replace(/\\/g, "\\b").replace(/:/g, "\\c").replace(/\n/g,"\\n"); }

	function eth_to_aoa(str/*:string*/, opts)/*:AOA*/ {
		var records = str.split('\n'), R = -1, C = -1, ri = 0, arr/*:AOA*/ = [];
		for (; ri !== records.length; ++ri) {
			var record = records[ri].trim().split(":");
			if(record[0] !== 'cell') continue;
			var addr = decode_cell(record[1]);
			if(arr.length <= addr.r) for(R = arr.length; R <= addr.r; ++R) if(!arr[R]) arr[R] = [];
			R = addr.r; C = addr.c;
			switch(record[2]) {
				case 't': arr[R][C] = decode(record[3]); break;
				case 'v': arr[R][C] = +record[3]; break;
				case 'vtf': var _f = record[record.length - 1];
					/* falls through */
				case 'vtc':
					switch(record[3]) {
						case 'nl': arr[R][C] = +record[4] ? true : false; break;
						default: arr[R][C] = +record[4]; break;
					}
					if(record[2] == 'vtf') arr[R][C] = [arr[R][C], _f];
			}
		}
		if(opts && opts.sheetRows) arr = arr.slice(0, opts.sheetRows);
		return arr;
	}

	function eth_to_sheet(d/*:string*/, opts)/*:Worksheet*/ { return aoa_to_sheet(eth_to_aoa(d, opts), opts); }
	function eth_to_workbook(d/*:string*/, opts)/*:Workbook*/ { return sheet_to_workbook(eth_to_sheet(d, opts), opts); }

	var header = [
		"socialcalc:version:1.5",
		"MIME-Version: 1.0",
		"Content-Type: multipart/mixed; boundary=SocialCalcSpreadsheetControlSave"
	].join("\n");

	var sep = [
		"--SocialCalcSpreadsheetControlSave",
		"Content-type: text/plain; charset=UTF-8"
	].join("\n") + "\n";

	/* TODO: the other parts */
	var meta = [
		"# SocialCalc Spreadsheet Control Save",
		"part:sheet"
	].join("\n");

	var end = "--SocialCalcSpreadsheetControlSave--";

	function sheet_to_eth_data(ws/*:Worksheet*/)/*:string*/ {
		if(!ws || !ws['!ref']) return "";
		var o/*:Array<string>*/ = [], oo/*:Array<string>*/ = [], cell, coord = "";
		var r = decode_range(ws['!ref']);
		var dense = Array.isArray(ws);
		for(var R = r.s.r; R <= r.e.r; ++R) {
			for(var C = r.s.c; C <= r.e.c; ++C) {
				coord = encode_cell({r:R,c:C});
				cell = dense ? (ws[R]||[])[C] : ws[coord];
				if(!cell || cell.v == null || cell.t === 'z') continue;
				oo = ["cell", coord, 't'];
				switch(cell.t) {
					case 's': case 'str': oo.push(encode(cell.v)); break;
					case 'n':
						if(!cell.f) { oo[2]='v'; oo[3]=cell.v; }
						else { oo[2]='vtf'; oo[3]='n'; oo[4]=cell.v; oo[5]=encode(cell.f); }
						break;
					case 'b':
						oo[2] = 'vt'+(cell.f?'f':'c'); oo[3]='nl'; oo[4]=cell.v?"1":"0";
						oo[5] = encode(cell.f||(cell.v?'TRUE':'FALSE'));
						break;
					case 'd':
						var t = datenum(parseDate(cell.v));
						oo[2] = 'vtc'; oo[3] = 'nd'; oo[4] = ""+t;
						oo[5] = cell.w || SSF_format(cell.z || table_fmt[14], t);
						break;
					case 'e': continue;
				}
				o.push(oo.join(":"));
			}
		}
		o.push("sheet:c:" + (r.e.c-r.s.c+1) + ":r:" + (r.e.r-r.s.r+1) + ":tvf:1");
		o.push("valueformat:1:text-wiki");
		//o.push("copiedfrom:" + ws['!ref']); // clipboard only
		return o.join("\n");
	}

	function sheet_to_eth(ws/*:Worksheet*//*::, opts:?any*/)/*:string*/ {
		return [header, sep, meta, sep, sheet_to_eth_data(ws), end].join("\n");
		// return ["version:1.5", sheet_to_eth_data(ws)].join("\n"); // clipboard form
	}

	return {
		to_workbook: eth_to_workbook,
		to_sheet: eth_to_sheet,
		from_sheet: sheet_to_eth
	};
})();

var PRN = /*#__PURE__*/(function() {
	function set_text_arr(data/*:string*/, arr/*:AOA*/, R/*:number*/, C/*:number*/, o/*:any*/) {
		if(o.raw) arr[R][C] = data;
		else if(data === ""){/* empty */}
		else if(data === 'TRUE') arr[R][C] = true;
		else if(data === 'FALSE') arr[R][C] = false;
		else if(!isNaN(fuzzynum(data))) arr[R][C] = fuzzynum(data);
		else if(!isNaN(fuzzydate(data).getDate())) arr[R][C] = parseDate(data);
		else arr[R][C] = data;
	}

	function prn_to_aoa_str(f/*:string*/, opts)/*:AOA*/ {
		var o = opts || {};
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
			set_text_arr(lines[R].slice(0, start).trim(), arr, R, C, o);
			for(C = 1; C <= (lines[R].length - start)/10 + 1; ++C)
				set_text_arr(lines[R].slice(start+(C-1)*10,start+C*10).trim(),arr,R,C,o);
		}
		if(o.sheetRows) arr = arr.slice(0, o.sheetRows);
		return arr;
	}

	// List of accepted CSV separators
	var guess_seps = {
		/*::[*/0x2C/*::]*/: ',',
		/*::[*/0x09/*::]*/: "\t",
		/*::[*/0x3B/*::]*/: ';',
		/*::[*/0x7C/*::]*/: '|'
	};

	// CSV separator weights to be used in case of equal numbers
	var guess_sep_weights = {
		/*::[*/0x2C/*::]*/: 3,
		/*::[*/0x09/*::]*/: 2,
		/*::[*/0x3B/*::]*/: 1,
		/*::[*/0x7C/*::]*/: 0
	};

	function guess_sep(str) {
		var cnt = {}, instr = false, end = 0, cc = 0;
		for(;end < str.length;++end) {
			if((cc=str.charCodeAt(end)) == 0x22) instr = !instr;
			else if(!instr && cc in guess_seps) cnt[cc] = (cnt[cc]||0)+1;
		}

		cc = [];
		for(end in cnt) if ( Object.prototype.hasOwnProperty.call(cnt, end) ) {
			cc.push([ cnt[end], end ]);
		}

		if ( !cc.length ) {
			cnt = guess_sep_weights;
			for(end in cnt) if ( Object.prototype.hasOwnProperty.call(cnt, end) ) {
				cc.push([ cnt[end], end ]);
			}
		}

		cc.sort(function(a, b) { return a[0] - b[0] || guess_sep_weights[a[1]] - guess_sep_weights[b[1]]; });

		return guess_seps[cc.pop()[1]] || 0x2C;
	}

	function dsv_to_sheet_str(str/*:string*/, opts)/*:Worksheet*/ {
		var o = opts || {};
		var sep = "";
		if(DENSE != null && o.dense == null) o.dense = DENSE;
		var ws/*:Worksheet*/ = o.dense ? ([]/*:any*/) : ({}/*:any*/);
		var range/*:Range*/ = ({s: {c:0, r:0}, e: {c:0, r:0}}/*:any*/);

		if(str.slice(0,4) == "sep=") {
			// If the line ends in \r\n
			if(str.charCodeAt(5) == 13 && str.charCodeAt(6) == 10 ) {
				sep = str.charAt(4); str = str.slice(7);
			}
			// If line ends in \r OR \n
			else if(str.charCodeAt(5) == 13 || str.charCodeAt(5) == 10 ) {
				sep = str.charAt(4); str = str.slice(6);
			}
			else sep = guess_sep(str.slice(0,1024));
		}
		else if(o && o.FS) sep = o.FS;
		else sep = guess_sep(str.slice(0,1024));
		var R = 0, C = 0, v = 0;
		var start = 0, end = 0, sepcc = sep.charCodeAt(0), instr = false, cc=0, startcc=str.charCodeAt(0);
		var _re/*:?RegExp*/ = o.dateNF != null ? dateNF_regex(o.dateNF) : null;
		function finish_cell() {
			var s = str.slice(start, end); if(s.slice(-1) == "\r") s = s.slice(0, -1);
			var cell = ({}/*:any*/);
			if(s.charAt(0) == '"' && s.charAt(s.length - 1) == '"') s = s.slice(1,-1).replace(/""/g,'"');
			if(s.length === 0) cell.t = 'z';
			else if(o.raw) { cell.t = 's'; cell.v = s; }
			else if(s.trim().length === 0) { cell.t = 's'; cell.v = s; }
			else if(s.charCodeAt(0) == 0x3D) {
				if(s.charCodeAt(1) == 0x22 && s.charCodeAt(s.length - 1) == 0x22) { cell.t = 's'; cell.v = s.slice(2,-1).replace(/""/g,'"'); }
				else if(fuzzyfmla(s)) { cell.t = 'n'; cell.f = s.slice(1); }
				else { cell.t = 's'; cell.v = s; } }
			else if(s == "TRUE") { cell.t = 'b'; cell.v = true; }
			else if(s == "FALSE") { cell.t = 'b'; cell.v = false; }
			else if(!isNaN(v = fuzzynum(s))) { cell.t = 'n'; if(o.cellText !== false) cell.w = s; cell.v = v; }
			else if(!isNaN((v = fuzzydate(s)).getDate()) || _re && s.match(_re)) {
				cell.z = o.dateNF || table_fmt[14];
				var k = 0;
				if(_re && s.match(_re)){ s=dateNF_fix(s, o.dateNF, (s.match(_re)||[])); k=1; v = parseDate(s, k); }
				if(o.cellDates) { cell.t = 'd'; cell.v = v; }
				else { cell.t = 'n'; cell.v = datenum(v); }
				if(o.cellText !== false) cell.w = SSF_format(cell.z, cell.v instanceof Date ? datenum(cell.v):cell.v);
				if(!o.cellNF) delete cell.z;
			} else {
				cell.t = 's';
				cell.v = s;
			}
			if(cell.t == 'z'){}
			else if(o.dense) { if(!ws[R]) ws[R] = []; ws[R][C] = cell; }
			else ws[encode_cell({c:C,r:R})] = cell;
			start = end+1; startcc = str.charCodeAt(start);
			if(range.e.c < C) range.e.c = C;
			if(range.e.r < R) range.e.r = R;
			if(cc == sepcc) ++C; else { C = 0; ++R; if(o.sheetRows && o.sheetRows <= R) return true; }
		}
		outer: for(;end < str.length;++end) switch((cc=str.charCodeAt(end))) {
			case 0x22: if(startcc === 0x22) instr = !instr; break;
			case 0x0d:
				if(instr) break;
				if(str.charCodeAt(end+1) == 0x0a) ++end;
				/* falls through */
			case sepcc: case 0x0a: if(!instr && finish_cell()) break outer; break;
			default: break;
		}
		if(end - start > 0) finish_cell();

		ws['!ref'] = encode_range(range);
		return ws;
	}

	function prn_to_sheet_str(str/*:string*/, opts)/*:Worksheet*/ {
		if(!(opts && opts.PRN)) return dsv_to_sheet_str(str, opts);
		if(opts.FS) return dsv_to_sheet_str(str, opts);
		if(str.slice(0,4) == "sep=") return dsv_to_sheet_str(str, opts);
		if(str.indexOf("\t") >= 0 || str.indexOf(",") >= 0 || str.indexOf(";") >= 0) return dsv_to_sheet_str(str, opts);
		return aoa_to_sheet(prn_to_aoa_str(str, opts), opts);
	}

	function prn_to_sheet(d/*:RawData*/, opts)/*:Worksheet*/ {
		var str = "", bytes = opts.type == 'string' ? [0,0,0,0] : firstbyte(d, opts);
		switch(opts.type) {
			case 'base64': str = Base64_decode(d); break;
			case 'binary': str = d; break;
			case 'buffer':
				if(opts.codepage == 65001) str = d.toString('utf8'); // TODO: test if buf
				else if(opts.codepage && typeof $cptable !== 'undefined') str = $cptable.utils.decode(opts.codepage, d);
				else str = has_buf && Buffer.isBuffer(d) ? d.toString('binary') : a2s(d);
				break;
			case 'array': str = cc2str(d); break;
			case 'string': str = d; break;
			default: throw new Error("Unrecognized type " + opts.type);
		}
		if(bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) str = utf8read(str.slice(3));
		else if(opts.type != 'string' && opts.type != 'buffer' && opts.codepage == 65001) str = utf8read(str);
		else if((opts.type == 'binary') && typeof $cptable !== 'undefined' && opts.codepage)  str = $cptable.utils.decode(opts.codepage, $cptable.utils.encode(28591,str));
		if(str.slice(0,19) == "socialcalc:version:") return ETH.to_sheet(opts.type == 'string' ? str : utf8read(str), opts);
		return prn_to_sheet_str(str, opts);
	}

	function prn_to_workbook(d/*:RawData*/, opts)/*:Workbook*/ { return sheet_to_workbook(prn_to_sheet(d, opts), opts); }

	function sheet_to_prn(ws/*:Worksheet*//*::, opts:?any*/)/*:string*/ {
		var o/*:Array<string>*/ = [];
		var r = safe_decode_range(ws['!ref']), cell/*:Cell*/;
		var dense = Array.isArray(ws);
		for(var R = r.s.r; R <= r.e.r; ++R) {
			var oo/*:Array<string>*/ = [];
			for(var C = r.s.c; C <= r.e.c; ++C) {
				var coord = encode_cell({r:R,c:C});
				cell = dense ? (ws[R]||[])[C] : ws[coord];
				if(!cell || cell.v == null) { oo.push("          "); continue; }
				var w = (cell.w || (format_cell(cell), cell.w) || "").slice(0,10);
				while(w.length < 10) w += " ";
				oo.push(w + (C === 0 ? " " : ""));
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

