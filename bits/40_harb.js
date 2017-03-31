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
		case 0x03: break;
		case 0x30: vfp = true; memo = true; break;
		case 0x31: vfp = true; break;
		case 0x83: memo = true; break;
		case 0x8B: memo = true; break;
		case 0xF5: memo = true; break;
		default: throw new Error("DBF Unsupported Version: " + ft.toString(16));
	}
	var filedate = new Date(d.read_shift(1) + 1900, d.read_shift(1) - 1, d.read_shift(1));
	var nrow = d.read_shift(4);
	var fpos = d.read_shift(2);
	var rlen = d.read_shift(2);
	d.l+=16;

	var flags = d.read_shift(1);
	//if(memo && ((flags & 0x02) === 0)) throw new Error("DBF Flags " + flags.toString(16) + " ft " + ft.toString(16));

	/* codepage present in FoxPro */
	var current_cp = 1252;
	if(d[d.l] !== 0) current_cp = dbf_codepage_map[d[d.l]];
	d.l+=1;

	d.l+=2;
	var fields = [], field = {};
	var hend = fpos - 10 - (vfp ? 264 : 0);
	while(d.l < hend) {
		field = {};
		field.name = cptable.utils.decode(current_cp, d.slice(d.l, d.l+10)).replace(/[\u0000\r\n].*$/g,"");
		d.l += 11;
		field.type = String.fromCharCode(d.read_shift(1));
		field.offset = d.read_shift(4);
		field.len = d.read_shift(1);
		field.dec = d.read_shift(1);
		if(field.name.length) fields.push(field);
		d.l += 14;
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
	if(d.read_shift(1) !== 0x0D) throw new Error("DBF Terminator not found " + d.l + " " + d[d.l]);
	d.l = fpos;
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
					//out[R][C] = new Date(); // FIXME!!!
					//break;
				case 'Y': out[R][C] = dd.read(4,'i')/1e4; break;
				case '0':
					if(fields[C].name === '_NullFlags') break;
					/* falls through */
				default: throw new Error("DBF Unsupported data type " + fields[C].type);
			}
		}
	}
	if(d.l < d.length && d[d.l++] != 0x1A) throw new Error("DBF EOF Marker missing " + (d.l-1) + " of " + d.length + " " + d[d.l-1].toString(16));
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
