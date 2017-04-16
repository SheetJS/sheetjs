var WK_ = (function() {
	function lotushopper(data, cb/*:RecordHopperCB*/, opts/*:any*/) {
		if(!data) return;
		prep_blob(data, data.l || 0);
		var Enum = opts.Enum || WK1Enum;
		while(data.l < data.length) {
			var RT = data.read_shift(2);
			var R = Enum[RT] || Enum[0xFF];
			var length = data.read_shift(2);
			var tgt = data.l + length;
			var d = R.f(data, length, opts);
			data.l = tgt;
			if(cb(d, R.n, RT)) return;
		}
	}

	function lotus_to_workbook(d/*:RawData*/, opts) {
		switch(opts.type) {
			case 'base64': return lotus_to_workbook_buf(s2a(Base64.decode(d)), opts);
			case 'binary': return lotus_to_workbook_buf(s2a(d), opts);
			case 'buffer':
			case 'array': return lotus_to_workbook_buf(d, opts);
		}
		throw "Unsupported type " + opts.type;
	}

	function lotus_to_workbook_buf(d,opts)/*:Workbook*/ {
		if(!d) return d;
		var o = opts || {};
		if(DENSE != null && o.dense == null) o.dense = DENSE;
		var s/*:Worksheet*/ = ((o.dense ? [] : {})/*:any*/), n = "Sheet1", sidx = 0;
		var sheets = {}, snames = [n];

		var refguess = {s: {r:0, c:0}, e: {r:0, c:0} };

		if(d[2] == 0x02) o.Enum = WK1Enum;
		else if(d[2] == 0x1a) o.Enum = WK3Enum;
		else if(d[2] == 0x0e) { o.Enum = WK3Enum; o.qpro = true; d.l = 0; }
		else throw new Error("Unrecognized LOTUS BOF " + d[2]);
		lotushopper(d, function(val, Rn, RT) {
			if(d[2] == 0x02) switch(RT) {
				case 0x00:
					o.vers = val;
					if(val >= 0x1000) o.qpro = true;
					break;
				case 0x06: refguess = val; break; /* RANGE */
				case 0x0F: /* LABEL */
					if(!o.qpro) val[1].v = val[1].v.substr(1);
					/* falls through */
				case 0x0D: /* INTEGER */
				case 0x0E: /* NUMBER */
				case 0x10: /* FORMULA */
				case 0x33: /* STRING */
					if(o.dense) {
						if(!s[val[0].r]) s[val[0].r] = [];
						s[val[0].r][val[0].c] = val[1];
					} else s[encode_cell(val[0])] = val[1];
					/* TODO: FORMAT */
					break;
			} else switch(RT) {
				case 0x16: /* LABEL16 */
					val[1].v = val[1].v.substr(1);
					/* falls through */
				case 0x17: /* NUMBER17 */
				case 0x18: /* NUMBER18 */
				case 0x19: /* FORMULA19 */
				case 0x25: /* NUMBER25 */
				case 0x27: /* NUMBER27 */
				case 0x28: /* FORMULA28 */
					if(val[3] > sidx) {
						s["!ref"] = encode_range(refguess);
						sheets[n] = s;
						s = (o.dense ? [] : {});
						refguess = {s: {r:0, c:0}, e: {r:0, c:0} };
						sidx = val[3]; n = "Sheet" + (sidx + 1);
						snames.push(n);
					}
					s[encode_cell(val[0])] = val[1];
					if(refguess.e.c < val[0].c) refguess.e.c = val[0].c;
					if(refguess.e.r < val[0].r) refguess.e.r = val[0].r;
					break;
				default: break;
			}
		}, o);

		s["!ref"] = encode_range(refguess);
		sheets[n] = s;
		return { SheetNames: snames, Sheets:sheets };
	}

	function parse_RANGE(blob, length) {
		var o = {s:{c:0,r:0},e:{c:0,r:0}};
		o.s.c = blob.read_shift(2);
		o.s.r = blob.read_shift(2);
		o.e.c = blob.read_shift(2);
		o.e.r = blob.read_shift(2);
		if(o.s.c == 0xFFFF) o.s.c = o.e.c = o.s.r = o.e.r = 0;
		return o;
	}

	function parse_cell(blob, length, opts) {
		var o = [{c:0,r:0}, {t:'n',v:0}, 0];
		if(opts.qpro && opts.vers != 0x5120) {
			o[0].c = blob.read_shift(1);
			blob.l++;
			o[0].r = blob.read_shift(2);
			blob.l+=2;
		} else {
			o[2] = blob.read_shift(1);
			o[0].c = blob.read_shift(2); o[0].r = blob.read_shift(2);
		}
		return o;
	}

	function parse_LABEL(blob, length, opts) {
		var tgt = blob.l + length;
		var o = parse_cell(blob, length, opts);
		o[1].t = 's';
		if(opts.vers == 0x5120) {
			blob.l++;
			var len = blob.read_shift(1);
			o[1].v = blob.read_shift(len, 'utf8');
			return o;
		}
		if(opts.qpro) blob.l++;
		o[1].v = blob.read_shift(tgt - blob.l, 'cstr');
		return o;
	}

	function parse_INTEGER(blob, length, opts) {
		var o = parse_cell(blob, length, opts);
		o[1].v = blob.read_shift(2, 'i');
		return o;
	}

	function parse_NUMBER(blob, length, opts) {
		var o = parse_cell(blob, length, opts);
		o[1].v = blob.read_shift(8, 'f');
		return o;
	}

	function parse_FORMULA(blob, length, opts) {
		var tgt = blob.l + length;
		var o = parse_cell(blob, length, opts);
		/* TODO: formula */
		o[1].v = blob.read_shift(8, 'f');
		if(opts.qpro) blob.l = tgt;
		else {
			var flen = blob.read_shift(2);
			blob.l += flen;
		}
		return o;
	}

	function parse_cell_3(blob, length) {
		var o = [{c:0,r:0}, {t:'n',v:0}, 0];
		o[0].r = blob.read_shift(2); o[3] = blob[blob.l++]; o[0].c = blob[blob.l++];
		return o;
	}

	function parse_LABEL_16(blob, length) {
		var o = parse_cell_3(blob, length);
		o[1].t = 's';
		o[1].v = blob.read_shift(length - 4, 'cstr');
		return o;
	}

	function parse_NUMBER_18(blob, length) {
		var o = parse_cell_3(blob, length);
		o[1].v = blob.read_shift(2);
		var v = o[1].v >> 1;
		/* TODO: figure out all of the corner cases */
		if(o[1].v & 0x1) {
			switch(v & 0x07) {
				case 1: v = (v >> 3) * 500; break;
				case 2: v = (v >> 3) / 20; break;
				case 4: v = (v >> 3) / 2000; break;
				case 6: v = (v >> 3) / 16; break;
				case 7: v = (v >> 3) / 64; break;
				default: throw "unknown NUMBER_18 encoding " + (v & 0x07);
			}
		}
		o[1].v = v;
		return o;
	}

	function parse_NUMBER_17(blob, length) {
		var o = parse_cell_3(blob, length);
		var v1 = blob.read_shift(4);
		var v2 = blob.read_shift(4);
		var e = blob.read_shift(2);
		if(e == 0xFFFF) { o[1].v = 0; return o; }
		var s = e & 0x8000; e = (e&0x7FFF) - 16446;
		o[1].v = ((e > 0 ? (v2 << e) : (v2 >>> -e)) + (e > -32 ? (v1 << (e + 32)) : (v1 >>> -(e + 32))));
		return o;
	}

	function parse_FORMULA_19(blob, length) {
		var o = parse_NUMBER_17(blob, 14);
		blob.l += length - 14; /* TODO: formula */
		return o;
	}

	function parse_NUMBER_25(blob, length) {
		var o = parse_cell_3(blob, length);
		var v1 = blob.read_shift(4);
		o[1].v = v1 >> 6;
		return o;
	}

	function parse_NUMBER_27(blob, length) {
		var o = parse_cell_3(blob, length);
		var v1 = blob.read_shift(8,'f');
		o[1].v = v1;
		return o;
	}

	function parse_FORMULA_28(blob, length) {
		var o = parse_NUMBER_27(blob, 14);
		blob.l += length - 10; /* TODO: formula */
		return o;
	}

	var WK1Enum = {
		/*::[*/0x0000/*::]*/: { n:"BOF", f:parseuint16 },
		/*::[*/0x0001/*::]*/: { n:"EOF", f:parsenoop },
		/*::[*/0x0002/*::]*/: { n: "CALCMODE", f:parsenoop },
		/*::[*/0x0003/*::]*/: { n:"CALCORDER", f:parsenoop },
		/*::[*/0x0004/*::]*/: { n:"SPLIT", f:parsenoop },
		/*::[*/0x0005/*::]*/: { n:"SYNC", f:parsenoop },
		/*::[*/0x0006/*::]*/: { n:"RANGE", f:parse_RANGE },
		/*::[*/0x0007/*::]*/: { n:"WINDOW1", f:parsenoop },
		/*::[*/0x0008/*::]*/: { n:"COLW1", f:parsenoop },
		/*::[*/0x0009/*::]*/: { n:"WINTWO", f:parsenoop },
		/*::[*/0x000A/*::]*/: { n:"COLW2", f:parsenoop },
		/*::[*/0x000B/*::]*/: { n:"NAME", f:parsenoop },
		/*::[*/0x000C/*::]*/: { n:"BLANK", f:parsenoop },
		/*::[*/0x000D/*::]*/: { n:"INTEGER", f:parse_INTEGER },
		/*::[*/0x000E/*::]*/: { n:"NUMBER", f:parse_NUMBER },
		/*::[*/0x000F/*::]*/: { n:"LABEL", f:parse_LABEL },
		/*::[*/0x0010/*::]*/: { n:"FORMULA", f:parse_FORMULA },
		/*::[*/0x0018/*::]*/: { n:"TABLE", f:parsenoop },
		/*::[*/0x0019/*::]*/: { n:"ORANGE", f:parsenoop },
		/*::[*/0x001A/*::]*/: { n:"PRANGE", f:parsenoop },
		/*::[*/0x001B/*::]*/: { n:"SRANGE", f:parsenoop },
		/*::[*/0x001C/*::]*/: { n:"FRANGE", f:parsenoop },
		/*::[*/0x001D/*::]*/: { n:"KRANGE1", f:parsenoop },
		/*::[*/0x0020/*::]*/: { n:"HRANGE", f:parsenoop },
		/*::[*/0x0023/*::]*/: { n:"KRANGE2", f:parsenoop },
		/*::[*/0x0024/*::]*/: { n:"PROTEC", f:parsenoop },
		/*::[*/0x0025/*::]*/: { n:"FOOTER", f:parsenoop },
		/*::[*/0x0026/*::]*/: { n:"HEADER", f:parsenoop },
		/*::[*/0x0027/*::]*/: { n:"SETUP", f:parsenoop },
		/*::[*/0x0028/*::]*/: { n:"MARGINS", f:parsenoop },
		/*::[*/0x0029/*::]*/: { n:"LABELFMT", f:parsenoop },
		/*::[*/0x002A/*::]*/: { n:"TITLES", f:parsenoop },
		/*::[*/0x002B/*::]*/: { n:"SHEETJS", f:parsenoop },
		/*::[*/0x002D/*::]*/: { n:"GRAPH", f:parsenoop },
		/*::[*/0x002E/*::]*/: { n:"NGRAPH", f:parsenoop },
		/*::[*/0x002F/*::]*/: { n:"CALCCOUNT", f:parsenoop },
		/*::[*/0x0030/*::]*/: { n:"UNFORMATTED", f:parsenoop },
		/*::[*/0x0031/*::]*/: { n:"CURSORW12", f:parsenoop },
		/*::[*/0x0032/*::]*/: { n:"WINDOW", f:parsenoop },
		/*::[*/0x0033/*::]*/: { n:"STRING", f:parse_LABEL },
		/*::[*/0x0037/*::]*/: { n:"PASSWORD", f:parsenoop },
		/*::[*/0x0038/*::]*/: { n:"LOCKED", f:parsenoop },
		/*::[*/0x003C/*::]*/: { n:"QUERY", f:parsenoop },
		/*::[*/0x003D/*::]*/: { n:"QUERYNAME", f:parsenoop },
		/*::[*/0x003E/*::]*/: { n:"PRINT", f:parsenoop },
		/*::[*/0x003F/*::]*/: { n:"PRINTNAME", f:parsenoop },
		/*::[*/0x0040/*::]*/: { n:"GRAPH2", f:parsenoop },
		/*::[*/0x0041/*::]*/: { n:"GRAPHNAME", f:parsenoop },
		/*::[*/0x0042/*::]*/: { n:"ZOOM", f:parsenoop },
		/*::[*/0x0043/*::]*/: { n:"SYMSPLIT", f:parsenoop },
		/*::[*/0x0044/*::]*/: { n:"NSROWS", f:parsenoop },
		/*::[*/0x0045/*::]*/: { n:"NSCOLS", f:parsenoop },
		/*::[*/0x0046/*::]*/: { n:"RULER", f:parsenoop },
		/*::[*/0x0047/*::]*/: { n:"NNAME", f:parsenoop },
		/*::[*/0x0048/*::]*/: { n:"ACOMM", f:parsenoop },
		/*::[*/0x0049/*::]*/: { n:"AMACRO", f:parsenoop },
		/*::[*/0x004A/*::]*/: { n:"PARSE", f:parsenoop },
		/*::[*/0x00FF/*::]*/: { n:"", f:parsenoop }
	};

	var WK3Enum = {
		/*::[*/0x0000/*::]*/: { n:"BOF", f:parsenoop },
		/*::[*/0x0001/*::]*/: { n:"EOF", f:parsenoop },
		/*::[*/0x0003/*::]*/: { n:"??", f:parsenoop },
		/*::[*/0x0004/*::]*/: { n:"??", f:parsenoop },
		/*::[*/0x0005/*::]*/: { n:"??", f:parsenoop },
		/*::[*/0x0006/*::]*/: { n:"??", f:parsenoop },
		/*::[*/0x0007/*::]*/: { n:"??", f:parsenoop },
		/*::[*/0x0009/*::]*/: { n:"??", f:parsenoop },
		/*::[*/0x000a/*::]*/: { n:"??", f:parsenoop },
		/*::[*/0x000b/*::]*/: { n:"??", f:parsenoop },
		/*::[*/0x000c/*::]*/: { n:"??", f:parsenoop },
		/*::[*/0x000e/*::]*/: { n:"??", f:parsenoop },
		/*::[*/0x000f/*::]*/: { n:"??", f:parsenoop },
		/*::[*/0x0010/*::]*/: { n:"??", f:parsenoop },
		/*::[*/0x0011/*::]*/: { n:"??", f:parsenoop },
		/*::[*/0x0012/*::]*/: { n:"??", f:parsenoop },
		/*::[*/0x0013/*::]*/: { n:"??", f:parsenoop },
		/*::[*/0x0015/*::]*/: { n:"??", f:parsenoop },
		/*::[*/0x0016/*::]*/: { n:"LABEL16", f:parse_LABEL_16},
		/*::[*/0x0017/*::]*/: { n:"NUMBER17", f:parse_NUMBER_17 },
		/*::[*/0x0018/*::]*/: { n:"NUMBER18", f:parse_NUMBER_18 },
		/*::[*/0x0019/*::]*/: { n:"FORMULA19", f:parse_FORMULA_19},
		/*::[*/0x001a/*::]*/: { n:"??", f:parsenoop },
		/*::[*/0x001b/*::]*/: { n:"??", f:parsenoop },
		/*::[*/0x001c/*::]*/: { n:"??", f:parsenoop },
		/*::[*/0x001d/*::]*/: { n:"??", f:parsenoop },
		/*::[*/0x001e/*::]*/: { n:"??", f:parsenoop },
		/*::[*/0x001f/*::]*/: { n:"??", f:parsenoop },
		/*::[*/0x0021/*::]*/: { n:"??", f:parsenoop },
		/*::[*/0x0025/*::]*/: { n:"NUMBER25", f:parse_NUMBER_25 },
		/*::[*/0x0027/*::]*/: { n:"NUMBER27", f:parse_NUMBER_27 },
		/*::[*/0x0028/*::]*/: { n:"FORMULA28", f:parse_FORMULA_28 },
		/*::[*/0x00FF/*::]*/: { n:"", f:parsenoop }
	};
	return {
		to_workbook: lotus_to_workbook
	};
})();
