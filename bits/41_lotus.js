var WK_ = /*#__PURE__*/ (function() {
	function lotushopper(data, cb/*:RecordHopperCB*/, opts/*:any*/) {
		if(!data) return;
		prep_blob(data, data.l || 0);
		var Enum = opts.Enum || WK1Enum;
		while(data.l < data.length) {
			var RT = data.read_shift(2);
			var R = Enum[RT] || Enum[0xFFFF];
			var length = data.read_shift(2);
			var tgt = data.l + length;
			var d = R.f && R.f(data, length, opts);
			data.l = tgt;
			if(cb(d, R, RT)) return;
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

	function lotus_to_workbook_buf(d, opts)/*:Workbook*/ {
		if(!d) return d;
		var o = opts || {};
		if(DENSE != null && o.dense == null) o.dense = DENSE;
		var s/*:Worksheet*/ = ((o.dense ? [] : {})/*:any*/), n = "Sheet1", sidx = 0;
		var sheets = {}, snames = [n];

		var refguess = {s: {r:0, c:0}, e: {r:0, c:0} };
		var sheetRows = o.sheetRows || 0;

		if(d[2] == 0x02) {
			o.Enum = WK1Enum;
			lotushopper(d, function(val, R, RT) { switch(RT) {
				case 0x00:
					o.vers = val;
					if(val >= 0x1000) o.qpro = true;
					break;
				case 0x06: refguess = val; break; /* RANGE */
				case 0x0F: /* LABEL */
					if(!o.qpro) val[1].v = val[1].v.slice(1);
					/* falls through */
				case 0x0D: /* INTEGER */
				case 0x0E: /* NUMBER */
				case 0x10: /* FORMULA */
				case 0x33: /* STRING */
					/* TODO: actual translation of the format code */
					if(RT == 0x0E && (val[2] & 0x70) == 0x70 && (val[2] & 0x0F) > 1 && (val[2] & 0x0F) < 15) {
						val[1].z = o.dateNF || SSF._table[14];
						if(o.cellDates) { val[1].t = 'd'; val[1].v = numdate(val[1].v); }
					}
					if(o.dense) {
						if(!s[val[0].r]) s[val[0].r] = [];
						s[val[0].r][val[0].c] = val[1];
					} else s[encode_cell(val[0])] = val[1];
					break;
			}}, o);
		} else if(d[2] == 0x1A || d[2] == 0x0E) {
			o.Enum = WK3Enum;
			if(d[2] == 0x0E) { o.qpro = true; d.l = 0; }
			lotushopper(d, function(val, R, RT) { switch(RT) {
				case 0x16: /* LABEL16 */
					val[1].v = val[1].v.slice(1);
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
					if(sheetRows > 0 && val[0].r >= sheetRows) break;
					if(o.dense) {
						if(!s[val[0].r]) s[val[0].r] = [];
						s[val[0].r][val[0].c] = val[1];
					} else s[encode_cell(val[0])] = val[1];
					if(refguess.e.c < val[0].c) refguess.e.c = val[0].c;
					if(refguess.e.r < val[0].r) refguess.e.r = val[0].r;
					break;
				default: break;
			}}, o);
		} else throw new Error("Unrecognized LOTUS BOF " + d[2]);

		s["!ref"] = encode_range(refguess);
		sheets[n] = s;
		return { SheetNames: snames, Sheets:sheets };
	}

	function sheet_to_wk1(ws/*:Worksheet*/, opts/*:WriteOpts*/) {
		var o = opts || {};
		if(+o.codepage >= 0) set_cp(+o.codepage);
		if(o.type == "string") throw new Error("Cannot write DBF to JS string");
		var ba = buf_array();
		var range = safe_decode_range(ws["!ref"]);
		var dense = Array.isArray(ws);
		var cols = [];

		write_biff_rec(ba, 0x00, write_BOF_WK1(0x0406));
		write_biff_rec(ba, 0x06, write_RANGE(range));
		for(var R = range.s.r; R <= range.e.r; ++R) {
			var rr = encode_row(R);
			for(var C = range.s.c; C <= range.e.c; ++C) {
				if(R === range.s.r) cols[C] = encode_col(C);
				var ref = cols[C] + rr;
				var cell = dense ? (ws[R]||[])[C] : ws[ref];
				if(!cell || cell.t == "z") continue;
				/* write cell */
				if(cell.t == "n") {
					if((cell.v|0)==cell.v && cell.v >= -32768 && cell.v <= 32767) write_biff_rec(ba, 0x0d, write_INTEGER(R, C, cell.v));
					else write_biff_rec(ba, 0x0e, write_NUMBER(R, C, cell.v));
				} else {
					var str = format_cell(cell);
					write_biff_rec(ba, 0x0F, write_LABEL(R, C, str.slice(0, 239)));
				}
			}
		}

		write_biff_rec(ba, 0x01);
		return ba.end();
	}

	function write_BOF_WK1(v/*:number*/) {
		var out = new_buf(2);
		out.write_shift(2, v);
		return out;
	}

	function parse_RANGE(blob) {
		var o = {s:{c:0,r:0},e:{c:0,r:0}};
		o.s.c = blob.read_shift(2);
		o.s.r = blob.read_shift(2);
		o.e.c = blob.read_shift(2);
		o.e.r = blob.read_shift(2);
		if(o.s.c == 0xFFFF) o.s.c = o.e.c = o.s.r = o.e.r = 0;
		return o;
	}
	function write_RANGE(range) {
		var out = new_buf(8);
		out.write_shift(2, range.s.c);
		out.write_shift(2, range.s.r);
		out.write_shift(2, range.e.c);
		out.write_shift(2, range.e.r);
		return out;
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
	function write_LABEL(R, C, s) {
		/* TODO: encoding */
		var o = new_buf(7 + s.length);
		o.write_shift(1, 0xFF);
		o.write_shift(2, C);
		o.write_shift(2, R);
		o.write_shift(1, 0x27); // ??
		for(var i = 0; i < o.length; ++i) {
			var cc = s.charCodeAt(i);
			o.write_shift(1, cc >= 0x80 ? 0x5F : cc);
		}
		o.write_shift(1, 0);
		return o;
	}

	function parse_INTEGER(blob, length, opts) {
		var o = parse_cell(blob, length, opts);
		o[1].v = blob.read_shift(2, 'i');
		return o;
	}
	function write_INTEGER(R, C, v) {
		var o = new_buf(7);
		o.write_shift(1, 0xFF);
		o.write_shift(2, C);
		o.write_shift(2, R);
		o.write_shift(2, v, 'i');
		return o;
	}

	function parse_NUMBER(blob, length, opts) {
		var o = parse_cell(blob, length, opts);
		o[1].v = blob.read_shift(8, 'f');
		return o;
	}
	function write_NUMBER(R, C, v) {
		var o = new_buf(13);
		o.write_shift(1, 0xFF);
		o.write_shift(2, C);
		o.write_shift(2, R);
		o.write_shift(8, v, 'f');
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

	function parse_cell_3(blob/*::, length*/) {
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
		if(o[1].v & 0x1) {
			switch(v & 0x07) {
				case 0: v = (v >> 3) * 5000; break;
				case 1: v = (v >> 3) * 500; break;
				case 2: v = (v >> 3) / 20; break;
				case 3: v = (v >> 3) / 200; break;
				case 4: v = (v >> 3) / 2000; break;
				case 5: v = (v >> 3) / 20000; break;
				case 6: v = (v >> 3) / 16; break;
				case 7: v = (v >> 3) / 64; break;
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
		if(e == 0xFFFF) {
			if(v1 === 0 && v2 === 0xC0000000) { o[1].t = "e"; o[1].v = 0x0F; } // ERR -> #VALUE!
			else if(v1 === 0 && v2 === 0xD0000000) { o[1].t = "e"; o[1].v = 0x2A; } // NA -> #N/A
			else o[1].v = 0;
			return o;
		}
		var s = e & 0x8000; e = (e&0x7FFF) - 16446;
		o[1].v = (1 - s*2) * ( /*(e > 0 ? (v2 << e) : (v2 >>> -e))*/ v2 * Math.pow(2, e+32) + v1 * Math.pow(2, e));
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
		/*::[*/0x0001/*::]*/: { n:"EOF" },
		/*::[*/0x0002/*::]*/: { n:"CALCMODE" },
		/*::[*/0x0003/*::]*/: { n:"CALCORDER" },
		/*::[*/0x0004/*::]*/: { n:"SPLIT" },
		/*::[*/0x0005/*::]*/: { n:"SYNC" },
		/*::[*/0x0006/*::]*/: { n:"RANGE", f:parse_RANGE },
		/*::[*/0x0007/*::]*/: { n:"WINDOW1" },
		/*::[*/0x0008/*::]*/: { n:"COLW1" },
		/*::[*/0x0009/*::]*/: { n:"WINTWO" },
		/*::[*/0x000A/*::]*/: { n:"COLW2" },
		/*::[*/0x000B/*::]*/: { n:"NAME" },
		/*::[*/0x000C/*::]*/: { n:"BLANK" },
		/*::[*/0x000D/*::]*/: { n:"INTEGER", f:parse_INTEGER },
		/*::[*/0x000E/*::]*/: { n:"NUMBER", f:parse_NUMBER },
		/*::[*/0x000F/*::]*/: { n:"LABEL", f:parse_LABEL },
		/*::[*/0x0010/*::]*/: { n:"FORMULA", f:parse_FORMULA },
		/*::[*/0x0018/*::]*/: { n:"TABLE" },
		/*::[*/0x0019/*::]*/: { n:"ORANGE" },
		/*::[*/0x001A/*::]*/: { n:"PRANGE" },
		/*::[*/0x001B/*::]*/: { n:"SRANGE" },
		/*::[*/0x001C/*::]*/: { n:"FRANGE" },
		/*::[*/0x001D/*::]*/: { n:"KRANGE1" },
		/*::[*/0x0020/*::]*/: { n:"HRANGE" },
		/*::[*/0x0023/*::]*/: { n:"KRANGE2" },
		/*::[*/0x0024/*::]*/: { n:"PROTEC" },
		/*::[*/0x0025/*::]*/: { n:"FOOTER" },
		/*::[*/0x0026/*::]*/: { n:"HEADER" },
		/*::[*/0x0027/*::]*/: { n:"SETUP" },
		/*::[*/0x0028/*::]*/: { n:"MARGINS" },
		/*::[*/0x0029/*::]*/: { n:"LABELFMT" },
		/*::[*/0x002A/*::]*/: { n:"TITLES" },
		/*::[*/0x002B/*::]*/: { n:"SHEETJS" },
		/*::[*/0x002D/*::]*/: { n:"GRAPH" },
		/*::[*/0x002E/*::]*/: { n:"NGRAPH" },
		/*::[*/0x002F/*::]*/: { n:"CALCCOUNT" },
		/*::[*/0x0030/*::]*/: { n:"UNFORMATTED" },
		/*::[*/0x0031/*::]*/: { n:"CURSORW12" },
		/*::[*/0x0032/*::]*/: { n:"WINDOW" },
		/*::[*/0x0033/*::]*/: { n:"STRING", f:parse_LABEL },
		/*::[*/0x0037/*::]*/: { n:"PASSWORD" },
		/*::[*/0x0038/*::]*/: { n:"LOCKED" },
		/*::[*/0x003C/*::]*/: { n:"QUERY" },
		/*::[*/0x003D/*::]*/: { n:"QUERYNAME" },
		/*::[*/0x003E/*::]*/: { n:"PRINT" },
		/*::[*/0x003F/*::]*/: { n:"PRINTNAME" },
		/*::[*/0x0040/*::]*/: { n:"GRAPH2" },
		/*::[*/0x0041/*::]*/: { n:"GRAPHNAME" },
		/*::[*/0x0042/*::]*/: { n:"ZOOM" },
		/*::[*/0x0043/*::]*/: { n:"SYMSPLIT" },
		/*::[*/0x0044/*::]*/: { n:"NSROWS" },
		/*::[*/0x0045/*::]*/: { n:"NSCOLS" },
		/*::[*/0x0046/*::]*/: { n:"RULER" },
		/*::[*/0x0047/*::]*/: { n:"NNAME" },
		/*::[*/0x0048/*::]*/: { n:"ACOMM" },
		/*::[*/0x0049/*::]*/: { n:"AMACRO" },
		/*::[*/0x004A/*::]*/: { n:"PARSE" },
		/*::[*/0x0066/*::]*/: { n:"PRANGES??" },
		/*::[*/0x0067/*::]*/: { n:"RRANGES??" },
		/*::[*/0x0068/*::]*/: { n:"FNAME??" },
		/*::[*/0x0069/*::]*/: { n:"MRANGES??" },
		/*::[*/0xFFFF/*::]*/: { n:"" }
	};

	var WK3Enum = {
		/*::[*/0x0000/*::]*/: { n:"BOF" },
		/*::[*/0x0001/*::]*/: { n:"EOF" },
		/*::[*/0x0002/*::]*/: { n:"PASSWORD" },
		/*::[*/0x0003/*::]*/: { n:"CALCSET" },
		/*::[*/0x0004/*::]*/: { n:"WINDOWSET" },
		/*::[*/0x0005/*::]*/: { n:"SHEETCELLPTR" },
		/*::[*/0x0006/*::]*/: { n:"SHEETLAYOUT" },
		/*::[*/0x0007/*::]*/: { n:"COLUMNWIDTH" },
		/*::[*/0x0008/*::]*/: { n:"HIDDENCOLUMN" },
		/*::[*/0x0009/*::]*/: { n:"USERRANGE" },
		/*::[*/0x000A/*::]*/: { n:"SYSTEMRANGE" },
		/*::[*/0x000B/*::]*/: { n:"ZEROFORCE" },
		/*::[*/0x000C/*::]*/: { n:"SORTKEYDIR" },
		/*::[*/0x000D/*::]*/: { n:"FILESEAL" },
		/*::[*/0x000E/*::]*/: { n:"DATAFILLNUMS" },
		/*::[*/0x000F/*::]*/: { n:"PRINTMAIN" },
		/*::[*/0x0010/*::]*/: { n:"PRINTSTRING" },
		/*::[*/0x0011/*::]*/: { n:"GRAPHMAIN" },
		/*::[*/0x0012/*::]*/: { n:"GRAPHSTRING" },
		/*::[*/0x0013/*::]*/: { n:"??" },
		/*::[*/0x0014/*::]*/: { n:"ERRCELL" },
		/*::[*/0x0015/*::]*/: { n:"NACELL" },
		/*::[*/0x0016/*::]*/: { n:"LABEL16", f:parse_LABEL_16},
		/*::[*/0x0017/*::]*/: { n:"NUMBER17", f:parse_NUMBER_17 },
		/*::[*/0x0018/*::]*/: { n:"NUMBER18", f:parse_NUMBER_18 },
		/*::[*/0x0019/*::]*/: { n:"FORMULA19", f:parse_FORMULA_19},
		/*::[*/0x001A/*::]*/: { n:"FORMULA1A" },
		/*::[*/0x001B/*::]*/: { n:"XFORMAT" },
		/*::[*/0x001C/*::]*/: { n:"DTLABELMISC" },
		/*::[*/0x001D/*::]*/: { n:"DTLABELCELL" },
		/*::[*/0x001E/*::]*/: { n:"GRAPHWINDOW" },
		/*::[*/0x001F/*::]*/: { n:"CPA" },
		/*::[*/0x0020/*::]*/: { n:"LPLAUTO" },
		/*::[*/0x0021/*::]*/: { n:"QUERY" },
		/*::[*/0x0022/*::]*/: { n:"HIDDENSHEET" },
		/*::[*/0x0023/*::]*/: { n:"??" },
		/*::[*/0x0025/*::]*/: { n:"NUMBER25", f:parse_NUMBER_25 },
		/*::[*/0x0026/*::]*/: { n:"??" },
		/*::[*/0x0027/*::]*/: { n:"NUMBER27", f:parse_NUMBER_27 },
		/*::[*/0x0028/*::]*/: { n:"FORMULA28", f:parse_FORMULA_28 },
		/*::[*/0x008E/*::]*/: { n:"??" },
		/*::[*/0x0093/*::]*/: { n:"??" },
		/*::[*/0x0096/*::]*/: { n:"??" },
		/*::[*/0x0097/*::]*/: { n:"??" },
		/*::[*/0x0098/*::]*/: { n:"??" },
		/*::[*/0x0099/*::]*/: { n:"??" },
		/*::[*/0x009A/*::]*/: { n:"??" },
		/*::[*/0x009B/*::]*/: { n:"??" },
		/*::[*/0x009C/*::]*/: { n:"??" },
		/*::[*/0x00A3/*::]*/: { n:"??" },
		/*::[*/0x00AE/*::]*/: { n:"??" },
		/*::[*/0x00AF/*::]*/: { n:"??" },
		/*::[*/0x00B0/*::]*/: { n:"??" },
		/*::[*/0x00B1/*::]*/: { n:"??" },
		/*::[*/0x00B8/*::]*/: { n:"??" },
		/*::[*/0x00B9/*::]*/: { n:"??" },
		/*::[*/0x00BA/*::]*/: { n:"??" },
		/*::[*/0x00BB/*::]*/: { n:"??" },
		/*::[*/0x00BC/*::]*/: { n:"??" },
		/*::[*/0x00C3/*::]*/: { n:"??" },
		/*::[*/0x00C9/*::]*/: { n:"??" },
		/*::[*/0x00CD/*::]*/: { n:"??" },
		/*::[*/0x00CE/*::]*/: { n:"??" },
		/*::[*/0x00CF/*::]*/: { n:"??" },
		/*::[*/0x00D0/*::]*/: { n:"??" },
		/*::[*/0x0100/*::]*/: { n:"??" },
		/*::[*/0x0103/*::]*/: { n:"??" },
		/*::[*/0x0104/*::]*/: { n:"??" },
		/*::[*/0x0105/*::]*/: { n:"??" },
		/*::[*/0x0106/*::]*/: { n:"??" },
		/*::[*/0x0107/*::]*/: { n:"??" },
		/*::[*/0x0109/*::]*/: { n:"??" },
		/*::[*/0x010A/*::]*/: { n:"??" },
		/*::[*/0x010B/*::]*/: { n:"??" },
		/*::[*/0x010C/*::]*/: { n:"??" },
		/*::[*/0x010E/*::]*/: { n:"??" },
		/*::[*/0x010F/*::]*/: { n:"??" },
		/*::[*/0x0180/*::]*/: { n:"??" },
		/*::[*/0x0185/*::]*/: { n:"??" },
		/*::[*/0x0186/*::]*/: { n:"??" },
		/*::[*/0x0189/*::]*/: { n:"??" },
		/*::[*/0x018C/*::]*/: { n:"??" },
		/*::[*/0x0200/*::]*/: { n:"??" },
		/*::[*/0x0202/*::]*/: { n:"??" },
		/*::[*/0x0201/*::]*/: { n:"??" },
		/*::[*/0x0204/*::]*/: { n:"??" },
		/*::[*/0x0205/*::]*/: { n:"??" },
		/*::[*/0x0280/*::]*/: { n:"??" },
		/*::[*/0x0281/*::]*/: { n:"??" },
		/*::[*/0x0282/*::]*/: { n:"??" },
		/*::[*/0x0283/*::]*/: { n:"??" },
		/*::[*/0x0284/*::]*/: { n:"??" },
		/*::[*/0x0285/*::]*/: { n:"??" },
		/*::[*/0x0286/*::]*/: { n:"??" },
		/*::[*/0x0287/*::]*/: { n:"??" },
		/*::[*/0x0288/*::]*/: { n:"??" },
		/*::[*/0x0292/*::]*/: { n:"??" },
		/*::[*/0x0293/*::]*/: { n:"??" },
		/*::[*/0x0294/*::]*/: { n:"??" },
		/*::[*/0x0295/*::]*/: { n:"??" },
		/*::[*/0x0296/*::]*/: { n:"??" },
		/*::[*/0x0299/*::]*/: { n:"??" },
		/*::[*/0x029A/*::]*/: { n:"??" },
		/*::[*/0x0300/*::]*/: { n:"??" },
		/*::[*/0x0304/*::]*/: { n:"??" },
		/*::[*/0x0640/*::]*/: { n:"??" },
		/*::[*/0x0642/*::]*/: { n:"??" },
		/*::[*/0x0701/*::]*/: { n:"??" },
		/*::[*/0x0702/*::]*/: { n:"??" },
		/*::[*/0x0703/*::]*/: { n:"??" },
		/*::[*/0x0704/*::]*/: { n:"??" },
		/*::[*/0x0780/*::]*/: { n:"??" },
		/*::[*/0x0800/*::]*/: { n:"??" },
		/*::[*/0x0801/*::]*/: { n:"??" },
		/*::[*/0x0804/*::]*/: { n:"??" },
		/*::[*/0x0A80/*::]*/: { n:"??" },
		/*::[*/0x2AF6/*::]*/: { n:"??" },
		/*::[*/0x3231/*::]*/: { n:"??" },
		/*::[*/0x6E49/*::]*/: { n:"??" },
		/*::[*/0x6F44/*::]*/: { n:"??" },
		/*::[*/0xFFFF/*::]*/: { n:"" }
	};
	return {
		sheet_to_wk1: sheet_to_wk1,
		to_workbook: lotus_to_workbook
	};
})();
