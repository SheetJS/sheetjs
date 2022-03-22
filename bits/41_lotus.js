var WK_ = /*#__PURE__*/(function() {
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
			case 'base64': return lotus_to_workbook_buf(s2a(Base64_decode(d)), opts);
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
		var s/*:Worksheet*/ = ((o.dense ? [] : {})/*:any*/), n = "Sheet1", next_n = "", sidx = 0;
		var sheets = {}, snames = [], realnames = [];

		var refguess = {s: {r:0, c:0}, e: {r:0, c:0} };
		var sheetRows = o.sheetRows || 0;

		if(d[2] == 0x00) {
			if(d[3] == 0x08 || d[3] == 0x09) {
				if(d.length >= 16 && d[14] == 0x05 && d[15] === 0x6c) throw new Error("Unsupported Works 3 for Mac file");
			}
		}

		if(d[2] == 0x02) {
			o.Enum = WK1Enum;
			lotushopper(d, function(val, R, RT) { switch(RT) {
				case 0x00: /* BOF */
					o.vers = val;
					if(val >= 0x1000) o.qpro = true;
					break;
				case 0x06: refguess = val; break; /* RANGE */
				case 0xCC: if(val) next_n = val; break; /* SHEETNAMECS */
				case 0xDE: next_n = val; break; /* SHEETNAMELP */
				case 0x0F: /* LABEL */
				case 0x33: /* STRING */
					if(!o.qpro) val[1].v = val[1].v.slice(1);
					/* falls through */
				case 0x0D: /* INTEGER */
				case 0x0E: /* NUMBER */
				case 0x10: /* FORMULA */
					/* TODO: actual translation of the format code */
					if(RT == 0x0E && (val[2] & 0x70) == 0x70 && (val[2] & 0x0F) > 1 && (val[2] & 0x0F) < 15) {
						val[1].z = o.dateNF || table_fmt[14];
						if(o.cellDates) { val[1].t = 'd'; val[1].v = numdate(val[1].v); }
					}

					if(o.qpro) {
						if(val[3] > sidx) {
							s["!ref"] = encode_range(refguess);
							sheets[n] = s;
							snames.push(n);
							s = (o.dense ? [] : {});
							refguess = {s: {r:0, c:0}, e: {r:0, c:0} };
							sidx = val[3]; n = next_n || "Sheet" + (sidx + 1); next_n = "";
						}
					}

					var tmpcell = o.dense ? (s[val[0].r]||[])[val[0].c] : s[encode_cell(val[0])];
					if(tmpcell) {
						tmpcell.t = val[1].t; tmpcell.v = val[1].v;
						if(val[1].z != null) tmpcell.z = val[1].z;
						if(val[1].f != null) tmpcell.f = val[1].f;
						break;
					}
					if(o.dense) {
						if(!s[val[0].r]) s[val[0].r] = [];
						s[val[0].r][val[0].c] = val[1];
					} else s[encode_cell(val[0])] = val[1];
					break;
				default:
			}}, o);
		} else if(d[2] == 0x1A || d[2] == 0x0E) {
			o.Enum = WK3Enum;
			if(d[2] == 0x0E) { o.qpro = true; d.l = 0; }
			lotushopper(d, function(val, R, RT) { switch(RT) {
				case 0xCC: n = val; break; /* SHEETNAMECS */
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
						snames.push(n);
						s = (o.dense ? [] : {});
						refguess = {s: {r:0, c:0}, e: {r:0, c:0} };
						sidx = val[3]; n = "Sheet" + (sidx + 1);
					}
					if(sheetRows > 0 && val[0].r >= sheetRows) break;
					if(o.dense) {
						if(!s[val[0].r]) s[val[0].r] = [];
						s[val[0].r][val[0].c] = val[1];
					} else s[encode_cell(val[0])] = val[1];
					if(refguess.e.c < val[0].c) refguess.e.c = val[0].c;
					if(refguess.e.r < val[0].r) refguess.e.r = val[0].r;
					break;
				case 0x1B: /* XFORMAT */
					if(val[0x36b0]) realnames[val[0x36b0][0]] = val[0x36b0][1];
					break;
				case 0x0601: /* SHEETINFOQP */
					realnames[val[0]] = val[1]; if(val[0] == sidx) n = val[1]; break;
				default: break;
			}}, o);
		} else throw new Error("Unrecognized LOTUS BOF " + d[2]);
		s["!ref"] = encode_range(refguess);
		sheets[next_n || n] = s;
		snames.push(next_n || n);
		if(!realnames.length) return { SheetNames: snames, Sheets: sheets };
		var osheets = {}, rnames = [];
		/* TODO: verify no collisions */
		for(var i = 0; i < realnames.length; ++i) if(sheets[snames[i]]) {
			rnames.push(realnames[i] || snames[i]);
			osheets[realnames[i]] = sheets[realnames[i]] || sheets[snames[i]];
		} else {
			rnames.push(realnames[i]);
			osheets[realnames[i]] = ({ "!ref": "A1" });
		}
		return { SheetNames: rnames, Sheets: osheets };
	}

	function sheet_to_wk1(ws/*:Worksheet*/, opts/*:WriteOpts*/) {
		var o = opts || {};
		if(+o.codepage >= 0) set_cp(+o.codepage);
		if(o.type == "string") throw new Error("Cannot write WK1 to JS string");
		var ba = buf_array();
		var range = safe_decode_range(ws["!ref"]);
		var dense = Array.isArray(ws);
		var cols = [];

		write_biff_rec(ba, 0x00, write_BOF_WK1(0x0406));
		write_biff_rec(ba, 0x06, write_RANGE(range));
		var max_R = Math.min(range.e.r, 8191);
		for(var R = range.s.r; R <= max_R; ++R) {
			var rr = encode_row(R);
			for(var C = range.s.c; C <= range.e.c; ++C) {
				if(R === range.s.r) cols[C] = encode_col(C);
				var ref = cols[C] + rr;
				var cell = dense ? (ws[R]||[])[C] : ws[ref];
				if(!cell || cell.t == "z") continue;
				/* TODO: formula records */
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

	function book_to_wk3(wb/*:Workbook*/, opts/*:WriteOpts*/) {
		var o = opts || {};
		if(+o.codepage >= 0) set_cp(+o.codepage);
		if(o.type == "string") throw new Error("Cannot write WK3 to JS string");
		var ba = buf_array();

		write_biff_rec(ba, 0x00, write_BOF_WK3(wb));

		for(var i = 0, cnt = 0; i < wb.SheetNames.length; ++i) if((wb.Sheets[wb.SheetNames[i]] || {})["!ref"]) write_biff_rec(ba, 0x1b, write_XFORMAT_SHEETNAME(wb.SheetNames[i], cnt++));

		var wsidx = 0;
		for(i = 0; i < wb.SheetNames.length; ++i) {
			var ws = wb.Sheets[wb.SheetNames[i]];
			if(!ws || !ws["!ref"]) continue;
			var range = safe_decode_range(ws["!ref"]);
			var dense = Array.isArray(ws);
			var cols = [];
			var max_R = Math.min(range.e.r, 8191);
			for(var R = range.s.r; R <= max_R; ++R) {
				var rr = encode_row(R);
				for(var C = range.s.c; C <= range.e.c; ++C) {
					if(R === range.s.r) cols[C] = encode_col(C);
					var ref = cols[C] + rr;
					var cell = dense ? (ws[R]||[])[C] : ws[ref];
					if(!cell || cell.t == "z") continue;
					/* TODO: FORMULA19 NUMBER18 records */
					if(cell.t == "n") {
						write_biff_rec(ba, 0x17, write_NUMBER_17(R, C, wsidx, cell.v));
					} else {
						var str = format_cell(cell);
						/* TODO: max len? */
						write_biff_rec(ba, 0x16, write_LABEL_16(R, C, wsidx, str.slice(0, 239)));
					}
				}
			}
			++wsidx;
		}

		write_biff_rec(ba, 0x01);
		return ba.end();
	}


	function write_BOF_WK1(v/*:number*/) {
		var out = new_buf(2);
		out.write_shift(2, v);
		return out;
	}

	function write_BOF_WK3(wb/*:Workbook*/) {
		var out = new_buf(26);
		out.write_shift(2, 0x1000);
		out.write_shift(2, 0x0004);
		out.write_shift(4, 0x0000);
		var rows = 0, cols = 0, wscnt = 0;
		for(var i = 0; i < wb.SheetNames.length; ++i) {
			var name = wb.SheetNames[i];
			var ws = wb.Sheets[name];
			if(!ws || !ws["!ref"]) continue;
			++wscnt;
			var range = decode_range(ws["!ref"]);
			if(rows < range.e.r) rows = range.e.r;
			if(cols < range.e.c) cols = range.e.c;
		}
		if(rows > 8191) rows = 8191;
		out.write_shift(2, rows);
		out.write_shift(1, wscnt);
		out.write_shift(1, cols);
		out.write_shift(2, 0x00);
		out.write_shift(2, 0x00);
		out.write_shift(1, 0x01);
		out.write_shift(1, 0x02);
		out.write_shift(4, 0);
		out.write_shift(4, 0);
		return out;
	}

	function parse_RANGE(blob, length, opts) {
		var o = {s:{c:0,r:0},e:{c:0,r:0}};
		if(length == 8 && opts.qpro) {
			o.s.c = blob.read_shift(1);
			blob.l++;
			o.s.r = blob.read_shift(2);
			o.e.c = blob.read_shift(1);
			blob.l++;
			o.e.r = blob.read_shift(2);
			return o;
		}
		o.s.c = blob.read_shift(2);
		o.s.r = blob.read_shift(2);
		if(length == 12 && opts.qpro) blob.l += 2;
		o.e.c = blob.read_shift(2);
		o.e.r = blob.read_shift(2);
		if(length == 12 && opts.qpro) blob.l += 2;
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
		var o = [{c:0,r:0}, {t:'n',v:0}, 0, 0];
		if(opts.qpro && opts.vers != 0x5120) {
			o[0].c = blob.read_shift(1);
			o[3] = blob.read_shift(1);
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
			wk1_fmla_to_csf(blob.slice(blob.l, blob.l + flen), o);
			blob.l += flen;
		}
		return o;
	}

	function wk1_parse_rc(B, V, col) {
		var rel = V & 0x8000;
		V &= ~0x8000;
		V = (rel ? B : 0) + ((V >= 0x2000) ? V - 0x4000 : V);
		return (rel ? "" : "$") + (col ? encode_col(V) : encode_row(V));
	}
	/* var oprec = [
		8, 8, 8, 8, 8, 8, 8, 8, 6, 4, 4, 5, 5, 7, 3, 3,
		3, 3, 3, 3, 1, 1, 2, 6, 8, 8, 8, 8, 8, 8, 8, 8
	]; */
	/* TODO: flesh out */
	var FuncTab = {
		0x33: ["FALSE", 0],
		0x34: ["TRUE", 0],
		0x46: ["LEN", 1],
		0x50: ["SUM", 69],
		0x51: ["AVERAGEA", 69],
		0x52: ["COUNTA", 69],
		0x53: ["MINA", 69],
		0x54: ["MAXA", 69],
		0x6F: ["T", 1]
	};
	var BinOpTab = [
		  "",   "",   "",   "",   "",   "",   "",   "", // eslint-disable-line no-mixed-spaces-and-tabs
		  "",  "+",  "-",  "*",  "/",  "^",  "=", "<>", // eslint-disable-line no-mixed-spaces-and-tabs
		"<=", ">=",  "<",  ">",   "",   "",   "",   "", // eslint-disable-line no-mixed-spaces-and-tabs
		 "&",   "",   "",   "",   "",   "",   "",   ""  // eslint-disable-line no-mixed-spaces-and-tabs
	];

	function wk1_fmla_to_csf(blob, o) {
		prep_blob(blob, 0);
		var out = [], argc = 0, R = "", C = "", argL = "", argR = "";
		while(blob.l < blob.length) {
			var cc = blob[blob.l++];
			switch(cc) {
				case 0x00: out.push(blob.read_shift(8, 'f')); break;
				case 0x01: {
					C = wk1_parse_rc(o[0].c, blob.read_shift(2), true);
					R = wk1_parse_rc(o[0].r, blob.read_shift(2), false);
					out.push(C + R);
				} break;
				case 0x02: {
					var c = wk1_parse_rc(o[0].c, blob.read_shift(2), true);
					var r = wk1_parse_rc(o[0].r, blob.read_shift(2), false);
					C = wk1_parse_rc(o[0].c, blob.read_shift(2), true);
					R = wk1_parse_rc(o[0].r, blob.read_shift(2), false);
					out.push(c + r + ":" + C + R);
				} break;
				case 0x03:
					if(blob.l < blob.length) { console.error("WK1 premature formula end"); return; }
					break;
				case 0x04: out.push("(" + out.pop() + ")"); break;
				case 0x05: out.push(blob.read_shift(2)); break;
				case 0x06: {
					/* TODO: text encoding */
					var Z = ""; while((cc = blob[blob.l++])) Z += String.fromCharCode(cc);
					out.push('"' + Z.replace(/"/g, '""') + '"');
				} break;

				case 0x08: out.push("-" + out.pop()); break;
				case 0x17: out.push("+" + out.pop()); break;
				case 0x16: out.push("NOT(" + out.pop() + ")"); break;

				case 0x14: case 0x15: {
					argR = out.pop(); argL = out.pop();
					out.push(["AND", "OR"][cc - 0x14] + "(" + argL + "," + argR + ")");
				} break;

				default:
					if(cc < 0x20 && BinOpTab[cc]) {
						argR = out.pop(); argL = out.pop();
						out.push(argL + BinOpTab[cc] + argR);
					} else if(FuncTab[cc]) {
						argc = FuncTab[cc][1];
						if(argc == 69) argc = blob[blob.l++];
						if(argc > out.length) { console.error("WK1 bad formula parse 0x" + cc.toString(16) + ":|" + out.join("|") + "|"); return; }
						var args = out.slice(-argc);
						out.length -= argc;
						out.push(FuncTab[cc][0] + "(" + args.join(",") + ")");
					}
					else if(cc <= 0x07) return console.error("WK1 invalid opcode " + cc.toString(16));
					else if(cc <= 0x18) return console.error("WK1 unsupported op " + cc.toString(16));
					else if(cc <= 0x1E) return console.error("WK1 invalid opcode " + cc.toString(16));
					else if(cc <= 0x73) return console.error("WK1 unsupported function opcode " + cc.toString(16));
					// possible future functions ??
					else return console.error("WK1 unrecognized opcode " + cc.toString(16));
			}
		}
		if(out.length == 1) o[1].f = "" + out[0];
		else console.error("WK1 bad formula parse |" + out.join("|") + "|");
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
	function write_LABEL_16(R, C, wsidx, s) {
		/* TODO: encoding */
		var o = new_buf(6 + s.length);
		o.write_shift(2, R);
		o.write_shift(1, wsidx);
		o.write_shift(1, C);
		o.write_shift(1, 0x27);
		for(var i = 0; i < s.length; ++i) {
			var cc = s.charCodeAt(i);
			o.write_shift(1, cc >= 0x80 ? 0x5F : cc);
		}
		o.write_shift(1, 0);
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
		o[1].v = (1 - s*2) * (v2 * Math.pow(2, e+32) + v1 * Math.pow(2, e));
		return o;
	}
	function write_NUMBER_17(R, C, wsidx, v) {
		var o = new_buf(14);
		o.write_shift(2, R);
		o.write_shift(1, wsidx);
		o.write_shift(1, C);
		if(v == 0) {
			o.write_shift(4, 0);
			o.write_shift(4, 0);
			o.write_shift(2, 0xFFFF);
			return o;
		}
		var s = 0, e = 0, v1 = 0, v2 = 0;
		if(v < 0) { s = 1; v = -v; }
		e = Math.log2(v) | 0;
		v /= Math.pow(2, e-31);
		v2 = (v)>>>0;
		if((v2&0x80000000) == 0) { v/=2; ++e; v2 = v >>> 0; }
		v -= v2;
		v2 |= 0x80000000;
		v2 >>>= 0;
		v *= Math.pow(2, 32);
		v1 = v>>>0;
		o.write_shift(4, v1);
		o.write_shift(4, v2);
		e += 0x3FFF + (s ? 0x8000 : 0);
		o.write_shift(2, e);
		return o;
	}

	function parse_FORMULA_19(blob, length) {
		var o = parse_NUMBER_17(blob, 14);
		blob.l += length - 14; /* TODO: WK3 formula */
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

	function parse_SHEETNAMECS(blob, length) {
		return blob[blob.l + length - 1] == 0 ? blob.read_shift(length, 'cstr') : "";
	}

	function parse_SHEETNAMELP(blob, length) {
		var len = blob[blob.l++];
		if(len > length - 1) len = length - 1;
		var o = ""; while(o.length < len) o += String.fromCharCode(blob[blob.l++]);
		return o;
	}

	function parse_SHEETINFOQP(blob, length, opts) {
		if(!opts.qpro || length < 21) return;
		var id = blob.read_shift(1);
		blob.l += 17;
		blob.l += 1; //var len = blob.read_shift(1);
		blob.l += 2;
		var nm = blob.read_shift(length - 21, 'cstr');
		return [id, nm];
	}

	function parse_XFORMAT(blob, length) {
		var o = {}, tgt = blob.l + length;
		while(blob.l < tgt) {
			var dt = blob.read_shift(2);
			if(dt == 0x36b0) {
				o[dt] = [0, ""];
				o[dt][0] = blob.read_shift(2);
				while(blob[blob.l]) { o[dt][1] += String.fromCharCode(blob[blob.l]); blob.l++; } blob.l++;
			}
			// TODO: 0x3a99 ??
		}
		return o;
	}
	function write_XFORMAT_SHEETNAME(name, wsidx) {
		var out = new_buf(5 + name.length);
		out.write_shift(2, 0x36b0);
		out.write_shift(2, wsidx);
		for(var i = 0; i < name.length; ++i) {
			var cc = name.charCodeAt(i);
			out[out.l++] = cc > 0x7F ? 0x5F : cc;
		}
		out[out.l++] = 0;
		return out;
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
		/*::[*/0x00CC/*::]*/: { n:"SHEETNAMECS", f:parse_SHEETNAMECS },
		/*::[*/0x00DE/*::]*/: { n:"SHEETNAMELP", f:parse_SHEETNAMELP },
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
		/*::[*/0x001B/*::]*/: { n:"XFORMAT", f:parse_XFORMAT },
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
		/*::[*/0x00CC/*::]*/: { n:"SHEETNAMECS", f:parse_SHEETNAMECS },
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
		/*::[*/0x0601/*::]*/: { n:"SHEETINFOQP", f:parse_SHEETINFOQP },
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
		book_to_wk3: book_to_wk3,
		to_workbook: lotus_to_workbook
	};
})();
