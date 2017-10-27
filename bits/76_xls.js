/* [MS-OLEDS] 2.3.8 CompObjStream */
function parse_compobj(obj/*:CFBEntry*/) {
	var v = {};
	var o = obj.content;
	/*:: if(o == null) return; */

	/* [MS-OLEDS] 2.3.7 CompObjHeader -- All fields MUST be ignored */
	var l = 28, m;
	m = __lpstr(o, l);
	l += 4 + __readUInt32LE(o,l);
	v.UserType = m;

	/* [MS-OLEDS] 2.3.1 ClipboardFormatOrAnsiString */
	m = __readUInt32LE(o,l); l+= 4;
	switch(m) {
		case 0x00000000: break;
		case 0xffffffff: case 0xfffffffe: l+=4; break;
		default:
			if(m > 0x190) throw new Error("Unsupported Clipboard: " + m.toString(16));
			l += m;
	}

	m = __lpstr(o, l); l += m.length === 0 ? 0 : 5 + m.length; v.Reserved1 = m;

	if((m = __readUInt32LE(o,l)) !== 0x71b2e9f4) return v;
	throw new Error("Unsupported Unicode Extension");
}

/*
	Continue logic for:
	- 2.4.58 Continue
	- 2.4.59 ContinueBigName
	- 2.4.60 ContinueFrt
	- 2.4.61 ContinueFrt11
	- 2.4.62 ContinueFrt12
*/
function slurp(R, blob, length/*:number*/, opts) {
	var l = length;
	var bufs = [];
	var d = blob.slice(blob.l,blob.l+l);
	if(opts && opts.enc && opts.enc.insitu) switch(R.n) {
	case 'BOF': case 'FilePass': case 'FileLock': case 'InterfaceHdr': case 'RRDInfo': case 'RRDHead': case 'UsrExcl': break;
	default:
		if(d.length === 0) break;
		opts.enc.insitu(d);
	}
	bufs.push(d);
	blob.l += l;
	var next = (XLSRecordEnum[__readUInt16LE(blob,blob.l)]);
	var start = 0;
	while(next != null && next.n.slice(0,8) === 'Continue') {
		l = __readUInt16LE(blob,blob.l+2);
		start = blob.l + 4;
		if(next.n == 'ContinueFrt') start += 4;
		else if(next.n.slice(0,11) == 'ContinueFrt') start += 12;
		bufs.push(blob.slice(start,blob.l+4+l));
		blob.l += 4+l;
		next = (XLSRecordEnum[__readUInt16LE(blob, blob.l)]);
	}
	var b = (bconcat(bufs)/*:any*/);
	prep_blob(b, 0);
	var ll = 0; b.lens = [];
	for(var j = 0; j < bufs.length; ++j) { b.lens.push(ll); ll += bufs[j].length; }
	return R.f(b, b.length, opts);
}

function safe_format_xf(p/*:any*/, opts/*:ParseOpts*/, date1904/*:?boolean*/) {
	if(p.t === 'z') return;
	if(!p.XF) return;
	var fmtid = 0;
	try {
		fmtid = p.z || p.XF.numFmtId || 0;
		if(opts.cellNF) p.z = SSF._table[fmtid];
	} catch(e) { if(opts.WTF) throw e; }
	if(!opts || opts.cellText !== false) try {
		if(p.t === 'e') { p.w = p.w || BErr[p.v]; }
		else if(fmtid === 0 || fmtid == "General") {
			if(p.t === 'n') {
				if((p.v|0) === p.v) p.w = SSF._general_int(p.v);
				else p.w = SSF._general_num(p.v);
			}
			else p.w = SSF._general(p.v);
		}
		else p.w = SSF.format(fmtid,p.v, {date1904:!!date1904});
	} catch(e) { if(opts.WTF) throw e; }
	if(opts.cellDates && fmtid && p.t == 'n' && SSF.is_date(SSF._table[fmtid] || String(fmtid))) {
		var _d = SSF.parse_date_code(p.v); if(_d) { p.t = 'd'; p.v = new Date(_d.y, _d.m-1,_d.d,_d.H,_d.M,_d.S,_d.u); }
	}
}

function make_cell(val, ixfe, t)/*:Cell*/ {
	return ({v:val, ixfe:ixfe, t:t}/*:any*/);
}

// 2.3.2
function parse_workbook(blob, options/*:ParseOpts*/)/*:Workbook*/ {
	var wb = ({opts:{}}/*:any*/);
	var Sheets = {};
	if(DENSE != null && options.dense == null) options.dense = DENSE;
	var out/*:Worksheet*/ = ((options.dense ? [] : {})/*:any*/);
	var Directory = {};
	var found_sheet = false;
	var range/*:Range*/ = ({}/*:any*/);
	var last_formula = null;
	var sst = [];
	var cur_sheet = "";
	var Preamble = {};
	var lastcell, last_cell = "", cc, cmnt, rng, rngC, rngR;
	var shared_formulae = {};
	var array_formulae = []; /* TODO: something more clever */
	var temp_val/*:Cell*/;
	var country;
	var cell_valid = true;
	var XFs = []; /* XF records */
	var palette = [];
	var Workbook/*:WBWBProps*/ = ({ Sheets:[], WBProps:{date1904:false} }/*:any*/), wsprops = {};
	var get_rgb = function getrgb(icv) {
		if(icv < 8) return XLSIcv[icv];
		if(icv < 64) return palette[icv-8] || XLSIcv[icv];
		return XLSIcv[icv];
	};
	var process_cell_style = function pcs(cell, line/*:any*/, options) {
		var xfd = line.XF.data;
		if(!xfd || !xfd.patternType || !options || !options.cellStyles) return;
		line.s = ({}/*:any*/);
		line.s.patternType = xfd.patternType;
		var t;
		if((t = rgb2Hex(get_rgb(xfd.icvFore)))) { line.s.fgColor = {rgb:t}; }
		if((t = rgb2Hex(get_rgb(xfd.icvBack)))) { line.s.bgColor = {rgb:t}; }
	};
	var addcell = function addcell(cell/*:any*/, line/*:any*/, options/*:any*/) {
		if(file_depth > 1) return;
		if(!cell_valid) return;
		if(options.cellStyles && line.XF && line.XF.data) process_cell_style(cell, line, options);
		delete line.ixfe; delete line.XF;
		lastcell = cell;
		last_cell = encode_cell(cell);
		if(range.s) {
			if(cell.r < range.s.r) range.s.r = cell.r;
			if(cell.c < range.s.c) range.s.c = cell.c;
		}
		if(range.e) {
			if(cell.r + 1 > range.e.r) range.e.r = cell.r + 1;
			if(cell.c + 1 > range.e.c) range.e.c = cell.c + 1;
		}
		if(options.cellFormula && line.f) {
			for(var afi = 0; afi < array_formulae.length; ++afi) {
				if(array_formulae[afi][0].s.c > cell.c) continue;
				if(array_formulae[afi][0].s.r > cell.r) continue;
				if(array_formulae[afi][0].e.c < cell.c) continue;
				if(array_formulae[afi][0].e.r < cell.r) continue;
				line.F = encode_range(array_formulae[afi][0]);
				if(array_formulae[afi][0].s.c != cell.c) delete line.f;
				if(array_formulae[afi][0].s.r != cell.r) delete line.f;
				if(line.f) line.f = "" + stringify_formula(array_formulae[afi][1], range, cell, supbooks, opts);
				break;
			}
		}
		if(options.sheetRows && lastcell.r >= options.sheetRows) cell_valid = false;
		else {
			if(options.dense) {
				if(!out[cell.r]) out[cell.r] = [];
				out[cell.r][cell.c] = line;
			} else out[last_cell] = line;
		}
	};
	var opts = ({
		enc: false, // encrypted
		sbcch: 0, // cch in the preceding SupBook
		snames: [], // sheetnames
		sharedf: shared_formulae, // shared formulae by address
		arrayf: array_formulae, // array formulae array
		rrtabid: [], // RRTabId
		lastuser: "", // Last User from WriteAccess
		biff: 8, // BIFF version
		codepage: 0, // CP from CodePage record
		winlocked: 0, // fLockWn from WinProtect
		cellStyles: !!options && !!options.cellStyles,
		WTF: !!options && !!options.wtf
	}/*:any*/);
	if(options.password) opts.password = options.password;
	var mergecells = [];
	var objects = [];
	var colinfo = [], rowinfo = [];
	var defwidth = 0, defheight = 0; // twips / MDW respectively
	var seencol = false;
	var supbooks = ([]/*:any*/); // 1-indexed, will hold extern names
	supbooks.SheetNames = opts.snames;
	supbooks.sharedf = opts.sharedf;
	supbooks.arrayf = opts.arrayf;
	supbooks.names = [];
	supbooks.XTI = [];
	var last_Rn = '';
	var file_depth = 0; /* TODO: make a real stack */
	var BIFF2Fmt = 0;
	var BIFF2FmtTable/*:Array<string>*/ = [];
	var FilterDatabases = []; /* TODO: sort out supbooks and process elsewhere */
	var last_lbl/*:?DefinedName*/;

	/* explicit override for some broken writers */
	opts.codepage = 1200;
	set_cp(1200);
	while(blob.l < blob.length - 1) {
		var s = blob.l;
		var RecordType = blob.read_shift(2);
		if(RecordType === 0 && last_Rn === 'EOF') break;
		var length = (blob.l === blob.length ? 0 : blob.read_shift(2)), y;
		var R = XLSRecordEnum[RecordType];
		//console.log(RecordType.toString(16), RecordType, R, blob.l, length, blob.length);
		//if(!R) console.log(blob.slice(blob.l, blob.l + length));
		if(R && R.f) {
			if(options.bookSheets) {
				if(last_Rn === 'BoundSheet8' && R.n !== 'BoundSheet8') break;
			}
			last_Rn = R.n;
			if(R.r === 2 || R.r == 12) {
				var rt = blob.read_shift(2); length -= 2;
				if(!opts.enc && rt !== RecordType) throw new Error("rt mismatch: " + rt + "!=" + RecordType);
				if(R.r == 12){ blob.l += 10; length -= 10; } // skip FRT
			}
			//console.error(R,blob.l,length,blob.length);
			var val;
			if(R.n === 'EOF') val = R.f(blob, length, opts);
			else val = slurp(R, blob, length, opts);
			var Rn = R.n;
			/* nested switch statements to workaround V8 128 limit */
			switch(Rn) {
				/* Workbook Options */
				case 'Date1904':
					/*:: if(!Workbook.WBProps) Workbook.WBProps = {}; */
					wb.opts.Date1904 = Workbook.WBProps.date1904 = val; break;
				case 'WriteProtect': wb.opts.WriteProtect = true; break;
				case 'FilePass':
					if(!opts.enc) blob.l = 0;
					opts.enc = val;
					if(opts.WTF) console.error(val);
					if(!options.password) throw new Error("File is password-protected");
					if(val.valid == null) throw new Error("Encryption scheme unsupported");
					if(!val.valid) throw new Error("Password is incorrect");
					break;
				case 'WriteAccess': opts.lastuser = val; break;
				case 'FileSharing': break; //TODO
				case 'CodePage':
					/* overrides based on test cases */
					switch(val) {
						case 0x5212: val =  1200; break;
						case 0x8000: val = 10000; break;
						case 0x8001: val =  1252; break;
					}
					opts.codepage = val;
					set_cp(val);
					break;
				case 'RRTabId': opts.rrtabid = val; break;
				case 'WinProtect': opts.winlocked = val; break;
				case 'Template': break; // TODO
				case 'RefreshAll': wb.opts.RefreshAll = val; break;
				case 'BookBool': break; // TODO
				case 'UsesELFs': break;
				case 'MTRSettings': break;
				case 'CalcCount': wb.opts.CalcCount = val; break;
				case 'CalcDelta': wb.opts.CalcDelta = val; break;
				case 'CalcIter': wb.opts.CalcIter = val; break;
				case 'CalcMode': wb.opts.CalcMode = val; break;
				case 'CalcPrecision': wb.opts.CalcPrecision = val; break;
				case 'CalcSaveRecalc': wb.opts.CalcSaveRecalc = val; break;
				case 'CalcRefMode': opts.CalcRefMode = val; break; // TODO: implement R1C1
				case 'Uncalced': break;
				case 'ForceFullCalculation': wb.opts.FullCalc = val; break;
				case 'WsBool': break; // TODO
				case 'XF': XFs.push(val); break;
				case 'ExtSST': break; // TODO
				case 'BookExt': break; // TODO
				case 'RichTextStream': break;
				case 'BkHim': break;

				case 'SupBook':
					supbooks.push([val]);
					supbooks[supbooks.length-1].XTI = [];
					break;
				case 'ExternName':
					supbooks[supbooks.length-1].push(val);
					break;
				case 'Index': break; // TODO
				case 'Lbl':
					last_lbl = ({
						Name: val.Name,
						Ref: stringify_formula(val.rgce,range,null,supbooks,opts)
					}/*:DefinedName*/);
					if(val.itab > 0) last_lbl.Sheet = val.itab - 1;
					supbooks.names.push(last_lbl);
					if(!supbooks[0]) supbooks[0] = [];
					supbooks[supbooks.length-1].push(val);
					if(val.Name == "_xlnm._FilterDatabase" && val.itab > 0)
						if(val.rgce && val.rgce[0] && val.rgce[0][0] && val.rgce[0][0][0] == 'PtgArea3d')
							FilterDatabases[val.itab - 1] = { ref: encode_range(val.rgce[0][0][1][2]) };
					break;
				case 'ExternCount': opts.ExternCount = val; break;
				case 'ExternSheet':
					if(supbooks.length == 0) { supbooks[0] = []; supbooks[0].XTI = []; }
					supbooks[supbooks.length - 1].XTI = supbooks[supbooks.length - 1].XTI.concat(val); supbooks.XTI = supbooks.XTI.concat(val); break;
				case 'NameCmt':
					/* TODO: search for correct name */
					if(opts.biff < 8) break;
					if(last_lbl != null) last_lbl.Comment = val[1];
					break;

				case 'Protect': out["!protect"] = val; break; /* for sheet or book */
				case 'Password': if(val !== 0 && opts.WTF) console.error("Password verifier: " + val); break;
				case 'Prot4Rev': case 'Prot4RevPass': break; /*TODO: Revision Control*/

				case 'BoundSheet8': {
					Directory[val.pos] = val;
					opts.snames.push(val.name);
				} break;
				case 'EOF': {
					if(--file_depth) break;
					if(range.e) {
						if(range.e.r > 0 && range.e.c > 0) {
							range.e.r--; range.e.c--;
							out["!ref"] = encode_range(range);
							range.e.r++; range.e.c++;
						}
						if(mergecells.length > 0) out["!merges"] = mergecells;
						if(objects.length > 0) out["!objects"] = objects;
						if(colinfo.length > 0) out["!cols"] = colinfo;
						if(rowinfo.length > 0) out["!rows"] = rowinfo;
						Workbook.Sheets.push(wsprops);
					}
					if(cur_sheet === "") Preamble = out; else Sheets[cur_sheet] = out;
					out = ((options.dense ? [] : {})/*:any*/);
				} break;
				case 'BOF': {
					if(opts.biff === 8) switch(RecordType) {
						case 0x0009: opts.biff = 2; break;
						case 0x0209: opts.biff = 3; break;
						case 0x0409: opts.biff = 4; break;
						default: switch(val.BIFFVer) {
						case 0x0500: opts.biff = 5; break;
						case 0x0600: opts.biff = 8; break;
						case 0x0002: opts.biff = 2; break;
						case 0x0007: opts.biff = 2; break;
					}}
					if(file_depth++) break;
					cell_valid = true;
					out = ((options.dense ? [] : {})/*:any*/);

					if(opts.biff < 5) {
						if(cur_sheet === "") cur_sheet = "Sheet1";
						range = {s:{r:0,c:0},e:{r:0,c:0}};
						/* fake BoundSheet8 */
						var fakebs8 = {pos: blob.l - length, name:cur_sheet};
						Directory[fakebs8.pos] = fakebs8;
						opts.snames.push(cur_sheet);
					}
					else cur_sheet = (Directory[s] || {name:""}).name;
					if(val.dt == 0x20) out["!type"] = "chart";
					mergecells = [];
					objects = [];
					array_formulae = []; opts.arrayf = array_formulae;
					colinfo = []; rowinfo = [];
					defwidth = defheight = 0;
					seencol = false;
					wsprops = {Hidden:(Directory[s]||{hs:0}).hs, name:cur_sheet };
				} break;

				case 'Number': case 'BIFF2NUM': case 'BIFF2INT': {
					if(out["!type"] == "chart") if(options.dense ? (out[val.r]||[])[val.c]: out[encode_cell({c:val.c, r:val.r})]) ++val.c;
					temp_val = ({ixfe: val.ixfe, XF: XFs[val.ixfe]||{}, v:val.val, t:'n'}/*:any*/);
					if(BIFF2Fmt > 0) temp_val.z = BIFF2FmtTable[(temp_val.ixfe>>8) & 0x1F];
					safe_format_xf(temp_val, options, wb.opts.Date1904);
					addcell({c:val.c, r:val.r}, temp_val, options);
				} break;
				case 'BoolErr': {
					temp_val = ({ixfe: val.ixfe, XF: XFs[val.ixfe], v:val.val, t:val.t}/*:any*/);
					if(BIFF2Fmt > 0) temp_val.z = BIFF2FmtTable[(temp_val.ixfe>>8) & 0x1F];
					safe_format_xf(temp_val, options, wb.opts.Date1904);
					addcell({c:val.c, r:val.r}, temp_val, options);
				} break;
				case 'RK': {
					temp_val = ({ixfe: val.ixfe, XF: XFs[val.ixfe], v:val.rknum, t:'n'}/*:any*/);
					if(BIFF2Fmt > 0) temp_val.z = BIFF2FmtTable[(temp_val.ixfe>>8) & 0x1F];
					safe_format_xf(temp_val, options, wb.opts.Date1904);
					addcell({c:val.c, r:val.r}, temp_val, options);
				} break;
				case 'MulRk': {
					for(var j = val.c; j <= val.C; ++j) {
						var ixfe = val.rkrec[j-val.c][0];
						temp_val= ({ixfe:ixfe, XF:XFs[ixfe], v:val.rkrec[j-val.c][1], t:'n'}/*:any*/);
						if(BIFF2Fmt > 0) temp_val.z = BIFF2FmtTable[(temp_val.ixfe>>8) & 0x1F];
						safe_format_xf(temp_val, options, wb.opts.Date1904);
						addcell({c:j, r:val.r}, temp_val, options);
					}
				} break;
				case 'Formula': {
					if(val.val == 'String') { last_formula = val; break; }
					temp_val = make_cell(val.val, val.cell.ixfe, val.tt);
					temp_val.XF = XFs[temp_val.ixfe];
					if(options.cellFormula) {
						var _f = val.formula;
						if(_f && _f[0] && _f[0][0] && _f[0][0][0] == 'PtgExp') {
							var _fr = _f[0][0][1][0], _fc = _f[0][0][1][1];
							var _fe = encode_cell({r:_fr, c:_fc});
							if(shared_formulae[_fe]) temp_val.f = ""+stringify_formula(val.formula,range,val.cell,supbooks, opts);
							else temp_val.F = ((options.dense ? (out[_fr]||[])[_fc]: out[_fe]) || {}).F;
						} else temp_val.f = ""+stringify_formula(val.formula,range,val.cell,supbooks, opts);
					}
					if(BIFF2Fmt > 0) temp_val.z = BIFF2FmtTable[(temp_val.ixfe>>8) & 0x1F];
					safe_format_xf(temp_val, options, wb.opts.Date1904);
					addcell(val.cell, temp_val, options);
					last_formula = val;
				} break;
				case 'String': {
					if(last_formula) { /* technically always true */
						last_formula.val = val;
						temp_val = make_cell(val, last_formula.cell.ixfe, 's');
						temp_val.XF = XFs[temp_val.ixfe];
						if(options.cellFormula) {
							temp_val.f = ""+stringify_formula(last_formula.formula, range, last_formula.cell, supbooks, opts);
						}
						if(BIFF2Fmt > 0) temp_val.z = BIFF2FmtTable[(temp_val.ixfe>>8) & 0x1F];
						safe_format_xf(temp_val, options, wb.opts.Date1904);
						addcell(last_formula.cell, temp_val, options);
						last_formula = null;
					} else throw new Error("String record expects Formula");
				} break;
				case 'Array': {
					array_formulae.push(val);
					var _arraystart = encode_cell(val[0].s);
					cc = options.dense ? (out[val[0].s.r]||[])[val[0].s.c] : out[_arraystart];
					if(options.cellFormula && cc) {
						if(!last_formula) break; /* technically unreachable */
						if(!_arraystart || !cc) break;
						cc.f = ""+stringify_formula(val[1], range, val[0], supbooks, opts);
						cc.F = encode_range(val[0]);
					}
				} break;
				case 'ShrFmla': {
					if(!cell_valid) break;
					if(!options.cellFormula) break;
					if(last_cell) {
						/* TODO: capture range */
						if(!last_formula) break; /* technically unreachable */
						shared_formulae[encode_cell(last_formula.cell)]= val[0];
						cc = options.dense ? (out[last_formula.cell.r]||[])[last_formula.cell.c] : out[encode_cell(last_formula.cell)];
						(cc||{}).f = ""+stringify_formula(val[0], range, lastcell, supbooks, opts);
					}
				} break;
				case 'LabelSst':
					temp_val=make_cell(sst[val.isst].t, val.ixfe, 's');
					temp_val.XF = XFs[temp_val.ixfe];
					if(BIFF2Fmt > 0) temp_val.z = BIFF2FmtTable[(temp_val.ixfe>>8) & 0x1F];
					safe_format_xf(temp_val, options, wb.opts.Date1904);
					addcell({c:val.c, r:val.r}, temp_val, options);
					break;
				case 'Blank': if(options.sheetStubs) {
					temp_val = ({ixfe: val.ixfe, XF: XFs[val.ixfe], t:'z'}/*:any*/);
					if(BIFF2Fmt > 0) temp_val.z = BIFF2FmtTable[(temp_val.ixfe>>8) & 0x1F];
					safe_format_xf(temp_val, options, wb.opts.Date1904);
					addcell({c:val.c, r:val.r}, temp_val, options);
				} break;
				case 'MulBlank': if(options.sheetStubs) {
					for(var _j = val.c; _j <= val.C; ++_j) {
						var _ixfe = val.ixfe[_j-val.c];
						temp_val= ({ixfe:_ixfe, XF:XFs[_ixfe], t:'z'}/*:any*/);
						if(BIFF2Fmt > 0) temp_val.z = BIFF2FmtTable[(temp_val.ixfe>>8) & 0x1F];
						safe_format_xf(temp_val, options, wb.opts.Date1904);
						addcell({c:_j, r:val.r}, temp_val, options);
					}
				} break;
				case 'RString':
				case 'Label': case 'BIFF2STR':
					temp_val=make_cell(val.val, val.ixfe, 's');
					temp_val.XF = XFs[temp_val.ixfe];
					if(BIFF2Fmt > 0) temp_val.z = BIFF2FmtTable[(temp_val.ixfe>>8) & 0x1F];
					safe_format_xf(temp_val, options, wb.opts.Date1904);
					addcell({c:val.c, r:val.r}, temp_val, options);
					break;

				case 'Dimensions': {
					if(file_depth === 1) range = val; /* TODO: stack */
				} break;
				case 'SST': {
					sst = val;
				} break;
				case 'Format': { /* val = [id, fmt] */
					SSF.load(val[1], val[0]);
				} break;
				case 'BIFF2FORMAT': {
					BIFF2FmtTable[BIFF2Fmt++] = val;
					for(var b2idx = 0; b2idx < BIFF2Fmt + 163; ++b2idx) if(SSF._table[b2idx] == val) break;
					if(b2idx >= 163) SSF.load(val, BIFF2Fmt + 163);
				} break;

				case 'MergeCells': mergecells = mergecells.concat(val); break;

				case 'Obj': objects[val.cmo[0]] = opts.lastobj = val; break;
				case 'TxO': opts.lastobj.TxO = val; break;
				case 'ImData': opts.lastobj.ImData = val; break;

				case 'HLink': {
					for(rngR = val[0].s.r; rngR <= val[0].e.r; ++rngR)
						for(rngC = val[0].s.c; rngC <= val[0].e.c; ++rngC) {
							cc = options.dense ? (out[rngR]||[])[rngC] : out[encode_cell({c:rngC,r:rngR})];
							if(cc) cc.l = val[1];
						}
				} break;
				case 'HLinkTooltip': {
					for(rngR = val[0].s.r; rngR <= val[0].e.r; ++rngR)
						for(rngC = val[0].s.c; rngC <= val[0].e.c; ++rngC) {
							cc = options.dense ? (out[rngR]||[])[rngC] : out[encode_cell({c:rngC,r:rngR})];
							if(cc) cc.l.Tooltip = val[1];
							}
				} break;

				/* Comments */
				case 'Note': {
					if(opts.biff <= 5 && opts.biff >= 2) break; /* TODO: BIFF5 */
					cc = options.dense ? (out[val[0].r]||[])[val[0].c] : out[encode_cell(val[0])];
					var noteobj = objects[val[2]];
					if(!cc) break;
					if(!cc.c) cc.c = [];
					cmnt = {a:val[1],t:noteobj.TxO.t};
					cc.c.push(cmnt);
				} break;

				default: switch(R.n) { /* nested */
				case 'ClrtClient': break;
				case 'XFExt': update_xfext(XFs[val.ixfe], val.ext); break;

				case 'DefColWidth': defwidth = val; break;
				case 'DefaultRowHeight': defheight = val[1]; break; // TODO: flags

				case 'ColInfo': {
					if(!opts.cellStyles) break;
					while(val.e >= val.s) {
						colinfo[val.e--] = { width: val.w/256 };
						if(!seencol) { seencol = true; find_mdw_colw(val.w/256); }
						process_col(colinfo[val.e+1]);
					}
				} break;
				case 'Row': {
					var rowobj = {};
					if(val.level != null) { rowinfo[val.r] = rowobj; rowobj.level = val.level; }
					if(val.hidden) { rowinfo[val.r] = rowobj; rowobj.hidden = true; }
					if(val.hpt) {
						rowinfo[val.r] = rowobj;
						rowobj.hpt = val.hpt; rowobj.hpx = pt2px(val.hpt);
					}
				} break;

				case 'LeftMargin':
				case 'RightMargin':
				case 'TopMargin':
				case 'BottomMargin':
					if(!out['!margins']) default_margins(out['!margins'] = {});
					out['!margins'][Rn.slice(0,-6).toLowerCase()] = val;
					break;

				case 'Setup': // TODO
					if(!out['!margins']) default_margins(out['!margins'] = {});
					out['!margins'].header = val.header;
					out['!margins'].footer = val.footer;
					break;

				case 'Header': break; // TODO
				case 'Footer': break; // TODO
				case 'HCenter': break; // TODO
				case 'VCenter': break; // TODO
				case 'Pls': break; // TODO
				case 'GCW': break;
				case 'LHRecord': break;
				case 'DBCell': break; // TODO
				case 'EntExU2': break; // TODO
				case 'SxView': break; // TODO
				case 'Sxvd': break; // TODO
				case 'SXVI': break; // TODO
				case 'SXVDEx': break; // TODO
				case 'SxIvd': break; // TODO
				case 'SXString': break; // TODO
				case 'Sync': break;
				case 'Addin': break;
				case 'SXDI': break; // TODO
				case 'SXLI': break; // TODO
				case 'SXEx': break; // TODO
				case 'QsiSXTag': break; // TODO
				case 'Selection': break;
				case 'Feat': break;
				case 'FeatHdr': case 'FeatHdr11': break;
				case 'Feature11': case 'Feature12': case 'List12': break;
				case 'Country': country = val; break;
				case 'RecalcId': break;
				case 'DxGCol': break; // TODO: htmlify
				case 'Fbi': case 'Fbi2': case 'GelFrame': break;
				case 'Font': break; // TODO
				case 'XFCRC': break; // TODO
				case 'Style': break; // TODO
				case 'StyleExt': break; // TODO
				case 'Palette': palette = val; break; // TODO
				case 'Theme': break; // TODO
				/* Protection */
				case 'ScenarioProtect': break;
				case 'ObjProtect': break;

				/* Conditional Formatting */
				case 'CondFmt12': break;

				/* Table */
				case 'Table': break; // TODO
				case 'TableStyles': break; // TODO
				case 'TableStyle': break; // TODO
				case 'TableStyleElement': break; // TODO

				/* PivotTable */
				case 'SXStreamID': break; // TODO
				case 'SXVS': break; // TODO
				case 'DConRef': break; // TODO
				case 'SXAddl': break; // TODO
				case 'DConBin': break; // TODO
				case 'DConName': break; // TODO
				case 'SXPI': break; // TODO
				case 'SxFormat': break; // TODO
				case 'SxSelect': break; // TODO
				case 'SxRule': break; // TODO
				case 'SxFilt': break; // TODO
				case 'SxItm': break; // TODO
				case 'SxDXF': break; // TODO

				/* Scenario Manager */
				case 'ScenMan': break;

				/* Data Consolidation */
				case 'DCon': break;

				/* Watched Cell */
				case 'CellWatch': break;

				/* Print Settings */
				case 'PrintRowCol': break;
				case 'PrintGrid': break;
				case 'PrintSize': break;

				case 'XCT': break;
				case 'CRN': break;

				case 'Scl': {
					//console.log("Zoom Level:", val[0]/val[1],val);
				} break;
				case 'SheetExt': {
					/* empty */
				} break;
				case 'SheetExtOptional': {
					/* empty */
				} break;

				/* VBA */
				case 'ObNoMacros': {
					/* empty */
				} break;
				case 'ObProj': {
					/* empty */
				} break;
				case 'CodeName': {
					/*:: if(!Workbook.WBProps) Workbook.WBProps = {}; */
					if(!cur_sheet) Workbook.WBProps.codeName = val;
				} break;
				case 'GUIDTypeLib': {
					/* empty */
				} break;

				case 'WOpt': break; // TODO: WTF?
				case 'PhoneticInfo': break;

				case 'OleObjectSize': break;

				/* Differential Formatting */
				case 'DXF': case 'DXFN': case 'DXFN12': case 'DXFN12List': case 'DXFN12NoCB': break;

				/* Data Validation */
				case 'Dv': case 'DVal': break;

				/* Data Series */
				case 'BRAI': case 'Series': case 'SeriesText': break;

				/* Data Connection */
				case 'DConn': break;
				case 'DbOrParamQry': break;
				case 'DBQueryExt': break;

				case 'OleDbConn': break;
				case 'ExtString': break;

				/* Formatting */
				case 'IFmtRecord': break;
				case 'CondFmt': case 'CF': case 'CF12': case 'CFEx': break;

				/* Explicitly Ignored */
				case 'Excel9File': break;
				case 'Units': break;
				case 'InterfaceHdr': case 'Mms': case 'InterfaceEnd': case 'DSF': break;
				case 'BuiltInFnGroupCount': /* 2.4.30 0x0E or 0x10 but excel 2011 generates 0x11? */ break;
				/* View Stuff */
				case 'Window1': case 'Window2': case 'HideObj': case 'GridSet': case 'Guts':
				case 'UserBView': case 'UserSViewBegin': case 'UserSViewEnd':
				case 'Pane': break;
				default: switch(R.n) { /* nested */
				/* Chart */
				case 'Dat':
				case 'Begin': case 'End':
				case 'StartBlock': case 'EndBlock':
				case 'Frame': case 'Area':
				case 'Axis': case 'AxisLine': case 'Tick': break;
				case 'AxesUsed':
				case 'CrtLayout12': case 'CrtLayout12A': case 'CrtLink': case 'CrtLine': case 'CrtMlFrt': case 'CrtMlFrtContinue': break;
				case 'LineFormat': case 'AreaFormat':
				case 'Chart': case 'Chart3d': case 'Chart3DBarShape': case 'ChartFormat': case 'ChartFrtInfo': break;
				case 'PlotArea': case 'PlotGrowth': break;
				case 'SeriesList': case 'SerParent': case 'SerAuxTrend': break;
				case 'DataFormat': case 'SerToCrt': case 'FontX': break;
				case 'CatSerRange': case 'AxcExt': case 'SerFmt': break;
				case 'ShtProps': break;
				case 'DefaultText': case 'Text': case 'CatLab': break;
				case 'DataLabExtContents': break;
				case 'Legend': case 'LegendException': break;
				case 'Pie': case 'Scatter': break;
				case 'PieFormat': case 'MarkerFormat': break;
				case 'StartObject': case 'EndObject': break;
				case 'AlRuns': case 'ObjectLink': break;
				case 'SIIndex': break;
				case 'AttachedLabel': case 'YMult': break;

				/* Chart Group */
				case 'Line': case 'Bar': break;
				case 'Surf': break;

				/* Axis Group */
				case 'AxisParent': break;
				case 'Pos': break;
				case 'ValueRange': break;

				/* Pivot Chart */
				case 'SXViewEx9': break; // TODO
				case 'SXViewLink': break;
				case 'PivotChartBits': break;
				case 'SBaseRef': break;
				case 'TextPropsStream': break;

				/* Chart Misc */
				case 'LnExt': break;
				case 'MkrExt': break;
				case 'CrtCoopt': break;

				/* Query Table */
				case 'Qsi': case 'Qsif': case 'Qsir': case 'QsiSXTag': break;
				case 'TxtQry': break;

				/* Filter */
				case 'FilterMode': break;
				case 'AutoFilter': case 'AutoFilterInfo': break;
				case 'AutoFilter12': break;
				case 'DropDownObjIds': break;
				case 'Sort': break;
				case 'SortData': break;

				/* Drawing */
				case 'ShapePropsStream': break;
				case 'MsoDrawing': case 'MsoDrawingGroup': case 'MsoDrawingSelection': break;
				/* Pub Stuff */
				case 'WebPub': case 'AutoWebPub': break;

				/* Print Stuff */
				case 'HeaderFooter': case 'HFPicture': case 'PLV':
				case 'HorizontalPageBreaks': case 'VerticalPageBreaks': break;
				/* Behavioral */
				case 'Backup': case 'CompressPictures': case 'Compat12': break;

				/* Should not Happen */
				case 'Continue': case 'ContinueFrt12': break;

				/* Future Records */
				case 'FrtFontList': case 'FrtWrapper': break;

				default: switch(R.n) { /* nested */
				/* BIFF5 records */
				case 'TabIdConf': case 'Radar': case 'RadarArea': case 'DropBar': case 'Intl': case 'CoordList': case 'SerAuxErrBar': break;

				/* BIFF2-4 records */
				case 'BIFF2FONTCLR': case 'BIFF2FMTCNT': case 'BIFF2FONTXTRA': break;
				case 'BIFF2XF': case 'BIFF3XF': case 'BIFF4XF': break;
				case 'BIFF4FMTCNT': case 'BIFF2ROW': case 'BIFF2WINDOW2': break;

				/* Miscellaneous */
				case 'SCENARIO': case 'DConBin': case 'PicF': case 'DataLabExt':
				case 'Lel': case 'BopPop': case 'BopPopCustom': case 'RealTimeData':
				case 'Name': break;
				case 'LHNGraph': case 'FnGroupName': case 'AddMenu': case 'LPr': break;
				case 'ListObj': case 'ListField': break;
				case 'RRSort': break;
				case 'BigName': break;
				case 'ToolbarHdr': case 'ToolbarEnd': break;
				case 'DDEObjName': break;
				case 'FRTArchId$': break;
				default: if(options.WTF) throw 'Unrecognized Record ' + R.n;
			}}}}
		} else blob.l += length;
	}
	var sheetnamesraw = Object.keys(Directory).sort(function(a,b) { return Number(a) - Number(b); }).map(function(x){return Directory[x].name;});
	var sheetnames = sheetnamesraw.slice();
	wb.Directory=sheetnamesraw;
	wb.SheetNames=sheetnamesraw;
	if(!options.bookSheets) wb.Sheets=Sheets;
	if(wb.Sheets) FilterDatabases.forEach(function(r,i) { wb.Sheets[wb.SheetNames[i]]['!autofilter'] = r; });
	wb.Preamble=Preamble;
	wb.Strings = sst;
	wb.SSF = SSF.get_table();
	if(opts.enc) wb.Encryption = opts.enc;
	wb.Metadata = {};
	if(country !== undefined) wb.Metadata.Country = country;
	if(supbooks.names.length > 0) Workbook.Names = supbooks.names;
	wb.Workbook = Workbook;
	return wb;
}

/* TODO: WTF */
function parse_props(cfb/*:CFBContainer*/, props, o) {
	/* [MS-OSHARED] 2.3.3.2.2 Document Summary Information Property Set */
	var DSI = CFB.find(cfb, '!DocumentSummaryInformation');
	if(DSI) try {
		var DocSummary = parse_PropertySetStream(DSI, DocSummaryPIDDSI);
		for(var d in DocSummary) props[d] = DocSummary[d];
	} catch(e) {if(o.WTF == 2) throw e;/* empty */}

	/* [MS-OSHARED] 2.3.3.2.1 Summary Information Property Set*/
	var SI = CFB.find(cfb, '!SummaryInformation');
	if(SI) try {
		var Summary = parse_PropertySetStream(SI, SummaryPIDSI);
		for(var s in Summary) if(props[s] == null) props[s] = Summary[s];
	} catch(e) {if(o.WTF == 2) throw e;/* empty */}
}

function parse_xlscfb(cfb/*:any*/, options/*:?ParseOpts*/)/*:Workbook*/ {
if(!options) options = {};
fix_read_opts(options);
reset_cp();
var CompObj/*:?CFBEntry*/, Summary, WB/*:?any*/;
if(cfb.FullPaths) {
	CompObj = CFB.find(cfb, '!CompObj');
	Summary = CFB.find(cfb, '!SummaryInformation');
	WB = CFB.find(cfb, '/Workbook') || CFB.find(cfb, '/Book');
} else {
	switch(options.type) {
		case 'base64': cfb = s2a(Base64.decode(cfb)); break;
		case 'binary': cfb = s2a(cfb); break;
		case 'buffer': break;
		case 'array': if(!Array.isArray(cfb)) cfb = Array.prototype.slice.call(cfb); break;
	}
	prep_blob(cfb, 0);
	WB = ({content: cfb}/*:any*/);
}
var CompObjP, SummaryP, WorkbookP/*:: :Workbook = XLSX.utils.book_new(); */;

var _data/*:?any*/;
if(CompObj) CompObjP = parse_compobj(CompObj);
if(options.bookProps && !options.bookSheets) WorkbookP = ({}/*:any*/);
else/*:: if(cfb instanceof CFBContainer) */ {
	var T = has_buf ? 'buffer' : 'array';
	if(WB && WB.content) WorkbookP = parse_workbook(WB.content, options);
	/* Quattro Pro 7-8 */
	else if((_data=CFB.find(cfb, 'PerfectOffice_MAIN')) && _data.content) WorkbookP = WK_.to_workbook(_data.content, (options.type = T, options));
	/* Quattro Pro 9 */
	else if((_data=CFB.find(cfb, 'NativeContent_MAIN')) && _data.content) WorkbookP = WK_.to_workbook(_data.content, (options.type = T, options));
	else throw new Error("Cannot find Workbook stream");
	if(options.bookVBA && CFB.find(cfb, '/_VBA_PROJECT_CUR/VBA/dir')) WorkbookP.vbaraw = make_vba_xls(cfb);
}

var props = {};
if(cfb.FullPaths) parse_props(/*::((*/cfb/*:: :any):CFBContainer)*/, props, options);

WorkbookP.Props = WorkbookP.Custprops = props; /* TODO: split up properties */
if(options.bookFiles) WorkbookP.cfb = cfb;
/*WorkbookP.CompObjP = CompObjP; // TODO: storage? */
return WorkbookP;
}


function write_xlscfb(wb/*:Workbook*/, opts/*:WriteOpts*/)/*:CFBContainer*/ {
	var o = opts || {};
	var cfb = CFB.utils.cfb_new({root:"R"});
	var wbpath = "/Workbook";
	switch(o.bookType || "xls") {
		case "xls": o.bookType = "biff8";
		/* falls through */
		case "biff8": wbpath = "/Workbook"; o.biff = 8; break;
		case "biff5": wbpath = "/Book"; o.biff = 5; break;
		default: throw new Error("invalid type " + o.bookType + " for XLS CFB");
	}
	CFB.utils.cfb_add(cfb, wbpath, write_biff_buf(wb, o));
	// TODO: SI, DSI, CO
	if(o.biff == 8 && wb.vbaraw) fill_vba_xls(cfb, CFB.read(wb.vbaraw, {type: typeof wb.vbaraw == "string" ? "binary" : "buffer"}));
	return cfb;
}
