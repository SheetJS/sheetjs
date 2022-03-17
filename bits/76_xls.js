/* [MS-OLEDS] 2.3.8 CompObjStream */
function parse_compobj(obj/*:CFBEntry*/) {
	var v = {};
	var o = obj.content;
	/*:: if(o == null) return; */

	/* [MS-OLEDS] 2.3.7 CompObjHeader -- All fields MUST be ignored */
	o.l = 28;

	v.AnsiUserType = o.read_shift(0, "lpstr-ansi");
	v.AnsiClipboardFormat = parse_ClipboardFormatOrAnsiString(o);

	if(o.length - o.l <= 4) return v;

	var m/*:number*/ = o.read_shift(4);
	if(m == 0 || m > 40) return v;
	o.l-=4; v.Reserved1 = o.read_shift(0, "lpstr-ansi");

	if(o.length - o.l <= 4) return v;
	m = o.read_shift(4);
	if(m !== 0x71b239f4) return v;
	v.UnicodeClipboardFormat = parse_ClipboardFormatOrUnicodeString(o);

	m = o.read_shift(4);
	if(m == 0 || m > 40) return v;
	o.l-=4; v.Reserved2 = o.read_shift(0, "lpwstr");
}

/*
	Continue logic for:
	- 2.4.58 Continue          0x003c
	- 2.4.59 ContinueBigName   0x043c
	- 2.4.60 ContinueFrt       0x0812
	- 2.4.61 ContinueFrt11     0x0875
	- 2.4.62 ContinueFrt12     0x087f
*/
var CONTINUE_RT = [ 0x003c, 0x043c, 0x0812, 0x0875, 0x087f ];
function slurp(RecordType, R, blob, length/*:number*/, opts)/*:any*/ {
	var l = length;
	var bufs = [];
	var d = blob.slice(blob.l,blob.l+l);
	if(opts && opts.enc && opts.enc.insitu && d.length > 0) switch(RecordType) {
	case 0x0009: case 0x0209: case 0x0409: case 0x0809/* BOF */: case 0x002f /* FilePass */: case 0x0195 /* FileLock */: case 0x00e1 /* InterfaceHdr */: case 0x0196 /* RRDInfo */: case 0x0138 /* RRDHead */: case 0x0194 /* UsrExcl */: case 0x000a /* EOF */:
		break;
	case 0x0085 /* BoundSheet8 */:
		break;
	default:
		opts.enc.insitu(d);
	}
	bufs.push(d);
	blob.l += l;
	var nextrt = __readUInt16LE(blob,blob.l), next = XLSRecordEnum[nextrt];
	var start = 0;
	while(next != null && CONTINUE_RT.indexOf(nextrt) > -1) {
		l = __readUInt16LE(blob,blob.l+2);
		start = blob.l + 4;
		if(nextrt == 0x0812 /* ContinueFrt */) start += 4;
		else if(nextrt == 0x0875 || nextrt == 0x087f) {
			start += 12;
		}
		d = blob.slice(start,blob.l+4+l);
		bufs.push(d);
		blob.l += 4+l;
		next = (XLSRecordEnum[nextrt = __readUInt16LE(blob, blob.l)]);
	}
	var b = (bconcat(bufs)/*:any*/);
	prep_blob(b, 0);
	var ll = 0; b.lens = [];
	for(var j = 0; j < bufs.length; ++j) { b.lens.push(ll); ll += bufs[j].length; }
	if(b.length < length) throw "XLS Record 0x" + RecordType.toString(16) + " Truncated: " + b.length + " < " + length;
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
				if((p.v|0) === p.v) p.w = p.v.toString(10);
				else p.w = SSF._general_num(p.v);
			}
			else p.w = SSF._general(p.v);
		}
		else p.w = SSF.format(fmtid,p.v, {date1904:!!date1904, dateNF: opts && opts.dateNF});
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
	var range/*:Range*/ = ({}/*:any*/);
	var last_formula = null;
	var sst/*:SST*/ = ([]/*:any*/);
	var cur_sheet = "";
	var Preamble = {};
	var lastcell, last_cell = "", cc/*:Cell*/, cmnt, rngC, rngR;
	var sharedf = {};
	var arrayf/*:Array<[Range, string]>*/ = [];
	var temp_val/*:Cell*/;
	var country;
	var XFs = []; /* XF records */
	var palette/*:Array<[number, number, number]>*/ = [];
	var Workbook/*:WBWBProps*/ = ({ Sheets:[], WBProps:{date1904:false}, Views:[{}] }/*:any*/), wsprops = {};
	var get_rgb = function getrgb(icv/*:number*/)/*:[number, number, number]*/ {
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
		if(options.sheetRows && cell.r >= options.sheetRows) return;
		if(options.cellStyles && line.XF && line.XF.data) process_cell_style(cell, line, options);
		delete line.ixfe; delete line.XF;
		lastcell = cell;
		last_cell = encode_cell(cell);
		if(!range || !range.s || !range.e) range = {s:{r:0,c:0},e:{r:0,c:0}};
		if(cell.r < range.s.r) range.s.r = cell.r;
		if(cell.c < range.s.c) range.s.c = cell.c;
		if(cell.r + 1 > range.e.r) range.e.r = cell.r + 1;
		if(cell.c + 1 > range.e.c) range.e.c = cell.c + 1;
		if(options.cellFormula && line.f) {
			for(var afi = 0; afi < arrayf.length; ++afi) {
				if(arrayf[afi][0].s.c > cell.c || arrayf[afi][0].s.r > cell.r) continue;
				if(arrayf[afi][0].e.c < cell.c || arrayf[afi][0].e.r < cell.r) continue;
				line.F = encode_range(arrayf[afi][0]);
				if(arrayf[afi][0].s.c != cell.c || arrayf[afi][0].s.r != cell.r) delete line.f;
				if(line.f) line.f = "" + stringify_formula(arrayf[afi][1], range, cell, supbooks, opts);
				break;
			}
		}
		{
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
		sharedf: sharedf, // shared formulae by address
		arrayf: arrayf, // array formulae array
		rrtabid: [], // RRTabId
		lastuser: "", // Last User from WriteAccess
		biff: 8, // BIFF version
		codepage: 0, // CP from CodePage record
		winlocked: 0, // fLockWn from WinProtect
		cellStyles: !!options && !!options.cellStyles,
		WTF: !!options && !!options.wtf
	}/*:any*/);
	if(options.password) opts.password = options.password;
	var themes;
	var merges/*:Array<Range>*/ = [];
	var objects = [];
	var colinfo/*:Array<ColInfo>*/ = [], rowinfo/*:Array<RowInfo>*/ = [];
	var seencol = false;
	var supbooks = ([]/*:any*/); // 1-indexed, will hold extern names
	supbooks.SheetNames = opts.snames;
	supbooks.sharedf = opts.sharedf;
	supbooks.arrayf = opts.arrayf;
	supbooks.names = [];
	supbooks.XTI = [];
	var last_RT = 0;
	var file_depth = 0; /* TODO: make a real stack */
	var BIFF2Fmt = 0, BIFF2FmtTable/*:Array<string>*/ = [];
	var FilterDatabases = []; /* TODO: sort out supbooks and process elsewhere */
	var last_lbl/*:?DefinedName*/;

	/* explicit override for some broken writers */
	opts.codepage = 1200;
	set_cp(1200);
	var seen_codepage = false;
	while(blob.l < blob.length - 1) {
		var s = blob.l;
		var RecordType = blob.read_shift(2);
		if(RecordType === 0 && last_RT === 0x000a /* EOF */) break;
		var length = (blob.l === blob.length ? 0 : blob.read_shift(2));
		var R = XLSRecordEnum[RecordType];
		//console.log(RecordType.toString(16), RecordType, R, blob.l, length, blob.length);
		//if(!R) console.log(blob.slice(blob.l, blob.l + length));
		if(R && R.f) {
			if(options.bookSheets) {
				if(last_RT === 0x0085 /* BoundSheet8 */ && RecordType !== 0x0085 /* R.n !== 'BoundSheet8' */) break;
			}
			last_RT = RecordType;
			if(R.r === 2 || R.r == 12) {
				var rt = blob.read_shift(2); length -= 2;
				if(!opts.enc && rt !== RecordType && (((rt&0xFF)<<8)|(rt>>8)) !== RecordType) throw new Error("rt mismatch: " + rt + "!=" + RecordType);
				if(R.r == 12){
					blob.l += 10; length -= 10;
				} // skip FRT
			}
			//console.error(R,blob.l,length,blob.length);
			var val/*:any*/ = ({}/*:any*/);
			if(RecordType === 0x000a /* EOF */) val = /*::(*/R.f(blob, length, opts)/*:: :any)*/;
			else val = /*::(*/slurp(RecordType, R, blob, length, opts)/*:: :any)*/;
			/*:: val = (val:any); */
			if(file_depth == 0 && [0x0009, 0x0209, 0x0409, 0x0809].indexOf(last_RT) === -1 /* 'BOF' */) continue;
			switch(RecordType) {
				case 0x0022 /* Date1904 */:
					/*:: if(!Workbook.WBProps) Workbook.WBProps = {}; */
					wb.opts.Date1904 = Workbook.WBProps.date1904 = val; break;
				case 0x0086 /* WriteProtect */: wb.opts.WriteProtect = true; break;
				case 0x002f /* FilePass */:
					if(!opts.enc) blob.l = 0;
					opts.enc = val;
					if(!options.password) throw new Error("File is password-protected");
					if(val.valid == null) throw new Error("Encryption scheme unsupported");
					if(!val.valid) throw new Error("Password is incorrect");
					break;
				case 0x005c /* WriteAccess */: opts.lastuser = val; break;
				case 0x0042 /* CodePage */:
					var cpval = Number(val);
					/* overrides based on test cases */
					switch(cpval) {
						case 0x5212: cpval =  1200; break;
						case 0x8000: cpval = 10000; break;
						case 0x8001: cpval =  1252; break;
					}
					set_cp(opts.codepage = cpval);
					seen_codepage = true;
					break;
				case 0x013d /* RRTabId */: opts.rrtabid = val; break;
				case 0x0019 /* WinProtect */: opts.winlocked = val; break;
				case 0x01b7 /* RefreshAll */: wb.opts["RefreshAll"] = val; break;
				case 0x000c /* CalcCount */: wb.opts["CalcCount"] = val; break;
				case 0x0010 /* CalcDelta */: wb.opts["CalcDelta"] = val; break;
				case 0x0011 /* CalcIter */: wb.opts["CalcIter"] = val; break;
				case 0x000d /* CalcMode */: wb.opts["CalcMode"] = val; break;
				case 0x000e /* CalcPrecision */: wb.opts["CalcPrecision"] = val; break;
				case 0x005f /* CalcSaveRecalc */: wb.opts["CalcSaveRecalc"] = val; break;
				case 0x000f /* CalcRefMode */: opts.CalcRefMode = val; break; // TODO: implement R1C1
				case 0x08a3 /* ForceFullCalculation */: wb.opts.FullCalc = val; break;
				case 0x0081 /* WsBool */:
					if(val.fDialog) out["!type"] = "dialog";
					if(!val.fBelow) (out["!outline"] || (out["!outline"] = {})).above = true;
					if(!val.fRight) (out["!outline"] || (out["!outline"] = {})).left = true;
					break; // TODO
				case 0x00e0 /* XF */:
					XFs.push(val); break;
				case 0x01ae /* SupBook */:
					supbooks.push([val]);
					supbooks[supbooks.length-1].XTI = [];
					break;
				case 0x0023: case 0x0223 /* ExternName */:
					supbooks[supbooks.length-1].push(val);
					break;
				case 0x0018: case 0x0218 /* Lbl */:
					last_lbl = ({
						Name: val.Name,
						Ref: stringify_formula(val.rgce,range,null,supbooks,opts)
					}/*:DefinedName*/);
					if(val.itab > 0) last_lbl.Sheet = val.itab - 1;
					supbooks.names.push(last_lbl);
					if(!supbooks[0]) { supbooks[0] = []; supbooks[0].XTI = []; }
					supbooks[supbooks.length-1].push(val);
					if(val.Name == "_xlnm._FilterDatabase" && val.itab > 0)
						if(val.rgce && val.rgce[0] && val.rgce[0][0] && val.rgce[0][0][0] == 'PtgArea3d')
							FilterDatabases[val.itab - 1] = { ref: encode_range(val.rgce[0][0][1][2]) };
					break;
				case 0x0016 /* ExternCount */: opts.ExternCount = val; break;
				case 0x0017 /* ExternSheet */:
					if(supbooks.length == 0) { supbooks[0] = []; supbooks[0].XTI = []; }
					supbooks[supbooks.length - 1].XTI = supbooks[supbooks.length - 1].XTI.concat(val); supbooks.XTI = supbooks.XTI.concat(val); break;
				case 0x0894 /* NameCmt */:
					/* TODO: search for correct name */
					if(opts.biff < 8) break;
					if(last_lbl != null) last_lbl.Comment = val[1];
					break;
				case 0x0012 /* Protect */: out["!protect"] = val; break; /* for sheet or book */
				case 0x0013 /* Password */: if(val !== 0 && opts.WTF) console.error("Password verifier: " + val); break;
				case 0x0085 /* BoundSheet8 */: {
					Directory[val.pos] = val;
					opts.snames.push(val.name);
				} break;
				case 0x000a /* EOF */: {
					if(--file_depth) break;
					if(range.e) {
						if(range.e.r > 0 && range.e.c > 0) {
							range.e.r--; range.e.c--;
							out["!ref"] = encode_range(range);
							if(options.sheetRows && options.sheetRows <= range.e.r) {
								var tmpri = range.e.r;
								range.e.r = options.sheetRows - 1;
								out["!fullref"] = out["!ref"];
								out["!ref"] = encode_range(range);
								range.e.r = tmpri;
							}
							range.e.r++; range.e.c++;
						}
						if(merges.length > 0) out["!merges"] = merges;
						if(objects.length > 0) out["!objects"] = objects;
						if(colinfo.length > 0) out["!cols"] = colinfo;
						if(rowinfo.length > 0) out["!rows"] = rowinfo;
						Workbook.Sheets.push(wsprops);
					}
					if(cur_sheet === "") Preamble = out; else Sheets[cur_sheet] = out;
					out = ((options.dense ? [] : {})/*:any*/);
				} break;
				case 0x0009: case 0x0209: case 0x0409: case 0x0809 /* BOF */: {
					if(opts.biff === 8) opts.biff = {
						/*::[*/0x0009/*::]*/:2,
						/*::[*/0x0209/*::]*/:3,
						/*::[*/0x0409/*::]*/:4
					}[RecordType] || {
						/*::[*/0x0200/*::]*/:2,
						/*::[*/0x0300/*::]*/:3,
						/*::[*/0x0400/*::]*/:4,
						/*::[*/0x0500/*::]*/:5,
						/*::[*/0x0600/*::]*/:8,
						/*::[*/0x0002/*::]*/:2,
						/*::[*/0x0007/*::]*/:2
					}[val.BIFFVer] || 8;
					opts.biffguess = val.BIFFVer == 0;
					if(val.BIFFVer == 0 && val.dt == 0x1000) { opts.biff = 5; seen_codepage = true; set_cp(opts.codepage = 28591); }
					if(opts.biff == 8 && val.BIFFVer == 0 && val.dt == 16) opts.biff = 2;
					if(file_depth++) break;
					out = ((options.dense ? [] : {})/*:any*/);

					if(opts.biff < 8 && !seen_codepage) { seen_codepage = true; set_cp(opts.codepage = options.codepage || 1252); }

					if(opts.biff < 5 || val.BIFFVer == 0 && val.dt == 0x1000) {
						if(cur_sheet === "") cur_sheet = "Sheet1";
						range = {s:{r:0,c:0},e:{r:0,c:0}};
						/* fake BoundSheet8 */
						var fakebs8 = {pos: blob.l - length, name:cur_sheet};
						Directory[fakebs8.pos] = fakebs8;
						opts.snames.push(cur_sheet);
					}
					else cur_sheet = (Directory[s] || {name:""}).name;
					if(val.dt == 0x20) out["!type"] = "chart";
					if(val.dt == 0x40) out["!type"] = "macro";
					merges = [];
					objects = [];
					opts.arrayf = arrayf = [];
					colinfo = []; rowinfo = [];
					seencol = false;
					wsprops = {Hidden:(Directory[s]||{hs:0}).hs, name:cur_sheet };
				} break;
				case 0x0203 /* Number */: case 0x0003 /* BIFF2NUM */: case 0x0002 /* BIFF2INT */: {
					if(out["!type"] == "chart") if(options.dense ? (out[val.r]||[])[val.c]: out[encode_cell({c:val.c, r:val.r})]) ++val.c;
					temp_val = ({ixfe: val.ixfe, XF: XFs[val.ixfe]||{}, v:val.val, t:'n'}/*:any*/);
					if(BIFF2Fmt > 0) temp_val.z = BIFF2FmtTable[(temp_val.ixfe>>8) & 0x3F];
					safe_format_xf(temp_val, options, wb.opts.Date1904);
					addcell({c:val.c, r:val.r}, temp_val, options);
				} break;
				case 0x0005: case 0x0205 /* BoolErr */: {
					temp_val = ({ixfe: val.ixfe, XF: XFs[val.ixfe], v:val.val, t:val.t}/*:any*/);
					if(BIFF2Fmt > 0) temp_val.z = BIFF2FmtTable[(temp_val.ixfe>>8) & 0x3F];
					safe_format_xf(temp_val, options, wb.opts.Date1904);
					addcell({c:val.c, r:val.r}, temp_val, options);
				} break;
				case 0x027e /* RK */: {
					temp_val = ({ixfe: val.ixfe, XF: XFs[val.ixfe], v:val.rknum, t:'n'}/*:any*/);
					if(BIFF2Fmt > 0) temp_val.z = BIFF2FmtTable[(temp_val.ixfe>>8) & 0x3F];
					safe_format_xf(temp_val, options, wb.opts.Date1904);
					addcell({c:val.c, r:val.r}, temp_val, options);
				} break;
				case 0x00bd /* MulRk */: {
					for(var j = val.c; j <= val.C; ++j) {
						var ixfe = val.rkrec[j-val.c][0];
						temp_val= ({ixfe:ixfe, XF:XFs[ixfe], v:val.rkrec[j-val.c][1], t:'n'}/*:any*/);
						if(BIFF2Fmt > 0) temp_val.z = BIFF2FmtTable[(temp_val.ixfe>>8) & 0x3F];
						safe_format_xf(temp_val, options, wb.opts.Date1904);
						addcell({c:j, r:val.r}, temp_val, options);
					}
				} break;
				case 0x0006: case 0x0206: case 0x0406 /* Formula */: {
					if(val.val == 'String') { last_formula = val; break; }
					temp_val = make_cell(val.val, val.cell.ixfe, val.tt);
					temp_val.XF = XFs[temp_val.ixfe];
					if(options.cellFormula) {
						var _f = val.formula;
						if(_f && _f[0] && _f[0][0] && _f[0][0][0] == 'PtgExp') {
							var _fr = _f[0][0][1][0], _fc = _f[0][0][1][1];
							var _fe = encode_cell({r:_fr, c:_fc});
							if(sharedf[_fe]) temp_val.f = ""+stringify_formula(val.formula,range,val.cell,supbooks, opts);
							else temp_val.F = ((options.dense ? (out[_fr]||[])[_fc]: out[_fe]) || {}).F;
						} else temp_val.f = ""+stringify_formula(val.formula,range,val.cell,supbooks, opts);
					}
					if(BIFF2Fmt > 0) temp_val.z = BIFF2FmtTable[(temp_val.ixfe>>8) & 0x3F];
					safe_format_xf(temp_val, options, wb.opts.Date1904);
					addcell(val.cell, temp_val, options);
					last_formula = val;
				} break;
				case 0x0007: case 0x0207 /* String */: {
					if(last_formula) { /* technically always true */
						last_formula.val = val;
						temp_val = make_cell(val, last_formula.cell.ixfe, 's');
						temp_val.XF = XFs[temp_val.ixfe];
						if(options.cellFormula) {
							temp_val.f = ""+stringify_formula(last_formula.formula, range, last_formula.cell, supbooks, opts);
						}
						if(BIFF2Fmt > 0) temp_val.z = BIFF2FmtTable[(temp_val.ixfe>>8) & 0x3F];
						safe_format_xf(temp_val, options, wb.opts.Date1904);
						addcell(last_formula.cell, temp_val, options);
						last_formula = null;
					} else throw new Error("String record expects Formula");
				} break;
				case 0x0021: case 0x0221 /* Array */: {
					arrayf.push(val);
					var _arraystart = encode_cell(val[0].s);
					cc = options.dense ? (out[val[0].s.r]||[])[val[0].s.c] : out[_arraystart];
					if(options.cellFormula && cc) {
						if(!last_formula) break; /* technically unreachable */
						if(!_arraystart || !cc) break;
						cc.f = ""+stringify_formula(val[1], range, val[0], supbooks, opts);
						cc.F = encode_range(val[0]);
					}
				} break;
				case 0x04bc /* ShrFmla */: {
					if(!options.cellFormula) break;
					if(last_cell) {
						/* TODO: capture range */
						if(!last_formula) break; /* technically unreachable */
						sharedf[encode_cell(last_formula.cell)]= val[0];
						cc = options.dense ? (out[last_formula.cell.r]||[])[last_formula.cell.c] : out[encode_cell(last_formula.cell)];
						(cc||{}).f = ""+stringify_formula(val[0], range, lastcell, supbooks, opts);
					}
				} break;
				case 0x00fd /* LabelSst */:
					temp_val=make_cell(sst[val.isst].t, val.ixfe, 's');
					if(sst[val.isst].h) temp_val.h = sst[val.isst].h;
					temp_val.XF = XFs[temp_val.ixfe];
					if(BIFF2Fmt > 0) temp_val.z = BIFF2FmtTable[(temp_val.ixfe>>8) & 0x3F];
					safe_format_xf(temp_val, options, wb.opts.Date1904);
					addcell({c:val.c, r:val.r}, temp_val, options);
					break;
				case 0x0201 /* Blank */: if(options.sheetStubs) {
					temp_val = ({ixfe: val.ixfe, XF: XFs[val.ixfe], t:'z'}/*:any*/);
					if(BIFF2Fmt > 0) temp_val.z = BIFF2FmtTable[(temp_val.ixfe>>8) & 0x3F];
					safe_format_xf(temp_val, options, wb.opts.Date1904);
					addcell({c:val.c, r:val.r}, temp_val, options);
				} break;
				case 0x00be /* MulBlank */: if(options.sheetStubs) {
					for(var _j = val.c; _j <= val.C; ++_j) {
						var _ixfe = val.ixfe[_j-val.c];
						temp_val= ({ixfe:_ixfe, XF:XFs[_ixfe], t:'z'}/*:any*/);
						if(BIFF2Fmt > 0) temp_val.z = BIFF2FmtTable[(temp_val.ixfe>>8) & 0x3F];
						safe_format_xf(temp_val, options, wb.opts.Date1904);
						addcell({c:_j, r:val.r}, temp_val, options);
					}
				} break;
				case 0x00d6 /* RString */:
				case 0x0204 /* Label */: case 0x0004 /* BIFF2STR */:
					temp_val=make_cell(val.val, val.ixfe, 's');
					temp_val.XF = XFs[temp_val.ixfe];
					if(BIFF2Fmt > 0) temp_val.z = BIFF2FmtTable[(temp_val.ixfe>>8) & 0x3F];
					safe_format_xf(temp_val, options, wb.opts.Date1904);
					addcell({c:val.c, r:val.r}, temp_val, options);
					break;

				case 0x0000: case 0x0200 /* Dimensions */: {
					if(file_depth === 1) range = val; /* TODO: stack */
				} break;
				case 0x00fc /* SST */: {
					sst = val;
				} break;
				case 0x041e /* Format */: { /* val = [id, fmt] */
					if(opts.biff == 4) {
						BIFF2FmtTable[BIFF2Fmt++] = val[1];
						for(var b4idx = 0; b4idx < BIFF2Fmt + 163; ++b4idx) if(SSF._table[b4idx] == val[1]) break;
						if(b4idx >= 163) SSF.load(val[1], BIFF2Fmt + 163);
					}
					else SSF.load(val[1], val[0]);
				} break;
				case 0x001e /* BIFF2FORMAT */: {
					BIFF2FmtTable[BIFF2Fmt++] = val;
					for(var b2idx = 0; b2idx < BIFF2Fmt + 163; ++b2idx) if(SSF._table[b2idx] == val) break;
					if(b2idx >= 163) SSF.load(val, BIFF2Fmt + 163);
				} break;

				case 0x00e5 /* MergeCells */: merges = merges.concat(val); break;

				case 0x005d /* Obj */: objects[val.cmo[0]] = opts.lastobj = val; break;
				case 0x01b6 /* TxO */: opts.lastobj.TxO = val; break;
				case 0x007f /* ImData */: opts.lastobj.ImData = val; break;

				case 0x01b8 /* HLink */: {
					for(rngR = val[0].s.r; rngR <= val[0].e.r; ++rngR)
						for(rngC = val[0].s.c; rngC <= val[0].e.c; ++rngC) {
							cc = options.dense ? (out[rngR]||[])[rngC] : out[encode_cell({c:rngC,r:rngR})];
							if(cc) cc.l = val[1];
						}
				} break;
				case 0x0800 /* HLinkTooltip */: {
					for(rngR = val[0].s.r; rngR <= val[0].e.r; ++rngR)
						for(rngC = val[0].s.c; rngC <= val[0].e.c; ++rngC) {
							cc = options.dense ? (out[rngR]||[])[rngC] : out[encode_cell({c:rngC,r:rngR})];
							if(cc && cc.l) cc.l.Tooltip = val[1];
							}
				} break;
				case 0x001c /* Note */: {
					if(opts.biff <= 5 && opts.biff >= 2) break; /* TODO: BIFF5 */
					cc = options.dense ? (out[val[0].r]||[])[val[0].c] : out[encode_cell(val[0])];
					var noteobj = objects[val[2]];
					if(!cc) {
						if(options.dense) {
							if(!out[val[0].r]) out[val[0].r] = [];
							cc = out[val[0].r][val[0].c] = ({t:"z"}/*:any*/);
						} else {
							cc = out[encode_cell(val[0])] = ({t:"z"}/*:any*/);
						}
						range.e.r = Math.max(range.e.r, val[0].r);
						range.s.r = Math.min(range.s.r, val[0].r);
						range.e.c = Math.max(range.e.c, val[0].c);
						range.s.c = Math.min(range.s.c, val[0].c);
					}
					if(!cc.c) cc.c = [];
					cmnt = {a:val[1],t:noteobj.TxO.t};
					cc.c.push(cmnt);
				} break;
				case 0x087d /* XFExt */: update_xfext(XFs[val.ixfe], val.ext); break;
				case 0x007d /* ColInfo */: {
					if(!opts.cellStyles) break;
					while(val.e >= val.s) {
						colinfo[val.e--] = { width: val.w/256, level: (val.level || 0), hidden: !!(val.flags & 1) };
						if(!seencol) { seencol = true; find_mdw_colw(val.w/256); }
						process_col(colinfo[val.e+1]);
					}
				} break;
				case 0x0208 /* Row */: {
					var rowobj = {};
					if(val.level != null) { rowinfo[val.r] = rowobj; rowobj.level = val.level; }
					if(val.hidden) { rowinfo[val.r] = rowobj; rowobj.hidden = true; }
					if(val.hpt) {
						rowinfo[val.r] = rowobj;
						rowobj.hpt = val.hpt; rowobj.hpx = pt2px(val.hpt);
					}
				} break;
				case 0x0026 /* LeftMargin */:
				case 0x0027 /* RightMargin */:
				case 0x0028 /* TopMargin */:
				case 0x0029 /* BottomMargin */:
					if(!out['!margins']) default_margins(out['!margins'] = {});
					out['!margins'][({0x26: "left", 0x27:"right", 0x28:"top", 0x29:"bottom"})[RecordType]] = val;
					break;
				case 0x00a1 /* Setup */: // TODO
					if(!out['!margins']) default_margins(out['!margins'] = {});
					out['!margins'].header = val.header;
					out['!margins'].footer = val.footer;
					break;
				case 0x023e /* Window2 */: // TODO
					// $FlowIgnore
					if(val.RTL) Workbook.Views[0].RTL = true;
					break;
				case 0x0092 /* Palette */: palette = val; break;
				case 0x0896 /* Theme */: themes = val; break;
				case 0x008c /* Country */: country = val; break;
				case 0x01ba /* CodeName */: {
					/*:: if(!Workbook.WBProps) Workbook.WBProps = {}; */
					if(!cur_sheet) Workbook.WBProps.CodeName = val || "ThisWorkbook";
					else wsprops.CodeName = val || wsprops.name;
				} break;
				case 0x0055 /* DefColWidth */:
				case 0x0225 /* DefaultRowHeight */:
				case 0x005e /* Uncalced */:
				case 0x01af /* Prot4Rev */: case 0x01bc /* Prot4RevPass */: /*TODO: Revision Control*/
				case 0x005b /* FileSharing */:
				case 0x00ff /* ExtSST */:
				case 0x0863 /* BookExt */:
				case 0x08a6 /* RichTextStream */:
				case 0x00e9 /* BkHim */:
				case 0x0060 /* Template */:
				case 0x00da /* BookBool */:
				case 0x0160 /* UsesELFs */:
				case 0x089a /* MTRSettings */:
				case 0x000b: case 0x020b /* Index */:
				case 0x105c /* ClrtClient */:
				case 0x001d /* Selection */:
				case 0x0014 /* Header */:
				case 0x0015 /* Footer */:
				case 0x0083 /* HCenter */:
				case 0x0084 /* VCenter */:
				case 0x004d /* Pls */:
				case 0x00ab /* GCW */:
				case 0x0094 /* LHRecord */:
				case 0x00d7 /* DBCell */:
				case 0x01c2 /* EntExU2 */:
				case 0x00b0 /* SxView */:
				case 0x00b1 /* Sxvd */:
				case 0x00b2 /* SXVI */:
				case 0x0100 /* SXVDEx */:
				case 0x00b4 /* SxIvd */:
				case 0x00cd /* SXString */:
				case 0x0097 /* Sync */:
				case 0x0087 /* Addin */:
				case 0x00c5 /* SXDI */:
				case 0x00b5 /* SXLI */:
				case 0x00f1 /* SXEx */:
				case 0x0802 /* QsiSXTag */:
				case 0x0868 /* Feat */:
				case 0x0867 /* FeatHdr */: case 0x0871 /* FeatHdr11 */:
				case 0x0872 /* Feature11 */: case 0x0878 /* Feature12 */: case 0x0877 /* List12 */:
				case 0x01c1 /* RecalcId */:
				case 0x0099 /* DxGCol */:
				case 0x1060 /* Fbi */: case 0x1068 /* Fbi2 */: case 0x1066 /* GelFrame */:
				case 0x0031 /* Font */:
				case 0x087c /* XFCRC */:
				case 0x0293 /* Style */:
				case 0x0892 /* StyleExt */:
				case 0x00dd /* ScenarioProtect */:
				case 0x0063 /* ObjProtect */:
				case 0x0879 /* CondFmt12 */:
				case 0x0236 /* Table */:
				case 0x088e /* TableStyles */:
				case 0x088f /* TableStyle */:
				case 0x0890 /* TableStyleElement */:
				case 0x00d5 /* SXStreamID */:
				case 0x00e3 /* SXVS */:
				case 0x0051 /* DConRef */:
				case 0x0864 /* SXAddl */:
				case 0x01b5 /* DConBin */:
				case 0x0052 /* DConName */:
				case 0x00b6 /* SXPI */:
				case 0x00fb /* SxFormat */:
				case 0x00f7 /* SxSelect */:
				case 0x00f0 /* SxRule */:
				case 0x00f2 /* SxFilt */:
				case 0x00f5 /* SxItm */:
				case 0x00f4 /* SxDXF */:
				case 0x00ae /* ScenMan */:
				case 0x0050 /* DCon */:
				case 0x086c /* CellWatch */:
				case 0x002a /* PrintRowCol */:
				case 0x002b /* PrintGrid */:
				case 0x0033 /* PrintSize */:
				case 0x0059 /* XCT */:
				case 0x005a /* CRN */:
				case 0x00a0 /* Scl */:
				case 0x0862 /* SheetExt */:
				case 0x01bd /* ObNoMacros */:
				case 0x00d3 /* ObProj */:
				case 0x0897 /* GUIDTypeLib */:
				case 0x080b /* WOpt */:
				case 0x00ef /* PhoneticInfo */:
				case 0x00de /* OleObjectSize */:
				case 0x088d /* DXF */:
				case 0x01be /* Dv */: case 0x01b2 /* DVal */:
				case 0x1051 /* BRAI */: case 0x1003 /* Series */: case 0x100d /* SeriesText */:
				case 0x0876 /* DConn */:
				case 0x00dc /* DbOrParamQry */:
				case 0x0803 /* DBQueryExt */:
				case 0x080a /* OleDbConn */:
				case 0x0804 /* ExtString */:
				case 0x104e /* IFmtRecord */:
				case 0x01b0 /* CondFmt */: case 0x01b1 /* CF */: case 0x087a /* CF12 */: case 0x087b /* CFEx */:
				case 0x01c0 /* Excel9File */:
				case 0x1001 /* Units */:
				case 0x00e1 /* InterfaceHdr' */: case 0x00c1 /* Mms */: case 0x00e2 /* InterfaceEnd */: case 0x0161 /* DSF */:
				case 0x009c /* BuiltInFnGroupCount */: /* 2.4.30 0x0E or 0x10 but excel 2011 generates 0x11? */ break;
				case 0x003d /* Window1 */: case 0x008d /* HideObj */: case 0x0082 /* GridSet */: case 0x0080 /* Guts */:
				case 0x01a9 /* UserBView */: case 0x01aa /* UserSViewBegin */: case 0x01ab /* UserSViewEnd */:
				case 0x0041 /* Pane */:
				case 0x1063 /* Dat */:
				case 0x1033 /* Begin */: case 0x1034 /* End */:
				case 0x0852 /* StartBlock */: case 0x0853 /* EndBlock */:
				case 0x1032 /* Frame */: case 0x101a /* Area */:
				case 0x101d /* Axis */: case 0x1021 /* AxisLine */: case 0x101e /* Tick */:
				case 0x1046 /* AxesUsed */:
				case 0x089d /* CrtLayout12 */: case 0x08a7 /* CrtLayout12A */: case 0x1022 /* CrtLink */: case 0x101c /* CrtLine */: case 0x089e /* CrtMlFrt */: case 0x089f /* CrtMlFrtContinue */:
				case 0x1007 /* LineFormat */: case 0x100a /* AreaFormat */:
				case 0x1002 /* Chart */: case 0x103a /* Chart3d */: case 0x105f /* Chart3DBarShape */: case 0x1014 /* ChartFormat */: case 0x0850 /* ChartFrtInfo */:
				case 0x1035 /* PlotArea */: case 0x1064 /* PlotGrowth */:
				case 0x1016 /* SeriesList */: case 0x104a /* SerParent */: case 0x104b /* SerAuxTrend */:
				case 0x1006 /* DataFormat */: case 0x1045 /* SerToCrt */: case 0x1026 /* FontX */:
				case 0x1020 /* CatSerRange */: case 0x1062 /* AxcExt */: case 0x105d /* SerFmt */:
				case 0x1044 /* ShtProps */:
				case 0x1024 /* DefaultText */: case 0x1025 /* Text */: case 0x0856 /* CatLab */:
				case 0x086b /* DataLabExtContents */:
				case 0x1015 /* Legend */: case 0x1043 /* LegendException */:
				case 0x1019 /* Pie */: case 0x101b /* Scatter */:
				case 0x100b /* PieFormat */: case 0x1009 /* MarkerFormat */:
				case 0x0854 /* StartObject */: case 0x0855 /* EndObject */:
				case 0x1050 /* AlRuns */: case 0x1027 /* ObjectLink */:
				case 0x1065 /* SIIndex */:
				case 0x100c /* AttachedLabel */: case 0x0857 /* YMult */:
				case 0x1018 /* Line */: case 0x1017 /* Bar */:
				case 0x103f /* Surf */:
				case 0x1041 /* AxisParent */:
				case 0x104f /* Pos */:
				case 0x101f /* ValueRange */:
				case 0x0810 /* SXViewEx9 */:
				case 0x0858 /* SXViewLink */:
				case 0x0859 /* PivotChartBits */:
				case 0x1048 /* SBaseRef */:
				case 0x08a5 /* TextPropsStream */:
				case 0x08c9 /* LnExt */:
				case 0x08ca /* MkrExt */:
				case 0x08cb /* CrtCoopt */:
				case 0x01ad /* Qsi */: case 0x0807 /* Qsif */: case 0x0806 /* Qsir */:
				case 0x0805 /* TxtQry */:
				case 0x009b /* FilterMode */:
				case 0x009e /* AutoFilter */: case 0x009d /* AutoFilterInfo */:
				case 0x087e /* AutoFilter12 */:
				case 0x0874 /* DropDownObjIds */:
				case 0x0090 /* Sort */:
				case 0x0895 /* SortData */:
				case 0x08a4 /* ShapePropsStream */:
				case 0x00ec /* MsoDrawing */: case 0x00eb /* MsoDrawingGroup*/: case 0x00ed /* MsoDrawingSelection */:
				case 0x0801 /* WebPub */: case 0x08c0 /* AutoWebPub */:
				case 0x089c /* HeaderFooter */: case 0x0866 /* HFPicture */: case 0x088b /* PLV */:
				case 0x001b /* HorizontalPageBreaks */: case 0x001a /* VerticalPageBreaks */:
				case 0x0040 /* Backup */: case 0x089b /* CompressPictures */: case 0x088c /* Compat12 */:
				case 0x003c /* 'Continue' */: case 0x087f /* 'ContinueFrt12' */:
				case 0x085a /* FrtFontList */: case 0x0851 /* 'FrtWrapper' */:
				case 0x00ea /* TabIdConf */: case 0x103e /* Radar */: case 0x1040 /* RadarArea */: case 0x103d /* DropBar */: case 'Intl': case 'CoordList': case 'SerAuxErrBar':
				case 0x0045 /* BIFF2FONTCLR */: case 0x001f /* BIFF2FMTCNT */: case 0x0032 /* BIFF2FONTXTRA */:
				case 0x0043 /* BIFF2XF */: case 0x0243 /* BIFF3XF */: case 0x0443 /* BIFF4XF */:
				case 0x0044 /* BIFF2XFINDEX */:
				case 0x0056 /* BIFF4FMTCNT */: case 0x0008 /* BIFF2ROW */: case 0x003e /* BIFF2WINDOW2 */:
				case 0x00af /* SCENARIO */: case 0x103c /* PicF */: case 0x086a /* DataLabExt */:
				case 0x01b9 /* Lel */: case 0x1061 /* BopPop */: case 0x1067 /* BopPopCustom */: case 0x0813 /* RealTimeData */:
				case 0x0095 /* LHNGraph */: case 0x009a /* FnGroupName */: case 0x00c2 /* AddMenu */: case 0x0098 /* LPr */:
				case 0x08c1 /* ListObj */: case 0x08c2 /* ListField */:
				case 0x013f /* RRSort */:
				case 0x0418 /* BigName */:
				case 0x00bf /* ToolbarHdr */: case 0x00c0/* ToolbarEnd */:
				case 0x0034 /* DDEObjName */:
				case 0x08d6 /* FRTArchId$ */: break;
				default: if(options.WTF) throw 'Unrecognized Record 0x' + RecordType.toString(16);
			}
		} else blob.l += length;
	}
	wb.SheetNames=keys(Directory).sort(function(a,b) { return Number(a) - Number(b); }).map(function(x){return Directory[x].name;});
	if(!options.bookSheets) wb.Sheets=Sheets;
	if(!wb.SheetNames.length && Preamble["!ref"]) {
		wb.SheetNames.push("Sheet1");
		/*jshint -W069 */
		if(wb.Sheets) wb.Sheets["Sheet1"] = Preamble;
		/*jshint +W069 */
	} else wb.Preamble=Preamble;
	if(wb.Sheets) FilterDatabases.forEach(function(r,i) { wb.Sheets[wb.SheetNames[i]]['!autofilter'] = r; });
	wb.Strings = sst;
	wb.SSF = SSF.get_table();
	if(opts.enc) wb.Encryption = opts.enc;
	if(themes) wb.Themes = themes;
	wb.Metadata = {};
	if(country !== undefined) wb.Metadata.Country = country;
	if(supbooks.names.length > 0) Workbook.Names = supbooks.names;
	wb.Workbook = Workbook;
	return wb;
}

/* TODO: split props*/
var PSCLSID = {
	SI: "e0859ff2f94f6810ab9108002b27b3d9",
	DSI: "02d5cdd59c2e1b10939708002b2cf9ae",
	UDI: "05d5cdd59c2e1b10939708002b2cf9ae"
};
function parse_xls_props(cfb/*:CFBContainer*/, props, o) {
	/* [MS-OSHARED] 2.3.3.2.2 Document Summary Information Property Set */
	var DSI = CFB.find(cfb, '/!DocumentSummaryInformation');
	if(DSI && DSI.size > 0) try {
		var DocSummary = parse_PropertySetStream(DSI, DocSummaryPIDDSI, PSCLSID.DSI);
		for(var d in DocSummary) props[d] = DocSummary[d];
	} catch(e) {if(o.WTF) throw e;/* empty */}

	/* [MS-OSHARED] 2.3.3.2.1 Summary Information Property Set*/
	var SI = CFB.find(cfb, '/!SummaryInformation');
	if(SI && SI.size > 0) try {
		var Summary = parse_PropertySetStream(SI, SummaryPIDSI, PSCLSID.SI);
		for(var s in Summary) if(props[s] == null) props[s] = Summary[s];
	} catch(e) {if(o.WTF) throw e;/* empty */}

	if(props.HeadingPairs && props.TitlesOfParts) {
		load_props_pairs(props.HeadingPairs, props.TitlesOfParts, props, o);
		delete props.HeadingPairs; delete props.TitlesOfParts;
	}
}
function write_xls_props(wb/*:Workbook*/, cfb/*:CFBContainer*/) {
	var DSEntries = [], SEntries = [], CEntries = [];
	var i = 0, Keys;
	var DocSummaryRE/*:{[key:string]:string}*/ = evert_key(DocSummaryPIDDSI, "n");
	var SummaryRE/*:{[key:string]:string}*/ = evert_key(SummaryPIDSI, "n");
	if(wb.Props) {
		Keys = keys(wb.Props);
		// $FlowIgnore
		for(i = 0; i < Keys.length; ++i) (Object.prototype.hasOwnProperty.call(DocSummaryRE, Keys[i]) ? DSEntries : Object.prototype.hasOwnProperty.call(SummaryRE, Keys[i]) ? SEntries : CEntries).push([Keys[i], wb.Props[Keys[i]]]);
	}
	if(wb.Custprops) {
		Keys = keys(wb.Custprops);
		// $FlowIgnore
		for(i = 0; i < Keys.length; ++i) if(!Object.prototype.hasOwnProperty.call((wb.Props||{}), Keys[i])) (Object.prototype.hasOwnProperty.call(DocSummaryRE, Keys[i]) ? DSEntries : Object.prototype.hasOwnProperty.call(SummaryRE, Keys[i]) ? SEntries : CEntries).push([Keys[i], wb.Custprops[Keys[i]]]);
	}
	var CEntries2 = [];
	for(i = 0; i < CEntries.length; ++i) {
		if(XLSPSSkip.indexOf(CEntries[i][0]) > -1 || PseudoPropsPairs.indexOf(CEntries[i][0]) > -1) continue;
		if(CEntries[i][1] == null) continue;
		CEntries2.push(CEntries[i]);
	}
	if(SEntries.length) CFB.utils.cfb_add(cfb, "/\u0005SummaryInformation", write_PropertySetStream(SEntries, PSCLSID.SI, SummaryRE, SummaryPIDSI));
	if(DSEntries.length || CEntries2.length) CFB.utils.cfb_add(cfb, "/\u0005DocumentSummaryInformation", write_PropertySetStream(DSEntries, PSCLSID.DSI, DocSummaryRE, DocSummaryPIDDSI, CEntries2.length ? CEntries2 : null, PSCLSID.UDI));
}

function parse_xlscfb(cfb/*:any*/, options/*:?ParseOpts*/)/*:Workbook*/ {
if(!options) options = {};
fix_read_opts(options);
reset_cp();
if(options.codepage) set_ansi(options.codepage);
var CompObj/*:?CFBEntry*/, WB/*:?any*/;
if(cfb.FullPaths) {
	if(CFB.find(cfb, '/encryption')) throw new Error("File is password-protected");
	CompObj = CFB.find(cfb, '!CompObj');
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
var /*::CompObjP, */WorkbookP/*:: :Workbook = XLSX.utils.book_new(); */;

var _data/*:?any*/;
if(CompObj) /*::CompObjP = */parse_compobj(CompObj);
if(options.bookProps && !options.bookSheets) WorkbookP = ({}/*:any*/);
else/*:: if(cfb instanceof CFBContainer) */ {
	var T = has_buf ? 'buffer' : 'array';
	if(WB && WB.content) WorkbookP = parse_workbook(WB.content, options);
	/* Quattro Pro 7-8 */
	else if((_data=CFB.find(cfb, 'PerfectOffice_MAIN')) && _data.content) WorkbookP = WK_.to_workbook(_data.content, (options.type = T, options));
	/* Quattro Pro 9 */
	else if((_data=CFB.find(cfb, 'NativeContent_MAIN')) && _data.content) WorkbookP = WK_.to_workbook(_data.content, (options.type = T, options));
	/* Works 4 for Mac */
	else if((_data=CFB.find(cfb, 'MN0')) && _data.content) throw new Error("Unsupported Works 4 for Mac file");
	else throw new Error("Cannot find Workbook stream");
	if(options.bookVBA && cfb.FullPaths && CFB.find(cfb, '/_VBA_PROJECT_CUR/VBA/dir')) WorkbookP.vbaraw = make_vba_xls(cfb);
}

var props = {};
if(cfb.FullPaths) parse_xls_props(/*::((*/cfb/*:: :any):CFBContainer)*/, props, options);

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
		case "xla": if(!o.bookType) o.bookType = "xla";
		/* falls through */
		case "biff8": wbpath = "/Workbook"; o.biff = 8; break;
		case "biff5": wbpath = "/Book"; o.biff = 5; break;
		default: throw new Error("invalid type " + o.bookType + " for XLS CFB");
	}
	CFB.utils.cfb_add(cfb, wbpath, write_biff_buf(wb, o));
	if(o.biff == 8 && (wb.Props || wb.Custprops)) write_xls_props(wb, cfb);
	// TODO: SI, DSI, CO
	if(o.biff == 8 && wb.vbaraw) fill_vba_xls(cfb, CFB.read(wb.vbaraw, {type: typeof wb.vbaraw == "string" ? "binary" : "buffer"}));
	return cfb;
}
