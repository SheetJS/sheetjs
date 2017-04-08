var attregexg2=/([\w:]+)=((?:")([^"]*)(?:")|(?:')([^']*)(?:'))/g;
var attregex2=/([\w:]+)=((?:")(?:[^"]*)(?:")|(?:')(?:[^']*)(?:'))/;
var _chr = function(c) { return String.fromCharCode(c); };
function xlml_parsexmltag(tag/*:string*/, skip_root/*:?boolean*/) {
	var words = tag.split(/\s+/);
	var z/*:any*/ = ([]/*:any*/); if(!skip_root) z[0] = words[0];
	if(words.length === 1) return z;
	var m = tag.match(attregexg2), y, j, w, i;
	if(m) for(i = 0; i != m.length; ++i) {
		y = m[i].match(attregex2);
/*:: if(!y || !y[2]) continue; */
		if((j=y[1].indexOf(":")) === -1) z[y[1]] = y[2].substr(1,y[2].length-2);
		else {
			if(y[1].substr(0,6) === "xmlns:") w = "xmlns"+y[1].substr(6);
			else w = y[1].substr(j+1);
			z[w] = y[2].substr(1,y[2].length-2);
		}
	}
	return z;
}
function xlml_parsexmltagobj(tag/*:string*/) {
	var words = tag.split(/\s+/);
	var z = {};
	if(words.length === 1) return z;
	var m = tag.match(attregexg2), y, j, w, i;
	if(m) for(i = 0; i != m.length; ++i) {
		y = m[i].match(attregex2);
/*:: if(!y || !y[2]) continue; */
		if((j=y[1].indexOf(":")) === -1) z[y[1]] = y[2].substr(1,y[2].length-2);
		else {
			if(y[1].substr(0,6) === "xmlns:") w = "xmlns"+y[1].substr(6);
			else w = y[1].substr(j+1);
			z[w] = y[2].substr(1,y[2].length-2);
		}
	}
	return z;
}

// ----

function xlml_format(format, value)/*:string*/ {
	var fmt = XLMLFormatMap[format] || unescapexml(format);
	if(fmt === "General") return SSF._general(value);
	return SSF.format(fmt, value);
}

function xlml_set_custprop(Custprops, Rn, cp, val/*:string*/) {
	var oval/*:any*/ = val;
	switch((cp[0].match(/dt:dt="([\w.]+)"/)||["",""])[1]) {
		case "boolean": oval = parsexmlbool(val); break;
		case "i2": case "int": oval = parseInt(val, 10); break;
		case "r4": case "float": oval = parseFloat(val); break;
		case "date": case "dateTime.tz": oval = parseDate(val); break;
		case "i8": case "string": case "fixed": case "uuid": case "bin.base64": break;
		default: throw new Error("bad custprop:" + cp[0]);
	}
	Custprops[unescapexml(Rn[3])] = oval;
}

function safe_format_xlml(cell/*:Cell*/, nf, o) {
	if(cell.t === 'z') return;
	try {
		if(cell.t === 'e') { cell.w = cell.w || BErr[cell.v]; }
		else if(nf === "General") {
			if(cell.t === 'n') {
				if((cell.v|0) === cell.v) cell.w = SSF._general_int(cell.v);
				else cell.w = SSF._general_num(cell.v);
			}
			else cell.w = SSF._general(cell.v);
		}
		else cell.w = xlml_format(nf||"General", cell.v);
		var z = XLMLFormatMap[nf]||nf||"General";
		if(o.cellNF) cell.z = z;
		if(o.cellDates && cell.t == 'n' && SSF.is_date(z)) {
			var _d = SSF.parse_date_code(cell.v); if(_d) { cell.t = 'd'; cell.v = new Date(Date.UTC(_d.y, _d.m-1,_d.d,_d.H,_d.M,_d.S,_d.u)); }
		}
	} catch(e) { if(o.WTF) throw e; }
}

function process_style_xlml(styles, stag, opts) {
	if(opts.cellStyles) {
		if(stag.Interior) {
			var I = stag.Interior;
			if(I.Pattern) I.patternType = XLMLPatternTypeMap[I.Pattern] || I.Pattern;
		}
	}
	styles[stag.ID] = stag;
}

/* TODO: there must exist some form of OSP-blessed spec */
function parse_xlml_data(xml, ss, data, cell/*:any*/, base, styles, csty, row, arrayf, o) {
	var nf = "General", sid = cell.StyleID, S = {}; o = o || {};
	var interiors = [];
	var i = 0;
	if(sid === undefined && row) sid = row.StyleID;
	if(sid === undefined && csty) sid = csty.StyleID;
	while(styles[sid] !== undefined) {
		if(styles[sid].nf) nf = styles[sid].nf;
		if(styles[sid].Interior) interiors.push(styles[sid].Interior);
		if(!styles[sid].Parent) break;
		sid = styles[sid].Parent;
	}
	switch(data.Type) {
		case 'Boolean':
			cell.t = 'b';
			cell.v = parsexmlbool(xml);
			break;
		case 'String':
			cell.t = 's'; cell.r = xlml_fixstr(unescapexml(xml));
			cell.v = xml.indexOf("<") > -1 ? unescapexml(ss) : cell.r;
			break;
		case 'DateTime':
			cell.v = (parseDate(xml) - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
			if(cell.v !== cell.v) cell.v = unescapexml(xml);
			else if(cell.v<60) cell.v = cell.v -1;
			if(!nf || nf == "General") nf = "yyyy-mm-dd";
			/* falls through */
		case 'Number':
			if(cell.v === undefined) cell.v=+xml;
			if(!cell.t) cell.t = 'n';
			break;
		case 'Error': cell.t = 'e'; cell.v = RBErr[xml]; cell.w = xml; break;
		default: cell.t = 's'; cell.v = xlml_fixstr(ss||xml); break;
	}
	safe_format_xlml(cell, nf, o);
	if(o.cellFormula != null) {
		if(cell.Formula) {
			var fstr = unescapexml(cell.Formula);
			/* strictly speaking, the leading = is required but some writers omit */
			if(fstr.charCodeAt(0) == 61 /* = */) fstr = fstr.substr(1);
			cell.f = rc_to_a1(fstr, base);
			delete cell.Formula;
			if(cell.ArrayRange == "RC") cell.F = rc_to_a1("RC:RC", base);
			else if(cell.ArrayRange) {
				cell.F = rc_to_a1(cell.ArrayRange, base);
				arrayf.push([safe_decode_range(cell.F), cell.F]);
			}
		} else {
			for(i = 0; i < arrayf.length; ++i)
				if(base.r >= arrayf[i][0].s.r && base.r <= arrayf[i][0].e.r)
					if(base.c >= arrayf[i][0].s.c && base.c <= arrayf[i][0].e.c)
						cell.F = arrayf[i][1];
		}
	}
	if(o.cellStyles) {
		interiors.forEach(function(x) {
			if(!S.patternType && x.patternType) S.patternType = x.patternType;
		});
		cell.s = S;
	}
	cell.ixfe = cell.StyleID !== undefined ? cell.StyleID : 'Default';
}

function xlml_clean_comment(comment/*:any*/) {
	comment.t = comment.v || "";
	comment.t = comment.t.replace(/\r\n/g,"\n").replace(/\r/g,"\n");
	comment.v = comment.w = comment.ixfe = undefined;
}

function xlml_normalize(d)/*:string*/ {
	if(has_buf &&/*::typeof Buffer !== "undefined" && d != null &&*/ Buffer.isBuffer(d)) return d.toString('utf8');
	if(typeof d === 'string') return d;
	throw new Error("Bad input format: expected Buffer or string");
}

/* TODO: Everything */
/* UOS uses CJK in tags */
var xlmlregex = /<(\/?)([^\s?>!\/:]*:|)([^\s?>]*[^\s?>\/])[^>]*>/mg;
//var xlmlregex = /<(\/?)([a-z0-9]*:|)(\w+)[^>]*>/mg;
function parse_xlml_xml(d, opts)/*:Workbook*/ {
	make_ssf(SSF);
	var str = debom(xlml_normalize(d));
	if(opts && opts.type == 'binary' && typeof cptable !== 'undefined') str = cptable.utils.decode(65001, char_codes(str));
	if(str.substr(0,1000).indexOf("<html") >= 0) return parse_html(str, opts);
	var Rn;
	var state = [], tmp;
	if(DENSE != null) opts.dense = DENSE;
	var sheets = {}, sheetnames = [], cursheet = (opts.dense ? [] : {}), sheetname = "";
	var table = {}, cell = ({}/*:any*/), row = {};
	var dtag = xlml_parsexmltag('<Data ss:Type="String">'), didx = 0;
	var c = 0, r = 0;
	var refguess = {s: {r:2000000, c:2000000}, e: {r:0, c:0} };
	var styles = {}, stag = {};
	var ss = "", fidx = 0;
	var mergecells = [];
	var Props = {}, Custprops = {}, pidx = 0, cp = {};
	var comments = [], comment = {};
	var cstys = [], csty, seencol = false;
	var arrayf = [];
	var rowinfo = [];
	var Workbook = { Sheets:[] }, wsprops = {};
	xlmlregex.lastIndex = 0;
	str = str.replace(/<!--([^\u2603]*?)-->/mg,"");
	while((Rn = xlmlregex.exec(str))) switch(Rn[3]) {
		case 'Data':
			if(state[state.length-1][1]) break;
			if(Rn[1]==='/') parse_xlml_data(str.slice(didx, Rn.index), ss, dtag, state[state.length-1][0]=="Comment"?comment:cell, {c:c,r:r}, styles, cstys[c], row, arrayf, opts);
			else { ss = ""; dtag = xlml_parsexmltag(Rn[0]); didx = Rn.index + Rn[0].length; }
			break;
		case 'Cell':
			if(Rn[1]==='/'){
				if(comments.length > 0) cell.c = comments;
				if((!opts.sheetRows || opts.sheetRows > r) && cell.v !== undefined) {
					if(opts.dense) {
						if(!cursheet[r]) cursheet[r] = [];
						cursheet[r][c] = cell;
					} else cursheet[encode_col(c) + encode_row(r)] = cell;
				}
				if(cell.HRef) {
					cell.l = {Target:cell.HRef, Tooltip:cell.HRefScreenTip};
					delete cell.HRef; delete cell.HRefScreenTip;
				}
				if(cell.MergeAcross || cell.MergeDown) {
					var cc = c + (parseInt(cell.MergeAcross,10)|0);
					var rr = r + (parseInt(cell.MergeDown,10)|0);
					mergecells.push({s:{c:c,r:r},e:{c:cc,r:rr}});
				}
				if(!opts.sheetStubs) { if(cell.MergeAcross) c = cc + 1; else ++c; }
				else if(cell.MergeAcross || cell.MergeDown) {
					/*:: if(!cc) cc = 0; if(!rr) rr = 0; */
					for(var cma = c; cma <= cc; ++cma) {
						for(var cmd = r; cmd <= rr; ++cmd) {
							if(cma > c || cmd > r) {
								if(opts.dense) {
									if(!cursheet[cmd]) cursheet[cmd] = [];
									cursheet[cmd][cma] = {t:'z'};
								} else cursheet[encode_col(cma) + encode_row(cmd)] = {t:'z'};
							}
						}
					}
					c = cc + 1;
				}
				else ++c;
			} else {
				cell = xlml_parsexmltagobj(Rn[0]);
				if(cell.Index) c = +cell.Index - 1;
				if(c < refguess.s.c) refguess.s.c = c;
				if(c > refguess.e.c) refguess.e.c = c;
				if(Rn[0].slice(-2) === "/>") ++c;
				comments = [];
			}
			break;
		case 'Row':
			if(Rn[1]==='/' || Rn[0].slice(-2) === "/>") {
				if(r < refguess.s.r) refguess.s.r = r;
				if(r > refguess.e.r) refguess.e.r = r;
				if(Rn[0].slice(-2) === "/>") {
					row = xlml_parsexmltag(Rn[0]);
					if(row.Index) r = +row.Index - 1;
				}
				c = 0; ++r;
			} else {
				row = xlml_parsexmltag(Rn[0]);
				if(row.Index) r = +row.Index - 1;
			}
			break;
		case 'Worksheet': /* TODO: read range from FullRows/FullColumns */
			if(Rn[1]==='/'){
				if((tmp=state.pop())[0]!==Rn[3]) throw new Error("Bad state: "+tmp.join("|"));
				sheetnames.push(sheetname);
				if(refguess.s.r <= refguess.e.r && refguess.s.c <= refguess.e.c) cursheet["!ref"] = encode_range(refguess);
				if(mergecells.length) cursheet["!merges"] = mergecells;
				if(cstys.length > 0) cursheet["!cols"] = cstys;
				if(rowinfo.length > 0) cursheet["!rows"] = rowinfo;
				sheets[sheetname] = cursheet;
			} else {
				refguess = {s: {r:2000000, c:2000000}, e: {r:0, c:0} };
				r = c = 0;
				state.push([Rn[3], false]);
				tmp = xlml_parsexmltag(Rn[0]);
				sheetname = unescapexml(tmp.Name);
				cursheet = (opts.dense ? [] : {});
				mergecells = [];
				arrayf = [];
				rowinfo = [];
				wsprops = {name:sheetname, Hidden:0};
				Workbook.Sheets.push(wsprops);
			}
			break;
		case 'Table':
			if(Rn[1]==='/'){if((tmp=state.pop())[0]!==Rn[3]) throw new Error("Bad state: "+tmp.join("|"));}
			else if(Rn[0].slice(-2) == "/>") break;
			else {
				table = xlml_parsexmltag(Rn[0]);
				state.push([Rn[3], false]);
				cstys = []; seencol = false;
			}
			break;

		case 'Style':
			if(Rn[1]==='/') process_style_xlml(styles, stag, opts);
			else stag = xlml_parsexmltag(Rn[0]);
			break;

		case 'NumberFormat':
			stag.nf = unescapexml(xlml_parsexmltag(Rn[0]).Format || "General");
			if(XLMLFormatMap[stag.nf]) stag.nf = XLMLFormatMap[stag.nf];
			for(var ssfidx = 0; ssfidx != 0x188; ++ssfidx) if(SSF._table[ssfidx] == stag.nf) break;
			if(ssfidx == 0x188) for(ssfidx = 0x39; ssfidx != 0x188; ++ssfidx) if(SSF._table[ssfidx] == null) { SSF.load(stag.nf, ssfidx); break; }
			break;

		case 'Column':
			if(state[state.length-1][0] !== 'Table') break;
			csty = xlml_parsexmltag(Rn[0]);
			csty.wpx = parseInt(csty.Width, 10);
			if(!seencol && csty.wpx > 10) {
				seencol = true; find_mdw_wpx(csty.wpx);
				for(var _col = 0; _col < cstys.length; ++_col) if(cstys[_col]) process_col(cstys[_col]);
			}
			if(seencol) process_col(csty);
			cstys[(csty.Index-1||cstys.length)] = csty;
			for(var i = 0; i < +csty.Span; ++i) cstys[cstys.length] = dup(csty);
			break;

		case 'NamedRange': break;
		case 'NamedCell': break;
		case 'B': break;
		case 'I': break;
		case 'U': break;
		case 'S': break;
		case 'Sub': break;
		case 'Sup': break;
		case 'Span': break;
		case 'Border': break;
		case 'Alignment': break;
		case 'Borders': break;
		case 'Font':
			if(Rn[0].slice(-2) === "/>") break;
			else if(Rn[1]==="/") ss += str.slice(fidx, Rn.index);
			else fidx = Rn.index + Rn[0].length;
			break;
		case 'Interior':
			if(!opts.cellStyles) break;
			stag.Interior = xlml_parsexmltag(Rn[0]);
			break;
		case 'Protection': break;

		case 'Author':
		case 'Title':
		case 'Description':
		case 'Created':
		case 'Keywords':
		case 'Subject':
		case 'Category':
		case 'Company':
		case 'LastAuthor':
		case 'LastSaved':
		case 'LastPrinted':
		case 'Version':
		case 'Revision':
		case 'TotalTime':
		case 'HyperlinkBase':
		case 'Manager':
			if(Rn[0].slice(-2) === "/>") break;
			else if(Rn[1]==="/") xlml_set_prop(Props, Rn[3], str.slice(pidx, Rn.index));
			else pidx = Rn.index + Rn[0].length;
			break;
		case 'Paragraphs': break;

		case 'Styles':
		case 'Workbook':
			if(Rn[1]==='/'){if((tmp=state.pop())[0]!==Rn[3]) throw new Error("Bad state: "+tmp.join("|"));}
			else state.push([Rn[3], false]);
			break;

		case 'Comment':
			if(Rn[1]==='/'){
				if((tmp=state.pop())[0]!==Rn[3]) throw new Error("Bad state: "+tmp.join("|"));
				xlml_clean_comment(comment);
				comments.push(comment);
			} else {
				state.push([Rn[3], false]);
				tmp = xlml_parsexmltag(Rn[0]);
				comment = {a:tmp.Author};
			}
			break;

		case 'Name': break;

		case 'ComponentOptions':
		case 'DocumentProperties':
		case 'CustomDocumentProperties':
		case 'OfficeDocumentSettings':
		case 'PivotTable':
		case 'PivotCache':
		case 'Names':
		case 'MapInfo':
		case 'PageBreaks':
		case 'QueryTable':
		case 'DataValidation':
		case 'AutoFilter':
		case 'Sorting':
		case 'Schema':
		case 'data':
		case 'ConditionalFormatting':
		case 'SmartTagType':
		case 'SmartTags':
		case 'ExcelWorkbook':
		case 'WorkbookOptions':
		case 'WorksheetOptions':
			if(Rn[1]==='/'){if((tmp=state.pop())[0]!==Rn[3]) throw new Error("Bad state: "+tmp.join("|"));}
			else if(Rn[0].charAt(Rn[0].length-2) !== '/') state.push([Rn[3], true]);
			break;

		default:
			/* FODS file root is <office:document> */
			if(state.length == 0 && Rn[3] == "document") return parse_fods(str, opts);
			/* UOS file root is <uof:UOF> */
			if(state.length == 0 && Rn[3] == "UOF") return parse_fods(str, opts);

			var seen = true;
			switch(state[state.length-1][0]) {
				/* OfficeDocumentSettings */
				case 'OfficeDocumentSettings': switch(Rn[3]) {
					case 'AllowPNG': break;
					case 'RemovePersonalInformation': break;
					case 'DownloadComponents': break;
					case 'LocationOfComponents': break;
					case 'Colors': break;
					case 'Color': break;
					case 'Index': break;
					case 'RGB': break;
					case 'PixelsPerInch': break;
					case 'TargetScreenSize': break;
					case 'ReadOnlyRecommended': break;
					default: seen = false;
				} break;

				/* ComponentOptions */
				case 'ComponentOptions': switch(Rn[3]) {
					case 'Toolbar': break;
					case 'HideOfficeLogo': break;
					case 'SpreadsheetAutoFit': break;
					case 'Label': break;
					case 'Caption': break;
					case 'MaxHeight': break;
					case 'MaxWidth': break;
					case 'NextSheetNumber': break;
					default: seen = false;
				} break;

				/* ExcelWorkbook */
				case 'ExcelWorkbook': switch(Rn[3]) {
					case 'WindowHeight': break;
					case 'WindowWidth': break;
					case 'WindowTopX': break;
					case 'WindowTopY': break;
					case 'TabRatio': break;
					case 'ProtectStructure': break;
					case 'ProtectWindows': break;
					case 'ActiveSheet': break;
					case 'DisplayInkNotes': break;
					case 'FirstVisibleSheet': break;
					case 'SupBook': break;
					case 'SheetName': break;
					case 'SheetIndex': break;
					case 'SheetIndexFirst': break;
					case 'SheetIndexLast': break;
					case 'Dll': break;
					case 'AcceptLabelsInFormulas': break;
					case 'DoNotSaveLinkValues': break;
					case 'Date1904': break;
					case 'Iteration': break;
					case 'MaxIterations': break;
					case 'MaxChange': break;
					case 'Path': break;
					case 'Xct': break;
					case 'Count': break;
					case 'SelectedSheets': break;
					case 'Calculation': break;
					case 'Uncalced': break;
					case 'StartupPrompt': break;
					case 'Crn': break;
					case 'ExternName': break;
					case 'Formula': break;
					case 'ColFirst': break;
					case 'ColLast': break;
					case 'WantAdvise': break;
					case 'Boolean': break;
					case 'Error': break;
					case 'Text': break;
					case 'OLE': break;
					case 'NoAutoRecover': break;
					case 'PublishObjects': break;
					case 'DoNotCalculateBeforeSave': break;
					case 'Number': break;
					case 'RefModeR1C1': break;
					case 'EmbedSaveSmartTags': break;
					default: seen = false;
				} break;

				/* WorkbookOptions */
				case 'WorkbookOptions': switch(Rn[3]) {
					case 'OWCVersion': break;
					case 'Height': break;
					case 'Width': break;
					default: seen = false;
				} break;

				/* WorksheetOptions */
				case 'WorksheetOptions': switch(Rn[3]) {
					case 'Visible':
						if(Rn[0].slice(-2) === "/>"){}
						else if(Rn[1]==="/") switch(str.slice(pidx, Rn.index)) {
							case "SheetHidden": wsprops.Hidden = 1; break;
							case "SheetVeryHidden": wsprops.Hidden = 2; break;
						}
						else pidx = Rn.index + Rn[0].length;
						break;
					case 'Unsynced': break;
					case 'Print': break;
					case 'Panes': break;
					case 'Scale': break;
					case 'Pane': break;
					case 'Number': break;
					case 'Layout': break;
					case 'Header': break;
					case 'Footer': break;
					case 'PageSetup': break;
					case 'PageMargins': break;
					case 'Selected': break;
					case 'ProtectObjects': break;
					case 'EnableSelection': break;
					case 'ProtectScenarios': break;
					case 'ValidPrinterInfo': break;
					case 'HorizontalResolution': break;
					case 'VerticalResolution': break;
					case 'NumberofCopies': break;
					case 'ActiveRow': break;
					case 'ActiveCol': break;
					case 'ActivePane': break;
					case 'TopRowVisible': break;
					case 'TopRowBottomPane': break;
					case 'LeftColumnVisible': break;
					case 'LeftColumnRightPane': break;
					case 'FitToPage': break;
					case 'RangeSelection': break;
					case 'PaperSizeIndex': break;
					case 'PageLayoutZoom': break;
					case 'PageBreakZoom': break;
					case 'FilterOn': break;
					case 'DoNotDisplayGridlines': break;
					case 'SplitHorizontal': break;
					case 'SplitVertical': break;
					case 'FreezePanes': break;
					case 'FrozenNoSplit': break;
					case 'FitWidth': break;
					case 'FitHeight': break;
					case 'CommentsLayout': break;
					case 'Zoom': break;
					case 'LeftToRight': break;
					case 'Gridlines': break;
					case 'AllowSort': break;
					case 'AllowFilter': break;
					case 'AllowInsertRows': break;
					case 'AllowDeleteRows': break;
					case 'AllowInsertCols': break;
					case 'AllowDeleteCols': break;
					case 'AllowInsertHyperlinks': break;
					case 'AllowFormatCells': break;
					case 'AllowSizeCols': break;
					case 'AllowSizeRows': break;
					case 'NoSummaryRowsBelowDetail': break;
					case 'TabColorIndex': break;
					case 'DoNotDisplayHeadings': break;
					case 'ShowPageLayoutZoom': break;
					case 'NoSummaryColumnsRightDetail': break;
					case 'BlackAndWhite': break;
					case 'DoNotDisplayZeros': break;
					case 'DisplayPageBreak': break;
					case 'RowColHeadings': break;
					case 'DoNotDisplayOutline': break;
					case 'NoOrientation': break;
					case 'AllowUsePivotTables': break;
					case 'ZeroHeight': break;
					case 'ViewableRange': break;
					case 'Selection': break;
					case 'ProtectContents': break;
					default: seen = false;
				} break;

				/* PivotTable */
				case 'PivotTable': case 'PivotCache': switch(Rn[3]) {
					case 'ImmediateItemsOnDrop': break;
					case 'ShowPageMultipleItemLabel': break;
					case 'CompactRowIndent': break;
					case 'Location': break;
					case 'PivotField': break;
					case 'Orientation': break;
					case 'LayoutForm': break;
					case 'LayoutSubtotalLocation': break;
					case 'LayoutCompactRow': break;
					case 'Position': break;
					case 'PivotItem': break;
					case 'DataType': break;
					case 'DataField': break;
					case 'SourceName': break;
					case 'ParentField': break;
					case 'PTLineItems': break;
					case 'PTLineItem': break;
					case 'CountOfSameItems': break;
					case 'Item': break;
					case 'ItemType': break;
					case 'PTSource': break;
					case 'CacheIndex': break;
					case 'ConsolidationReference': break;
					case 'FileName': break;
					case 'Reference': break;
					case 'NoColumnGrand': break;
					case 'NoRowGrand': break;
					case 'BlankLineAfterItems': break;
					case 'Hidden': break;
					case 'Subtotal': break;
					case 'BaseField': break;
					case 'MapChildItems': break;
					case 'Function': break;
					case 'RefreshOnFileOpen': break;
					case 'PrintSetTitles': break;
					case 'MergeLabels': break;
					case 'DefaultVersion': break;
					case 'RefreshName': break;
					case 'RefreshDate': break;
					case 'RefreshDateCopy': break;
					case 'VersionLastRefresh': break;
					case 'VersionLastUpdate': break;
					case 'VersionUpdateableMin': break;
					case 'VersionRefreshableMin': break;
					case 'Calculation': break;
					default: seen = false;
				} break;

				/* PageBreaks */
				case 'PageBreaks': switch(Rn[3]) {
					case 'ColBreaks': break;
					case 'ColBreak': break;
					case 'RowBreaks': break;
					case 'RowBreak': break;
					case 'ColStart': break;
					case 'ColEnd': break;
					case 'RowEnd': break;
					default: seen = false;
				} break;

				/* AutoFilter */
				case 'AutoFilter': switch(Rn[3]) {
					case 'AutoFilterColumn': break;
					case 'AutoFilterCondition': break;
					case 'AutoFilterAnd': break;
					case 'AutoFilterOr': break;
					default: seen = false;
				} break;

				/* QueryTable */
				case 'QueryTable': switch(Rn[3]) {
					case 'Id': break;
					case 'AutoFormatFont': break;
					case 'AutoFormatPattern': break;
					case 'QuerySource': break;
					case 'QueryType': break;
					case 'EnableRedirections': break;
					case 'RefreshedInXl9': break;
					case 'URLString': break;
					case 'HTMLTables': break;
					case 'Connection': break;
					case 'CommandText': break;
					case 'RefreshInfo': break;
					case 'NoTitles': break;
					case 'NextId': break;
					case 'ColumnInfo': break;
					case 'OverwriteCells': break;
					case 'DoNotPromptForFile': break;
					case 'TextWizardSettings': break;
					case 'Source': break;
					case 'Number': break;
					case 'Decimal': break;
					case 'ThousandSeparator': break;
					case 'TrailingMinusNumbers': break;
					case 'FormatSettings': break;
					case 'FieldType': break;
					case 'Delimiters': break;
					case 'Tab': break;
					case 'Comma': break;
					case 'AutoFormatName': break;
					case 'VersionLastEdit': break;
					case 'VersionLastRefresh': break;
					default: seen = false;
				} break;

				/* Sorting */
				case 'Sorting':
				/* ConditionalFormatting */
				case 'ConditionalFormatting':
				/* DataValidation */
				case 'DataValidation': switch(Rn[3]) {
					case 'Range': break;
					case 'Type': break;
					case 'Min': break;
					case 'Max': break;
					case 'Sort': break;
					case 'Descending': break;
					case 'Order': break;
					case 'CaseSensitive': break;
					case 'Value': break;
					case 'ErrorStyle': break;
					case 'ErrorMessage': break;
					case 'ErrorTitle': break;
					case 'CellRangeList': break;
					case 'InputMessage': break;
					case 'InputTitle': break;
					case 'ComboHide': break;
					case 'InputHide': break;
					case 'Condition': break;
					case 'Qualifier': break;
					case 'UseBlank': break;
					case 'Value1': break;
					case 'Value2': break;
					case 'Format': break;
					default: seen = false;
				} break;

				/* MapInfo (schema) */
				case 'MapInfo': case 'Schema': case 'data': switch(Rn[3]) {
					case 'Map': break;
					case 'Entry': break;
					case 'Range': break;
					case 'XPath': break;
					case 'Field': break;
					case 'XSDType': break;
					case 'FilterOn': break;
					case 'Aggregate': break;
					case 'ElementType': break;
					case 'AttributeType': break;
				/* These are from xsd (XML Schema Definition) */
					case 'schema':
					case 'element':
					case 'complexType':
					case 'datatype':
					case 'all':
					case 'attribute':
					case 'extends': break;

					case 'row': break;
					default: seen = false;
				} break;

				/* SmartTags (can be anything) */
				case 'SmartTags': break;

				default: seen = false; break;
			}
			if(seen) break;
			/* CustomDocumentProperties */
			if(!state[state.length-1][1]) throw 'Unrecognized tag: ' + Rn[3] + "|" + state.join("|");
			if(state[state.length-1][0]==='CustomDocumentProperties') {
				if(Rn[0].slice(-2) === "/>") break;
				else if(Rn[1]==="/") xlml_set_custprop(Custprops, Rn, cp, str.slice(pidx, Rn.index));
				else { cp = Rn; pidx = Rn.index + Rn[0].length; }
				break;
			}
			if(opts.WTF) throw 'Unrecognized tag: ' + Rn[3] + "|" + state.join("|");
	}
	var out = ({}/*:any*/);
	if(!opts.bookSheets && !opts.bookProps) out.Sheets = sheets;
	out.SheetNames = sheetnames;
	out.Workbook = Workbook;
	out.SSF = SSF.get_table();
	out.Props = Props;
	out.Custprops = Custprops;
	return out;
}

function parse_xlml(data, opts)/*:Workbook*/ {
	fix_read_opts(opts=opts||{});
	switch(opts.type||"base64") {
		case "base64": return parse_xlml_xml(Base64.decode(data), opts);
		case "binary": case "buffer": case "file": return parse_xlml_xml(data, opts);
		case "array": return parse_xlml_xml(data.map(_chr).join(""), opts);
	}
	/*:: throw new Error("unsupported type " + opts.type); */
}

/* TODO */
function write_props_xlml(wb, opts) {
	var o = [];
	/* DocumentProperties */
	if(wb.Props) o.push(xlml_write_docprops(wb.Props));
	/* CustomDocumentProperties */
	if(wb.Custprops) o.push(xlml_write_custprops(wb.Props, wb.Custprops));
	return o.join("");
}
/* TODO */
function write_wb_xlml(wb, opts) {
	/* OfficeDocumentSettings */
	/* ExcelWorkbook */
	return "";
}
/* TODO */
function write_sty_xlml(wb, opts)/*:string*/ {
	/* Styles */
	return "";
}
/* WorksheetOptions */
function write_ws_xlml_wsopts(ws/*:Worksheet*/, opts, idx/*:number*/, wb/*:Workbook*/)/*:string*/ {
	var o = [];
	/* PageSetup */
	if(wb && wb.Workbook && wb.Workbook.Sheets && wb.Workbook.Sheets[idx]) {
		/* Visible */
		if(!!wb.Workbook.Sheets[idx].Hidden) o.push("<Visible>" + (wb.Workbook.Sheets[idx].Hidden == 1 ? "SheetHidden" : "SheetVeryHidden") + "</Visible>");
		else {
			/* Selected */
			for(var i = 0; i < idx; ++i) if(wb.Workbook.Sheets[i] && !wb.Workbook.Sheets[i].Hidden) break;
			if(i == idx) o.push("<Selected/>");
		}
	}
	if(o.length == 0) return "";
	return writextag("WorksheetOptions", o.join(""), {xmlns:XLMLNS.x});
}
function write_ws_xlml_comment(comments) {
	return comments.map(function(c) {
		// TODO: formatted text
		var t = xlml_unfixstr(c.t||"");
		var d =writextag("ss:Data", t, {"xmlns":"http://www.w3.org/TR/REC-html40"});
		return writextag("Comment", d, {"ss:Author":c.a});
	}).join("");
}
function write_ws_xlml_cell(cell, ref, ws, opts, idx, wb, addr)/*:string*/{
	if(!cell || cell.v == undefined && cell.f == undefined) return "<Cell></Cell>";

	var attr = {};
	if(cell.f) attr["ss:Formula"] = "=" + escapexml(a1_to_rc(cell.f, addr));
	if(cell.F && cell.F.substr(0, ref.length) == ref) {
		var end = decode_cell(cell.F.substr(ref.length + 1));
		attr["ss:ArrayRange"] = "RC:R" + (end.r == addr.r ? "" : "[" + (end.r - addr.r) + "]") + "C" + (end.c == addr.c ? "" : "[" + (end.c - addr.c) + "]");
	}

	if(cell.l && cell.l.Target) {
		attr["ss:HRef"] = escapexml(cell.l.Target);
		if(cell.l.Tooltip) attr["x:HRefScreenTip"] = escapexml(cell.l.Tooltip);
	}

	if(ws['!merges']) {
		var marr = ws['!merges'];
		for(var mi = 0; mi != marr.length; ++mi) {
			if(marr[mi].s.c != addr.c || marr[mi].s.r != addr.r) continue;
			if(marr[mi].e.c > marr[mi].s.c) attr['ss:MergeAcross'] = marr[mi].e.c - marr[mi].s.c;
			if(marr[mi].e.r > marr[mi].s.r) attr['ss:MergeDown'] = marr[mi].e.r - marr[mi].s.r;
		}
	}

	var t = "", p = "";
	switch(cell.t) {
		case 'z': return "";
		case 'n': t = 'Number'; p = String(cell.v); break;
		case 'b': t = 'Boolean'; p = (cell.v ? "1" : "0"); break;
		case 'e': t = 'Error'; p = BErr[cell.v]; break;
		case 'd': t = 'DateTime'; p = new Date(cell.v).toISOString(); break;
		case 's': t = 'String'; p = escapexml(cell.v||""); break;
	}
	var _v = (cell.v != null ? p : "");
	if(opts && opts.type == 'binary' && typeof cptable !== 'undefined' && cell.t == 's') {
		_v = cptable.utils.encode(65001, _v);
		var __v = "";
		for(var __i = 0; __i < _v.length; ++__i) __v += String.fromCharCode(_v[__i]);
		_v = __v;
	}
	var m = '<Data ss:Type="' + t + '">' + _v + '</Data>';

	if((cell.c||[]).length > 0) m += write_ws_xlml_comment(cell.c);

	return writextag("Cell", m, attr);
}
/* TODO */
function write_ws_xlml_table(ws/*:Worksheet*/, opts, idx/*:number*/, wb/*:Workbook*/)/*:string*/ {
	if(!ws['!ref']) return "";
	var range = safe_decode_range(ws['!ref']);
	var marr = ws['!merges'] || [], mi = 0;
	var o = [];
	if(ws['!cols']) ws['!cols'].forEach(function(n, i) {
		var p = col_obj_w(i, n);
		o.push(writextag("Column",null, {"ss:Index":i+1, "ss:Width":width2px(p.width)}));
	});
	var dense = Array.isArray(ws);
	for(var R = range.s.r; R <= range.e.r; ++R) {
		var row = ['<Row ss:Index="' + (R+1) + '">'];
		for(var C = range.s.c; C <= range.e.c; ++C) {
			var skip = false;
			for(mi = 0; mi != marr.length; ++mi) {
				if(marr[mi].s.c > C) continue;
				if(marr[mi].s.r > R) continue;
				if(marr[mi].e.c < C) continue;
				if(marr[mi].e.r < R) continue;
				if(marr[mi].s.c != C || marr[mi].s.r != R) skip = true;
				break;
			}
			if(skip) continue;
			var addr = {r:R,c:C};
			var ref = encode_cell(addr), cell = dense ? (ws[R]||[])[C] : ws[ref];
			row.push(write_ws_xlml_cell(cell, ref, ws, opts, idx, wb, addr));
		}
		row.push("</Row>");
		if(row.length > 2) o.push(row.join(""));
	}
	return o.join("");
}
function write_ws_xlml(idx/*:number*/, opts, wb/*:Workbook*/)/*:string*/ {
	var o = [];
	var s = wb.SheetNames[idx];
	var ws = wb.Sheets[s];

	/* Table */
	var t = ws ? write_ws_xlml_table(ws, opts, idx, wb) : "";
	if(t.length > 0) o.push("<Table>" + t + "</Table>");

	/* WorksheetOptions */
	o.push(write_ws_xlml_wsopts(ws, opts, idx, wb));

	return o.join("");
}
function write_xlml(wb, opts)/*:string*/ {
	var d = [];
	d.push(write_props_xlml(wb, opts));
	d.push(write_wb_xlml(wb, opts));
	d.push(write_sty_xlml(wb, opts));
	for(var i = 0; i < wb.SheetNames.length; ++i)
		d.push(writextag("Worksheet", write_ws_xlml(i, opts, wb), {"ss:Name":escapexml(wb.SheetNames[i])}));
	return XML_HEADER + writextag("Workbook", d.join(""), {
		'xmlns':      XLMLNS.ss,
		'xmlns:o':    XLMLNS.o,
		'xmlns:x':    XLMLNS.x,
		'xmlns:ss':   XLMLNS.ss,
		'xmlns:dt':   XLMLNS.dt,
		'xmlns:html': XLMLNS.html
	});
}
