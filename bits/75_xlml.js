var attregexg2=/([\w:]+)=((?:")([^"]*)(?:")|(?:')([^']*)(?:'))/g;
var attregex2=/([\w:]+)=((?:")(?:[^"]*)(?:")|(?:')(?:[^']*)(?:'))/;
function xlml_parsexmltag(tag/*:string*/, skip_root/*:?boolean*/) {
	var words = tag.split(/\s+/);
	var z/*:any*/ = ([]/*:any*/); if(!skip_root) z[0] = words[0];
	if(words.length === 1) return z;
	var m = tag.match(attregexg2), y, j, w, i;
	if(m) for(i = 0; i != m.length; ++i) {
		y = m[i].match(attregex2);
/*:: if(!y || !y[2]) continue; */
		if((j=y[1].indexOf(":")) === -1) z[y[1]] = y[2].slice(1,y[2].length-1);
		else {
			if(y[1].slice(0,6) === "xmlns:") w = "xmlns"+y[1].slice(6);
			else w = y[1].slice(j+1);
			z[w] = y[2].slice(1,y[2].length-1);
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
		if((j=y[1].indexOf(":")) === -1) z[y[1]] = y[2].slice(1,y[2].length-1);
		else {
			if(y[1].slice(0,6) === "xmlns:") w = "xmlns"+y[1].slice(6);
			else w = y[1].slice(j+1);
			z[w] = y[2].slice(1,y[2].length-1);
		}
	}
	return z;
}

// ----

/* map from xlml named formats to SSF TODO: localize */
var XLMLFormatMap/*: {[string]:string}*/;

function xlml_format(format, value)/*:string*/ {
	var fmt = XLMLFormatMap[format] || unescapexml(format);
	if(fmt === "General") return SSF_general(value);
	return SSF_format(fmt, value);
}

function xlml_set_custprop(Custprops, key, cp, val/*:string*/) {
	var oval/*:any*/ = val;
	switch((cp[0].match(/dt:dt="([\w.]+)"/)||["",""])[1]) {
		case "boolean": oval = parsexmlbool(val); break;
		case "i2": case "int": oval = parseInt(val, 10); break;
		case "r4": case "float": oval = parseFloat(val); break;
		case "date": case "dateTime.tz": oval = parseDate(val); break;
		case "i8": case "string": case "fixed": case "uuid": case "bin.base64": break;
		default: throw new Error("bad custprop:" + cp[0]);
	}
	Custprops[unescapexml(key)] = oval;
}

function safe_format_xlml(cell/*:Cell*/, nf, o) {
	if(cell.t === 'z') return;
	if(!o || o.cellText !== false) try {
		if(cell.t === 'e') { cell.w = cell.w || BErr[cell.v]; }
		else if(nf === "General") {
			if(cell.t === 'n') {
				if((cell.v|0) === cell.v) cell.w = cell.v.toString(10);
				else cell.w = SSF_general_num(cell.v);
			}
			else cell.w = SSF_general(cell.v);
		}
		else cell.w = xlml_format(nf||"General", cell.v);
	} catch(e) { if(o.WTF) throw e; }
	try {
		var z = XLMLFormatMap[nf]||nf||"General";
		if(o.cellNF) cell.z = z;
		if(o.cellDates && cell.t == 'n' && fmt_is_date(z)) {
			var _d = SSF_parse_date_code(cell.v); if(_d) { cell.t = 'd'; cell.v = new Date(_d.y, _d.m-1,_d.d,_d.H,_d.M,_d.S,_d.u); }
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
			cell.v = (xml.indexOf("<") > -1 ? unescapexml(ss||xml).replace(/<.*?>/g, "") : cell.r); // todo: BR etc
			break;
		case 'DateTime':
			if(xml.slice(-1) != "Z") xml += "Z";
			cell.v = (parseDate(xml) - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
			if(cell.v !== cell.v) cell.v = unescapexml(xml);
			else if(cell.v<60) cell.v = cell.v -1;
			if(!nf || nf == "General") nf = "yyyy-mm-dd";
			/* falls through */
		case 'Number':
			if(cell.v === undefined) cell.v=+xml;
			if(!cell.t) cell.t = 'n';
			break;
		case 'Error': cell.t = 'e'; cell.v = RBErr[xml]; if(o.cellText !== false) cell.w = xml; break;
		default:
			if(xml == "" && ss == "") { cell.t = 'z'; }
			else { cell.t = 's'; cell.v = xlml_fixstr(ss||xml); }
			break;
	}
	safe_format_xlml(cell, nf, o);
	if(o.cellFormula !== false) {
		if(cell.Formula) {
			var fstr = unescapexml(cell.Formula);
			/* strictly speaking, the leading = is required but some writers omit */
			if(fstr.charCodeAt(0) == 61 /* = */) fstr = fstr.slice(1);
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
	if(cell.StyleID !== undefined) cell.ixfe = cell.StyleID;
}

function xlml_prefix_dname(dname) {
	return XLSLblBuiltIn.indexOf("_xlnm." + dname) > -1 ? "_xlnm." + dname : dname;
}

function xlml_clean_comment(comment/*:any*/) {
	comment.t = comment.v || "";
	comment.t = comment.t.replace(/\r\n/g,"\n").replace(/\r/g,"\n");
	comment.v = comment.w = comment.ixfe = undefined;
}

/* TODO: Everything */
function parse_xlml_xml(d, _opts)/*:Workbook*/ {
	var opts = _opts || {};
	make_ssf();
	var str = debom(xlml_normalize(d));
	if(opts.type == 'binary' || opts.type == 'array' || opts.type == 'base64') {
		if(typeof $cptable !== 'undefined') str = $cptable.utils.decode(65001, char_codes(str));
		else str = utf8read(str);
	}
	var opening = str.slice(0, 1024).toLowerCase(), ishtml = false;
	opening = opening.replace(/".*?"/g, "");
	if((opening.indexOf(">") & 1023) > Math.min((opening.indexOf(",") & 1023), (opening.indexOf(";")&1023))) { var _o = dup(opts); _o.type = "string"; return PRN.to_workbook(str, _o); }
	if(opening.indexOf("<?xml") == -1) ["html", "table", "head", "meta", "script", "style", "div"].forEach(function(tag) { if(opening.indexOf("<" + tag) >= 0) ishtml = true; });
	if(ishtml) return html_to_workbook(str, opts);

	XLMLFormatMap = ({
		"General Number": "General",
		"General Date": table_fmt[22],
		"Long Date": "dddd, mmmm dd, yyyy",
		"Medium Date": table_fmt[15],
		"Short Date": table_fmt[14],
		"Long Time": table_fmt[19],
		"Medium Time": table_fmt[18],
		"Short Time": table_fmt[20],
		"Currency": '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
		"Fixed": table_fmt[2],
		"Standard": table_fmt[4],
		"Percent": table_fmt[10],
		"Scientific": table_fmt[11],
		"Yes/No": '"Yes";"Yes";"No";@',
		"True/False": '"True";"True";"False";@',
		"On/Off": '"Yes";"Yes";"No";@'
	}/*:any*/);


	var Rn;
	var state = [], tmp;
	if(DENSE != null && opts.dense == null) opts.dense = DENSE;
	var sheets = {}, sheetnames/*:Array<string>*/ = [], cursheet/*:Worksheet*/ = (opts.dense ? [] : {}), sheetname = "";
	var cell = ({}/*:any*/), row = {};// eslint-disable-line no-unused-vars
	var dtag = xlml_parsexmltag('<Data ss:Type="String">'), didx = 0;
	var c = 0, r = 0;
	var refguess/*:Range*/ = {s: {r:2000000, c:2000000}, e: {r:0, c:0} };
	var styles = {}, stag = {};
	var ss = "", fidx = 0;
	var merges/*:Array<Range>*/ = [];
	var Props = {}, Custprops = {}, pidx = 0, cp = [];
	var comments/*:Array<Comment>*/ = [], comment/*:Comment*/ = ({}/*:any*/);
	var cstys = [], csty, seencol = false;
	var arrayf/*:Array<[Range, string]>*/ = [];
	var rowinfo/*:Array<RowInfo>*/ = [], rowobj = {}, cc = 0, rr = 0;
	var Workbook/*:WBWBProps*/ = ({ Sheets:[], WBProps:{date1904:false} }/*:any*/), wsprops = {};
	xlmlregex.lastIndex = 0;
	str = str.replace(/<!--([\s\S]*?)-->/mg,"");
	var raw_Rn3 = "";
	while((Rn = xlmlregex.exec(str))) switch((Rn[3] = (raw_Rn3 = Rn[3]).toLowerCase())) {
		case 'data' /*case 'Data'*/:
			if(raw_Rn3 == "data") {
				if(Rn[1]==='/'){if((tmp=state.pop())[0]!==Rn[3]) throw new Error("Bad state: "+tmp.join("|"));}
				else if(Rn[0].charAt(Rn[0].length-2) !== '/') state.push([Rn[3], true]);
				break;
			}
			if(state[state.length-1][1]) break;
			if(Rn[1]==='/') parse_xlml_data(str.slice(didx, Rn.index), ss, dtag, state[state.length-1][0]==/*"Comment"*/"comment"?comment:cell, {c:c,r:r}, styles, cstys[c], row, arrayf, opts);
			else { ss = ""; dtag = xlml_parsexmltag(Rn[0]); didx = Rn.index + Rn[0].length; }
			break;
		case 'cell' /*case 'Cell'*/:
			if(Rn[1]==='/'){
				if(comments.length > 0) cell.c = comments;
				if((!opts.sheetRows || opts.sheetRows > r) && cell.v !== void 0) {
					if(opts.dense) {
						if(!cursheet[r]) cursheet[r] = [];
						cursheet[r][c] = cell;
					} else cursheet[encode_col(c) + encode_row(r)] = cell;
				}
				if(cell.HRef) {
					cell.l = ({Target:unescapexml(cell.HRef)}/*:any*/);
					if(cell.HRefScreenTip) cell.l.Tooltip = cell.HRefScreenTip;
					delete cell.HRef; delete cell.HRefScreenTip;
				}
				if(cell.MergeAcross || cell.MergeDown) {
					cc = c + (parseInt(cell.MergeAcross,10)|0);
					rr = r + (parseInt(cell.MergeDown,10)|0);
					if(cc > c || rr > r) merges.push({s:{c:c,r:r},e:{c:cc,r:rr}});
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
		case 'row' /*case 'Row'*/:
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
				rowobj = {};
				if(row.AutoFitHeight == "0" || row.Height) {
					rowobj.hpx = parseInt(row.Height, 10); rowobj.hpt = px2pt(rowobj.hpx);
					rowinfo[r] = rowobj;
				}
				if(row.Hidden == "1") { rowobj.hidden = true; rowinfo[r] = rowobj; }
			}
			break;
		case 'worksheet' /*case 'Worksheet'*/: /* TODO: read range from FullRows/FullColumns */
			if(Rn[1]==='/'){
				if((tmp=state.pop())[0]!==Rn[3]) throw new Error("Bad state: "+tmp.join("|"));
				sheetnames.push(sheetname);
				if(refguess.s.r <= refguess.e.r && refguess.s.c <= refguess.e.c) {
					cursheet["!ref"] = encode_range(refguess);
					if(opts.sheetRows && opts.sheetRows <= refguess.e.r) {
						cursheet["!fullref"] = cursheet["!ref"];
						refguess.e.r = opts.sheetRows - 1;
						cursheet["!ref"] = encode_range(refguess);
					}
				}
				if(merges.length) cursheet["!merges"] = merges;
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
				merges = [];
				arrayf = [];
				rowinfo = [];
				wsprops = {name:sheetname, Hidden:0};
				Workbook.Sheets.push(wsprops);
			}
			break;
		case 'table' /*case 'Table'*/:
			if(Rn[1]==='/'){if((tmp=state.pop())[0]!==Rn[3]) throw new Error("Bad state: "+tmp.join("|"));}
			else if(Rn[0].slice(-2) == "/>") break;
			else {
				state.push([Rn[3], false]);
				cstys = []; seencol = false;
			}
			break;

		case 'style' /*case 'Style'*/:
			if(Rn[1]==='/') process_style_xlml(styles, stag, opts);
			else stag = xlml_parsexmltag(Rn[0]);
			break;

		case 'numberformat' /*case 'NumberFormat'*/:
			stag.nf = unescapexml(xlml_parsexmltag(Rn[0]).Format || "General");
			if(XLMLFormatMap[stag.nf]) stag.nf = XLMLFormatMap[stag.nf];
			for(var ssfidx = 0; ssfidx != 0x188; ++ssfidx) if(table_fmt[ssfidx] == stag.nf) break;
			if(ssfidx == 0x188) for(ssfidx = 0x39; ssfidx != 0x188; ++ssfidx) if(table_fmt[ssfidx] == null) { SSF__load(stag.nf, ssfidx); break; }
			break;

		case 'column' /*case 'Column'*/:
			if(state[state.length-1][0] !== /*'Table'*/'table') break;
			if(Rn[1]==='/') break;
			csty = xlml_parsexmltag(Rn[0]);
			if(csty.Hidden) { csty.hidden = true; delete csty.Hidden; }
			if(csty.Width) csty.wpx = parseInt(csty.Width, 10);
			if(!seencol && csty.wpx > 10) {
				seencol = true; MDW = DEF_MDW; //find_mdw_wpx(csty.wpx);
				for(var _col = 0; _col < cstys.length; ++_col) if(cstys[_col]) process_col(cstys[_col]);
			}
			if(seencol) process_col(csty);
			cstys[(csty.Index-1||cstys.length)] = csty;
			for(var i = 0; i < +csty.Span; ++i) cstys[cstys.length] = dup(csty);
			break;

		case 'namedrange' /*case 'NamedRange'*/:
			if(Rn[1]==='/') break;
			if(!Workbook.Names) Workbook.Names = [];
			var _NamedRange = parsexmltag(Rn[0]);
			var _DefinedName/*:DefinedName*/ = ({
				Name: xlml_prefix_dname(_NamedRange.Name),
				Ref: rc_to_a1(_NamedRange.RefersTo.slice(1), {r:0, c:0})
			}/*:any*/);
			if(Workbook.Sheets.length>0) _DefinedName.Sheet=Workbook.Sheets.length-1;
			/*:: if(Workbook.Names) */Workbook.Names.push(_DefinedName);
			break;

		case 'namedcell' /*case 'NamedCell'*/: break;
		case 'b' /*case 'B'*/: break;
		case 'i' /*case 'I'*/: break;
		case 'u' /*case 'U'*/: break;
		case 's' /*case 'S'*/: break;
		case 'em' /*case 'EM'*/: break;
		case 'h2' /*case 'H2'*/: break;
		case 'h3' /*case 'H3'*/: break;
		case 'sub' /*case 'Sub'*/: break;
		case 'sup' /*case 'Sup'*/: break;
		case 'span' /*case 'Span'*/: break;
		case 'alignment' /*case 'Alignment'*/:
			break;
		case 'borders' /*case 'Borders'*/: break;
		case 'border' /*case 'Border'*/: break;
		case 'font' /*case 'Font'*/:
			if(Rn[0].slice(-2) === "/>") break;
			else if(Rn[1]==="/") ss += str.slice(fidx, Rn.index);
			else fidx = Rn.index + Rn[0].length;
			break;
		case 'interior' /*case 'Interior'*/:
			if(!opts.cellStyles) break;
			stag.Interior = xlml_parsexmltag(Rn[0]);
			break;
		case 'protection' /*case 'Protection'*/: break;

		case 'author' /*case 'Author'*/:
		case 'title' /*case 'Title'*/:
		case 'description' /*case 'Description'*/:
		case 'created' /*case 'Created'*/:
		case 'keywords' /*case 'Keywords'*/:
		case 'subject' /*case 'Subject'*/:
		case 'category' /*case 'Category'*/:
		case 'company' /*case 'Company'*/:
		case 'lastauthor' /*case 'LastAuthor'*/:
		case 'lastsaved' /*case 'LastSaved'*/:
		case 'lastprinted' /*case 'LastPrinted'*/:
		case 'version' /*case 'Version'*/:
		case 'revision' /*case 'Revision'*/:
		case 'totaltime' /*case 'TotalTime'*/:
		case 'hyperlinkbase' /*case 'HyperlinkBase'*/:
		case 'manager' /*case 'Manager'*/:
		case 'contentstatus' /*case 'ContentStatus'*/:
		case 'identifier' /*case 'Identifier'*/:
		case 'language' /*case 'Language'*/:
		case 'appname' /*case 'AppName'*/:
			if(Rn[0].slice(-2) === "/>") break;
			else if(Rn[1]==="/") xlml_set_prop(Props, raw_Rn3, str.slice(pidx, Rn.index));
			else pidx = Rn.index + Rn[0].length;
			break;
		case 'paragraphs' /*case 'Paragraphs'*/: break;

		case 'styles' /*case 'Styles'*/:
		case 'workbook' /*case 'Workbook'*/:
			if(Rn[1]==='/'){if((tmp=state.pop())[0]!==Rn[3]) throw new Error("Bad state: "+tmp.join("|"));}
			else state.push([Rn[3], false]);
			break;

		case 'comment' /*case 'Comment'*/:
			if(Rn[1]==='/'){
				if((tmp=state.pop())[0]!==Rn[3]) throw new Error("Bad state: "+tmp.join("|"));
				xlml_clean_comment(comment);
				comments.push(comment);
			} else {
				state.push([Rn[3], false]);
				tmp = xlml_parsexmltag(Rn[0]);
				comment = ({a:tmp.Author}/*:any*/);
			}
			break;

		case 'autofilter' /*case 'AutoFilter'*/:
			if(Rn[1]==='/'){if((tmp=state.pop())[0]!==Rn[3]) throw new Error("Bad state: "+tmp.join("|"));}
			else if(Rn[0].charAt(Rn[0].length-2) !== '/') {
				var AutoFilter = xlml_parsexmltag(Rn[0]);
				cursheet['!autofilter'] = { ref:rc_to_a1(AutoFilter.Range).replace(/\$/g,"") };
				state.push([Rn[3], true]);
			}
			break;

		case 'name' /*case 'Name'*/: break;

		case 'datavalidation' /*case 'DataValidation'*/:
			if(Rn[1]==='/'){
				if((tmp=state.pop())[0]!==Rn[3]) throw new Error("Bad state: "+tmp.join("|"));
			} else {
				if(Rn[0].charAt(Rn[0].length-2) !== '/') state.push([Rn[3], true]);
			}
			break;

		case 'pixelsperinch' /*case 'PixelsPerInch'*/:
			break;
		case 'componentoptions' /*case 'ComponentOptions'*/:
		case 'documentproperties' /*case 'DocumentProperties'*/:
		case 'customdocumentproperties' /*case 'CustomDocumentProperties'*/:
		case 'officedocumentsettings' /*case 'OfficeDocumentSettings'*/:
		case 'pivottable' /*case 'PivotTable'*/:
		case 'pivotcache' /*case 'PivotCache'*/:
		case 'names' /*case 'Names'*/:
		case 'mapinfo' /*case 'MapInfo'*/:
		case 'pagebreaks' /*case 'PageBreaks'*/:
		case 'querytable' /*case 'QueryTable'*/:
		case 'sorting' /*case 'Sorting'*/:
		case 'schema' /*case 'Schema'*/: //case 'data' /*case 'data'*/:
		case 'conditionalformatting' /*case 'ConditionalFormatting'*/:
		case 'smarttagtype' /*case 'SmartTagType'*/:
		case 'smarttags' /*case 'SmartTags'*/:
		case 'excelworkbook' /*case 'ExcelWorkbook'*/:
		case 'workbookoptions' /*case 'WorkbookOptions'*/:
		case 'worksheetoptions' /*case 'WorksheetOptions'*/:
			if(Rn[1]==='/'){if((tmp=state.pop())[0]!==Rn[3]) throw new Error("Bad state: "+tmp.join("|"));}
			else if(Rn[0].charAt(Rn[0].length-2) !== '/') state.push([Rn[3], true]);
			break;

		case 'null' /*case 'Null'*/: break;

		default:
			/* FODS file root is <office:document> */
			if(state.length == 0 && Rn[3] == "document") return parse_fods(str, opts);
			/* UOS file root is <uof:UOF> */
			if(state.length == 0 && Rn[3] == "uof"/*"UOF"*/) return parse_fods(str, opts);

			var seen = true;
			switch(state[state.length-1][0]) {
				/* OfficeDocumentSettings */
				case 'officedocumentsettings' /*case 'OfficeDocumentSettings'*/: switch(Rn[3]) {
					case 'allowpng' /*case 'AllowPNG'*/: break;
					case 'removepersonalinformation' /*case 'RemovePersonalInformation'*/: break;
					case 'downloadcomponents' /*case 'DownloadComponents'*/: break;
					case 'locationofcomponents' /*case 'LocationOfComponents'*/: break;
					case 'colors' /*case 'Colors'*/: break;
					case 'color' /*case 'Color'*/: break;
					case 'index' /*case 'Index'*/: break;
					case 'rgb' /*case 'RGB'*/: break;
					case 'targetscreensize' /*case 'TargetScreenSize'*/: break;
					case 'readonlyrecommended' /*case 'ReadOnlyRecommended'*/: break;
					default: seen = false;
				} break;

				/* ComponentOptions */
				case 'componentoptions' /*case 'ComponentOptions'*/: switch(Rn[3]) {
					case 'toolbar' /*case 'Toolbar'*/: break;
					case 'hideofficelogo' /*case 'HideOfficeLogo'*/: break;
					case 'spreadsheetautofit' /*case 'SpreadsheetAutoFit'*/: break;
					case 'label' /*case 'Label'*/: break;
					case 'caption' /*case 'Caption'*/: break;
					case 'maxheight' /*case 'MaxHeight'*/: break;
					case 'maxwidth' /*case 'MaxWidth'*/: break;
					case 'nextsheetnumber' /*case 'NextSheetNumber'*/: break;
					default: seen = false;
				} break;

				/* ExcelWorkbook */
				case 'excelworkbook' /*case 'ExcelWorkbook'*/: switch(Rn[3]) {
					case 'date1904' /*case 'Date1904'*/:
						/*:: if(!Workbook.WBProps) Workbook.WBProps = {}; */
						Workbook.WBProps.date1904 = true;
						break;
					case 'windowheight' /*case 'WindowHeight'*/: break;
					case 'windowwidth' /*case 'WindowWidth'*/: break;
					case 'windowtopx' /*case 'WindowTopX'*/: break;
					case 'windowtopy' /*case 'WindowTopY'*/: break;
					case 'tabratio' /*case 'TabRatio'*/: break;
					case 'protectstructure' /*case 'ProtectStructure'*/: break;
					case 'protectwindow' /*case 'ProtectWindow'*/: break;
					case 'protectwindows' /*case 'ProtectWindows'*/: break;
					case 'activesheet' /*case 'ActiveSheet'*/: break;
					case 'displayinknotes' /*case 'DisplayInkNotes'*/: break;
					case 'firstvisiblesheet' /*case 'FirstVisibleSheet'*/: break;
					case 'supbook' /*case 'SupBook'*/: break;
					case 'sheetname' /*case 'SheetName'*/: break;
					case 'sheetindex' /*case 'SheetIndex'*/: break;
					case 'sheetindexfirst' /*case 'SheetIndexFirst'*/: break;
					case 'sheetindexlast' /*case 'SheetIndexLast'*/: break;
					case 'dll' /*case 'Dll'*/: break;
					case 'acceptlabelsinformulas' /*case 'AcceptLabelsInFormulas'*/: break;
					case 'donotsavelinkvalues' /*case 'DoNotSaveLinkValues'*/: break;
					case 'iteration' /*case 'Iteration'*/: break;
					case 'maxiterations' /*case 'MaxIterations'*/: break;
					case 'maxchange' /*case 'MaxChange'*/: break;
					case 'path' /*case 'Path'*/: break;
					case 'xct' /*case 'Xct'*/: break;
					case 'count' /*case 'Count'*/: break;
					case 'selectedsheets' /*case 'SelectedSheets'*/: break;
					case 'calculation' /*case 'Calculation'*/: break;
					case 'uncalced' /*case 'Uncalced'*/: break;
					case 'startupprompt' /*case 'StartupPrompt'*/: break;
					case 'crn' /*case 'Crn'*/: break;
					case 'externname' /*case 'ExternName'*/: break;
					case 'formula' /*case 'Formula'*/: break;
					case 'colfirst' /*case 'ColFirst'*/: break;
					case 'collast' /*case 'ColLast'*/: break;
					case 'wantadvise' /*case 'WantAdvise'*/: break;
					case 'boolean' /*case 'Boolean'*/: break;
					case 'error' /*case 'Error'*/: break;
					case 'text' /*case 'Text'*/: break;
					case 'ole' /*case 'OLE'*/: break;
					case 'noautorecover' /*case 'NoAutoRecover'*/: break;
					case 'publishobjects' /*case 'PublishObjects'*/: break;
					case 'donotcalculatebeforesave' /*case 'DoNotCalculateBeforeSave'*/: break;
					case 'number' /*case 'Number'*/: break;
					case 'refmoder1c1' /*case 'RefModeR1C1'*/: break;
					case 'embedsavesmarttags' /*case 'EmbedSaveSmartTags'*/: break;
					default: seen = false;
				} break;

				/* WorkbookOptions */
				case 'workbookoptions' /*case 'WorkbookOptions'*/: switch(Rn[3]) {
					case 'owcversion' /*case 'OWCVersion'*/: break;
					case 'height' /*case 'Height'*/: break;
					case 'width' /*case 'Width'*/: break;
					default: seen = false;
				} break;

				/* WorksheetOptions */
				case 'worksheetoptions' /*case 'WorksheetOptions'*/: switch(Rn[3]) {
					case 'visible' /*case 'Visible'*/:
						if(Rn[0].slice(-2) === "/>"){/* empty */}
						else if(Rn[1]==="/") switch(str.slice(pidx, Rn.index)) {
							case "SheetHidden": wsprops.Hidden = 1; break;
							case "SheetVeryHidden": wsprops.Hidden = 2; break;
						}
						else pidx = Rn.index + Rn[0].length;
						break;
					case 'header' /*case 'Header'*/:
						if(!cursheet['!margins']) default_margins(cursheet['!margins']={}, 'xlml');
						if(!isNaN(+parsexmltag(Rn[0]).Margin)) cursheet['!margins'].header = +parsexmltag(Rn[0]).Margin;
						break;
					case 'footer' /*case 'Footer'*/:
						if(!cursheet['!margins']) default_margins(cursheet['!margins']={}, 'xlml');
						if(!isNaN(+parsexmltag(Rn[0]).Margin)) cursheet['!margins'].footer = +parsexmltag(Rn[0]).Margin;
						break;
					case 'pagemargins' /*case 'PageMargins'*/:
						var pagemargins = parsexmltag(Rn[0]);
						if(!cursheet['!margins']) default_margins(cursheet['!margins']={},'xlml');
						if(!isNaN(+pagemargins.Top)) cursheet['!margins'].top = +pagemargins.Top;
						if(!isNaN(+pagemargins.Left)) cursheet['!margins'].left = +pagemargins.Left;
						if(!isNaN(+pagemargins.Right)) cursheet['!margins'].right = +pagemargins.Right;
						if(!isNaN(+pagemargins.Bottom)) cursheet['!margins'].bottom = +pagemargins.Bottom;
						break;
					case 'displayrighttoleft' /*case 'DisplayRightToLeft'*/:
						if(!Workbook.Views) Workbook.Views = [];
						if(!Workbook.Views[0]) Workbook.Views[0] = {};
						Workbook.Views[0].RTL = true;
						break;

					case 'freezepanes' /*case 'FreezePanes'*/: break;
					case 'frozennosplit' /*case 'FrozenNoSplit'*/: break;

					case 'splithorizontal' /*case 'SplitHorizontal'*/:
					case 'splitvertical' /*case 'SplitVertical'*/:
						break;

					case 'donotdisplaygridlines' /*case 'DoNotDisplayGridlines'*/:
						break;

					case 'activerow' /*case 'ActiveRow'*/: break;
					case 'activecol' /*case 'ActiveCol'*/: break;
					case 'toprowbottompane' /*case 'TopRowBottomPane'*/: break;
					case 'leftcolumnrightpane' /*case 'LeftColumnRightPane'*/: break;

					case 'unsynced' /*case 'Unsynced'*/: break;
					case 'print' /*case 'Print'*/: break;
					case 'printerrors' /*case 'PrintErrors'*/: break;
					case 'panes' /*case 'Panes'*/: break;
					case 'scale' /*case 'Scale'*/: break;
					case 'pane' /*case 'Pane'*/: break;
					case 'number' /*case 'Number'*/: break;
					case 'layout' /*case 'Layout'*/: break;
					case 'pagesetup' /*case 'PageSetup'*/: break;
					case 'selected' /*case 'Selected'*/: break;
					case 'protectobjects' /*case 'ProtectObjects'*/: break;
					case 'enableselection' /*case 'EnableSelection'*/: break;
					case 'protectscenarios' /*case 'ProtectScenarios'*/: break;
					case 'validprinterinfo' /*case 'ValidPrinterInfo'*/: break;
					case 'horizontalresolution' /*case 'HorizontalResolution'*/: break;
					case 'verticalresolution' /*case 'VerticalResolution'*/: break;
					case 'numberofcopies' /*case 'NumberofCopies'*/: break;
					case 'activepane' /*case 'ActivePane'*/: break;
					case 'toprowvisible' /*case 'TopRowVisible'*/: break;
					case 'leftcolumnvisible' /*case 'LeftColumnVisible'*/: break;
					case 'fittopage' /*case 'FitToPage'*/: break;
					case 'rangeselection' /*case 'RangeSelection'*/: break;
					case 'papersizeindex' /*case 'PaperSizeIndex'*/: break;
					case 'pagelayoutzoom' /*case 'PageLayoutZoom'*/: break;
					case 'pagebreakzoom' /*case 'PageBreakZoom'*/: break;
					case 'filteron' /*case 'FilterOn'*/: break;
					case 'fitwidth' /*case 'FitWidth'*/: break;
					case 'fitheight' /*case 'FitHeight'*/: break;
					case 'commentslayout' /*case 'CommentsLayout'*/: break;
					case 'zoom' /*case 'Zoom'*/: break;
					case 'lefttoright' /*case 'LeftToRight'*/: break;
					case 'gridlines' /*case 'Gridlines'*/: break;
					case 'allowsort' /*case 'AllowSort'*/: break;
					case 'allowfilter' /*case 'AllowFilter'*/: break;
					case 'allowinsertrows' /*case 'AllowInsertRows'*/: break;
					case 'allowdeleterows' /*case 'AllowDeleteRows'*/: break;
					case 'allowinsertcols' /*case 'AllowInsertCols'*/: break;
					case 'allowdeletecols' /*case 'AllowDeleteCols'*/: break;
					case 'allowinserthyperlinks' /*case 'AllowInsertHyperlinks'*/: break;
					case 'allowformatcells' /*case 'AllowFormatCells'*/: break;
					case 'allowsizecols' /*case 'AllowSizeCols'*/: break;
					case 'allowsizerows' /*case 'AllowSizeRows'*/: break;
					case 'nosummaryrowsbelowdetail' /*case 'NoSummaryRowsBelowDetail'*/:
						if(!cursheet["!outline"]) cursheet["!outline"] = {};
						cursheet["!outline"].above = true;
						break;
					case 'tabcolorindex' /*case 'TabColorIndex'*/: break;
					case 'donotdisplayheadings' /*case 'DoNotDisplayHeadings'*/: break;
					case 'showpagelayoutzoom' /*case 'ShowPageLayoutZoom'*/: break;
					case 'nosummarycolumnsrightdetail' /*case 'NoSummaryColumnsRightDetail'*/:
						if(!cursheet["!outline"]) cursheet["!outline"] = {};
						cursheet["!outline"].left = true;
						break;
					case 'blackandwhite' /*case 'BlackAndWhite'*/: break;
					case 'donotdisplayzeros' /*case 'DoNotDisplayZeros'*/: break;
					case 'displaypagebreak' /*case 'DisplayPageBreak'*/: break;
					case 'rowcolheadings' /*case 'RowColHeadings'*/: break;
					case 'donotdisplayoutline' /*case 'DoNotDisplayOutline'*/: break;
					case 'noorientation' /*case 'NoOrientation'*/: break;
					case 'allowusepivottables' /*case 'AllowUsePivotTables'*/: break;
					case 'zeroheight' /*case 'ZeroHeight'*/: break;
					case 'viewablerange' /*case 'ViewableRange'*/: break;
					case 'selection' /*case 'Selection'*/: break;
					case 'protectcontents' /*case 'ProtectContents'*/: break;
					default: seen = false;
				} break;

				/* PivotTable */
				case 'pivottable' /*case 'PivotTable'*/: case 'pivotcache' /*case 'PivotCache'*/: switch(Rn[3]) {
					case 'immediateitemsondrop' /*case 'ImmediateItemsOnDrop'*/: break;
					case 'showpagemultipleitemlabel' /*case 'ShowPageMultipleItemLabel'*/: break;
					case 'compactrowindent' /*case 'CompactRowIndent'*/: break;
					case 'location' /*case 'Location'*/: break;
					case 'pivotfield' /*case 'PivotField'*/: break;
					case 'orientation' /*case 'Orientation'*/: break;
					case 'layoutform' /*case 'LayoutForm'*/: break;
					case 'layoutsubtotallocation' /*case 'LayoutSubtotalLocation'*/: break;
					case 'layoutcompactrow' /*case 'LayoutCompactRow'*/: break;
					case 'position' /*case 'Position'*/: break;
					case 'pivotitem' /*case 'PivotItem'*/: break;
					case 'datatype' /*case 'DataType'*/: break;
					case 'datafield' /*case 'DataField'*/: break;
					case 'sourcename' /*case 'SourceName'*/: break;
					case 'parentfield' /*case 'ParentField'*/: break;
					case 'ptlineitems' /*case 'PTLineItems'*/: break;
					case 'ptlineitem' /*case 'PTLineItem'*/: break;
					case 'countofsameitems' /*case 'CountOfSameItems'*/: break;
					case 'item' /*case 'Item'*/: break;
					case 'itemtype' /*case 'ItemType'*/: break;
					case 'ptsource' /*case 'PTSource'*/: break;
					case 'cacheindex' /*case 'CacheIndex'*/: break;
					case 'consolidationreference' /*case 'ConsolidationReference'*/: break;
					case 'filename' /*case 'FileName'*/: break;
					case 'reference' /*case 'Reference'*/: break;
					case 'nocolumngrand' /*case 'NoColumnGrand'*/: break;
					case 'norowgrand' /*case 'NoRowGrand'*/: break;
					case 'blanklineafteritems' /*case 'BlankLineAfterItems'*/: break;
					case 'hidden' /*case 'Hidden'*/: break;
					case 'subtotal' /*case 'Subtotal'*/: break;
					case 'basefield' /*case 'BaseField'*/: break;
					case 'mapchilditems' /*case 'MapChildItems'*/: break;
					case 'function' /*case 'Function'*/: break;
					case 'refreshonfileopen' /*case 'RefreshOnFileOpen'*/: break;
					case 'printsettitles' /*case 'PrintSetTitles'*/: break;
					case 'mergelabels' /*case 'MergeLabels'*/: break;
					case 'defaultversion' /*case 'DefaultVersion'*/: break;
					case 'refreshname' /*case 'RefreshName'*/: break;
					case 'refreshdate' /*case 'RefreshDate'*/: break;
					case 'refreshdatecopy' /*case 'RefreshDateCopy'*/: break;
					case 'versionlastrefresh' /*case 'VersionLastRefresh'*/: break;
					case 'versionlastupdate' /*case 'VersionLastUpdate'*/: break;
					case 'versionupdateablemin' /*case 'VersionUpdateableMin'*/: break;
					case 'versionrefreshablemin' /*case 'VersionRefreshableMin'*/: break;
					case 'calculation' /*case 'Calculation'*/: break;
					default: seen = false;
				} break;

				/* PageBreaks */
				case 'pagebreaks' /*case 'PageBreaks'*/: switch(Rn[3]) {
					case 'colbreaks' /*case 'ColBreaks'*/: break;
					case 'colbreak' /*case 'ColBreak'*/: break;
					case 'rowbreaks' /*case 'RowBreaks'*/: break;
					case 'rowbreak' /*case 'RowBreak'*/: break;
					case 'colstart' /*case 'ColStart'*/: break;
					case 'colend' /*case 'ColEnd'*/: break;
					case 'rowend' /*case 'RowEnd'*/: break;
					default: seen = false;
				} break;

				/* AutoFilter */
				case 'autofilter' /*case 'AutoFilter'*/: switch(Rn[3]) {
					case 'autofiltercolumn' /*case 'AutoFilterColumn'*/: break;
					case 'autofiltercondition' /*case 'AutoFilterCondition'*/: break;
					case 'autofilterand' /*case 'AutoFilterAnd'*/: break;
					case 'autofilteror' /*case 'AutoFilterOr'*/: break;
					default: seen = false;
				} break;

				/* QueryTable */
				case 'querytable' /*case 'QueryTable'*/: switch(Rn[3]) {
					case 'id' /*case 'Id'*/: break;
					case 'autoformatfont' /*case 'AutoFormatFont'*/: break;
					case 'autoformatpattern' /*case 'AutoFormatPattern'*/: break;
					case 'querysource' /*case 'QuerySource'*/: break;
					case 'querytype' /*case 'QueryType'*/: break;
					case 'enableredirections' /*case 'EnableRedirections'*/: break;
					case 'refreshedinxl9' /*case 'RefreshedInXl9'*/: break;
					case 'urlstring' /*case 'URLString'*/: break;
					case 'htmltables' /*case 'HTMLTables'*/: break;
					case 'connection' /*case 'Connection'*/: break;
					case 'commandtext' /*case 'CommandText'*/: break;
					case 'refreshinfo' /*case 'RefreshInfo'*/: break;
					case 'notitles' /*case 'NoTitles'*/: break;
					case 'nextid' /*case 'NextId'*/: break;
					case 'columninfo' /*case 'ColumnInfo'*/: break;
					case 'overwritecells' /*case 'OverwriteCells'*/: break;
					case 'donotpromptforfile' /*case 'DoNotPromptForFile'*/: break;
					case 'textwizardsettings' /*case 'TextWizardSettings'*/: break;
					case 'source' /*case 'Source'*/: break;
					case 'number' /*case 'Number'*/: break;
					case 'decimal' /*case 'Decimal'*/: break;
					case 'thousandseparator' /*case 'ThousandSeparator'*/: break;
					case 'trailingminusnumbers' /*case 'TrailingMinusNumbers'*/: break;
					case 'formatsettings' /*case 'FormatSettings'*/: break;
					case 'fieldtype' /*case 'FieldType'*/: break;
					case 'delimiters' /*case 'Delimiters'*/: break;
					case 'tab' /*case 'Tab'*/: break;
					case 'comma' /*case 'Comma'*/: break;
					case 'autoformatname' /*case 'AutoFormatName'*/: break;
					case 'versionlastedit' /*case 'VersionLastEdit'*/: break;
					case 'versionlastrefresh' /*case 'VersionLastRefresh'*/: break;
					default: seen = false;
				} break;

				case 'datavalidation' /*case 'DataValidation'*/:
				switch(Rn[3]) {
					case 'range' /*case 'Range'*/: break;

					case 'type' /*case 'Type'*/: break;
					case 'min' /*case 'Min'*/: break;
					case 'max' /*case 'Max'*/: break;
					case 'sort' /*case 'Sort'*/: break;
					case 'descending' /*case 'Descending'*/: break;
					case 'order' /*case 'Order'*/: break;
					case 'casesensitive' /*case 'CaseSensitive'*/: break;
					case 'value' /*case 'Value'*/: break;
					case 'errorstyle' /*case 'ErrorStyle'*/: break;
					case 'errormessage' /*case 'ErrorMessage'*/: break;
					case 'errortitle' /*case 'ErrorTitle'*/: break;
					case 'inputmessage' /*case 'InputMessage'*/: break;
					case 'inputtitle' /*case 'InputTitle'*/: break;
					case 'combohide' /*case 'ComboHide'*/: break;
					case 'inputhide' /*case 'InputHide'*/: break;
					case 'condition' /*case 'Condition'*/: break;
					case 'qualifier' /*case 'Qualifier'*/: break;
					case 'useblank' /*case 'UseBlank'*/: break;
					case 'value1' /*case 'Value1'*/: break;
					case 'value2' /*case 'Value2'*/: break;
					case 'format' /*case 'Format'*/: break;

					case 'cellrangelist' /*case 'CellRangeList'*/: break;
					default: seen = false;
				} break;

				case 'sorting' /*case 'Sorting'*/:
				case 'conditionalformatting' /*case 'ConditionalFormatting'*/:
				switch(Rn[3]) {
					case 'range' /*case 'Range'*/: break;
					case 'type' /*case 'Type'*/: break;
					case 'min' /*case 'Min'*/: break;
					case 'max' /*case 'Max'*/: break;
					case 'sort' /*case 'Sort'*/: break;
					case 'descending' /*case 'Descending'*/: break;
					case 'order' /*case 'Order'*/: break;
					case 'casesensitive' /*case 'CaseSensitive'*/: break;
					case 'value' /*case 'Value'*/: break;
					case 'errorstyle' /*case 'ErrorStyle'*/: break;
					case 'errormessage' /*case 'ErrorMessage'*/: break;
					case 'errortitle' /*case 'ErrorTitle'*/: break;
					case 'cellrangelist' /*case 'CellRangeList'*/: break;
					case 'inputmessage' /*case 'InputMessage'*/: break;
					case 'inputtitle' /*case 'InputTitle'*/: break;
					case 'combohide' /*case 'ComboHide'*/: break;
					case 'inputhide' /*case 'InputHide'*/: break;
					case 'condition' /*case 'Condition'*/: break;
					case 'qualifier' /*case 'Qualifier'*/: break;
					case 'useblank' /*case 'UseBlank'*/: break;
					case 'value1' /*case 'Value1'*/: break;
					case 'value2' /*case 'Value2'*/: break;
					case 'format' /*case 'Format'*/: break;
					default: seen = false;
				} break;

				/* MapInfo (schema) */
				case 'mapinfo' /*case 'MapInfo'*/: case 'schema' /*case 'Schema'*/: case 'data' /*case 'data'*/: switch(Rn[3]) {
					case 'map' /*case 'Map'*/: break;
					case 'entry' /*case 'Entry'*/: break;
					case 'range' /*case 'Range'*/: break;
					case 'xpath' /*case 'XPath'*/: break;
					case 'field' /*case 'Field'*/: break;
					case 'xsdtype' /*case 'XSDType'*/: break;
					case 'filteron' /*case 'FilterOn'*/: break;
					case 'aggregate' /*case 'Aggregate'*/: break;
					case 'elementtype' /*case 'ElementType'*/: break;
					case 'attributetype' /*case 'AttributeType'*/: break;
				/* These are from xsd (XML Schema Definition) */
					case 'schema' /*case 'schema'*/:
					case 'element' /*case 'element'*/:
					case 'complextype' /*case 'complexType'*/:
					case 'datatype' /*case 'datatype'*/:
					case 'all' /*case 'all'*/:
					case 'attribute' /*case 'attribute'*/:
					case 'extends' /*case 'extends'*/: break;

					case 'row' /*case 'row'*/: break;
					default: seen = false;
				} break;

				/* SmartTags (can be anything) */
				case 'smarttags' /*case 'SmartTags'*/: break;

				default: seen = false; break;
			}
			if(seen) break;
			/* CustomDocumentProperties */
			if(Rn[3].match(/!\[CDATA/)) break;
			if(!state[state.length-1][1]) throw 'Unrecognized tag: ' + Rn[3] + "|" + state.join("|");
			if(state[state.length-1][0]===/*'CustomDocumentProperties'*/'customdocumentproperties') {
				if(Rn[0].slice(-2) === "/>") break;
				else if(Rn[1]==="/") xlml_set_custprop(Custprops, raw_Rn3, cp, str.slice(pidx, Rn.index));
				else { cp = Rn; pidx = Rn.index + Rn[0].length; }
				break;
			}
			if(opts.WTF) throw 'Unrecognized tag: ' + Rn[3] + "|" + state.join("|");
	}
	var out = ({}/*:any*/);
	if(!opts.bookSheets && !opts.bookProps) out.Sheets = sheets;
	out.SheetNames = sheetnames;
	out.Workbook = Workbook;
	out.SSF = dup(table_fmt);
	out.Props = Props;
	out.Custprops = Custprops;
	out.bookType = "xlml";
	return out;
}

function parse_xlml(data/*:RawBytes|string*/, opts)/*:Workbook*/ {
	fix_read_opts(opts=opts||{});
	switch(opts.type||"base64") {
		case "base64": return parse_xlml_xml(Base64_decode(data), opts);
		case "binary": case "buffer": case "file": return parse_xlml_xml(data, opts);
		case "array": return parse_xlml_xml(a2s(data), opts);
	}
	/*:: throw new Error("unsupported type " + opts.type); */
}

/* TODO */
function write_props_xlml(wb/*:Workbook*/, opts)/*:string*/ {
	var o/*:Array<string>*/ = [];
	/* DocumentProperties */
	if(wb.Props) o.push(xlml_write_docprops(wb.Props, opts));
	/* CustomDocumentProperties */
	if(wb.Custprops) o.push(xlml_write_custprops(wb.Props, wb.Custprops, opts));
	return o.join("");
}
/* TODO */
function write_wb_xlml(wb/*::, opts*/)/*:string*/ {
	/* OfficeDocumentSettings */
	/* ExcelWorkbook */
	if((((wb||{}).Workbook||{}).WBProps||{}).date1904) return '<ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel"><Date1904/></ExcelWorkbook>';
	return "";
}
/* TODO */
function write_sty_xlml(wb, opts)/*:string*/ {
	/* Styles */
	var styles/*:Array<string>*/ = ['<Style ss:ID="Default" ss:Name="Normal"><NumberFormat/></Style>'];
	opts.cellXfs.forEach(function(xf, id) {
		var payload/*:Array<string>*/ = [];
		payload.push(writextag('NumberFormat', null, {"ss:Format": escapexml(table_fmt[xf.numFmtId])}));

		var o = /*::(*/{"ss:ID": "s" + (21+id)}/*:: :any)*/;
		styles.push(writextag('Style', payload.join(""), o));
	});
	return writextag("Styles", styles.join(""));
}
function write_name_xlml(n) { return writextag("NamedRange", null, {"ss:Name": n.Name.slice(0,6) == "_xlnm." ? n.Name.slice(6) : n.Name, "ss:RefersTo":"=" + a1_to_rc(n.Ref, {r:0,c:0})}); }
function write_names_xlml(wb/*::, opts*/)/*:string*/ {
	if(!((wb||{}).Workbook||{}).Names) return "";
	/*:: if(!wb || !wb.Workbook || !wb.Workbook.Names) throw new Error("unreachable"); */
	var names/*:Array<any>*/ = wb.Workbook.Names;
	var out/*:Array<string>*/ = [];
	for(var i = 0; i < names.length; ++i) {
		var n = names[i];
		if(n.Sheet != null) continue;
		if(n.Name.match(/^_xlfn\./)) continue;
		out.push(write_name_xlml(n));
	}
	return writextag("Names", out.join(""));
}
function write_ws_xlml_names(ws/*:Worksheet*/, opts, idx/*:number*/, wb/*:Workbook*/)/*:string*/ {
	if(!ws) return "";
	if(!((wb||{}).Workbook||{}).Names) return "";
	/*:: if(!wb || !wb.Workbook || !wb.Workbook.Names) throw new Error("unreachable"); */
	var names/*:Array<any>*/ = wb.Workbook.Names;
	var out/*:Array<string>*/ = [];
	for(var i = 0; i < names.length; ++i) {
		var n = names[i];
		if(n.Sheet != idx) continue;
		/*switch(n.Name) {
			case "_": continue;
		}*/
		if(n.Name.match(/^_xlfn\./)) continue;
		out.push(write_name_xlml(n));
	}
	return out.join("");
}
/* WorksheetOptions */
function write_ws_xlml_wsopts(ws/*:Worksheet*/, opts, idx/*:number*/, wb/*:Workbook*/)/*:string*/ {
	if(!ws) return "";
	var o/*:Array<string>*/ = [];
	/* NOTE: spec technically allows any order, but stick with implied order */

	/* FitToPage */
	/* DoNotDisplayColHeaders */
	/* DoNotDisplayRowHeaders */
	/* ViewableRange */
	/* Selection */
	/* GridlineColor */
	/* Name */
	/* ExcelWorksheetType */
	/* IntlMacro */
	/* Unsynced */
	/* Selected */
	/* CodeName */

	if(ws['!margins']) {
		o.push("<PageSetup>");
		if(ws['!margins'].header) o.push(writextag("Header", null, {'x:Margin':ws['!margins'].header}));
		if(ws['!margins'].footer) o.push(writextag("Footer", null, {'x:Margin':ws['!margins'].footer}));
		o.push(writextag("PageMargins", null, {
			'x:Bottom': ws['!margins'].bottom || "0.75",
			'x:Left': ws['!margins'].left || "0.7",
			'x:Right': ws['!margins'].right || "0.7",
			'x:Top': ws['!margins'].top || "0.75"
		}));
		o.push("</PageSetup>");
	}

	/* PageSetup */
	/* DisplayPageBreak */
	/* TransitionExpressionEvaluation */
	/* TransitionFormulaEntry */
	/* Print */
	/* Zoom */
	/* PageLayoutZoom */
	/* PageBreakZoom */
	/* ShowPageBreakZoom */
	/* DefaultRowHeight */
	/* DefaultColumnWidth */
	/* StandardWidth */

	if(wb && wb.Workbook && wb.Workbook.Sheets && wb.Workbook.Sheets[idx]) {
		/* Visible */
		if(wb.Workbook.Sheets[idx].Hidden) o.push(writextag("Visible", (wb.Workbook.Sheets[idx].Hidden == 1 ? "SheetHidden" : "SheetVeryHidden"), {}));
		else {
			/* Selected */
			for(var i = 0; i < idx; ++i) if(wb.Workbook.Sheets[i] && !wb.Workbook.Sheets[i].Hidden) break;
			if(i == idx) o.push("<Selected/>");
		}
	}

	/* LeftColumnVisible */

	if(((((wb||{}).Workbook||{}).Views||[])[0]||{}).RTL) o.push("<DisplayRightToLeft/>");

	/* GridlineColorIndex */
	/* DisplayFormulas */
	/* DoNotDisplayGridlines */
	/* DoNotDisplayHeadings */
	/* DoNotDisplayOutline */
	/* ApplyAutomaticOutlineStyles */
	/* NoSummaryRowsBelowDetail */
	/* NoSummaryColumnsRightDetail */
	/* DoNotDisplayZeros */
	/* ActiveRow */
	/* ActiveColumn */
	/* FilterOn */
	/* RangeSelection */
	/* TopRowVisible */
	/* TopRowBottomPane */
	/* LeftColumnRightPane */
	/* ActivePane */
	/* SplitHorizontal */
	/* SplitVertical */
	/* FreezePanes */
	/* FrozenNoSplit */
	/* TabColorIndex */
	/* Panes */

	/* NOTE: Password not supported in XLML Format */
	if(ws['!protect']) {
		o.push(writetag("ProtectContents", "True"));
		if(ws['!protect'].objects) o.push(writetag("ProtectObjects", "True"));
		if(ws['!protect'].scenarios) o.push(writetag("ProtectScenarios", "True"));
		if(ws['!protect'].selectLockedCells != null && !ws['!protect'].selectLockedCells) o.push(writetag("EnableSelection", "NoSelection"));
		else if(ws['!protect'].selectUnlockedCells != null && !ws['!protect'].selectUnlockedCells) o.push(writetag("EnableSelection", "UnlockedCells"));
	[
		[ "formatCells", "AllowFormatCells" ],
		[ "formatColumns", "AllowSizeCols" ],
		[ "formatRows", "AllowSizeRows" ],
		[ "insertColumns", "AllowInsertCols" ],
		[ "insertRows", "AllowInsertRows" ],
		[ "insertHyperlinks", "AllowInsertHyperlinks" ],
		[ "deleteColumns", "AllowDeleteCols" ],
		[ "deleteRows", "AllowDeleteRows" ],
		[ "sort", "AllowSort" ],
		[ "autoFilter", "AllowFilter" ],
		[ "pivotTables", "AllowUsePivotTables" ]
	].forEach(function(x) { if(ws['!protect'][x[0]]) o.push("<"+x[1]+"/>"); });
	}

	if(o.length == 0) return "";
	return writextag("WorksheetOptions", o.join(""), {xmlns:XLMLNS.x});
}
function write_ws_xlml_comment(comments/*:Array<any>*/)/*:string*/ {
	return comments.map(function(c) {
		// TODO: formatted text
		var t = xlml_unfixstr(c.t||"");
		var d =writextag("ss:Data", t, {"xmlns":"http://www.w3.org/TR/REC-html40"});
		return writextag("Comment", d, {"ss:Author":c.a});
	}).join("");
}
function write_ws_xlml_cell(cell, ref/*:string*/, ws, opts, idx/*:number*/, wb, addr)/*:string*/{
	if(!cell || (cell.v == undefined && cell.f == undefined)) return "";

	var attr = {};
	if(cell.f) attr["ss:Formula"] = "=" + escapexml(a1_to_rc(cell.f, addr));
	if(cell.F && cell.F.slice(0, ref.length) == ref) {
		var end = decode_cell(cell.F.slice(ref.length + 1));
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
		case 'z': if(!opts.sheetStubs) return ""; break;
		case 'n': t = 'Number'; p = String(cell.v); break;
		case 'b': t = 'Boolean'; p = (cell.v ? "1" : "0"); break;
		case 'e': t = 'Error'; p = BErr[cell.v]; break;
		case 'd': t = 'DateTime'; p = new Date(cell.v).toISOString(); if(cell.z == null) cell.z = cell.z || table_fmt[14]; break;
		case 's': t = 'String'; p = escapexlml(cell.v||""); break;
	}
	/* TODO: cell style */
	var os = get_cell_style(opts.cellXfs, cell, opts);
	attr["ss:StyleID"] = "s" + (21+os);
	attr["ss:Index"] = addr.c + 1;
	var _v = (cell.v != null ? p : "");
	var m = cell.t == 'z' ? "" : ('<Data ss:Type="' + t + '">' + _v + '</Data>');

	if((cell.c||[]).length > 0) m += write_ws_xlml_comment(cell.c);

	return writextag("Cell", m, attr);
}
function write_ws_xlml_row(R/*:number*/, row)/*:string*/ {
	var o = '<Row ss:Index="' + (R+1) + '"';
	if(row) {
		if(row.hpt && !row.hpx) row.hpx = pt2px(row.hpt);
		if(row.hpx) o += ' ss:AutoFitHeight="0" ss:Height="' + row.hpx + '"';
		if(row.hidden) o += ' ss:Hidden="1"';
	}
	return o + '>';
}
/* TODO */
function write_ws_xlml_table(ws/*:Worksheet*/, opts, idx/*:number*/, wb/*:Workbook*/)/*:string*/ {
	if(!ws['!ref']) return "";
	var range/*:Range*/ = safe_decode_range(ws['!ref']);
	var marr/*:Array<Range>*/ = ws['!merges'] || [], mi = 0;
	var o/*:Array<string>*/ = [];
	if(ws['!cols']) ws['!cols'].forEach(function(n, i) {
		process_col(n);
		var w = !!n.width;
		var p = col_obj_w(i, n);
		var k/*:any*/ = {"ss:Index":i+1};
		if(w) k['ss:Width'] = width2px(p.width);
		if(n.hidden) k['ss:Hidden']="1";
		o.push(writextag("Column",null,k));
	});
	var dense = Array.isArray(ws);
	for(var R = range.s.r; R <= range.e.r; ++R) {
		var row = [write_ws_xlml_row(R, (ws['!rows']||[])[R])];
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
	var o/*:Array<string>*/ = [];
	var s = wb.SheetNames[idx];
	var ws = wb.Sheets[s];

	var t/*:string*/ = ws ? write_ws_xlml_names(ws, opts, idx, wb) : "";
	if(t.length > 0) o.push("<Names>" + t + "</Names>");

	/* Table */
	t = ws ? write_ws_xlml_table(ws, opts, idx, wb) : "";
	if(t.length > 0) o.push("<Table>" + t + "</Table>");

	/* WorksheetOptions */
	o.push(write_ws_xlml_wsopts(ws, opts, idx, wb));

	if(ws["!autofilter"]) o.push('<AutoFilter x:Range="' + a1_to_rc(fix_range(ws["!autofilter"].ref), {r:0,c:0}) + '" xmlns="urn:schemas-microsoft-com:office:excel"></AutoFilter>');

	return o.join("");
}
function write_xlml(wb, opts)/*:string*/ {
	if(!opts) opts = {};
	if(!wb.SSF) wb.SSF = dup(table_fmt);
	if(wb.SSF) {
		make_ssf(); SSF_load_table(wb.SSF);
		// $FlowIgnore
		opts.revssf = evert_num(wb.SSF); opts.revssf[wb.SSF[65535]] = 0;
		opts.ssf = wb.SSF;
		opts.cellXfs = [];
		get_cell_style(opts.cellXfs, {}, {revssf:{"General":0}});
	}
	var d/*:Array<string>*/ = [];
	d.push(write_props_xlml(wb, opts));
	d.push(write_wb_xlml(wb, opts));
	d.push("");
	d.push("");
	for(var i = 0; i < wb.SheetNames.length; ++i)
		d.push(writextag("Worksheet", write_ws_xlml(i, opts, wb), {"ss:Name":escapexml(wb.SheetNames[i])}));
	d[2] = write_sty_xlml(wb, opts);
	d[3] = write_names_xlml(wb, opts);
	return XML_HEADER + writextag("Workbook", d.join(""), {
		'xmlns':      XLMLNS.ss,
		'xmlns:o':    XLMLNS.o,
		'xmlns:x':    XLMLNS.x,
		'xmlns:ss':   XLMLNS.ss,
		'xmlns:dt':   XLMLNS.dt,
		'xmlns:html': XLMLNS.html
	});
}
