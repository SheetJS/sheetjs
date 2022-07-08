/* OpenDocument */
function parse_text_p(text/*:string*//*::, tag*/)/*:Array<any>*/ {
	/* 6.1.2 White Space Characters */
	var fixed = text
		.replace(/[\t\r\n]/g, " ").trim().replace(/ +/g, " ")
		.replace(/<text:s\/>/g," ")
		.replace(/<text:s text:c="(\d+)"\/>/g, function($$,$1) { return Array(parseInt($1,10)+1).join(" "); })
		.replace(/<text:tab[^>]*\/>/g,"\t")
		.replace(/<text:line-break\/>/g,"\n");
	var v = unescapexml(fixed.replace(/<[^>]*>/g,""));

	return [v];
}

/* Note: ODS can stick styles in content.xml or styles.xml, FODS blurs lines */
function parse_ods_styles(d/*:string*/, _opts, _nfm) {
	var number_format_map = _nfm || {};
	var str = xlml_normalize(d);
	xlmlregex.lastIndex = 0;
	str = str.replace(/<!--([\s\S]*?)-->/mg,"").replace(/<!DOCTYPE[^\[]*\[[^\]]*\]>/gm,"");
	var Rn, NFtag, NF = "", tNF = "", y, etpos = 0, tidx = -1, infmt = false, payload = "";
	while((Rn = xlmlregex.exec(str))) {
		switch((Rn[3]=Rn[3].replace(/_.*$/,""))) {
		/* Number Format Definitions */
		case 'number-style': // <number:number-style> 16.29.2
		case 'currency-style': // <number:currency-style> 16.29.8
		case 'percentage-style': // <number:percentage-style> 16.29.10
		case 'date-style': // <number:date-style> 16.29.11
		case 'time-style': // <number:time-style> 16.29.19
		case 'text-style': // <number:text-style> 16.29.26
			if(Rn[1]==='/') {
				infmt = false;
				if(NFtag['truncate-on-overflow'] == "false") {
					if(NF.match(/h/)) NF = NF.replace(/h+/, "[$&]");
					else if(NF.match(/m/)) NF = NF.replace(/m+/, "[$&]");
					else if(NF.match(/s/)) NF = NF.replace(/s+/, "[$&]");
				}
				number_format_map[NFtag.name] = NF;
				NF = "";
			} else if(Rn[0].charAt(Rn[0].length-2) !== '/') {
				infmt = true;
				NF = "";
				NFtag = parsexmltag(Rn[0], false);
			} break;

		// LibreOffice bug https://bugs.documentfoundation.org/show_bug.cgi?id=149484
		case 'boolean-style': // <number:boolean-style> 16.29.24
			if(Rn[1]==='/') {
				infmt = false;
				number_format_map[NFtag.name] = "General";
				NF = "";
			} else if(Rn[0].charAt(Rn[0].length-2) !== '/') {
				infmt = true;
				NF = "";
				NFtag = parsexmltag(Rn[0], false);
			} break;

		/* Number Format Elements */
		case 'boolean': // <number:boolean> 16.29.25
			NF += "General"; // ODF spec is unfortunately underspecified here
			break;

		case 'text': // <number:text> 16.29.27
			if(Rn[1]==='/') {
				payload = str.slice(tidx, xlmlregex.lastIndex - Rn[0].length);
				// NOTE: Excel has a different interpretation of "%%" and friends
				if(payload == "%" && NFtag[0] == '<number:percentage-style') NF += "%";
				else NF += '"' + payload.replace(/"/g, '""') + '"';
			} else if(Rn[0].charAt(Rn[0].length-2) !== '/') {
				tidx = xlmlregex.lastIndex;
			} break;


		case 'day': { // <number:day> 16.29.12
			y = parsexmltag(Rn[0], false);
			switch(y["style"]) {
				case "short": NF += "d"; break;
				case "long": NF += "dd"; break;
				default: NF += "dd"; break; // TODO: error condition
			}
		} break;

		case 'day-of-week': { // <number:day-of-week> 16.29.16
			y = parsexmltag(Rn[0], false);
			switch(y["style"]) {
				case "short": NF += "ddd"; break;
				case "long": NF += "dddd"; break;
				default: NF += "ddd"; break;
			}
		} break;

		case 'era': { // <number:era> 16.29.15 TODO: proper mapping
			y = parsexmltag(Rn[0], false);
			switch(y["style"]) {
				case "short": NF += "ee"; break;
				case "long": NF += "eeee"; break;
				default: NF += "eeee"; break; // TODO: error condition
			}
		} break;

		case 'hours': { // <number:hours> 16.29.20
			y = parsexmltag(Rn[0], false);
			switch(y["style"]) {
				case "short": NF += "h"; break;
				case "long": NF += "hh"; break;
				default: NF += "hh"; break; // TODO: error condition
			}
		} break;

		case 'minutes': { // <number:minutes> 16.29.21
			y = parsexmltag(Rn[0], false);
			switch(y["style"]) {
				case "short": NF += "m"; break;
				case "long": NF += "mm"; break;
				default: NF += "mm"; break; // TODO: error condition
			}
		} break;

		case 'month': { // <number:month> 16.29.13
			y = parsexmltag(Rn[0], false);
			if(y["textual"]) NF += "mm";
			switch(y["style"]) {
				case "short": NF += "m"; break;
				case "long": NF += "mm"; break;
				default: NF += "m"; break;
			}
		} break;

		case 'seconds': { // <number:seconds> 16.29.22
			y = parsexmltag(Rn[0], false);
			switch(y["style"]) {
				case "short": NF += "s"; break;
				case "long": NF += "ss"; break;
				default: NF += "ss"; break; // TODO: error condition
			}
			if(y["decimal-places"]) NF += "." + fill("0", +y["decimal-places"]);
		} break;

		case 'year': { // <number:year> 16.29.14
			y = parsexmltag(Rn[0], false);
			switch(y["style"]) {
				case "short": NF += "yy"; break;
				case "long": NF += "yyyy"; break;
				default: NF += "yy"; break; // TODO: error condition
			}
		} break;

		case 'am-pm': // <number:am-pm> 16.29.23
			NF += "AM/PM"; // LO autocorrects A/P -> AM/PM
			break;

		case 'week-of-year': // <number:week-of-year> 16.29.17
		case 'quarter': // <number:quarter> 16.29.18
			console.error("Excel does not support ODS format token " + Rn[3]);
			break;

		case 'fill-character': // <number:fill-character> 16.29.5
			if(Rn[1]==='/') {
				payload = str.slice(tidx, xlmlregex.lastIndex - Rn[0].length);
				// NOTE: Excel has a different interpretation of "%%" and friends
				NF += '"' + payload.replace(/"/g, '""') + '"*';
			} else if(Rn[0].charAt(Rn[0].length-2) !== '/') {
				tidx = xlmlregex.lastIndex;
			} break;

		case 'scientific-number': // <number:scientific-number> 16.29.6
			// TODO: find a mapping for all parameters
			y = parsexmltag(Rn[0], false);
			NF += "0." + fill("0", +y["min-decimal-places"] || +y["decimal-places"] || 2) + fill("?", +y["decimal-places"] - +y["min-decimal-places"] || 0) + "E" + (parsexmlbool(y["forced-exponent-sign"]) ? "+" : "") + fill("0", +y["min-exponent-digits"] || 2);
			break;

		case 'fraction': // <number:fraction> 16.29.7
			// TODO: find a mapping for all parameters
			y = parsexmltag(Rn[0], false);
			if(!+y["min-integer-digits"]) NF += "#";
			else NF += fill("0", +y["min-integer-digits"]);
			NF += " ";
			NF += fill("?", +y["min-numerator-digits"] || 1);
			NF += "/";
			if(+y["denominator-value"]) NF += y["denominator-value"];
			else NF += fill("?", +y["min-denominator-digits"] || 1);
			break;

		case 'currency-symbol': // <number:currency-symbol> 16.29.9
			// TODO: localization with [$-...]
			if(Rn[1]==='/') {
				NF += '"' + str.slice(tidx, xlmlregex.lastIndex - Rn[0].length).replace(/"/g, '""') + '"';
			} else if(Rn[0].charAt(Rn[0].length-2) !== '/') {
				tidx = xlmlregex.lastIndex;
			} else NF += "$";
			break;

		case 'text-properties': // <style:text-properties> 16.29.29
			y = parsexmltag(Rn[0], false);
			switch((y["color"]||"").toLowerCase().replace("#", "")) {
				case "ff0000": case "red": NF = "[Red]" + NF; break;
			}
			break;

		case 'text-content': // <number:text-content> 16.29.28
			NF += "@";
			break;

		case 'map': // <style:map> 16.3
			// TODO: handle more complex maps
			y = parsexmltag(Rn[0], false);
			if(unescapexml(y["condition"]) == "value()>=0") NF = number_format_map[y["apply-style-name"]] + ";" + NF;
			else console.error("ODS number format may be incorrect: " + y["condition"]);
			break;

		case 'number': // <number:number> 16.29.3
			// TODO: handle all the attributes
			if(Rn[1]==='/') break;
			y = parsexmltag(Rn[0], false);
			tNF = "";
			tNF += fill("0", +y["min-integer-digits"] || 1);
			if(parsexmlbool(y["grouping"])) tNF = commaify(fill("#", Math.max(0, 4 - tNF.length)) + tNF);
			if(+y["min-decimal-places"] || +y["decimal-places"]) tNF += ".";
			if(+y["min-decimal-places"]) tNF += fill("0", +y["min-decimal-places"] || 1);
			if(+y["decimal-places"] - (+y["min-decimal-places"]||0)) tNF += fill("0", +y["decimal-places"] - (+y["min-decimal-places"]||0)); // TODO: should this be "#" ?
			NF += tNF;
			break;

		case 'embedded-text': // <number:embedded-text> 16.29.4
			// TODO: verify interplay with grouping et al
			if(Rn[1]==='/') {
				if(etpos == 0) NF += '"' + str.slice(tidx, xlmlregex.lastIndex - Rn[0].length).replace(/"/g, '""') + '"';
				else NF = NF.slice(0, etpos) + '"' + str.slice(tidx, xlmlregex.lastIndex - Rn[0].length).replace(/"/g, '""') + '"' + NF.slice(etpos);
			} else if(Rn[0].charAt(Rn[0].length-2) !== '/') {
				tidx = xlmlregex.lastIndex;
				etpos = -+parsexmltag(Rn[0], false)["position"] || 0;
			} break;

	}}
	return number_format_map;
}

function parse_content_xml(d/*:string*/, _opts, _nfm)/*:Workbook*/ {
		var opts = _opts || {};
		if(DENSE != null && opts.dense == null) opts.dense = DENSE;
		var str = xlml_normalize(d);
		var state/*:Array<any>*/ = [], tmp;
		var tag/*:: = {}*/;
		var nfidx, NF = "", pidx = 0;
		var sheetag/*:: = {name:"", '名称':""}*/;
		var rowtag/*:: = {'行号':""}*/;
		var Sheets = {}, SheetNames/*:Array<string>*/ = [];
		var ws = opts.dense ? ([]/*:any*/) : ({}/*:any*/);
		var Rn, q/*:: :any = ({t:"", v:null, z:null, w:"",c:[],}:any)*/;
		var ctag = ({value:""}/*:any*/);
		var textp = "", textpidx = 0, textptag/*:: = {}*/;
		var textR = [];
		var R = -1, C = -1, range = {s: {r:1000000,c:10000000}, e: {r:0, c:0}};
		var row_ol = 0;
		var number_format_map = _nfm || {}, styles = {};
		var merges/*:Array<Range>*/ = [], mrange = {}, mR = 0, mC = 0;
		var rowinfo/*:Array<RowInfo>*/ = [], rowpeat = 1, colpeat = 1;
		var arrayf/*:Array<[Range, string]>*/ = [];
		var WB = {Names:[], WBProps:{}};
		var atag = ({}/*:any*/);
		var _Ref/*:[string, string]*/ = ["", ""];
		var comments/*:Array<Comment>*/ = [], comment/*:Comment*/ = ({}/*:any*/);
		var creator = "", creatoridx = 0;
		var isstub = false, intable = false;
		var i = 0;
		var baddate = 0;
		xlmlregex.lastIndex = 0;
		str = str.replace(/<!--([\s\S]*?)-->/mg,"").replace(/<!DOCTYPE[^\[]*\[[^\]]*\]>/gm,"");
		while((Rn = xlmlregex.exec(str))) switch((Rn[3]=Rn[3].replace(/_.*$/,""))) {

			case 'table': case '工作表': // 9.1.2 <table:table>
				if(Rn[1]==='/') {
					if(range.e.c >= range.s.c && range.e.r >= range.s.r) ws['!ref'] = encode_range(range);
					else ws['!ref'] = "A1:A1";
					if(opts.sheetRows > 0 && opts.sheetRows <= range.e.r) {
						ws['!fullref'] = ws['!ref'];
						range.e.r = opts.sheetRows - 1;
						ws['!ref'] = encode_range(range);
					}
					if(merges.length) ws['!merges'] = merges;
					if(rowinfo.length) ws["!rows"] = rowinfo;
					sheetag.name = sheetag['名称'] || sheetag.name;
					if(typeof JSON !== 'undefined') JSON.stringify(sheetag);
					SheetNames.push(sheetag.name);
					Sheets[sheetag.name] = ws;
					intable = false;
				}
				else if(Rn[0].charAt(Rn[0].length-2) !== '/') {
					sheetag = parsexmltag(Rn[0], false);
					R = C = -1;
					range.s.r = range.s.c = 10000000; range.e.r = range.e.c = 0;
					ws = opts.dense ? ([]/*:any*/) : ({}/*:any*/); merges = [];
					rowinfo = [];
					intable = true;
				}
				break;

			case 'table-row-group': // 9.1.9 <table:table-row-group>
				if(Rn[1] === "/") --row_ol; else ++row_ol;
				break;
			case 'table-row': case '行': // 9.1.3 <table:table-row>
				if(Rn[1] === '/') { R+=rowpeat; rowpeat = 1; break; }
				rowtag = parsexmltag(Rn[0], false);
				if(rowtag['行号']) R = rowtag['行号'] - 1; else if(R == -1) R = 0;
				rowpeat = +rowtag['number-rows-repeated'] || 1;
				/* TODO: remove magic */
				if(rowpeat < 10) for(i = 0; i < rowpeat; ++i) if(row_ol > 0) rowinfo[R + i] = {level: row_ol};
				C = -1; break;
			case 'covered-table-cell': // 9.1.5 <table:covered-table-cell>
				if(Rn[1] !== '/') ++C;
				if(opts.sheetStubs) {
					if(opts.dense) { if(!ws[R]) ws[R] = []; ws[R][C] = {t:'z'}; }
					else ws[encode_cell({r:R,c:C})] = {t:'z'};
				}
				textp = ""; textR = [];
				break; /* stub */
			case 'table-cell': case '数据':
				if(Rn[0].charAt(Rn[0].length-2) === '/') {
					++C;
					ctag = parsexmltag(Rn[0], false);
					colpeat = parseInt(ctag['number-columns-repeated']||"1", 10);
					q = ({t:'z', v:null/*:: , z:null, w:"",c:[]*/}/*:any*/);
					if(ctag.formula && opts.cellFormula != false) q.f = ods_to_csf_formula(unescapexml(ctag.formula));
					if(ctag["style-name"] && styles[ctag["style-name"]]) q.z = styles[ctag["style-name"]];
					if((ctag['数据类型'] || ctag['value-type']) == "string") {
						q.t = "s"; q.v = unescapexml(ctag['string-value'] || "");
						if(opts.dense) {
							if(!ws[R]) ws[R] = [];
							ws[R][C] = q;
						} else {
							ws[encode_cell({r:R,c:C})] = q;
						}
					}
					C+= colpeat-1;
				} else if(Rn[1]!=='/') {
					++C;
					textp = ""; textpidx = 0; textR = [];
					colpeat = 1;
					var rptR = rowpeat ? R + rowpeat - 1 : R;
					if(C > range.e.c) range.e.c = C;
					if(C < range.s.c) range.s.c = C;
					if(R < range.s.r) range.s.r = R;
					if(rptR > range.e.r) range.e.r = rptR;
					ctag = parsexmltag(Rn[0], false);
					comments = []; comment = ({}/*:any*/);
					q = ({t:ctag['数据类型'] || ctag['value-type'], v:null/*:: , z:null, w:"",c:[]*/}/*:any*/);
					if(ctag["style-name"] && styles[ctag["style-name"]]) q.z = styles[ctag["style-name"]];
					if(opts.cellFormula) {
						if(ctag.formula) ctag.formula = unescapexml(ctag.formula);
						if(ctag['number-matrix-columns-spanned'] && ctag['number-matrix-rows-spanned']) {
							mR = parseInt(ctag['number-matrix-rows-spanned'],10) || 0;
							mC = parseInt(ctag['number-matrix-columns-spanned'],10) || 0;
							mrange = {s: {r:R,c:C}, e:{r:R + mR-1,c:C + mC-1}};
							q.F = encode_range(mrange);
							arrayf.push([mrange, q.F]);
						}
						if(ctag.formula) q.f = ods_to_csf_formula(ctag.formula);
						else for(i = 0; i < arrayf.length; ++i)
							if(R >= arrayf[i][0].s.r && R <= arrayf[i][0].e.r)
								if(C >= arrayf[i][0].s.c && C <= arrayf[i][0].e.c)
									q.F = arrayf[i][1];
					}
					if(ctag['number-columns-spanned'] || ctag['number-rows-spanned']) {
						mR = parseInt(ctag['number-rows-spanned'],10) || 0;
						mC = parseInt(ctag['number-columns-spanned'],10) || 0;
						mrange = {s: {r:R,c:C}, e:{r:R + mR-1,c:C + mC-1}};
						merges.push(mrange);
					}

					/* 19.675.2 table:number-columns-repeated */
					if(ctag['number-columns-repeated']) colpeat = parseInt(ctag['number-columns-repeated'], 10);

					/* 19.385 office:value-type */
					switch(q.t) {
						case 'boolean': q.t = 'b'; q.v = parsexmlbool(ctag['boolean-value']) || (+ctag['boolean-value'] >= 1); break;
						case 'float': q.t = 'n'; q.v = parseFloat(ctag.value); break;
						case 'percentage': q.t = 'n'; q.v = parseFloat(ctag.value); break;
						case 'currency': q.t = 'n'; q.v = parseFloat(ctag.value); break;
						case 'date': q.t = 'd'; q.v = parseDate(ctag['date-value']);
							if(!opts.cellDates) { q.t = 'n'; q.v = datenum(q.v, WB.WBProps.date1904) - baddate; }
							if(!q.z) q.z = 'm/d/yy'; break;
						case 'time': q.t = 'n'; q.v = parse_isodur(ctag['time-value'])/86400;
							if(opts.cellDates) { q.t = 'd'; q.v = numdate(q.v); }
							if(!q.z) q.z = 'HH:MM:SS'; break;
						case 'number': q.t = 'n'; q.v = parseFloat(ctag['数据数值']); break;
						default:
							if(q.t === 'string' || q.t === 'text' || !q.t) {
								q.t = 's';
								if(ctag['string-value'] != null) { textp = unescapexml(ctag['string-value']); textR = []; }
							} else throw new Error('Unsupported value type ' + q.t);
					}
				} else {
					isstub = false;
					if(q.t === 's') {
						q.v = textp || '';
						if(textR.length) q.R = textR;
						isstub = textpidx == 0;
					}
					if(atag.Target) q.l = atag;
					if(comments.length > 0) { q.c = comments; comments = []; }
					if(textp && opts.cellText !== false) q.w = textp;
					if(isstub) { q.t = "z"; delete q.v; }
					if(!isstub || opts.sheetStubs) {
						if(!(opts.sheetRows && opts.sheetRows <= R)) {
							for(var rpt = 0; rpt < rowpeat; ++rpt) {
								colpeat = parseInt(ctag['number-columns-repeated']||"1", 10);
								if(opts.dense) {
									if(!ws[R + rpt]) ws[R + rpt] = [];
									ws[R + rpt][C] = rpt == 0 ? q : dup(q);
									while(--colpeat > 0) ws[R + rpt][C + colpeat] = dup(q);
								} else {
									ws[encode_cell({r:R + rpt,c:C})] = q;
									while(--colpeat > 0) ws[encode_cell({r:R + rpt,c:C + colpeat})] = dup(q);
								}
								if(range.e.c <= C) range.e.c = C;
							}
						}
					}
					colpeat = parseInt(ctag['number-columns-repeated']||"1", 10);
					C += colpeat-1; colpeat = 0;
					q = {/*:: t:"", v:null, z:null, w:"",c:[]*/};
					textp = ""; textR = [];
				}
				atag = ({}/*:any*/);
				break; // 9.1.4 <table:table-cell>

			/* pure state */
			case 'document': // TODO: <office:document> is the root for FODS
			case 'document-content': case '电子表格文档': // 3.1.3.2 <office:document-content>
			case 'spreadsheet': case '主体': // 3.7 <office:spreadsheet>
			case 'scripts': // 3.12 <office:scripts>
			case 'styles': // TODO <office:styles>
			case 'font-face-decls': // 3.14 <office:font-face-decls>
			case 'master-styles': // 3.15.4 <office:master-styles> -- relevant for FODS
				if(Rn[1]==='/'){if((tmp=state.pop())[0]!==Rn[3]) throw "Bad state: "+tmp;}
				else if(Rn[0].charAt(Rn[0].length-2) !== '/') state.push([Rn[3], true]);
				break;

			case 'annotation': // 14.1 <office:annotation>
				if(Rn[1]==='/'){
					if((tmp=state.pop())[0]!==Rn[3]) throw "Bad state: "+tmp;
					comment.t = textp;
					if(textR.length) /*::(*/comment/*:: :any)*/.R = textR;
					comment.a = creator;
					comments.push(comment);
				}
				else if(Rn[0].charAt(Rn[0].length-2) !== '/') {state.push([Rn[3], false]);}
				creator = ""; creatoridx = 0;
				textp = ""; textpidx = 0; textR = [];
				break;

			case 'creator': // 4.3.2.7 <dc:creator>
				if(Rn[1]==='/') { creator = str.slice(creatoridx,Rn.index); }
				else creatoridx = Rn.index + Rn[0].length;
				break;

			/* ignore state */
			case 'meta': case '元数据': // TODO: <office:meta> <uof:元数据> FODS/UOF
			case 'settings': // TODO: <office:settings>
			case 'config-item-set': // TODO: <office:config-item-set>
			case 'config-item-map-indexed': // TODO: <office:config-item-map-indexed>
			case 'config-item-map-entry': // TODO: <office:config-item-map-entry>
			case 'config-item-map-named': // TODO: <office:config-item-map-entry>
			case 'shapes': // 9.2.8 <table:shapes>
			case 'frame': // 10.4.2 <draw:frame>
			case 'text-box': // 10.4.3 <draw:text-box>
			case 'image': // 10.4.4 <draw:image>
			case 'data-pilot-tables': // 9.6.2 <table:data-pilot-tables>
			case 'list-style': // 16.30 <text:list-style>
			case 'form': // 13.13 <form:form>
			case 'dde-links': // 9.8 <table:dde-links>
			case 'event-listeners': // TODO
			case 'chart': // TODO
				if(Rn[1]==='/'){if((tmp=state.pop())[0]!==Rn[3]) throw "Bad state: "+tmp;}
				else if(Rn[0].charAt(Rn[0].length-2) !== '/') state.push([Rn[3], false]);
				textp = ""; textpidx = 0; textR = [];
				break;

			case 'scientific-number': // <number:scientific-number>
			case 'currency-symbol': // <number:currency-symbol>
			case 'fill-character': // 16.29.5 <number:fill-character>
				break;

			case 'text-style': // 16.27.25 <number:text-style>
			case 'boolean-style': // 16.27.23 <number:boolean-style>
			case 'number-style': // 16.27.2 <number:number-style>
			case 'currency-style': // 16.29.8 <number:currency-style>
			case 'percentage-style': // 16.27.9 <number:percentage-style>
			case 'date-style': // 16.27.10 <number:date-style>
			case 'time-style': // 16.27.18 <number:time-style>
				if(Rn[1]==='/'){
					var xlmlidx = xlmlregex.lastIndex;
					parse_ods_styles(str.slice(nfidx, xlmlregex.lastIndex), _opts, number_format_map);
					xlmlregex.lastIndex = xlmlidx;
				} else if(Rn[0].charAt(Rn[0].length-2) !== '/') {
					nfidx = xlmlregex.lastIndex - Rn[0].length;
				} break;

			case 'script': break; // 3.13 <office:script>
			case 'libraries': break; // TODO: <ooo:libraries>
			case 'automatic-styles': break; // 3.15.3 <office:automatic-styles>

			case 'default-style': // TODO: <style:default-style>
			case 'page-layout': break; // TODO: <style:page-layout>
			case 'style': { // 16.2 <style:style>
				var styletag = parsexmltag(Rn[0], false);
				if(styletag["family"] == "table-cell" && number_format_map[styletag["data-style-name"]]) styles[styletag["name"]] = number_format_map[styletag["data-style-name"]];
			} break;
			case 'map': break; // 16.3 <style:map>
			case 'font-face': break; // 16.21 <style:font-face>

			case 'paragraph-properties': break; // 17.6 <style:paragraph-properties>
			case 'table-properties': break; // 17.15 <style:table-properties>
			case 'table-column-properties': break; // 17.16 <style:table-column-properties>
			case 'table-row-properties': break; // 17.17 <style:table-row-properties>
			case 'table-cell-properties': break; // 17.18 <style:table-cell-properties>

			case 'number': // 16.27.3 <number:number>
				break;

			case 'fraction': break; // TODO 16.27.6 <number:fraction>

			case 'day': // 16.27.11 <number:day>
			case 'month': // 16.27.12 <number:month>
			case 'year': // 16.27.13 <number:year>
			case 'era': // 16.27.14 <number:era>
			case 'day-of-week': // 16.27.15 <number:day-of-week>
			case 'week-of-year': // 16.27.16 <number:week-of-year>
			case 'quarter': // 16.27.17 <number:quarter>
			case 'hours': // 16.27.19 <number:hours>
			case 'minutes': // 16.27.20 <number:minutes>
			case 'seconds': // 16.27.21 <number:seconds>
			case 'am-pm': // 16.27.22 <number:am-pm>
				break;

			case 'boolean': break; // 16.27.24 <number:boolean>
			case 'text': // 16.27.26 <number:text>
				if(Rn[0].slice(-2) === "/>") break;
				else if(Rn[1]==="/") switch(state[state.length-1][0]) {
					case 'number-style':
					case 'date-style':
					case 'time-style':
						NF += str.slice(pidx, Rn.index);
						break;
				}
				else pidx = Rn.index + Rn[0].length;
				break;

			case 'named-range': // 9.4.12 <table:named-range>
				tag = parsexmltag(Rn[0], false);
				_Ref = ods_to_csf_3D(tag['cell-range-address']);
				var nrange = ({Name:tag.name, Ref:_Ref[0] + '!' + _Ref[1]}/*:any*/);
				if(intable) nrange.Sheet = SheetNames.length;
				WB.Names.push(nrange);
				break;

			case 'text-content': break; // 16.27.27 <number:text-content>
			case 'text-properties': break; // 16.27.27 <style:text-properties>
			case 'embedded-text': break; // 16.27.4 <number:embedded-text>

			case 'body': case '电子表格': break; // 3.3 16.9.6 19.726.3

			case 'forms': break; // 12.25.2 13.2
			case 'table-column': break; // 9.1.6 <table:table-column>
			case 'table-header-rows': break; // 9.1.7 <table:table-header-rows>
			case 'table-rows': break; // 9.1.12 <table:table-rows>
			/* TODO: outline levels */
			case 'table-column-group': break; // 9.1.10 <table:table-column-group>
			case 'table-header-columns': break; // 9.1.11 <table:table-header-columns>
			case 'table-columns': break; // 9.1.12 <table:table-columns>

			case 'null-date': // 9.4.2 <table:null-date>
				tag = parsexmltag(Rn[0], false);
				switch(tag["date-value"]) {
					case "1904-01-01": WB.WBProps.date1904 = true;
					/* falls through */
					case "1900-01-01": baddate = 0;
				}
				break;

			case 'graphic-properties': break; // 17.21 <style:graphic-properties>
			case 'calculation-settings': break; // 9.4.1 <table:calculation-settings>
			case 'named-expressions': break; // 9.4.11 <table:named-expressions>
			case 'label-range': break; // 9.4.9 <table:label-range>
			case 'label-ranges': break; // 9.4.10 <table:label-ranges>
			case 'named-expression': break; // 9.4.13 <table:named-expression>
			case 'sort': break; // 9.4.19 <table:sort>
			case 'sort-by': break; // 9.4.20 <table:sort-by>
			case 'sort-groups': break; // 9.4.22 <table:sort-groups>

			case 'tab': break; // 6.1.4 <text:tab>
			case 'line-break': break; // 6.1.5 <text:line-break>
			case 'span': break; // 6.1.7 <text:span>
			case 'p': case '文本串': // 5.1.3 <text:p>
				if(['master-styles'].indexOf(state[state.length-1][0]) > -1) break;
				if(Rn[1]==='/' && (!ctag || !ctag['string-value'])) {
					var ptp = parse_text_p(str.slice(textpidx,Rn.index), textptag);
					textp = (textp.length > 0 ? textp + "\n" : "") + ptp[0];
				} else { textptag = parsexmltag(Rn[0], false); textpidx = Rn.index + Rn[0].length; }
				break; // <text:p>
			case 's': break; // <text:s>

			case 'database-range': // 9.4.15 <table:database-range>
				if(Rn[1]==='/') break;
				try {
					_Ref = ods_to_csf_3D(parsexmltag(Rn[0])['target-range-address']);
					Sheets[_Ref[0]]['!autofilter'] = { ref:_Ref[1] };
				} catch(e) {/* empty */}
				break;

			case 'date': break; // <*:date>

			case 'object': break; // 10.4.6.2 <draw:object>
			case 'title': case '标题': break; // <*:title> OR <uof:标题>
			case 'desc': break; // <*:desc>
			case 'binary-data': break; // 10.4.5 TODO: b64 blob

			/* 9.2 Advanced Tables */
			case 'table-source': break; // 9.2.6
			case 'scenario': break; // 9.2.6

			case 'iteration': break; // 9.4.3 <table:iteration>
			case 'content-validations': break; // 9.4.4 <table:
			case 'content-validation': break; // 9.4.5 <table:
			case 'help-message': break; // 9.4.6 <table:
			case 'error-message': break; // 9.4.7 <table:
			case 'database-ranges': break; // 9.4.14 <table:database-ranges>
			case 'filter': break; // 9.5.2 <table:filter>
			case 'filter-and': break; // 9.5.3 <table:filter-and>
			case 'filter-or': break; // 9.5.4 <table:filter-or>
			case 'filter-condition': break; // 9.5.5 <table:filter-condition>

			case 'list-level-style-bullet': break; // 16.31 <text:
			case 'list-level-style-number': break; // 16.32 <text:
			case 'list-level-properties': break; // 17.19 <style:

			/* 7.3 Document Fields */
			case 'sender-firstname': // 7.3.6.2
			case 'sender-lastname': // 7.3.6.3
			case 'sender-initials': // 7.3.6.4
			case 'sender-title': // 7.3.6.5
			case 'sender-position': // 7.3.6.6
			case 'sender-email': // 7.3.6.7
			case 'sender-phone-private': // 7.3.6.8
			case 'sender-fax': // 7.3.6.9
			case 'sender-company': // 7.3.6.10
			case 'sender-phone-work': // 7.3.6.11
			case 'sender-street': // 7.3.6.12
			case 'sender-city': // 7.3.6.13
			case 'sender-postal-code': // 7.3.6.14
			case 'sender-country': // 7.3.6.15
			case 'sender-state-or-province': // 7.3.6.16
			case 'author-name': // 7.3.7.1
			case 'author-initials': // 7.3.7.2
			case 'chapter': // 7.3.8
			case 'file-name': // 7.3.9
			case 'template-name': // 7.3.9
			case 'sheet-name': // 7.3.9
				break;

			case 'event-listener':
				break;
			/* TODO: FODS Properties */
			case 'initial-creator':
			case 'creation-date':
			case 'print-date':
			case 'generator':
			case 'document-statistic':
			case 'user-defined':
			case 'editing-duration':
			case 'editing-cycles':
				break;

			/* TODO: FODS Config */
			case 'config-item':
				break;

			/* TODO: style tokens */
			case 'page-number': break; // TODO <text:page-number>
			case 'page-count': break; // TODO <text:page-count>
			case 'time': break; // TODO <text:time>

			/* 9.3 Advanced Table Cells */
			case 'cell-range-source': break; // 9.3.1 <table:
			case 'detective': break; // 9.3.2 <table:
			case 'operation': break; // 9.3.3 <table:
			case 'highlighted-range': break; // 9.3.4 <table:

			/* 9.6 Data Pilot Tables <table: */
			case 'data-pilot-table': // 9.6.3
			case 'source-cell-range': // 9.6.5
			case 'source-service': // 9.6.6
			case 'data-pilot-field': // 9.6.7
			case 'data-pilot-level': // 9.6.8
			case 'data-pilot-subtotals': // 9.6.9
			case 'data-pilot-subtotal': // 9.6.10
			case 'data-pilot-members': // 9.6.11
			case 'data-pilot-member': // 9.6.12
			case 'data-pilot-display-info': // 9.6.13
			case 'data-pilot-sort-info': // 9.6.14
			case 'data-pilot-layout-info': // 9.6.15
			case 'data-pilot-field-reference': // 9.6.16
			case 'data-pilot-groups': // 9.6.17
			case 'data-pilot-group': // 9.6.18
			case 'data-pilot-group-member': // 9.6.19
				break;

			/* 10.3 Drawing Shapes */
			case 'rect': // 10.3.2
				break;

			/* 14.6 DDE Connections */
			case 'dde-connection-decls': // 14.6.2 <text:
			case 'dde-connection-decl': // 14.6.3 <text:
			case 'dde-link': // 14.6.4 <table:
			case 'dde-source': // 14.6.5 <office:
				break;

			case 'properties': break; // 13.7 <form:properties>
			case 'property': break; // 13.8 <form:property>

			case 'a': // 6.1.8 hyperlink
				if(Rn[1]!== '/') {
					atag = parsexmltag(Rn[0], false);
					if(!atag.href) break;
					atag.Target = unescapexml(atag.href); delete atag.href;
					if(atag.Target.charAt(0) == "#" && atag.Target.indexOf(".") > -1) {
						_Ref = ods_to_csf_3D(atag.Target.slice(1));
						atag.Target = "#" + _Ref[0] + "!" + _Ref[1];
					} else if(atag.Target.match(/^\.\.[\\\/]/)) atag.Target = atag.Target.slice(3);
				}
				break;

			/* non-standard */
			case 'table-protection': break;
			case 'data-pilot-grand-total': break; // <table:
			case 'office-document-common-attrs': break; // bare
			default: switch(Rn[2]) {
				case 'dc:':       // TODO: properties
				case 'calcext:':  // ignore undocumented extensions
				case 'loext:':    // ignore undocumented extensions
				case 'ooo:':      // ignore undocumented extensions
				case 'chartooo:': // ignore undocumented extensions
				case 'draw:':     // TODO: drawing
				case 'style:':    // TODO: styles
				case 'chart:':    // TODO: charts
				case 'form:':     // TODO: forms
				case 'uof:':      // TODO: uof
				case '表:':       // TODO: uof
				case '字:':       // TODO: uof
					break;
				default: if(opts.WTF) throw new Error(Rn);
			}
		}
		var out/*:Workbook*/ = ({
			Sheets: Sheets,
			SheetNames: SheetNames,
			Workbook: WB
		}/*:any*/);
		if(opts.bookSheets) delete /*::(*/out/*:: :any)*/.Sheets;
		return out;
}

function parse_ods(zip/*:ZIPFile*/, opts/*:?ParseOpts*/)/*:Workbook*/ {
	opts = opts || ({}/*:any*/);
	if(safegetzipfile(zip, 'META-INF/manifest.xml')) parse_manifest(getzipdata(zip, 'META-INF/manifest.xml'), opts);
	var styles = getzipstr(zip, 'styles.xml');
	var Styles = styles && parse_ods_styles(utf8read(styles), opts);
	var content = getzipstr(zip, 'content.xml');
	if(!content) throw new Error("Missing content.xml in ODS / UOF file");
	var wb = parse_content_xml(utf8read(content), opts, Styles);
	if(safegetzipfile(zip, 'meta.xml')) wb.Props = parse_core_props(getzipdata(zip, 'meta.xml'));
	wb.bookType = "ods";
	return wb;
}
function parse_fods(data/*:string*/, opts/*:?ParseOpts*/)/*:Workbook*/ {
	var wb = parse_content_xml(data, opts);
	wb.bookType = "fods";
	return wb;
}

