var parse_content_xml = (function() {

	var parse_text_p = function(text, tag) {
		return unescapexml(text.replace(/<text:s\/>/g," ").replace(/<[^>]*>/g,""));
	};

	var number_formats = {
		/* ods name: [short ssf fmt, long ssf fmt] */
		day: ["d", "dd"],
		month: ["m", "mm"],
		year: ["y", "yy"],
		hours: ["h", "hh"],
		minutes: ["m", "mm"],
		seconds: ["s", "ss"],
		"am-pm": ["A/P", "AM/PM"],
		"day-of-week": ["ddd", "dddd"]
	};

	return function pcx(d/*:string*/, _opts)/*:Workbook*/ {
		var opts = _opts || {};
		var str = xlml_normalize(d);
		var state/*:Array<any>*/ = [], tmp;
		var tag/*:: = {}*/;
		var NFtag = {name:""}, NF = "", pidx = 0;
		var sheetag/*:: = {name:"", '名称':""}*/;
		var rowtag/*:: = {'行号':""}*/;
		var Sheets = {}, SheetNames/*:Array<string>*/ = [], ws = {};
		var Rn, q/*:: = ({t:"", v:null, z:null, w:""}:any)*/;
		var ctag = {value:""};
		var textp = "", textpidx = 0, textptag/*:: = {}*/;
		var R = -1, C = -1, range = {s: {r:1000000,c:10000000}, e: {r:0, c:0}};
		var number_format_map = {};
		var merges = [], mrange = {}, mR = 0, mC = 0;
		var arrayf = [];
		var rept = 1, isstub = false;
		var i = 0;
		xlmlregex.lastIndex = 0;
		while((Rn = xlmlregex.exec(str))) switch((Rn[3]=Rn[3].replace(/_.*$/,""))) {

			case 'table': case '工作表': // 9.1.2 <table:table>
				if(Rn[1]==='/') {
					if(range.e.c >= range.s.c && range.e.r >= range.s.r) ws['!ref'] = encode_range(range);
					if(merges.length) ws['!merges'] = merges;
					sheetag.name = utf8read(sheetag['名称'] || sheetag.name);
					SheetNames.push(sheetag.name);
					Sheets[sheetag.name] = ws;
				}
				else if(Rn[0].charAt(Rn[0].length-2) !== '/') {
					sheetag = parsexmltag(Rn[0], false);
					R = C = -1;
					range.s.r = range.s.c = 10000000; range.e.r = range.e.c = 0;
					ws = {}; merges = [];
				}
				break;

			case 'table-row': case '行': // 9.1.3 <table:table-row>
				if(Rn[1] === '/') break;
				rowtag = parsexmltag(Rn[0], false);
				if(rowtag['行号']) R = rowtag['行号'] - 1; else ++R;
				C = -1; break;
			case 'covered-table-cell': // 9.1.5 table:covered-table-cell
				++C; break; /* stub */
			case 'table-cell': case '数据':
				if(Rn[0].charAt(Rn[0].length-2) === '/') {
					ctag = parsexmltag(Rn[0], false);
					if(ctag['number-columns-repeated']) C+= parseInt(ctag['number-columns-repeated'], 10);
					else ++C;
				}
				else if(Rn[1]!=='/') {
					++C;
					rept = 1;
					if(C > range.e.c) range.e.c = C;
					if(R > range.e.r) range.e.r = R;
					if(C < range.s.c) range.s.c = C;
					if(R < range.s.r) range.s.r = R;
					ctag = parsexmltag(Rn[0], false);
					q = ({t:ctag['数据类型'] || ctag['value-type'], v:null/*:: , z:null, w:""*/}/*:any*/);
					if(opts.cellFormula) {
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
					if(ctag['number-columns-repeated']) rept = parseInt(ctag['number-columns-repeated'], 10);

					/* 19.385 office:value-type */
					switch(q.t) {
						case 'boolean': q.t = 'b'; q.v = parsexmlbool(ctag['boolean-value']); break;
						case 'float': q.t = 'n'; q.v = parseFloat(ctag.value); break;
						case 'percentage': q.t = 'n'; q.v = parseFloat(ctag.value); break;
						case 'currency': q.t = 'n'; q.v = parseFloat(ctag.value); break;
						case 'date': q.t = 'n'; q.v = datenum(new Date(ctag['date-value'])); q.z = 'm/d/yy'; break;
						case 'time': q.t = 'n'; q.v = parse_isodur(ctag['time-value'])/86400; break;
						case 'number': q.t = 'n'; q.v = parseFloat(ctag['数据数值']); break;
						default:
							if(q.t === 'string' || q.t === 'text' || !q.t) {
								q.t = 's';
								if(ctag['string-value'] != null) textp = ctag['string-value'];
							} else throw new Error('Unsupported value type ' + q.t);
					}
				} else {
					isstub = false;
					if(q.t === 's') {
						q.v = textp || '';
						isstub = textpidx == 0;
					}
					if(textp) q.w = textp;
					if(!isstub || opts.cellStubs) {
						if(!(opts.sheetRows && opts.sheetRows < R)) {
							ws[encode_cell({r:R,c:C})] = q;
							while(--rept > 0) ws[encode_cell({r:R,c:++C})] = dup(q);
							if(range.e.c <= C) range.e.c = C;
						}
					} else { C += rept; rept = 0; }
					q = {/*:: t:"", v:null, z:null, w:""*/};
					textp = "";
				}
				break; // 9.1.4 <table:table-cell>

			/* pure state */
			case 'document': // TODO: <office:document> is the root for FODS
			case 'document-content': case '电子表格文档': // 3.1.3.2 <office:document-content>
			case 'spreadsheet': case '主体': // 3.7 <office:spreadsheet>
			case 'scripts': // 3.12 <office:scripts>
			case 'styles': // TODO <office:styles>
			case 'font-face-decls': // 3.14 <office:font-face-decls>
				if(Rn[1]==='/'){if((tmp=state.pop())[0]!==Rn[3]) throw "Bad state: "+tmp;}
				else if(Rn[0].charAt(Rn[0].length-2) !== '/') state.push([Rn[3], true]);
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
			case 'annotation': // 14.1 <office:annotation>
			case 'event-listeners': // TODO
				if(Rn[1]==='/'){if((tmp=state.pop())[0]!==Rn[3]) throw "Bad state: "+tmp;}
				else if(Rn[0].charAt(Rn[0].length-2) !== '/') state.push([Rn[3], false]);
				textp = ""; textpidx = 0;
				break;

			case 'scientific-number': // TODO: <number:scientific-number>
				break;
			case 'currency-symbol': // TODO: <number:currency-symbol>
				break;
			case 'currency-style': // TODO: <number:currency-style>
				break;
			case 'number-style': // 16.27.2 <number:number-style>
			case 'percentage-style': // 16.27.9 <number:percentage-style>
			case 'date-style': // 16.27.10 <number:date-style>
			case 'time-style': // 16.27.18 <number:time-style>
				if(Rn[1]==='/'){
					number_format_map[NFtag.name] = NF;
					if((tmp=state.pop())[0]!==Rn[3]) throw "Bad state: "+tmp;
				} else if(Rn[0].charAt(Rn[0].length-2) !== '/') {
					NF = "";
					NFtag = parsexmltag(Rn[0], false);
					state.push([Rn[3], true]);
				} break;

			case 'script': break; // 3.13 <office:script>
			case 'libraries': break; // TODO: <ooo:libraries>
			case 'automatic-styles': break; // 3.15.3 <office:automatic-styles>
			case 'master-styles': break; // TODO: <office:automatic-styles>

			case 'default-style': // TODO: <style:default-style>
			case 'page-layout': break; // TODO: <style:page-layout>
			case 'style': break; // 16.2 <style:style>
			case 'map': break; // 16.3 <style:map>
			case 'font-face': break; // 16.21 <style:font-face>

			case 'paragraph-properties': break; // 17.6 <style:paragraph-properties>
			case 'table-properties': break; // 17.15 <style:table-properties>
			case 'table-column-properties': break; // 17.16 <style:table-column-properties>
			case 'table-row-properties': break; // 17.17 <style:table-row-properties>
			case 'table-cell-properties': break; // 17.18 <style:table-cell-properties>

			case 'number': // 16.27.3 <number:number>
				switch(state[state.length-1][0]) {
					case 'time-style':
					case 'date-style':
						tag = parsexmltag(Rn[0], false);
						NF += number_formats[Rn[3]][tag.style==='long'?1:0]; break;
				} break;

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
				switch(state[state.length-1][0]) {
					case 'time-style':
					case 'date-style':
						tag = parsexmltag(Rn[0], false);
						NF += number_formats[Rn[3]][tag.style==='long'?1:0]; break;
				} break;

			case 'boolean-style': break; // 16.27.23 <number:boolean-style>
			case 'boolean': break; // 16.27.24 <number:boolean>
			case 'text-style': break; // 16.27.25 <number:text-style>
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
			case 'text-content': break; // 16.27.27 <number:text-content>
			case 'text-properties': break; // 16.27.27 <style:text-properties>

			case 'body': case '电子表格': break; // 3.3 16.9.6 19.726.3

			case 'forms': break; // 12.25.2 13.2
			case 'table-column': break; // 9.1.6 <table:table-column>

			case 'null-date': break; // 9.4.2 <table:null-date> TODO: date1904

			case 'graphic-properties': break; // 17.21 <style:graphic-properties>
			case 'calculation-settings': break; // 9.4.1 <table:calculation-settings>
			case 'named-expressions': break; // 9.4.11 <table:named-expressions>
			case 'named-range': break; // 9.4.12 <table:named-range>
			case 'named-expression': break; // 9.4.13 <table:named-expression>
			case 'sort': break; // 9.4.19 <table:sort>
			case 'sort-by': break; // 9.4.20 <table:sort-by>
			case 'sort-groups': break; // 9.4.22 <table:sort-groups>

			case 'span': break; // <text:span>
			case 'line-break': break; // 6.1.5 <text:line-break>
			case 'p': case '文本串':
				if(Rn[1]==='/') textp = parse_text_p(str.slice(textpidx,Rn.index), textptag);
				else { textptag = parsexmltag(Rn[0], false); textpidx = Rn.index + Rn[0].length; }
				break; // <text:p>
			case 's': break; // <text:s>
			case 'date': break; // <*:date>

			case 'object': break; // 10.4.6.2 <draw:object>
			case 'title': case '标题': break; // <*:title> OR <uof:标题>
			case 'desc': break; // <*:desc>

			case 'table-source': break; // 9.2.6

			case 'iteration': break; // 9.4.3 <table:iteration>
			case 'content-validations': break; // 9.4.4 <table:
			case 'content-validation': break; // 9.4.5 <table:
			case 'error-message': break; // 9.4.7 <table:
			case 'database-ranges': break; // 9.4.14 <table:database-ranges>
			case 'database-range': break; // 9.4.15 <table:database-range>
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

			case 'event-listener': // TODO
			/* TODO: FODS Properties */
			case 'initial-creator':
			case 'creator':
			case 'creation-date':
			case 'generator':
			case 'document-statistic':
			case 'user-defined':
				break;

			/* TODO: FODS Config */
			case 'config-item':
				break;

			/* TODO: style tokens */
			case 'page-number': break; // TODO <text:page-number>
			case 'page-count': break; // TODO <text:page-count>
			case 'time': break; // TODO <text:time>

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

			case 'a': break; // 6.1.8 hyperlink

			/* non-standard */
			case 'table-protection': break;
			case 'data-pilot-grand-total': break; // <table:
			default:
				if(Rn[2] === 'dc:') break; // TODO: properties
				if(Rn[2] === 'draw:') break; // TODO: drawing
				if(Rn[2] === 'style:') break; // TODO: styles
				if(Rn[2] === 'calcext:') break; // ignore undocumented extensions
				if(Rn[2] === 'loext:') break; // ignore undocumented extensions
				if(Rn[2] === 'uof:') break; // TODO: uof
				if(Rn[2] === '表:') break; // TODO: uof
				if(Rn[2] === '字:') break; // TODO: uof
				if(opts.WTF) throw Rn;
		}
		var out = {
			Sheets: Sheets,
			SheetNames: SheetNames
		};
		return out;
	};
})();
