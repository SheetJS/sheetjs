var parse_content_xml = (function() {

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

	return function pcx(d, opts) {
		var str = xlml_normalize(d);
		var state = [], tmp;
		var tag;
		var NFtag, NF, pidx;
		var sheetag;
		var Sheets = {}, SheetNames = [], ws = {};
		var Rn, q;
		var ctag;
		var textp, textpidx, textptag;
		var R, C, range = {s: {r:1000000,c:10000000}, e: {r:0, c:0}};
		var number_format_map = {};
		var merges = [], mrange = {}, mR = 0, mC = 0;

		while((Rn = xlmlregex.exec(str))) switch(Rn[3]) {

			case 'table': // 9.1.2 <table:table>
				if(Rn[1]==='/') {
					if(range.e.c >= range.s.c && range.e.r >= range.s.r) ws['!ref'] = get_utils().encode_range(range);
					if(merges.length) ws['!merges'] = merges;
					SheetNames.push(sheetag.name);
					Sheets[sheetag.name] = ws;
				}
				else if(Rn[0].charAt(Rn[0].length-2) !== '/') {
					sheetag = parsexmltag(Rn[0]);
					R = C = -1;
					range.s.r = range.s.c = 10000000; range.e.r = range.e.c = 0;
					ws = {}; merges = [];
				}
				break;

			case 'table-row': // 9.1.3 <table:table-row>
				if(Rn[1] === '/') break;
				++R; C = -1; break;
			case 'covered-table-cell': // 9.1.5 table:covered-table-cell
				++C; break; /* stub */
			case 'table-cell':
				if(Rn[0].charAt(Rn[0].length-2) === '/') {
					ctag = parsexmltag(Rn[0]);
					if(ctag['number-columns-repeated']) C+= parseInt(ctag['number-columns-repeated'], 10);
					else ++C;
				}
				else if(Rn[1]!=='/') {
					++C;
					if(C > range.e.c) range.e.c = C;
					if(R > range.e.r) range.e.r = R;
					if(C < range.s.c) range.s.c = C;
					if(R < range.s.r) range.s.r = R;
					ctag = parsexmltag(Rn[0]);
					q = {t:ctag['value-type'], v:null};
					if(ctag['number-columns-spanned'] || ctag['number-rows-spanned']) {
						mR = parseInt(ctag['number-rows-spanned'],10) || 0;
						mC = parseInt(ctag['number-columns-spanned'],10) || 0;
						mrange = {s: {r:R,c:C}, e:{r:R + mR-1,c:C + mC-1}};
						merges.push(mrange);
					}
					/* 19.385 office:value-type */
					switch(q.t) {
						case 'boolean': q.t = 'b'; q.v = parsexmlbool(ctag['boolean-value']); break;
						case 'float': q.t = 'n'; q.v = parseFloat(ctag.value); break;
						case 'percentage': q.t = 'n'; q.v = parseFloat(ctag.value); break;
						case 'currency': q.t = 'n'; q.v = parseFloat(ctag.value); break;
						case 'date': q.t = 'n'; q.v = datenum(ctag['date-value']); q.z = 'm/d/yy'; break;
						case 'time': q.t = 'n'; q.v = parse_isodur(ctag['time-value'])/86400; break;
						case 'string': q.t = 's'; break;
						default: throw new Error('Unsupported value type ' + q.t);
					}
				} else {
					if(q.t === 's') q.v = textp;
					if(textp) q.w = textp;
					if(!(opts.sheetRows && opts.sheetRows < R)) ws[get_utils().encode_cell({r:R,c:C})] = q;
					q = null;
				}
				break; // 9.1.4 <table:table-cell>

			/* pure state */
			case 'document-content': // 3.1.3.2 <office:document-content>
			case 'spreadsheet': // 3.7 <office:spreadsheet>
			case 'scripts': // 3.12 <office:scripts>
			case 'font-face-decls': // 3.14 <office:font-face-decls>
				if(Rn[1]==='/'){if((tmp=state.pop())[0]!==Rn[3]) throw "Bad state: "+tmp;}
				else if(Rn[0].charAt(Rn[0].length-2) !== '/') state.push([Rn[3], true]);
				break;

			/* ignore state */
			case 'shapes': // 9.2.8 <table:shapes>
			case 'frame': // 10.4.2 <draw:frame>
				if(Rn[1]==='/'){if((tmp=state.pop())[0]!==Rn[3]) throw "Bad state: "+tmp;}
				else if(Rn[0].charAt(Rn[0].length-2) !== '/') state.push([Rn[3], false]);
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
					NFtag = parsexmltag(Rn[0]);
					state.push([Rn[3], true]);
				} break;

			case 'script': break; // 3.13 <office:script>
			case 'automatic-styles': break; // 3.15.3 <office:automatic-styles>

			case 'style': break; // 16.2 <style:style>
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
						tag = parsexmltag(Rn[0]);
						NF += number_formats[Rn[3]][tag.style==='long'?1:0]; break;
				} break;

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
						tag = parsexmltag(Rn[0]);
						NF += number_formats[Rn[3]][tag.style==='long'?1:0]; break;
				} break;

			case 'boolean-style': break; // 16.27.23 <number:boolean-style>
			case 'boolean': break; // 16.27.24 <number:boolean>
			case 'text-style': break; // 16.27.25 <number:text-style>
			case 'text': // 16.27.26 <number:text>
				if(Rn[0].substr(-2) === "/>") break;
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

			case 'body': break; // 3.3 16.9.6 19.726.3

			case 'forms': break; // 12.25.2 13.2
			case 'table-column': break; // 9.1.6 <table:table-column>

			case 'graphic-properties': break;
			case 'calculation-settings': break; // 9.4.1 <table:calculation-settings>
			case 'named-expressions': break; // 9.4.11 <table:named-expressions>
			case 'named-range': break; // 9.4.11 <table:named-range>
			case 'span': break; // <text:span>
			case 'p':
				if(Rn[1]==='/') textp = parse_text_p(str.slice(textpidx,Rn.index), textptag);
				else { textptag = parsexmltag(Rn[0]); textpidx = Rn.index + Rn[0].length; }
				break; // <text:p>
			case 's': break; // <text:s>
			case 'date': break; // <*:date>
			case 'annotation': break;

			case 'object': break; // 10.4.6.2 <draw:object>
			case 'title': break; // <*:title>
			case 'desc': break; // <*:desc>

			case 'database-ranges': break; // 9.4.14 <table:database-ranges>
			case 'database-range': break; // 9.4.15 <table:database-range>
			case 'filter': break; // 9.5.2 <table:filter>
			case 'filter-and': break; // 9.5.3 <table:filter-and>
			case 'filter-or': break; // 9.5.4 <table:filter-or>
			case 'filter-condition': break; // 9.5.5 <table:filter-condition>
			default: if(opts.WTF) throw Rn;
		}
		var out = {
			Sheets: Sheets,
			SheetNames: SheetNames
		};
		return out;
	};
})();
