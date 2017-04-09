/* TODO: in browser attach to DOM; in node use an html parser */
function parse_html(str/*:string*/, _opts)/*:Workbook*/ {
	var opts = _opts || {};
	if(DENSE != null && opts.dense == null) opts.dense = DENSE;
	var ws/*:Worksheet*/ = opts.dense ? ([]/*:any*/) : ({}/*:any*/);
	var o/*:Workbook*/ = { SheetNames: ["Sheet1"], Sheets: {Sheet1:ws} };
	var i = str.indexOf("<table"), j = str.indexOf("</table");
	if(i == -1 || j == -1) throw new Error("Invalid HTML: missing <table> / </table> pair");
	var rows = str.slice(i, j).split(/<tr[^>]*>/);
	var R = 0, C = 0;
	var range = {s:{r:10000000, c:10000000},e:{r:0,c:0}};
	for(i = 0; i < rows.length; ++i) {
		if(rows[i].substr(0,3) != "<td") continue;
		var cells = rows[i].split("</td>");
		for(j = 0; j < cells.length; ++j) {
			if(cells[j].substr(0,3) != "<td") continue;
			++C;
			var m = cells[j], cc = 0;
			/* TODO: parse styles etc */
			while(m.charAt(0) == "<" && (cc = m.indexOf(">")) > -1) m = m.slice(cc+1);
			while(m.indexOf(">") > -1) m = m.slice(0, m.lastIndexOf("<"));
			/* TODO: generate stub cells */
			if(!m.length) continue;
			if(range.s.r > R) range.s.r = R;
			if(range.e.r < R) range.e.r = R;
			if(range.s.c > C) range.s.c = C;
			if(range.e.c < C) range.e.c = C;
			if(opts.dense) {
				if(!ws[R]) ws[R] = [];
				if(Number(m) == Number(m)) ws[R][C] = {t:'n', v:+m};
				else ws[R][C] = {t:'s', v:m};
			} else {
				var coord/*:string*/ = encode_cell({r:R, c:C});
				/* TODO: value parsing */
				if(Number(m) == Number(m)) ws[coord] = {t:'n', v:+m};
				else ws[coord] = {t:'s', v:m};
			}
		}
		++R; C = 0;
	}
	ws['!ref'] = encode_range(range);
	return o;
}

function parse_dom_table(table/*:HTMLElement*/, _opts/*:?any*/)/*:Worksheet*/ {
	var opts = _opts || {};
	if(DENSE != null) opts.dense = DENSE;
	var ws/*:Worksheet*/ = opts.dense ? ([]/*:any*/) : ({}/*:any*/);
	var rows = table.getElementsByTagName('tr');
	var range = {s:{r:0,c:0},e:{r:rows.length - 1,c:0}};
	var merges = [], midx = 0;
	var R = 0, _C = 0, C = 0, RS = 0, CS = 0;
	for(; R < rows.length; ++R) {
		var row = rows[R];
		var elts = row.children;
		for(_C = C = 0; _C < elts.length; ++_C) {
			var elt = elts[_C], v = elts[_C].innerText;
			for(midx = 0; midx < merges.length; ++midx) {
				var m = merges[midx];
				if(m.s.c == C && m.s.r <= R && R <= m.e.r) { C = m.e.c+1; midx = -1; }
			}
			/* TODO: figure out how to extract nonstandard mso- style */
			CS = +elt.getAttribute("colspan") || 1;
			if((RS = +elt.getAttribute("rowspan"))>0) merges.push({s:{r:R,c:C},e:{r:R + RS - 1, c:C + CS - 1}});
			var o = {t:'s', v:v};
			if(v != null && v.length && !isNaN(Number(v))) o = {t:'n', v:Number(v)};
			if(opts.dense) { if(!ws[R]) ws[R] = []; ws[R][C] = o; }
			else ws[encode_cell({c:C, r:R})] = o;
			if(range.e.c < C) range.e.c = C;
			C += CS;
		}
	}
	ws['!merges'] = merges;
	ws['!ref'] = encode_range(range);
	return ws;
}

function table_to_book(table/*:HTMLElement*/, opts/*:?any*/)/*:Workbook*/ {
	return sheet_to_workbook(parse_dom_table(table, opts), opts);
}
