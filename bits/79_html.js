/* TODO: in browser attach to DOM; in node use an html parser */
function parse_html(str/*:string*/, opts)/*:Workbook*/ {
	var ws/*:Worksheet*/ = ({}/*:any*/);
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
			var coord/*:string*/ = encode_cell({r:R, c:C});
			/* TODO: value parsing */
			if(m == +m) ws[coord] = {t:'n', v:+m};
			else ws[coord] = {t:'s', v:m};
		}
		++R; C = 0;
	}
	ws['!ref'] = encode_range(range);
	return o;
}
