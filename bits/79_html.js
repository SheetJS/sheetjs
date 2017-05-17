/* note: browser DOM element cannot see mso- style attrs, must parse */
var HTML_ = (function() {
	function html_to_sheet(str/*:string*/, _opts)/*:Workbook*/ {
		var opts = _opts || {};
		if(DENSE != null && opts.dense == null) opts.dense = DENSE;
		var ws/*:Worksheet*/ = opts.dense ? ([]/*:any*/) : ({}/*:any*/);
		var i = str.indexOf("<table"), j = str.indexOf("</table");
		if(i == -1 || j == -1) throw new Error("Invalid HTML: missing <table> / </table> pair");
		var rows = str.slice(i, j).split(/(:?<tr[^>]*>)/);
		var R = -1, C = 0, RS = 0, CS = 0;
		var range = {s:{r:10000000, c:10000000},e:{r:0,c:0}};
		var merges = [], midx = 0;
		for(i = 0; i < rows.length; ++i) {
			var row = rows[i].trim();
			if(row.substr(0,3) == "<tr") { ++R; C = 0; continue; }
			if(row.substr(0,3) != "<td") continue;
			var cells = row.split("</td>");
			for(j = 0; j < cells.length; ++j) {
				var cell = cells[j].trim();
				if(cell.substr(0,3) != "<td") continue;
				var m = cell, cc = 0;
				/* TODO: parse styles etc */
				while(m.charAt(0) == "<" && (cc = m.indexOf(">")) > -1) m = m.slice(cc+1);
				while(m.indexOf(">") > -1) m = m.slice(0, m.lastIndexOf("<"));
				var tag = parsexmltag(cell.slice(0, cell.indexOf(">")));
				CS = tag.colspan ? +tag.colspan : 1;
				if((RS = +tag.rowspan)>0 || CS>1) merges.push({s:{r:R,c:C},e:{r:R + (RS||1) - 1, c:C + CS - 1}});
				/* TODO: generate stub cells */
				if(!m.length) { C += CS; continue; }
				m = unescapexml(m).replace(/[\r\n]/g,"");
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
				C += CS;
			}
		}
		ws['!ref'] = encode_range(range);
		return ws;
	}
	function html_to_book(str/*:string*/, opts)/*:Workbook*/ {
		return sheet_to_workbook(html_to_sheet(str, opts), opts);
	}
	function make_html_row(ws/*:Worksheet*/, r/*:Range*/, R/*:number*/, o/*:Sheet2HTMLOpts*/)/*:string*/ {
		var M = (ws['!merges'] ||[]);
		var oo = [];
		var nullcell = "<td" + (o.editable ? ' contenteditable="true"' : "" ) + "></td>";
		for(var C = r.s.c; C <= r.e.c; ++C) {
			var RS = 0, CS = 0;
			for(var j = 0; j < M.length; ++j) {
				if(M[j].s.r > R || M[j].s.c > C) continue;
				if(M[j].e.r < R || M[j].e.c < C) continue;
				if(M[j].s.r < R || M[j].s.c < C) { RS = -1; break; }
				RS = M[j].e.r - M[j].s.r + 1; CS = M[j].e.c - M[j].s.c + 1; break;
			}
			if(RS < 0) continue;
			var coord = encode_cell({r:R,c:C});
			var cell = o.dense ? (ws[R]||[])[C] : ws[coord];
			if(!cell || cell.v == null) { oo.push(nullcell); continue; }
			/* TODO: html entities */
			var w = cell.h || escapexml(cell.w || (format_cell(cell), cell.w) || "");
			var sp = {};
			if(RS > 1) sp.rowspan = RS;
			if(CS > 1) sp.colspan = CS;
			if(o.editable) sp.contenteditable = "true";
			oo.push(writextag('td', w, sp));
		}
		return "<tr>" + oo.join("") + "</tr>";
	}
	var _BEGIN = "<html><head><title>SheetJS Table Export</title></head><body><table>";
	var _END = "</table></body></html>";
	function sheet_to_html(ws/*:Worksheet*/, opts/*:?Sheet2HTMLOpts*/)/*:string*/ {
		var o = opts || {};
		var out/*:Array<string>*/ = [];
		var r = decode_range(ws['!ref']);
		o.dense = Array.isArray(ws);
		for(var R = r.s.r; R <= r.e.r; ++R) out.push(make_html_row(ws, r, R, o));
		var header = o.header != null ? o.header : _BEGIN;
		var footer = o.footer != null ? o.footer : _END;
		return header + out.join("") + footer ;
	}

	return {
		to_workbook: html_to_book,
		to_sheet: html_to_sheet,
		_row: make_html_row,
		BEGIN: _BEGIN,
		END: _END,
		from_sheet: sheet_to_html
	};
})();

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
			var elt = elts[_C], v = elts[_C].innerText || elts[_C].textContent;
			for(midx = 0; midx < merges.length; ++midx) {
				var m = merges[midx];
				if(m.s.c == C && m.s.r <= R && R <= m.e.r) { C = m.e.c+1; midx = -1; }
			}
			/* TODO: figure out how to extract nonstandard mso- style */
			CS = +elt.getAttribute("colspan") || 1;
			if((RS = +elt.getAttribute("rowspan"))>0 || CS>1) merges.push({s:{r:R,c:C},e:{r:R + (RS||1) - 1, c:C + CS - 1}});
			var o/*:Cell*/ = {t:'s', v:v};
			if(v != null && v.length) {
				if(!isNaN(Number(v))) o = {t:'n', v:Number(v)};
				else if(!isNaN(fuzzydate(v).getDate())) {
					o = ({t:'d', v:parseDate(v)}/*:any*/);
					if(!opts.cellDates) o = ({t:'n', v:datenum(o.v)}/*:any*/);
					o.z = opts.dateNF || SSF._table[14];
				}
			}
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
