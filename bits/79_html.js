/* note: browser DOM element cannot see mso- style attrs, must parse */
var HTML_ = (function() {
	function html_to_sheet(str/*:string*/, _opts)/*:Workbook*/ {
		var opts = _opts || {};
		if(DENSE != null && opts.dense == null) opts.dense = DENSE;
		var ws/*:Worksheet*/ = opts.dense ? ([]/*:any*/) : ({}/*:any*/);
		var mtch/*:any*/ = str.match(/<table/i);
		if(!mtch) throw new Error("Invalid HTML: could not find <table>");
		var mtch2/*:any*/ = str.match(/<\/table/i);
		var i/*:number*/ = mtch.index, j/*:number*/ = mtch2 && mtch2.index || str.length;
		var rows = split_regex(str.slice(i, j), /(:?<tr[^>]*>)/i, "<tr>");
		var R = -1, C = 0, RS = 0, CS = 0;
		var range/*:Range*/ = {s:{r:10000000, c:10000000},e:{r:0,c:0}};
		var merges/*:Array<Range>*/ = [];
		for(i = 0; i < rows.length; ++i) {
			var row = rows[i].trim();
			var hd = row.slice(0,3).toLowerCase();
			if(hd == "<tr") { ++R; if(opts.sheetRows && opts.sheetRows <= R) { --R; break; } C = 0; continue; }
			if(hd != "<td" && hd != "<th") continue;
			var cells = row.split(/<\/t[dh]>/i);
			for(j = 0; j < cells.length; ++j) {
				var cell = cells[j].trim();
				if(!cell.match(/<t[dh]/i)) continue;
				var m = cell, cc = 0;
				/* TODO: parse styles etc */
				while(m.charAt(0) == "<" && (cc = m.indexOf(">")) > -1) m = m.slice(cc+1);
				var tag = parsexmltag(cell.slice(0, cell.indexOf(">")));
				CS = tag.colspan ? +tag.colspan : 1;
				if((RS = +tag.rowspan)>1 || CS>1) merges.push({s:{r:R,c:C},e:{r:R + (RS||1) - 1, c:C + CS - 1}});
				var _t/*:string*/ = tag.t || "";
				/* TODO: generate stub cells */
				if(!m.length) { C += CS; continue; }
				m = htmldecode(m);
				if(range.s.r > R) range.s.r = R; if(range.e.r < R) range.e.r = R;
				if(range.s.c > C) range.s.c = C; if(range.e.c < C) range.e.c = C;
				if(!m.length) continue;
				var o/*:Cell*/ = {t:'s', v:m};
				if(opts.raw || !m.trim().length || _t == 's'){}
				else if(m === 'TRUE') o = {t:'b', v:true};
				else if(m === 'FALSE') o = {t:'b', v:false};
				else if(!isNaN(fuzzynum(m))) o = {t:'n', v:fuzzynum(m)};
				else if(!isNaN(fuzzydate(m).getDate())) {
					o = ({t:'d', v:parseDate(m)}/*:any*/);
					if(!opts.cellDates) o = ({t:'n', v:datenum(o.v)}/*:any*/);
					o.z = opts.dateNF || SSF._table[14];
				}
				if(opts.dense) { if(!ws[R]) ws[R] = []; ws[R][C] = o; }
				else ws[encode_cell({r:R, c:C})] = o;
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
		var M/*:Array<Range>*/ = (ws['!merges'] ||[]);
		var oo/*:Array<string>*/ = [];
		var nullcell = "<td>" + (o.editable ? '<span contenteditable="true"></span>' : "" ) + "</td>";
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
			sp.t = cell.t;
			if(o.editable) w = '<span contenteditable="true">' + w + '</span>';
			sp.id = "sjs-" + coord;
			oo.push(writextag('td', w, sp));
		}
		var preamble = "<tr>";
		return preamble + oo.join("") + "</tr>";
	}
	function make_html_preamble(ws/*:Worksheet*/, R/*:Range*/, o/*:Sheet2HTMLOpts*/)/*:string*/ {
		var out/*:Array<string>*/ = [];
		return out.join("") + '<table' + (o && o.id ? ' id="' + o.id + '"' : "") + '>';
	}
	var _BEGIN = '<html><head><meta charset="utf-8"/><title>SheetJS Table Export</title></head><body>';
	var _END = '</body></html>';
	function sheet_to_html(ws/*:Worksheet*/, opts/*:?Sheet2HTMLOpts*//*, wb:?Workbook*/)/*:string*/ {
		var o = opts || {};
		var header = o.header != null ? o.header : _BEGIN;
		var footer = o.footer != null ? o.footer : _END;
		var out/*:Array<string>*/ = [header];
		var r = decode_range(ws['!ref']);
		o.dense = Array.isArray(ws);
		out.push(make_html_preamble(ws, r, o));
		for(var R = r.s.r; R <= r.e.r; ++R) out.push(make_html_row(ws, r, R, o));
		out.push("</table>" + footer);
		return out.join("");
	}

	return {
		to_workbook: html_to_book,
		to_sheet: html_to_sheet,
		_row: make_html_row,
		BEGIN: _BEGIN,
		END: _END,
		_preamble: make_html_preamble,
		from_sheet: sheet_to_html
	};
})();

function parse_dom_table(table/*:HTMLElement*/, _opts/*:?any*/)/*:Worksheet*/ {
	var opts = _opts || {};
	if(DENSE != null) opts.dense = DENSE;
	var ws/*:Worksheet*/ = opts.dense ? ([]/*:any*/) : ({}/*:any*/);
	var rows/*:HTMLCollection<HTMLTableRowElement>*/ = table.getElementsByTagName('tr');
	var sheetRows = opts.sheetRows || 10000000;
	var range/*:Range*/ = {s:{r:0,c:0},e:{r:0,c:0}};
	var merges/*:Array<Range>*/ = [], midx = 0;
	var rowinfo/*:Array<RowInfo>*/ = [];
	var _R = 0, R = 0, _C, C, RS, CS;
	for(; _R < rows.length && R < sheetRows; ++_R) {
		var row/*:HTMLTableRowElement*/ = rows[_R];
		if (is_dom_element_hidden(row)) {
			if (opts.display) continue;
			rowinfo[R] = {hidden: true};
		}
		var elts/*:HTMLCollection<HTMLTableCellElement>*/ = (row.children/*:any*/);
		for(_C = C = 0; _C < elts.length; ++_C) {
			var elt/*:HTMLTableCellElement*/ = elts[_C];
			if (opts.display && is_dom_element_hidden(elt)) continue;
			var v/*:string*/ = htmldecode(elt.innerHTML);
			for(midx = 0; midx < merges.length; ++midx) {
				var m/*:Range*/ = merges[midx];
				if(m.s.c == C && m.s.r <= R && R <= m.e.r) { C = m.e.c+1; midx = -1; }
			}
			/* TODO: figure out how to extract nonstandard mso- style */
			CS = +elt.getAttribute("colspan") || 1;
			if((RS = +elt.getAttribute("rowspan"))>0 || CS>1) merges.push({s:{r:R,c:C},e:{r:R + (RS||1) - 1, c:C + CS - 1}});
			var o/*:Cell*/ = {t:'s', v:v};
			var _t/*:string*/ = elt.getAttribute("t") || "";
			if(v != null) {
				if(v.length == 0) o.t = _t || 'z';
				else if(opts.raw || v.trim().length == 0 || _t == "s"){}
				else if(v === 'TRUE') o = {t:'b', v:true};
				else if(v === 'FALSE') o = {t:'b', v:false};
				else if(!isNaN(fuzzynum(v))) o = {t:'n', v:fuzzynum(v)};
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
		++R;
	}
	if(merges.length) ws['!merges'] = merges;
	if(rowinfo.length) ws['!rows'] = rowinfo;
	range.e.r = R - 1;
	ws['!ref'] = encode_range(range);
	if(R >= sheetRows) ws['!fullref'] = encode_range((range.e.r = rows.length-_R+R-1,range)); // We can count the real number of rows to parse but we don't to improve the performance
	return ws;
}

function table_to_book(table/*:HTMLElement*/, opts/*:?any*/)/*:Workbook*/ {
	return sheet_to_workbook(parse_dom_table(table, opts), opts);
}

function is_dom_element_hidden(element/*:HTMLElement*/)/*:boolean*/ {
	var display/*:string*/ = '';
	var get_computed_style/*:?function*/ = get_get_computed_style_function(element);
	if(get_computed_style) display = get_computed_style(element).getPropertyValue('display');
	if(!display) display = element.style.display; // Fallback for cases when getComputedStyle is not available (e.g. an old browser or some Node.js environments) or doesn't work (e.g. if the element is not inserted to a document)
	return display === 'none';
}

/* global getComputedStyle */
function get_get_computed_style_function(element/*:HTMLElement*/)/*:?function*/ {
	// The proper getComputedStyle implementation is the one defined in the element window
	if(element.ownerDocument.defaultView && typeof element.ownerDocument.defaultView.getComputedStyle === 'function') return element.ownerDocument.defaultView.getComputedStyle;
	// If it is not available, try to get one from the global namespace
	if(typeof getComputedStyle === 'function') return getComputedStyle;
	return null;
}
