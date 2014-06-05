/* 18.3 Worksheets */
function parse_ws_xml(data, opts, rels) {
	if(!data) return data;
	/* 18.3.1.99 worksheet CT_Worksheet */
	var s = {}, mtch;

	/* 18.3.1.35 dimension CT_SheetDimension ? */
	var ref = data.match(/<dimension ref="([^"]*)"\s*\/>/);
	if(ref && ref.length == 2 && ref[1].indexOf(":") !== -1) s["!ref"] = ref[1];

	/* 18.3.1.55 mergeCells CT_MergeCells */
	var mergecells = [];
	if(data.match(/<\/mergeCells>/)) {
		var merges = data.match(/<mergeCell ref="([A-Z0-9:]+)"\s*\/>/g);
		mergecells = merges.map(function(range) {
			return decode_range(/<mergeCell ref="([A-Z0-9:]+)"\s*\/>/.exec(range)[1]);
		});
	}

	/* 18.3.1.17 cols CT_Cols */
	var columns = [];
	if(opts.cellStyles && data.match(/<\/cols>/)) {
		/* 18.3.1.13 col CT_Col */
		var cols = data.match(/<col[^>]*\/>/g);
		parse_ws_xml_cols(columns, cols);
	}

	var refguess = {s: {r:1000000, c:1000000}, e: {r:0, c:0} };
	var sidx = 0;

	var match_v = matchtag("v"), match_f = matchtag("f");
	/* 18.3.1.80 sheetData CT_SheetData ? */
	mtch=data.match(/<(?:\w+:)?sheetData>([^\u2603]*)<\/(?:\w+:)?sheetData>/m);
	if(mtch) for(var marr = mtch[1].split(/<\/(?:\w+:)?row>/), mt = 0; mt != marr.length; ++mt) {
		x = marr[mt];
		if(x.length === 0 || x.trim().length === 0) continue;

		/* 18.3.1.73 row CT_Row */
		for(var ri = 0; ri != x.length; ++ri) if(x.charCodeAt(ri) === 62) break; ++ri;
		var row = parsexmltag(x.substr(0,ri));
		if(opts.sheetRows && opts.sheetRows < +row.r) continue;
		if(refguess.s.r > row.r - 1) refguess.s.r = row.r - 1;
		if(refguess.e.r < row.r - 1) refguess.e.r = row.r - 1;

		/* 18.3.1.4 c CT_Cell */
		var cells = x.substr(ri).split(/<(?:\w+:)?c /);
		for(var ix = 0, c=cells[0]; ix != cells.length; ++ix) {
			c = cells[ix];
			if(c.length === 0 || c.trim().length === 0) continue;
			var cref = c.match(/r=["']([^"']*)["']/), idx = ix, i=0, cc=0, a1="";
			c = "<c " + c;
			if(cref && cref.length == 2) {
				idx = 0; a1=cref[1];
				for(i=0; i != a1.length; ++i) {
					if((cc=a1.charCodeAt(i)-64) < 1 || cc > 26) break;
					idx = 26*idx + cc;
				}
				--idx;
			}

			for(var ci = 0; ci != c.length; ++ci) if(c.charCodeAt(ci) === 62) break; ++ci;
			var cell = parsexmltag(c.substr(0,ci), true);
			var d = c.substr(ci);
			var p = {};

			var x=d.match(match_v);if(x)p.v=unescapexml(x[1]);
			if(opts.cellFormula) {x=d.match(match_f);if(x)p.f=unescapexml(x[1]);}

			/* SCHEMA IS ACTUALLY INCORRECT HERE.  IF A CELL HAS NO T, EMIT "" */
			if(cell.t === undefined && p.v === undefined) {
				if(!opts.sheetStubs) continue;
				p.t = "str"; p.v = undefined;
			}
			else p.t = cell.t || "n";
			if(refguess.s.c > idx) refguess.s.c = idx;
			if(refguess.e.c < idx) refguess.e.c = idx;
			/* 18.18.11 t ST_CellType */
			switch(p.t) {
				case 'n': p.v = parseFloat(p.v); break;
				case 's': {
					sidx = parseInt(p.v, 10);
					p.v = strs[sidx].t;
					p.r = strs[sidx].r;
					if(opts.cellHTML) p.h = strs[sidx].h;
				} break;
				case 'str': if(p.v) p.v = utf8read(p.v); break;
				case 'inlineStr':
					var is = d.match(/<is>(.*)<\/is>/);
					is = is ? parse_si(is[1]) : {t:"",r:""};
					p.t = 'str'; p.v = is.t;
					break; // inline string
				case 'b': p.v = parsexmlbool(p.v); break;
				case 'd':
					p.v = datenum(p.v);
					p.t = 'n';
					break;
				/* in case of error, stick value in .raw */
				case 'e': p.raw = RBErr[p.v]; break;
			}

			/* formatting */
			var fmtid = 0, fillid = 0;
			if(cell.s && styles.CellXf) {
				var cf = styles.CellXf[cell.s];
				if(cf && cf.numFmtId) fmtid = cf.numFmtId;
				if(opts.cellStyles && cf && cf.fillId) fillid = cf.fillId;
			}
			safe_format(p, fmtid, fillid, opts);
			s[cell.r] = p;
		}
	}

	/* 18.3.1.48 hyperlinks CT_Hyperlinks */
	if(data.match(/<\/hyperlinks>/)) parse_ws_xml_hlinks(s, data.match(/<hyperlink[^>]*\/>/g), rels);

	if(!s["!ref"] && refguess.e.c >= refguess.s.c && refguess.e.r >= refguess.s.r) s["!ref"] = encode_range(refguess);
	if(opts.sheetRows && s["!ref"]) {
		var tmpref = decode_range(s["!ref"]);
		if(opts.sheetRows < +tmpref.e.r) {
			tmpref.e.r = opts.sheetRows - 1;
			if(tmpref.e.r > refguess.e.r) tmpref.e.r = refguess.e.r;
			if(tmpref.e.r < tmpref.s.r) tmpref.s.r = tmpref.e.r;
			if(tmpref.e.c > refguess.e.c) tmpref.e.c = refguess.e.c;
			if(tmpref.e.c < tmpref.s.c) tmpref.s.c = tmpref.e.c;
			s["!fullref"] = s["!ref"];
			s["!ref"] = encode_range(tmpref);
		}
	}
	if(mergecells.length > 0) s["!merges"] = mergecells;
	if(columns.length > 0) s["!cols"] = columns;
	return s;
}


var parse_ws_xml_hlinks = function(s, data, rels) {
	data.forEach(function(h) {
		var val = parsexmltag(h, true);
		if(!val.ref) return;
		var rel = rels['!id'][val.id];
		if(rel) {
			val.Target = rel.Target;
			if(val.location) val.Target += "#"+val.location;
			val.Rel = rel;
		}
		var rng = decode_range(val.ref);
		for(var R=rng.s.r;R<=rng.e.r;++R) for(var C=rng.s.c;C<=rng.e.c;++C) {
			var addr = encode_cell({c:C,r:R});
			if(!s[addr]) s[addr] = {t:"str",v:undefined};
			s[addr].l = val;
		}
	});
};

var parse_ws_xml_cols = function(columns, cols) {
	var seencol = false;
	for(var coli = 0; coli != cols.length; ++coli) {
		var coll = parsexmltag(cols[coli], true);
		var colm=Number(coll.min)-1, colM=Number(coll.max)-1;
		delete coll.min; delete coll.max;
		if(!seencol && coll.width) { seencol = true; find_mdw(+coll.width, coll); }
		if(coll.width) {
			coll.wpx = width2px(+coll.width);
			coll.wch = px2char(coll.wpx);
			coll.MDW = MDW;
		}
		while(colm <= colM) columns[colm++] = coll;
	}
};

var write_ws_xml_cols = function(ws, cols) {
	var o = ["<cols>"], col, width;
	for(var i = 0; i != cols.length; ++i) {
		if(!(col = cols[i])) continue;
		var p = {min:i+1,max:i+1};
		/* wch (chars), wpx (pixels) */
		width = -1;
		if(col.wpx) width = px2char(col.wpx);
		else if(col.wch) width = col.wch;
		if(width > -1) { p.width = char2width(width); p.customWidth= 1; }
		o.push(writextag('col', null, p));
	}
	o.push("</cols>");
	return o.join("");
};

var write_ws_xml_cell = function(cell, ref, ws, opts, idx, wb) {
	var vv = cell.v; if(cell.t == 'b') vv = cell.v ? "1" : "0";
	var v = writextag('v', escapexml(String(vv))), o = {r:ref};
	o.s = get_cell_style(opts.cellXfs, cell, opts);
	if(o.s === 0 || o.s === "0") delete o.s;
	/* TODO: cell style */
	if(typeof cell.v === 'undefined') return "";
	switch(cell.t) {
		case 's': case 'str':
			if(opts.bookSST) {
				v = writextag('v', String(get_sst_id(opts.Strings, cell.v)));
				o.t = "s"; return writextag('c', v, o);
			}
			o.t = "str"; return writextag('c', v, o);
		case 'n': delete o.t; return writextag('c', v, o);
		case 'b': o.t = "b"; return writextag('c', v, o);
		case 'e': o.t = "e"; return writextag('c', v, o);
	}
};

var write_ws_xml_data = function(ws, opts, idx, wb) {
	var o = [], r = [], range = utils.decode_range(ws['!ref']), cell, ref;
	for(var R = range.s.r; R <= range.e.r; ++R) {
		r = [];
		for(var C = range.s.c; C <= range.e.c; ++C) {
			ref = utils.encode_cell({c:C, r:R});
			if(!ws[ref]) continue;
			if((cell = write_ws_xml_cell(ws[ref], ref, ws, opts, idx, wb))) r.push(cell);
		}
		if(r.length) o.push(writextag('row', r.join(""), {r:encode_row(R)}));
	}
	return o.join("");
};

var WS_XML_ROOT = writextag('worksheet', null, {
	'xmlns': XMLNS.main[0],
	'xmlns:r': XMLNS.r
});

var write_ws_xml = function(idx, opts, wb) {
	var o = [], s = wb.SheetNames[idx], ws = wb.Sheets[s] || {}, sidx = 0, rdata = "";
	o.push(XML_HEADER);
	o.push(WS_XML_ROOT);
	o.push(writextag('dimension', null, {'ref': ws['!ref'] || 'A1'}));
	if((ws['!cols']||[]).length > 0) o.push(write_ws_xml_cols(ws, ws['!cols']));
	sidx = o.length;
	o.push(writextag('sheetData', null));
	if(ws['!ref']) rdata = write_ws_xml_data(ws, opts, idx, wb);
	if(rdata.length) o.push(rdata);
	if(o.length>sidx+1){ o.push('</sheetData>'); o[sidx]=o[sidx].replace("/>",">"); }

	if(o.length>2){ o.push('</worksheet>'); o[1]=o[1].replace("/>",">"); }
	return o.join("");
};
