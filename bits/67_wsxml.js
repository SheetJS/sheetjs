function parse_ws_xml_dim(ws, s) {
	var d = safe_decode_range(s);
	if(d.s.r<=d.e.r && d.s.c<=d.e.c && d.s.r>=0 && d.s.c>=0) ws["!ref"] = encode_range(d);
}
var mergecregex = /<(?:\w:)?mergeCell ref="[A-Z0-9:]+"\s*\/>/g;
var sheetdataregex = /<(?:\w+:)?sheetData>([^\u2603]*)<\/(?:\w+:)?sheetData>/;
var hlinkregex = /<(?:\w*:)?hyperlink[^>]*\/>/g;
var dimregex = /"(\w*:\w*)"/;
var colregex = /<(?:\w*:)?col[^>]*\/>/g;
/* 18.3 Worksheets */
function parse_ws_xml(data/*:?string*/, opts, rels)/*:Worksheet*/ {
	if(!data) return data;
	/* 18.3.1.99 worksheet CT_Worksheet */
	var s = ({}/*:any*/);

	/* 18.3.1.35 dimension CT_SheetDimension ? */
	// $FlowIgnore
	var ridx = (data.match(/<(?:\w*:)?dimension/)||{index:-1}).index;
	if(ridx > 0) {
		var ref = data.substr(ridx,50).match(dimregex);
		if(ref != null) parse_ws_xml_dim(s, ref[1]);
	}

	/* 18.3.1.55 mergeCells CT_MergeCells */
	var mergecells = [];
	var merges = data.match(mergecregex);
	if(merges) for(ridx = 0; ridx != merges.length; ++ridx)
		mergecells[ridx] = safe_decode_range(merges[ridx].substr(merges[ridx].indexOf("\"")+1));

	/* 18.3.1.17 cols CT_Cols */
	var columns = [];
	if(opts.cellStyles) {
		/* 18.3.1.13 col CT_Col */
		var cols = data.match(colregex);
		if(cols) parse_ws_xml_cols(columns, cols);
	}

	var refguess/*:Range*/ = ({s: {r:2000000, c:2000000}, e: {r:0, c:0} }/*:any*/);

	/* 18.3.1.80 sheetData CT_SheetData ? */
	var mtch=data.match(sheetdataregex);
	if(mtch) parse_ws_xml_data(mtch[1], s, opts, refguess);

	/* 18.3.1.48 hyperlinks CT_Hyperlinks */
	var hlink = data.match(hlinkregex);
	if(hlink) parse_ws_xml_hlinks(s, hlink, rels);

	if(!s["!ref"] && refguess.e.c >= refguess.s.c && refguess.e.r >= refguess.s.r) s["!ref"] = encode_range(refguess);
	if(opts.sheetRows > 0 && s["!ref"]) {
		var tmpref = safe_decode_range(s["!ref"]);
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

function write_ws_xml_merges(merges) {
	if(merges.length == 0) return "";
	var o = '<mergeCells count="' + merges.length + '">';
	for(var i = 0; i != merges.length; ++i) o += '<mergeCell ref="' + encode_range(merges[i]) + '"/>';
	return o + '</mergeCells>';
}

function parse_ws_xml_hlinks(s, data, rels) {
	for(var i = 0; i != data.length; ++i) {
		var val = parsexmltag(data[i], true);
		if(!val.ref) return;
		var rel = rels ? rels['!id'][val.id] : null;
		if(rel) {
			val.Target = rel.Target;
			if(val.location) val.Target += "#"+val.location;
			val.Rel = rel;
		} else {
			val.Target = val.location;
			rel = {Target: val.location, TargetMode: 'Internal'};
			val.Rel = rel;
		}
		var rng = safe_decode_range(val.ref);
		for(var R=rng.s.r;R<=rng.e.r;++R) for(var C=rng.s.c;C<=rng.e.c;++C) {
			var addr = encode_cell({c:C,r:R});
			if(!s[addr]) s[addr] = {t:"z",v:undefined};
			s[addr].l = val;
		}
	}
}

function parse_ws_xml_cols(columns, cols) {
	var seencol = false;
	for(var coli = 0; coli != cols.length; ++coli) {
		var coll = parsexmltag(cols[coli], true);
		var colm=parseInt(coll.min, 10)-1, colM=parseInt(coll.max,10)-1;
		delete coll.min; delete coll.max;
		if(!seencol && coll.width) { seencol = true; find_mdw(+coll.width, coll); }
		if(coll.width) {
			coll.wpx = width2px(+coll.width);
			coll.wch = px2char(coll.wpx);
			coll.MDW = MDW;
		}
		while(colm <= colM) columns[colm++] = coll;
	}
}

function write_ws_xml_cols(ws, cols)/*:string*/ {
	var o = ["<cols>"], col, width;
	for(var i = 0; i != cols.length; ++i) {
		if(!(col = cols[i])) continue;
		var p = ({min:i+1,max:i+1}/*:any*/);
		/* wch (chars), wpx (pixels) */
		width = -1;
		if(col.wpx) width = px2char(col.wpx);
		else if(col.wch) width = col.wch;
		if(width > -1) { p.width = char2width(width); p.customWidth= 1; }
		o[o.length] = (writextag('col', null, p));
	}
	o[o.length] = "</cols>";
	return o.join("");
}

function write_ws_xml_cell(cell, ref, ws, opts, idx, wb) {
	if(cell.v === undefined && cell.f === undefined || cell.t === 'z') return "";
	var vv = "";
	var oldt = cell.t, oldv = cell.v;
	switch(cell.t) {
		case 'b': vv = cell.v ? "1" : "0"; break;
		case 'n': vv = ''+cell.v; break;
		case 'e': vv = BErr[cell.v]; break;
		case 'd':
			if(opts.cellDates) vv = new Date(cell.v).toISOString();
			else {
				cell.t = 'n';
				vv = ''+(cell.v = datenum(new Date(cell.v)));
				if(typeof cell.z === 'undefined') cell.z = SSF._table[14];
			}
			break;
		default: vv = cell.v; break;
	}
	var v = writetag('v', escapexml(vv)), o = ({r:ref}/*:any*/);
	/* TODO: cell style */
	var os = get_cell_style(opts.cellXfs, cell, opts);
	if(os !== 0) o.s = os;
	switch(cell.t) {
		case 'n': break;
		case 'd': o.t = "d"; break;
		case 'b': o.t = "b"; break;
		case 'e': o.t = "e"; break;
		default: if(cell.v == null) { delete cell.t; break; }
			if(opts.bookSST) {
				v = writetag('v', ''+get_sst_id(opts.Strings, cell.v));
				o.t = "s"; break;
			}
			o.t = "str"; break;
	}
	if(cell.t != oldt) { cell.t = oldt; cell.v = oldv; }
	if(cell.f) {
		var ff = cell.F && cell.F.substr(0, ref.length) == ref ? {t:"array", ref:cell.F} : null;
		v = writextag('f', escapexml(cell.f), ff) + (cell.v != null ? v : "");
	}
	return writextag('c', v, o);
}

var parse_ws_xml_data = (function parse_ws_xml_data_factory() {
	var cellregex = /<(?:\w+:)?c[ >]/, rowregex = /<\/(?:\w+:)?row>/;
	var rregex = /r=["']([^"']*)["']/, isregex = /<(?:\w+:)?is>([\S\s]*?)<\/(?:\w+:)?is>/;
	var refregex = /ref=["']([^"']*)["']/;
	var match_v = matchtag("v"), match_f = matchtag("f");

return function parse_ws_xml_data(sdata, s, opts, guess) {
	var ri = 0, x = "", cells = [], cref = [], idx = 0, i=0, cc=0, d="", p/*:any*/;
	var tag, tagr = 0, tagc = 0;
	var sstr, ftag;
	var fmtid = 0, fillid = 0, do_format = Array.isArray(styles.CellXf), cf;
	var arrayf = [];
	var sharedf = [];
	for(var marr = sdata.split(rowregex), mt = 0, marrlen = marr.length; mt != marrlen; ++mt) {
		x = marr[mt].trim();
		var xlen = x.length;
		if(xlen === 0) continue;

		/* 18.3.1.73 row CT_Row */
		for(ri = 0; ri < xlen; ++ri) if(x.charCodeAt(ri) === 62) break; ++ri;
		tag = parsexmltag(x.substr(0,ri), true);
		/* SpreadSheetGear uses implicit r/c */
		tagr = typeof tag.r !== 'undefined' ? parseInt(tag.r, 10) : tagr+1; tagc = -1;
		if(opts.sheetRows && opts.sheetRows < tagr) continue;
		if(guess.s.r > tagr - 1) guess.s.r = tagr - 1;
		if(guess.e.r < tagr - 1) guess.e.r = tagr - 1;

		/* 18.3.1.4 c CT_Cell */
		cells = x.substr(ri).split(cellregex);
		for(ri = 0; ri != cells.length; ++ri) {
			x = cells[ri].trim();
			if(x.length === 0) continue;
			cref = x.match(rregex); idx = ri; i=0; cc=0;
			x = "<c " + (x.substr(0,1)=="<"?">":"") + x;
			if(cref != null && cref.length === 2) {
				idx = 0; d=cref[1];
				for(i=0; i != d.length; ++i) {
					if((cc=d.charCodeAt(i)-64) < 1 || cc > 26) break;
					idx = 26*idx + cc;
				}
				--idx;
				tagc = idx;
			} else ++tagc;
			for(i = 0; i != x.length; ++i) if(x.charCodeAt(i) === 62) break; ++i;
			tag = parsexmltag(x.substr(0,i), true);
			if(!tag.r) tag.r = utils.encode_cell({r:tagr-1, c:tagc});
			d = x.substr(i);
			p = ({t:""}/*:any*/);

			if((cref=d.match(match_v))!= null && /*::cref != null && */cref[1] !== '') p.v=unescapexml(cref[1]);
			if(opts.cellFormula) {
				if((cref=d.match(match_f))!= null && /*::cref != null && */cref[1] !== '') {
					/* TODO: match against XLSXFutureFunctions */
					p.f=unescapexml(utf8read(cref[1])).replace(/_xlfn\./,"");
					if(/*::cref != null && cref[0] != null && */cref[0].indexOf('t="array"') > -1) {
						p.F = (d.match(refregex)||[])[1];
						if(p.F.indexOf(":") > -1) arrayf.push([safe_decode_range(p.F), p.F]);
					} else if(/*::cref != null && cref[0] != null && */cref[0].indexOf('t="shared"') > -1) {
						// TODO: parse formula
						ftag = parsexmltag(cref[0]);
						sharedf[parseInt(ftag.si, 10)] = [ftag, unescapexml(utf8read(cref[1]))];
					}
				} else if((cref=d.match(/<f[^>]*\/>/))) {
					ftag = parsexmltag(cref[0]);
					if(sharedf[ftag.si]) p.f = shift_formula_xlsx(sharedf[ftag.si][1], sharedf[ftag.si][0].ref, tag.r);
				}
				/* TODO: factor out contains logic */
				var _tag = decode_cell(tag.r);
				for(i = 0; i < arrayf.length; ++i)
					if(_tag.r >= arrayf[i][0].s.r && _tag.r <= arrayf[i][0].e.r)
						if(_tag.c >= arrayf[i][0].s.c && _tag.c <= arrayf[i][0].e.c)
							p.F = arrayf[i][1];
			}

			/* SCHEMA IS ACTUALLY INCORRECT HERE.  IF A CELL HAS NO T, EMIT "" */
			if(tag.t === undefined && p.v === undefined) {
				if(!opts.sheetStubs) continue;
				p.t = "z";
			}
			else p.t = tag.t || "n";
			if(guess.s.c > idx) guess.s.c = idx;
			if(guess.e.c < idx) guess.e.c = idx;
			/* 18.18.11 t ST_CellType */
			switch(p.t) {
				case 'n': p.v = parseFloat(p.v); break;
				case 's':
					sstr = strs[parseInt(p.v, 10)];
					if(typeof p.v == 'undefined') {
						if(!opts.sheetStubs) continue;
						p.t = 'z';
					}
					p.v = sstr.t;
					p.r = sstr.r;
					if(opts.cellHTML) p.h = sstr.h;
					break;
				case 'str':
					p.t = "s";
					p.v = (p.v!=null) ? utf8read(p.v) : '';
					if(opts.cellHTML) p.h = p.v;
					break;
				case 'inlineStr':
					cref = d.match(isregex);
					p.t = 's';
					if(cref != null && (sstr = parse_si(cref[1]))) p.v = sstr.t; else p.v = "";
					break; // inline string
				case 'b': p.v = parsexmlbool(p.v); break;
				case 'd':
					if(!opts.cellDates) { p.v = datenum(new Date(p.v)); p.t = 'n'; }
					break;
				/* error string in .v, number in .v */
				case 'e': p.w = p.v; p.v = RBErr[p.v]; break;
			}
			/* formatting */
			fmtid = fillid = 0;
			if(do_format && tag.s !== undefined) {
				cf = styles.CellXf[tag.s];
				if(cf != null) {
					if(cf.numFmtId != null) fmtid = cf.numFmtId;
					if(opts.cellStyles && cf.fillId != null) fillid = cf.fillId;
				}
			}
			safe_format(p, fmtid, fillid, opts);
			s[tag.r] = p;
		}
	}
}; })();

function write_ws_xml_data(ws/*:Worksheet*/, opts, idx/*:number*/, wb/*:Workbook*/)/*:string*/ {
	var o = [], r = [], range = safe_decode_range(ws['!ref']), cell, ref, rr = "", cols = [], R=0, C=0;
	for(C = range.s.c; C <= range.e.c; ++C) cols[C] = encode_col(C);
	for(R = range.s.r; R <= range.e.r; ++R) {
		r = [];
		rr = encode_row(R);
		for(C = range.s.c; C <= range.e.c; ++C) {
			ref = cols[C] + rr;
			if(ws[ref] === undefined) continue;
			if((cell = write_ws_xml_cell(ws[ref], ref, ws, opts, idx, wb)) != null) r.push(cell);
		}
		if(r.length > 0) o[o.length] = (writextag('row', r.join(""), {r:rr}));
	}
	return o.join("");
}

var WS_XML_ROOT = writextag('worksheet', null, {
	'xmlns': XMLNS.main[0],
	'xmlns:r': XMLNS.r
});

function write_ws_xml(idx/*:number*/, opts, wb/*:Workbook*/)/*:string*/ {
	var o = [XML_HEADER, WS_XML_ROOT];
	var s = wb.SheetNames[idx], sidx = 0, rdata = "";
	var ws = wb.Sheets[s];
	if(ws === undefined) ws = {};
	var ref = ws['!ref']; if(ref === undefined) ref = 'A1';
	o[o.length] = (writextag('dimension', null, {'ref': ref}));

	if(ws['!cols'] !== undefined && ws['!cols'].length > 0) o[o.length] = (write_ws_xml_cols(ws, ws['!cols']));
	o[sidx = o.length] = '<sheetData/>';
	if(ws['!ref'] !== undefined) {
		rdata = write_ws_xml_data(ws, opts, idx, wb);
		if(rdata.length > 0) o[o.length] = (rdata);
	}
	if(o.length>sidx+1) { o[o.length] = ('</sheetData>'); o[sidx]=o[sidx].replace("/>",">"); }

	if(ws['!merges'] !== undefined && ws['!merges'].length > 0) o[o.length] = (write_ws_xml_merges(ws['!merges']));

	if(o.length>2) { o[o.length] = ('</worksheet>'); o[1]=o[1].replace("/>",">"); }
	return o.join("");
}
