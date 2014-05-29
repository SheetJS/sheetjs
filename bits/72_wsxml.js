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

	var refguess = {s: {r:1000000, c:1000000}, e: {r:0, c:0} };
	var q = (opts.cellFormula ? ["v","f"] : ["v"]);
	var sidx = 0;

	/* 18.3.1.80 sheetData CT_SheetData ? */
	if((mtch=data.match(/<(?:\w+:)?sheetData>([^\u2603]*)<\/(?:\w+:)?sheetData>/m))) mtch[1].split(/<\/(?:\w+:)?row>/).forEach(function(x) {
		if(x === "" || x.trim() === "") return;

		/* 18.3.1.73 row CT_Row */
		var row = parsexmltag(x.match(/<(?:\w+:)?row[^>]*>/)[0]);
		if(opts.sheetRows && opts.sheetRows < +row.r) return;
		if(refguess.s.r > row.r - 1) refguess.s.r = row.r - 1;
		if(refguess.e.r < row.r - 1) refguess.e.r = row.r - 1;
		/* 18.3.1.4 c CT_Cell */
		var cells = x.substr(x.indexOf('>')+1).split(/<(?:\w+:)?c /);
		cells.forEach(function(c, idx) { if(c === "" || c.trim() === "") return;
			var cref = c.match(/r=["']([^"']*)["']/);
			c = "<c " + c;
			if(cref && cref.length == 2) idx = decode_cell(cref[1]).c;
			var cell = parsexmltag((c.match(/<c[^>]*>/)||[c])[0]); delete cell[0];
			var d = c.substr(c.indexOf('>')+1);
			var p = {};
			q.forEach(function(f){var x=d.match(matchtag(f));if(x)p[f]=unescapexml(x[1]);});
			/* SCHEMA IS ACTUALLY INCORRECT HERE.  IF A CELL HAS NO T, EMIT "" */
			if(cell.t === undefined && p.v === undefined) {
				if(!opts.sheetStubs) return;
				p.t = "str"; p.v = undefined;
			}
			else p.t = (cell.t ? cell.t : "n"); // default is "n" in schema
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
				case 'b': if(typeof p.v !== 'boolean') p.v = parsexmlbool(p.v); break;
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
			try {
				p.w = SSF.format(fmtid,p.v,_ssfopts);
				if(opts.cellNF) p.z = SSF._table[fmtid];
				if(fillid) {
					p.s = styles.Fills[fillid];
					if (p.s.fgColor && p.s.fgColor.theme) {
						p.s.fgColor.rgb = rgb_tint(themes.themeElements.clrScheme[p.s.fgColor.theme].rgb, p.s.fgColor.tint || 0);
					}
					if (p.s.bgColor && p.s.bgColor.theme) {
						p.s.bgColor.rgb = rgb_tint(themes.themeElements.clrScheme[p.s.bgColor.theme].rgb, p.s.bgColor.tint || 0);
					}
				}
			} catch(e) { if(opts.WTF) throw e; }
			s[cell.r] = p;
		});
	});

	/* 18.3.1.48 hyperlinks CT_Hyperlinks */
	if(data.match(/<\/hyperlinks>/)) data.match(/<hyperlink[^>]*\/>/g).forEach(function(h) {
		var val = parsexmltag(h); delete val[0];
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
	return s;
}

var WS_XML_ROOT = writextag('worksheet', null, {
	'xmlns': XMLNS.main[0],
	'xmlns:r': XMLNS.r
});

var write_ws_xml_cell = function(cell, ref, ws, opts, idx, wb) {
	var vv = cell.v; if(cell.t == 'b') vv = cell.v ? "1" : "0";
	var v = writextag('v', escapexml(String(vv))), o = {r:ref};
	o.s = get_cell_style(opts.cellXfs, cell, opts);
	if(o.s === 0 || o.s === "0") delete o.s;
	/* TODO: cell style */
	if(typeof cell.v === 'undefined') return "";
	switch(cell.t) {
		case 's': case 'str': {
			if(opts.bookSST) {
				v = writextag('v', String(get_sst_id(opts.Strings, cell.v)));
				o.t = "s"; return writextag('c', v, o);
			} else { o.t = "str"; return writextag('c', v, o); }
		} break;
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

var write_ws_xml = function(idx, opts, wb) {
	var o = [], s = wb.SheetNames[idx], ws = wb.Sheets[s] || {}, sidx = 0, rdata = "";
	o.push(XML_HEADER);
	o.push(WS_XML_ROOT);
	o.push(writextag('dimension', null, {'ref': ws['!ref'] || 'A1'}));

	sidx = o.length;
	o.push(writextag('sheetData', null));
	if(ws['!ref']) rdata = write_ws_xml_data(ws, opts, idx, wb);
	if(rdata.length) o.push(rdata);
	if(o.length>sidx+1){ o.push('</sheetData>'); o[sidx]=o[sidx].replace("/>",">"); }

	if(o.length>2){ o.push('</worksheet>'); o[1]=o[1].replace("/>",">"); }
	return o.join("");
};
