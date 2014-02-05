/* 18.3 Worksheets */
function parse_worksheet(data) {
	if(!data) return data;
	/* 18.3.1.99 worksheet CT_Worksheet */
	var s = {};

	/* 18.3.1.35 dimension CT_SheetDimension ? */
	var ref = data.match(/<dimension ref="([^"]*)"\s*\/>/);
	if(ref && ref.length == 2 && ref[1].indexOf(":") !== -1) s["!ref"] = ref[1];

	var refguess = {s: {r:1000000, c:1000000}, e: {r:0, c:0} };
	var q = ["v","f"];
	var sidx = 0;
	/* 18.3.1.80 sheetData CT_SheetData ? */
	if(!data.match(/<sheetData *\/>/))
	data.match(/<sheetData>([^\u2603]*)<\/sheetData>/m)[1].split("</row>").forEach(function(x) {
		if(x === "" || x.trim() === "") return;

		/* 18.3.1.73 row CT_Row */
		var row = parsexmltag(x.match(/<row[^>]*>/)[0]);
		if(refguess.s.r > row.r - 1) refguess.s.r = row.r - 1;
		if(refguess.e.r < row.r - 1) refguess.e.r = row.r - 1;

		/* 18.3.1.4 c CT_Cell */
		var cells = x.substr(x.indexOf('>')+1).split(/<c /);
		cells.forEach(function(c, idx) { if(c === "" || c.trim() === "") return;
			var cref = c.match(/r=["']([^"']*)["']/);
			c = "<c " + c;
			if(cref && cref.length == 2) {
				var cref_cell = decode_cell(cref[1]);
				idx = cref_cell.c;
			}
			if(refguess.s.c > idx) refguess.s.c = idx;
			if(refguess.e.c < idx) refguess.e.c = idx;
			var cell = parsexmltag((c.match(/<c[^>]*>/)||[c])[0]); delete cell[0];
			var d = c.substr(c.indexOf('>')+1);
			var p = {};
			q.forEach(function(f){var x=d.match(matchtag(f));if(x)p[f]=unescapexml(x[1]);});

			/* SCHEMA IS ACTUALLY INCORRECT HERE.  IF A CELL HAS NO T, EMIT "" */
			if(cell.t === undefined && p.v === undefined) { p.t = "str"; p.v = undefined; }
			else p.t = (cell.t ? cell.t : "n"); // default is "n" in schema
			switch(p.t) {
				case 'n': p.v = parseFloat(p.v); break;
				case 's': {
					sidx = parseInt(p.v, 10);
					p.v = strs[sidx].t;
					p.r = strs[sidx].r;
					p.h = strs[sidx].h;
				} break;
				case 'str': if(p.v) p.v = utf8read(p.v); break;
				case 'inlineStr':
					var is = d.match(/<is>(.*)<\/is>/);
					is = is ? parse_si(is[1]) : {t:"",r:""};
					p.t = 'str'; p.v = is.t;
					break; // inline string
				case 'b':
					switch(p.v) {
						case '0': case 'FALSE': case "false": case false: p.v=false; break;
						case '1': case 'TRUE':  case "true":  case true:  p.v=true;  break;
						default: throw "Unrecognized boolean: " + p.v;
					} break;
				/* in case of error, stick value in .raw */
				case 'e': p.raw = RBErr[p.v]; break;
				default: throw "Unrecognized cell type: " + p.t;
			}

			/* formatting */
			var fmtid = 0;
			if(cell.s && styles.CellXf) {
				var cf = styles.CellXf[cell.s];
				if(cf && cf.numFmtId) fmtid = cf.numFmtId;
			}
			try { p.w = SSF.format(fmtid,p.v,_ssfopts); } catch(e) { }
			s[cell.r] = p;
		});
	});
	if(!s["!ref"]) s["!ref"] = encode_range(refguess);
	return s;
}

