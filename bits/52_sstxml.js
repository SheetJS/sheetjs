/* Parse a list of <r> tags */
var parse_rs = (function() {
	var tregex = matchtag("t"), rpregex = matchtag("rPr");
	/* 18.4.7 rPr CT_RPrElt */
	var parse_rpr = function(rpr, intro, outro) {
		var font = {};
		(rpr.match(/<[^>]*>/g)||[]).forEach(function(x) {
			var y = parsexmltag(x);
			switch(y[0]) {
				/* 18.8.12 condense CT_BooleanProperty */
				/* ** not required . */
				case '<condense': break;
				/* 18.8.17 extend CT_BooleanProperty */
				/* ** not required . */
				case '<extend': break;
				/* 18.8.36 shadow CT_BooleanProperty */
				/* ** not required . */
				case '<shadow': break;

				/* 18.4.1 charset CT_IntProperty TODO */
				case '<charset': break;

				/* 18.4.2 outline CT_BooleanProperty TODO */
				case '<outline': break;

				/* 18.4.5 rFont CT_FontName */
				case '<rFont': font.name = y.val; break;

				/* 18.4.11 sz CT_FontSize */
				case '<sz': font.sz = y.val; break;

				/* 18.4.10 strike CT_BooleanProperty */
				case '<strike':
					if(!y.val) break;
					/* falls through */
				case '<strike/>': font.strike = 1; break;
				case '</strike>': break;

				/* 18.4.13 u CT_UnderlineProperty */
				case '<u':
					if(!y.val) break;
					/* falls through */
				case '<u/>': font.u = 1; break;
				case '</u>': break;

				/* 18.8.2 b */
				case '<b':
					if(!y.val) break;
					/* falls through */
				case '<b/>': font.b = 1; break;
				case '</b>': break;

				/* 18.8.26 i */
				case '<i':
					if(!y.val) break;
					/* falls through */
				case '<i/>': font.i = 1; break;
				case '</i>': break;

				/* 18.3.1.15 color CT_Color TODO: tint, theme, auto, indexed */
				case '<color':
					if(y.rgb) font.color = y.rgb.substr(2,6);
					break;

				/* 18.8.18 family ST_FontFamily */
				case '<family': font.family = y.val; break;

				/* 18.4.14 vertAlign CT_VerticalAlignFontProperty TODO */
				case '<vertAlign': break;

				/* 18.8.35 scheme CT_FontScheme TODO */
				case '<scheme': break;

				default:
					if(y[0][2] !== '/') throw 'Unrecognized rich format ' + y[0];
			}
		});
		/* TODO: These should be generated styles, not inline */
		var style = [];
		if(font.b) style.push("font-weight: bold;");
		if(font.i) style.push("font-style: italic;");
		intro.push('<span style="' + style.join("") + '">');
		outro.push("</span>");
	};

	/* 18.4.4 r CT_RElt */
	function parse_r(r) {
		var terms = [[],"",[]];
		/* 18.4.12 t ST_Xstring */
		var t = r.match(tregex);
		if(!isval(t)) return "";
		terms[1] = t[1];

		var rpr = r.match(rpregex);
		if(isval(rpr)) parse_rpr(rpr[1], terms[0], terms[2]);
		return terms[0].join("") + terms[1].replace(/\r\n/g,'<br/>') + terms[2].join("");
	}
	return function(rs) {
		return rs.replace(/<r>/g,"").split(/<\/r>/).map(parse_r).join("");
	};
})();

/* 18.4.8 si CT_Rst */
var parse_si = function(x) {
	var z = {};
	if(!x) return z;
	var y;
	/* 18.4.12 t ST_Xstring (Plaintext String) */
	if(x[1] === 't') {
		z.t = utf8read(unescapexml(x.replace(/<[^>]*>/g,"")));
		z.raw = x;
		z.r = z.t;
	}
	/* 18.4.4 r CT_RElt (Rich Text Run) */
	else if((y = x.match(/<r>/))) {
		z.raw = x;
		/* TODO: properly parse (note: no other valid child can have body text) */
		z.t = utf8read(unescapexml(x.replace(/<[^>]*>/gm,"")));
		z.r = parse_rs(x);
	}
	/* 18.4.3 phoneticPr CT_PhoneticPr (TODO: needed for Asian support) */
	/* 18.4.6 rPh CT_PhoneticRun (TODO: needed for Asian support) */
	return z;
};

/* 18.4 Shared String Table */
var parse_sst_xml = function(data) {
	var s = [];
	/* 18.4.9 sst CT_Sst */
	var sst = data.match(new RegExp("<sst([^>]*)>([\\s\\S]*)<\/sst>","m"));
	if(isval(sst)) {
		s = sst[2].replace(/<si>/g,"").split(/<\/si>/).map(parse_si);
		sst = parsexmltag(sst[1]); s.Count = sst.count; s.Unique = sst.uniqueCount;
	}
	return s;
};

