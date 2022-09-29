/* 18.4.7 rPr CT_RPrElt */
function parse_rpr(rpr) {
	var font = {}, m = rpr.match(tagregex), i = 0;
	var pass = false;
	if(m) for(;i!=m.length; ++i) {
		var y = parsexmltag(m[i]);
		switch(y[0].replace(/\w*:/g,"")) {
			/* 18.8.12 condense CT_BooleanProperty */
			/* ** not required . */
			case '<condense': break;
			/* 18.8.17 extend CT_BooleanProperty */
			/* ** not required . */
			case '<extend': break;
			/* 18.8.36 shadow CT_BooleanProperty */
			/* ** not required . */
			case '<shadow':
				if(!y.val) break;
				/* falls through */
			case '<shadow>':
			case '<shadow/>': font.shadow = 1; break;
			case '</shadow>': break;

			/* 18.4.1 charset CT_IntProperty TODO */
			case '<charset':
				if(y.val == '1') break;
				font.cp = CS2CP[parseInt(y.val, 10)];
				break;

			/* 18.4.2 outline CT_BooleanProperty TODO */
			case '<outline':
				if(!y.val) break;
				/* falls through */
			case '<outline>':
			case '<outline/>': font.outline = 1; break;
			case '</outline>': break;

			/* 18.4.5 rFont CT_FontName */
			case '<rFont': font.name = y.val; break;

			/* 18.4.11 sz CT_FontSize */
			case '<sz': font.sz = y.val; break;

			/* 18.4.10 strike CT_BooleanProperty */
			case '<strike':
				if(!y.val) break;
				/* falls through */
			case '<strike>':
			case '<strike/>': font.strike = 1; break;
			case '</strike>': break;

			/* 18.4.13 u CT_UnderlineProperty */
			case '<u':
				if(!y.val) break;
				switch(y.val) {
					case 'double': font.uval = "double"; break;
					case 'singleAccounting': font.uval = "single-accounting"; break;
					case 'doubleAccounting': font.uval = "double-accounting"; break;
				}
				/* falls through */
			case '<u>':
			case '<u/>': font.u = 1; break;
			case '</u>': break;

			/* 18.8.2 b */
			case '<b':
				if(y.val == '0') break;
				/* falls through */
			case '<b>':
			case '<b/>': font.b = 1; break;
			case '</b>': break;

			/* 18.8.26 i */
			case '<i':
				if(y.val == '0') break;
				/* falls through */
			case '<i>':
			case '<i/>': font.i = 1; break;
			case '</i>': break;

			/* 18.3.1.15 color CT_Color TODO: tint, theme, auto, indexed */
			case '<color':
				if(y.rgb) font.color = y.rgb.slice(2,8);
				break;
			case '<color>': case '<color/>': case '</color>': break;

			/* 18.8.18 family ST_FontFamily */
			case '<family': font.family = y.val; break;
			case '<family>': case '<family/>': case '</family>': break;

			/* 18.4.14 vertAlign CT_VerticalAlignFontProperty TODO */
			case '<vertAlign': font.valign = y.val; break;
			case '<vertAlign>': case '<vertAlign/>': case '</vertAlign>': break;

			/* 18.8.35 scheme CT_FontScheme TODO */
			case '<scheme': break;
			case '<scheme>': case '<scheme/>': case '</scheme>': break;

			/* 18.2.10 extLst CT_ExtensionList ? */
			case '<extLst': case '<extLst>': case '</extLst>': break;
			case '<ext': pass = true; break;
			case '</ext>': pass = false; break;
			default:
				if(y[0].charCodeAt(1) !== 47 && !pass) throw new Error('Unrecognized rich format ' + y[0]);
		}
	}
	return font;
}

var parse_rs = /*#__PURE__*/(function() {
	var tregex = matchtag("t"), rpregex = matchtag("rPr");
	/* 18.4.4 r CT_RElt */
	function parse_r(r) {
		/* 18.4.12 t ST_Xstring */
		var t = r.match(tregex)/*, cp = 65001*/;
		if(!t) return {t:"s", v:""};

		var o/*:Cell*/ = ({t:'s', v:unescapexml(t[1])}/*:any*/);
		var rpr = r.match(rpregex);
		if(rpr) o.s = parse_rpr(rpr[1]);
		return o;
	}
	var rregex = /<(?:\w+:)?r>/g, rend = /<\/(?:\w+:)?r>/;
	return function parse_rs(rs) {
		return rs.replace(rregex,"").split(rend).map(parse_r).filter(function(r) { return r.v; });
	};
})();


/* Parse a list of <r> tags */
var rs_to_html = /*#__PURE__*/(function parse_rs_factory() {
	var nlregex = /(\r\n|\n)/g;
	function parse_rpr2(font, intro, outro) {
		var style/*:Array<string>*/ = [];

		if(font.u) style.push("text-decoration: underline;");
		if(font.uval) style.push("text-underline-style:" + font.uval + ";");
		if(font.sz) style.push("font-size:" + font.sz + "pt;");
		if(font.outline) style.push("text-effect: outline;");
		if(font.shadow) style.push("text-shadow: auto;");
		intro.push('<span style="' + style.join("") + '">');

		if(font.b) { intro.push("<b>"); outro.push("</b>"); }
		if(font.i) { intro.push("<i>"); outro.push("</i>"); }
		if(font.strike) { intro.push("<s>"); outro.push("</s>"); }

		var align = font.valign || "";
		if(align == "superscript" || align == "super") align = "sup";
		else if(align == "subscript") align = "sub";
		if(align != "") { intro.push("<" + align + ">"); outro.push("</" + align + ">"); }

		outro.push("</span>");
		return font;
	}

	/* 18.4.4 r CT_RElt */
	function r_to_html(r) {
		var terms/*:[Array<string>, string, Array<string>]*/ = [[],r.v,[]];
		if(!r.v) return "";

		if(r.s) parse_rpr2(r.s, terms[0], terms[2]);

		return terms[0].join("") + terms[1].replace(nlregex,'<br/>') + terms[2].join("");
	}

	return function parse_rs(rs) {
		return rs.map(r_to_html).join("");
	};
})();

/* 18.4.8 si CT_Rst */
var sitregex = /<(?:\w+:)?t[^>]*>([^<]*)<\/(?:\w+:)?t>/g, sirregex = /<(?:\w+:)?r>/;
var sirphregex = /<(?:\w+:)?rPh.*?>([\s\S]*?)<\/(?:\w+:)?rPh>/g;
function parse_si(x, opts) {
	var html = opts ? opts.cellHTML : true;
	var z = {};
	if(!x) return { t: "" };
	//var y;
	/* 18.4.12 t ST_Xstring (Plaintext String) */
	// TODO: is whitespace actually valid here?
	if(x.match(/^\s*<(?:\w+:)?t[^>]*>/)) {
		z.t = unescapexml(utf8read(x.slice(x.indexOf(">")+1).split(/<\/(?:\w+:)?t>/)[0]||""), true);
		z.r = utf8read(x);
		if(html) z.h = escapehtml(z.t);
	}
	/* 18.4.4 r CT_RElt (Rich Text Run) */
	else if((/*y = */x.match(sirregex))) {
		z.r = utf8read(x);
		z.t = unescapexml(utf8read((x.replace(sirphregex, '').match(sitregex)||[]).join("").replace(tagregex,"")), true);
		if(html) z.h = rs_to_html(parse_rs(z.r));
	}
	/* 18.4.3 phoneticPr CT_PhoneticPr (TODO: needed for Asian support) */
	/* 18.4.6 rPh CT_PhoneticRun (TODO: needed for Asian support) */
	return z;
}

/* 18.4 Shared String Table */
var sstr0 = /<(?:\w+:)?sst([^>]*)>([\s\S]*)<\/(?:\w+:)?sst>/;
var sstr1 = /<(?:\w+:)?(?:si|sstItem)>/g;
var sstr2 = /<\/(?:\w+:)?(?:si|sstItem)>/;
function parse_sst_xml(data/*:string*/, opts)/*:SST*/ {
	var s/*:SST*/ = ([]/*:any*/), ss = "";
	if(!data) return s;
	/* 18.4.9 sst CT_Sst */
	var sst = data.match(sstr0);
	if(sst) {
		ss = sst[2].replace(sstr1,"").split(sstr2);
		for(var i = 0; i != ss.length; ++i) {
			var o = parse_si(ss[i].trim(), opts);
			if(o != null) s[s.length] = o;
		}
		sst = parsexmltag(sst[1]); s.Count = sst.count; s.Unique = sst.uniqueCount;
	}
	return s;
}

var straywsregex = /^\s|\s$|[\t\n\r]/;
function write_sst_xml(sst/*:SST*/, opts)/*:string*/ {
	if(!opts.bookSST) return "";
	var o = [XML_HEADER];
	o[o.length] = (writextag('sst', null, {
		xmlns: XMLNS_main[0],
		count: sst.Count,
		uniqueCount: sst.Unique
	}));
	for(var i = 0; i != sst.length; ++i) { if(sst[i] == null) continue;
		var s/*:XLString*/ = sst[i];
		var sitag = "<si>";
		if(s.r) sitag += s.r;
		else {
			sitag += "<t";
			if(!s.t) s.t = "";
			if(typeof s.t !== "string") s.t = String(s.t);
			if(s.t.match(straywsregex)) sitag += ' xml:space="preserve"';
			sitag += ">" + escapexml(s.t) + "</t>";
		}
		sitag += "</si>";
		o[o.length] = (sitag);
	}
	if(o.length>2){ o[o.length] = ('</sst>'); o[1]=o[1].replace("/>",">"); }
	return o.join("");
}
