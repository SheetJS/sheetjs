RELS.THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

/* 20.1.6.2 clrScheme CT_ColorScheme */
function parse_clrScheme(t, opts) {
	themes.themeElements.clrScheme = [];
	var color = {};
	(t[0].match(tagregex)||[]).forEach(function(x) {
		var y = parsexmltag(x);
		switch(y[0]) {
			case '<a:clrScheme': case '</a:clrScheme>': break;

			/* 20.1.2.3.32 srgbClr CT_SRgbColor */
			case '<a:srgbClr': color.rgb = y.val; break;

			/* 20.1.2.3.33 sysClr CT_SystemColor */
			case '<a:sysClr': color.rgb = y.lastClr; break;

			/* 20.1.4.1.9 dk1 (Dark 1) */
			case '<a:dk1>':
			case '</a:dk1>':
			/* 20.1.4.1.10 dk2 (Dark 2) */
			case '<a:dk2>':
			case '</a:dk2>':
			/* 20.1.4.1.22 lt1 (Light 1) */
			case '<a:lt1>':
			case '</a:lt1>':
			/* 20.1.4.1.23 lt2 (Light 2) */
			case '<a:lt2>':
			case '</a:lt2>':
			/* 20.1.4.1.1 accent1 (Accent 1) */
			case '<a:accent1>':
			case '</a:accent1>':
			/* 20.1.4.1.2 accent2 (Accent 2) */
			case '<a:accent2>':
			case '</a:accent2>':
			/* 20.1.4.1.3 accent3 (Accent 3) */
			case '<a:accent3>':
			case '</a:accent3>':
			/* 20.1.4.1.4 accent4 (Accent 4) */
			case '<a:accent4>':
			case '</a:accent4>':
			/* 20.1.4.1.5 accent5 (Accent 5) */
			case '<a:accent5>':
			case '</a:accent5>':
			/* 20.1.4.1.6 accent6 (Accent 6) */
			case '<a:accent6>':
			case '</a:accent6>':
			/* 20.1.4.1.19 hlink (Hyperlink) */
			case '<a:hlink>':
			case '</a:hlink>':
			/* 20.1.4.1.15 folHlink (Followed Hyperlink) */
			case '<a:folHlink>':
			case '</a:folHlink>':
				if (y[0][1] === '/') {
					themes.themeElements.clrScheme.push(color);
					color = {};
				} else {
					color.name = y[0].substring(3, y[0].length - 1);
				}
				break;

			default: if(opts.WTF) throw 'unrecognized ' + y[0] + ' in clrScheme';
		}
	});
}

/* 20.1.4.1.18 fontScheme CT_FontScheme */
function parse_fontScheme(t, opts) { }

/* 20.1.4.1.15 fmtScheme CT_StyleMatrix */
function parse_fmtScheme(t, opts) { }

var clrsregex = /<a:clrScheme([^>]*)>[^\u2603]*<\/a:clrScheme>/;
var fntsregex = /<a:fontScheme([^>]*)>[^\u2603]*<\/a:fontScheme>/;
var fmtsregex = /<a:fmtScheme([^>]*)>[^\u2603]*<\/a:fmtScheme>/;

/* 20.1.6.10 themeElements CT_BaseStyles */
function parse_themeElements(data, opts) {
	themes.themeElements = {};

	var t;

	[
		/* clrScheme CT_ColorScheme */
		['clrScheme', clrsregex, parse_clrScheme],
		/* fontScheme CT_FontScheme */
		['fontScheme', fntsregex, parse_fontScheme],
		/* fmtScheme CT_StyleMatrix */
		['fmtScheme', fmtsregex, parse_fmtScheme]
	].forEach(function(m) {
		if(!(t=data.match(m[1]))) throw new Error(m[0] + ' not found in themeElements');
		m[2](t, opts);
	});
}

var themeltregex = /<a:themeElements([^>]*)>[^\u2603]*<\/a:themeElements>/;

/* 14.2.7 Theme Part */
function parse_theme_xml(data, opts) {
	/* 20.1.6.9 theme CT_OfficeStyleSheet */
	if(!data || data.length === 0) return themes;

	var t;

	/* themeElements CT_BaseStyles */
	if(!(t=data.match(themeltregex))) throw 'themeElements not found in theme';
	parse_themeElements(t[0], opts);

	return themes;
}

function write_theme() {
	var o = [XML_HEADER];
	o[o.length] = '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">';
	o[o.length] =  '<a:themeElements>';

	o[o.length] =   '<a:clrScheme name="Office">';
	o[o.length] =    '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>';
	o[o.length] =    '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>';
	o[o.length] =    '<a:dk2><a:srgbClr val="1F497D"/></a:dk2>';
	o[o.length] =    '<a:lt2><a:srgbClr val="EEECE1"/></a:lt2>';
	o[o.length] =    '<a:accent1><a:srgbClr val="4F81BD"/></a:accent1>';
	o[o.length] =    '<a:accent2><a:srgbClr val="C0504D"/></a:accent2>';
	o[o.length] =    '<a:accent3><a:srgbClr val="9BBB59"/></a:accent3>';
	o[o.length] =    '<a:accent4><a:srgbClr val="8064A2"/></a:accent4>';
	o[o.length] =    '<a:accent5><a:srgbClr val="4BACC6"/></a:accent5>';
	o[o.length] =    '<a:accent6><a:srgbClr val="F79646"/></a:accent6>';
	o[o.length] =    '<a:hlink><a:srgbClr val="0000FF"/></a:hlink>';
	o[o.length] =    '<a:folHlink><a:srgbClr val="800080"/></a:folHlink>';
	o[o.length] =   '</a:clrScheme>';

	o[o.length] =   '<a:fontScheme name="Office">';
	o[o.length] =    '<a:majorFont>';
	o[o.length] =     '<a:latin typeface="Cambria"/>';
	o[o.length] =     '<a:ea typeface=""/>';
	o[o.length] =     '<a:cs typeface=""/>';
	o[o.length] =     '<a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>';
	o[o.length] =     '<a:font script="Hang" typeface="맑은 고딕"/>';
	o[o.length] =     '<a:font script="Hans" typeface="宋体"/>';
	o[o.length] =     '<a:font script="Hant" typeface="新細明體"/>';
	o[o.length] =     '<a:font script="Arab" typeface="Times New Roman"/>';
	o[o.length] =     '<a:font script="Hebr" typeface="Times New Roman"/>';
	o[o.length] =     '<a:font script="Thai" typeface="Tahoma"/>';
	o[o.length] =     '<a:font script="Ethi" typeface="Nyala"/>';
	o[o.length] =     '<a:font script="Beng" typeface="Vrinda"/>';
	o[o.length] =     '<a:font script="Gujr" typeface="Shruti"/>';
	o[o.length] =     '<a:font script="Khmr" typeface="MoolBoran"/>';
	o[o.length] =     '<a:font script="Knda" typeface="Tunga"/>';
	o[o.length] =     '<a:font script="Guru" typeface="Raavi"/>';
	o[o.length] =     '<a:font script="Cans" typeface="Euphemia"/>';
	o[o.length] =     '<a:font script="Cher" typeface="Plantagenet Cherokee"/>';
	o[o.length] =     '<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>';
	o[o.length] =     '<a:font script="Tibt" typeface="Microsoft Himalaya"/>';
	o[o.length] =     '<a:font script="Thaa" typeface="MV Boli"/>';
	o[o.length] =     '<a:font script="Deva" typeface="Mangal"/>';
	o[o.length] =     '<a:font script="Telu" typeface="Gautami"/>';
	o[o.length] =     '<a:font script="Taml" typeface="Latha"/>';
	o[o.length] =     '<a:font script="Syrc" typeface="Estrangelo Edessa"/>';
	o[o.length] =     '<a:font script="Orya" typeface="Kalinga"/>';
	o[o.length] =     '<a:font script="Mlym" typeface="Kartika"/>';
	o[o.length] =     '<a:font script="Laoo" typeface="DokChampa"/>';
	o[o.length] =     '<a:font script="Sinh" typeface="Iskoola Pota"/>';
	o[o.length] =     '<a:font script="Mong" typeface="Mongolian Baiti"/>';
	o[o.length] =     '<a:font script="Viet" typeface="Times New Roman"/>';
	o[o.length] =     '<a:font script="Uigh" typeface="Microsoft Uighur"/>';
	o[o.length] =     '<a:font script="Geor" typeface="Sylfaen"/>';
	o[o.length] =    '</a:majorFont>';
	o[o.length] =    '<a:minorFont>';
	o[o.length] =     '<a:latin typeface="Calibri"/>';
	o[o.length] =     '<a:ea typeface=""/>';
	o[o.length] =     '<a:cs typeface=""/>';
	o[o.length] =     '<a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>';
	o[o.length] =     '<a:font script="Hang" typeface="맑은 고딕"/>';
	o[o.length] =     '<a:font script="Hans" typeface="宋体"/>';
	o[o.length] =     '<a:font script="Hant" typeface="新細明體"/>';
	o[o.length] =     '<a:font script="Arab" typeface="Arial"/>';
	o[o.length] =     '<a:font script="Hebr" typeface="Arial"/>';
	o[o.length] =     '<a:font script="Thai" typeface="Tahoma"/>';
	o[o.length] =     '<a:font script="Ethi" typeface="Nyala"/>';
	o[o.length] =     '<a:font script="Beng" typeface="Vrinda"/>';
	o[o.length] =     '<a:font script="Gujr" typeface="Shruti"/>';
	o[o.length] =     '<a:font script="Khmr" typeface="DaunPenh"/>';
	o[o.length] =     '<a:font script="Knda" typeface="Tunga"/>';
	o[o.length] =     '<a:font script="Guru" typeface="Raavi"/>';
	o[o.length] =     '<a:font script="Cans" typeface="Euphemia"/>';
	o[o.length] =     '<a:font script="Cher" typeface="Plantagenet Cherokee"/>';
	o[o.length] =     '<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>';
	o[o.length] =     '<a:font script="Tibt" typeface="Microsoft Himalaya"/>';
	o[o.length] =     '<a:font script="Thaa" typeface="MV Boli"/>';
	o[o.length] =     '<a:font script="Deva" typeface="Mangal"/>';
	o[o.length] =     '<a:font script="Telu" typeface="Gautami"/>';
	o[o.length] =     '<a:font script="Taml" typeface="Latha"/>';
	o[o.length] =     '<a:font script="Syrc" typeface="Estrangelo Edessa"/>';
	o[o.length] =     '<a:font script="Orya" typeface="Kalinga"/>';
	o[o.length] =     '<a:font script="Mlym" typeface="Kartika"/>';
	o[o.length] =     '<a:font script="Laoo" typeface="DokChampa"/>';
	o[o.length] =     '<a:font script="Sinh" typeface="Iskoola Pota"/>';
	o[o.length] =     '<a:font script="Mong" typeface="Mongolian Baiti"/>';
	o[o.length] =     '<a:font script="Viet" typeface="Arial"/>';
	o[o.length] =     '<a:font script="Uigh" typeface="Microsoft Uighur"/>';
	o[o.length] =     '<a:font script="Geor" typeface="Sylfaen"/>';
	o[o.length] =    '</a:minorFont>';
	o[o.length] =   '</a:fontScheme>';

	o[o.length] =   '<a:fmtScheme name="Office">';
	o[o.length] =    '<a:fillStyleLst>';
	o[o.length] =     '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
	o[o.length] =     '<a:gradFill rotWithShape="1">';
	o[o.length] =      '<a:gsLst>';
	o[o.length] =       '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs>';
	o[o.length] =       '<a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs>';
	o[o.length] =       '<a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs>';
	o[o.length] =      '</a:gsLst>';
	o[o.length] =      '<a:lin ang="16200000" scaled="1"/>';
	o[o.length] =     '</a:gradFill>';
	o[o.length] =     '<a:gradFill rotWithShape="1">';
	o[o.length] =      '<a:gsLst>';
	o[o.length] =       '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="100000"/><a:shade val="100000"/><a:satMod val="130000"/></a:schemeClr></a:gs>';
	o[o.length] =       '<a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="50000"/><a:shade val="100000"/><a:satMod val="350000"/></a:schemeClr></a:gs>';
	o[o.length] =      '</a:gsLst>';
	o[o.length] =      '<a:lin ang="16200000" scaled="0"/>';
	o[o.length] =     '</a:gradFill>';
	o[o.length] =    '</a:fillStyleLst>';
	o[o.length] =    '<a:lnStyleLst>';
	o[o.length] =     '<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln>';
	o[o.length] =     '<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>';
	o[o.length] =     '<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>';
	o[o.length] =    '</a:lnStyleLst>';
	o[o.length] =    '<a:effectStyleLst>';
	o[o.length] =     '<a:effectStyle>';
	o[o.length] =      '<a:effectLst>';
	o[o.length] =       '<a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw>';
	o[o.length] =      '</a:effectLst>';
	o[o.length] =     '</a:effectStyle>';
	o[o.length] =     '<a:effectStyle>';
	o[o.length] =      '<a:effectLst>';
	o[o.length] =       '<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw>';
	o[o.length] =      '</a:effectLst>';
	o[o.length] =     '</a:effectStyle>';
	o[o.length] =     '<a:effectStyle>';
	o[o.length] =      '<a:effectLst>';
	o[o.length] =       '<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw>';
	o[o.length] =      '</a:effectLst>';
	o[o.length] =      '<a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d>';
	o[o.length] =      '<a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d>';
	o[o.length] =     '</a:effectStyle>';
	o[o.length] =    '</a:effectStyleLst>';
	o[o.length] =    '<a:bgFillStyleLst>';
	o[o.length] =     '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
	o[o.length] =     '<a:gradFill rotWithShape="1">';
	o[o.length] =      '<a:gsLst>';
	o[o.length] =       '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs>';
	o[o.length] =       '<a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs>';
	o[o.length] =       '<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs>';
	o[o.length] =      '</a:gsLst>';
	o[o.length] =      '<a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path>';
	o[o.length] =     '</a:gradFill>';
	o[o.length] =     '<a:gradFill rotWithShape="1">';
	o[o.length] =      '<a:gsLst>';
	o[o.length] =       '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs>';
	o[o.length] =       '<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs>';
	o[o.length] =      '</a:gsLst>';
	o[o.length] =      '<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>';
	o[o.length] =     '</a:gradFill>';
	o[o.length] =    '</a:bgFillStyleLst>';
	o[o.length] =   '</a:fmtScheme>';

	o[o.length] =  '</a:themeElements>';
	o[o.length] =  '<a:objectDefaults>';
	o[o.length] =   '<a:spDef>';
	o[o.length] =    '<a:spPr/><a:bodyPr/><a:lstStyle/><a:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></a:style>';
	o[o.length] =   '</a:spDef>';
	o[o.length] =   '<a:lnDef>';
	o[o.length] =    '<a:spPr/><a:bodyPr/><a:lstStyle/><a:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="1"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></a:style>';
	o[o.length] =     '</a:lnDef>';
	o[o.length] =  '</a:objectDefaults>';
	o[o.length] =  '<a:extraClrSchemeLst/>';
	o[o.length] = '</a:theme>';
	return o.join("");
}
