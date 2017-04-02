/* L.5.5.2 SpreadsheetML Comments + VML Schema */
var _shapeid = 1024;
function write_comments_vml(rId, comments) {
	var csize = [21600, 21600];
	/* L.5.2.1.2 Path Attribute */
	var bbox = ["m0,0l0",csize[1],csize[0],csize[1],csize[0],"0xe"].join(",");
	var o = [
		writextag("xml", null, { 'xmlns:v': XLMLNS.v, 'xmlns:o': XLMLNS.o, 'xmlns:x': XLMLNS.x, 'xmlns:mv': XLMLNS.mv }).replace(/\/>/,">"),
		writextag("o:shapelayout", writextag("o:idmap", null, {'v:ext':"edit", 'data':rId}), {'v:ext':"edit"}),
		writextag("v:shapetype", [
			writextag("v:stroke", null, {joinstyle:"miter"}),
			writextag("v:path", null, {gradientshapeok:"t", 'o:connecttype':"rect"})
		].join(""), {id:"_x0000_t202", 'o:spt':202, coordsize:csize.join(","),path:bbox})
	];
	while(_shapeid < rId * 1000) _shapeid += 1000;

	comments.map(function(x) { return decode_cell(x[0]); }).forEach(function(c,i) { o = o.concat([
	'<v:shape' + wxt_helper({
		id:'_x0000_s' + (++_shapeid),
		type:"#_x0000_t202",
		style:"position:absolute; margin-left:80pt;margin-top:5pt;width:104pt;height:64pt;z-index:10;visibility:hidden",
		fillcolor:"#ECFAD4",
		strokecolor:"#edeaa1"
	}) + '>',
		writextag('v:fill', writextag("o:fill", null, {type:"gradientUnscaled", 'v:ext':"view"}), {'color2':"#BEFF82", 'angle':"-180", 'type':"gradient"}),
		writextag("v:shadow", null, {on:"t", 'obscured':"t"}),
		writextag("v:path", null, {'o:connecttype':"none"}),
		'<v:textbox><div style="text-align:left"></div></v:textbox>',
		'<x:ClientData ObjectType="Note">',
			'<x:MoveWithCells/>',
			'<x:SizeWithCells/>',
			/* Part 4 19.4.2.3 Anchor (Anchor) */
			writetag('x:Anchor', [c.c, 0, c.r, 0, c.c+3, 100, c.r+5, 100].join(",")),
			writetag('x:AutoFill', "False"),
			writetag('x:Row', String(c.r)),
			writetag('x:Column', String(c.c)),
		'</x:ClientData>',
	'</v:shape>'
	]); });
	o.push('</xml>');
	return o.join("");
}

