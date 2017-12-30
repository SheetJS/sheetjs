/* 18.7 Comments */
function parse_comments_xml(data/*:string*/, opts)/*:Array<RawComment>*/ {
	/* 18.7.6 CT_Comments */
	if(data.match(/<(?:\w+:)?comments *\/>/)) return [];
	var authors/*:Array<string>*/ = [];
	var commentList/*:Array<RawComment>*/ = [];
	var authtag = data.match(/<(?:\w+:)?authors>([\s\S]*)<\/(?:\w+:)?authors>/);
	if(authtag && authtag[1]) authtag[1].split(/<\/\w*:?author>/).forEach(function(x) {
		if(x === "" || x.trim() === "") return;
		var a = x.match(/<(?:\w+:)?author[^>]*>(.*)/);
		if(a) authors.push(a[1]);
	});
	var cmnttag = data.match(/<(?:\w+:)?commentList>([\s\S]*)<\/(?:\w+:)?commentList>/);
	if(cmnttag && cmnttag[1]) cmnttag[1].split(/<\/\w*:?comment>/).forEach(function(x, index) {
		if(x === "" || x.trim() === "") return;
		var cm = x.match(/<(?:\w+:)?comment[^>]*>/);
		if(!cm) return;
		var y = parsexmltag(cm[0]);
		var comment/*:RawComment*/ = ({ author: y.authorId && authors[y.authorId] || "sheetjsghost", ref: y.ref, guid: y.guid }/*:any*/);
		var cell = decode_cell(y.ref);
		if(opts.sheetRows && opts.sheetRows <= cell.r) return;
		var textMatch = x.match(/<(?:\w+:)?text>([\s\S]*)<\/(?:\w+:)?text>/);
		var rt = !!textMatch && !!textMatch[1] && parse_si(textMatch[1]) || {r:"",t:"",h:""};
		comment.r = rt.r;
		if(rt.r == "<t></t>") rt.t = rt.h = "";
		comment.t = rt.t.replace(/\r\n/g,"\n").replace(/\r/g,"\n");
		if(opts.cellHTML) comment.h = rt.h;
		commentList.push(comment);
	});
	return commentList;
}

var CMNT_XML_ROOT = writextag('comments', null, { 'xmlns': XMLNS.main[0] });
function write_comments_xml(data, opts) {
	var o = [XML_HEADER, CMNT_XML_ROOT];

	var iauthor/*:Array<string>*/ = [];
	o.push("<authors>");
	data.map(function(x) { return x[1]; }).forEach(function(comment) {
		comment.map(function(x) { return escapexml(x.a); }).forEach(function(a) {
			if(iauthor.indexOf(a) > -1) return;
			iauthor.push(a);
			o.push("<author>" + a + "</author>");
		});
	});
	o.push("</authors>");
	o.push("<commentList>");
	data.forEach(function(d) {
		d[1].forEach(function(c) {
			/* 18.7.3 CT_Comment */
			o.push('<comment ref="' + d[0] + '" authorId="' + iauthor.indexOf(escapexml(c.a)) + '"><text>');
			o.push(writetag("t", c.t == null ? "" : c.t));
			o.push('</text></comment>');
		});
	});
	o.push("</commentList>");
	if(o.length>2) { o[o.length] = ('</comments>'); o[1]=o[1].replace("/>",">"); }
	return o.join("");
}
