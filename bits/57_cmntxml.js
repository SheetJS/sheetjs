/* 18.7.3 CT_Comment */
function parse_comments_xml(data/*:string*/, opts)/*:Array<Comment>*/ {
	if(data.match(/<(?:\w+:)?comments *\/>/)) return [];
	var authors = [];
	var commentList = [];
	var authtag = data.match(/<(?:\w+:)?authors>([^\u2603]*)<\/(?:\w+:)?authors>/);
	if(authtag && authtag[1]) authtag[1].split(/<\/\w*:?author>/).forEach(function(x) {
		if(x === "" || x.trim() === "") return;
		var a = x.match(/<(?:\w+:)?author[^>]*>(.*)/);
		if(a) authors.push(a[1]);
	});
	var cmnttag = data.match(/<(?:\w+:)?commentList>([^\u2603]*)<\/(?:\w+:)?commentList>/);
	if(cmnttag && cmnttag[1]) cmnttag[1].split(/<\/\w*:?comment>/).forEach(function(x, index) {
		if(x === "" || x.trim() === "") return;
		var cm = x.match(/<(?:\w+:)?comment[^>]*>/);
		if(!cm) return;
		var y = parsexmltag(cm[0]);
		var comment/*:Comment*/ = ({ author: y.authorId && authors[y.authorId] ? authors[y.authorId] : undefined, ref: y.ref, guid: y.guid }/*:any*/);
		var cell = decode_cell(y.ref);
		if(opts.sheetRows && opts.sheetRows <= cell.r) return;
		var textMatch = x.match(/<(?:\w+:)?text>([^\u2603]*)<\/(?:\w+:)?text>/);
		if (!textMatch || !textMatch[1]) return; // a comment may contain an empty text tag.
		var rt = parse_si(textMatch[1]);
		if(!rt) return;
		comment.r = rt.r;
		comment.t = rt.t;
		if(opts.cellHTML) comment.h = rt.h;
		commentList.push(comment);
	});
	return commentList;
}

function write_comments_xml(data, opts) { }
