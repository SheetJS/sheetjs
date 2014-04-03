/* 18.7.3 CT_Comment */
function parse_comments_xml(data, opts) {
	if(data.match(/<(?:\w+:)?comments *\/>/)) return [];
	var authors = [];
	var commentList = [];
	data.match(/<(?:\w+:)?authors>([^\u2603]*)<\/(?:\w+:)?authors>/m)[1].split(/<\/(?:\w+:)?author>/).forEach(function(x) {
		if(x === "" || x.trim() === "") return;
		authors.push(x.match(/<(?:\w+:)?author[^>]*>(.*)/)[1]);
	});
	(data.match(/<(?:\w+:)?commentList>([^\u2603]*)<\/(?:\w+:)?commentList>/m)||["",""])[1].split(/<\/(?:\w+:)?comment>/).forEach(function(x, index) {
		if(x === "" || x.trim() === "") return;
		var y = parsexmltag(x.match(/<(?:\w+:)?comment[^>]*>/)[0]);
		var comment = { author: y.authorId && authors[y.authorId] ? authors[y.authorId] : undefined, ref: y.ref, guid: y.guid };
		var cell = decode_cell(y.ref);
		if(opts.sheetRows && opts.sheetRows <= cell.r) return;
		var textMatch = x.match(/<text>([^\u2603]*)<\/text>/m);
		if (!textMatch || !textMatch[1]) return; // a comment may contain an empty text tag.
		var rt = parse_si(textMatch[1]);
		comment.r = rt.r;
		comment.t = rt.t;
		if(opts.cellHTML) comment.h = rt.h;
		commentList.push(comment);
	});
	return commentList;
}
