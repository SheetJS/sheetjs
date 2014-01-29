/* 18.7.3 CT_Comment */
function parseComments(data) {
	if(data.match(/<comments *\/>/)) {
		throw new Error('Not a valid comments xml');
	}
	var authors = [];
	var commentList = [];
	data.match(/<authors>([^\u2603]*)<\/authors>/m)[1].split('</author>').forEach(function(x) {
		if(x === "" || x.trim() === "") return;
		authors.push(x.match(/<author[^>]*>(.*)/)[1]);
	});
	data.match(/<commentList>([^\u2603]*)<\/commentList>/m)[1].split('</comment>').forEach(function(x, index) {
		if(x === "" || x.trim() === "") return;
		var y = parsexmltag(x.match(/<comment[^>]*>/)[0]);
		var comment = { author: y.authorId && authors[y.authorId] ? authors[y.authorId] : undefined, ref: y.ref, guid: y.guid };
		var textMatch = x.match(/<text>([^\u2603]*)<\/text>/m);
		if (!textMatch || !textMatch[1]) return; // a comment may contain an empty text tag.
		var rt = parse_si(textMatch[1]);
		comment.raw = rt.raw;
		comment.t = rt.t;
		comment.r = rt.r;
		commentList.push(comment);
	});
	return commentList;
}

function parseCommentsAddToSheets(zip, dirComments, sheets, sheetRels) {
	for(var i = 0; i != dirComments.length; ++i) {
		var canonicalpath=dirComments[i];
		var comments=parseComments(getdata(getzipfile(zip, canonicalpath.replace(/^\//,''))));
		// find the sheets targeted by these comments
		var sheetNames = Object.keys(sheets);
		for(var j = 0; j != sheetNames.length; ++j) {
			var sheetName = sheetNames[j];
			var rels = sheetRels[sheetName];
			if (rels) {
				var rel = rels[canonicalpath];
				if (rel) {
					insertCommentsIntoSheet(sheetName, sheets[sheetName], comments);
				}
			}
		}
	}
}

function insertCommentsIntoSheet(sheetName, sheet, comments) {
	comments.forEach(function(comment) {
		var cell = sheet[comment.ref];
		if (!cell) {
			cell = {};
			sheet[comment.ref] = cell;
			var range = decode_range(sheet["!ref"]);
			var thisCell = decode_cell(comment.ref);
			if(range.s.r > thisCell.r) range.s.r = thisCell.r;
			if(range.e.r < thisCell.r) range.e.r = thisCell.r;
			if(range.s.c > thisCell.c) range.s.c = thisCell.c;
			if(range.e.c < thisCell.c) range.e.c = thisCell.c;
			var encoded = encode_range(range);
			if (encoded !== sheet["!ref"]) sheet["!ref"] = encoded;
		}

		if (!cell.c) {
			cell.c = [];
		}
		cell.c.push({a: comment.author, t: comment.t, raw: comment.raw, r: comment.r});
	});
}

