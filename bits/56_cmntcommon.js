RELS.CMNT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

function parse_comments(zip, dirComments, sheets, sheetRels, opts) {
	for(var i = 0; i != dirComments.length; ++i) {
		var canonicalpath=dirComments[i];
		var comments=parse_cmnt(getzipdata(zip, canonicalpath.replace(/^\//,''), true), canonicalpath, opts);
		if(!comments || !comments.length) continue;
		// find the sheets targeted by these comments
		var sheetNames = keys(sheets);
		for(var j = 0; j != sheetNames.length; ++j) {
			var sheetName = sheetNames[j];
			var rels = sheetRels[sheetName];
			if(rels) {
				var rel = rels[canonicalpath];
				if(rel) insertCommentsIntoSheet(sheetName, sheets[sheetName], comments);
			}
		}
	}
}

function insertCommentsIntoSheet(sheetName, sheet, comments/*:Array<RawComment>*/) {
	var dense = Array.isArray(sheet);
	var cell/*:Cell*/, r;
	comments.forEach(function(comment) {
		if(dense) {
			r = decode_cell(comment.ref);
			if(!sheet[r.r]) sheet[r.r] = [];
			cell = sheet[r.r][r.c];
		} else cell = sheet[comment.ref];
		if (!cell) {
			cell = {};
			if(dense) sheet[r.r][r.c] = cell;
			else sheet[comment.ref] = cell;
			var range = safe_decode_range(sheet["!ref"]||"BDWGO1000001:A1");
			var thisCell = decode_cell(comment.ref);
			if(range.s.r > thisCell.r) range.s.r = thisCell.r;
			if(range.e.r < thisCell.r) range.e.r = thisCell.r;
			if(range.s.c > thisCell.c) range.s.c = thisCell.c;
			if(range.e.c < thisCell.c) range.e.c = thisCell.c;
			var encoded = encode_range(range);
			if (encoded !== sheet["!ref"]) sheet["!ref"] = encoded;
		}

		if (!cell.c) cell.c = [];
		var o/*:Comment*/ = ({a: comment.author, t: comment.t, r: comment.r});
		if(comment.h) o.h = comment.h;
		cell.c.push(o);
	});
}

