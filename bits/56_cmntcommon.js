RELS.CMNT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

function sheet_insert_comments(sheet, comments/*:Array<RawComment>*/) {
	var dense = Array.isArray(sheet);
	var cell/*:Cell*/;
	comments.forEach(function(comment) {
		var r = decode_cell(comment.ref);
		if(dense) {
			if(!sheet[r.r]) sheet[r.r] = [];
			cell = sheet[r.r][r.c];
		} else cell = sheet[comment.ref];
		if (!cell) {
			cell = ({t:"z"}/*:any*/);
			if(dense) sheet[r.r][r.c] = cell;
			else sheet[comment.ref] = cell;
			var range = safe_decode_range(sheet["!ref"]||"BDWGO1000001:A1");
			if(range.s.r > r.r) range.s.r = r.r;
			if(range.e.r < r.r) range.e.r = r.r;
			if(range.s.c > r.c) range.s.c = r.c;
			if(range.e.c < r.c) range.e.c = r.c;
			var encoded = encode_range(range);
			if (encoded !== sheet["!ref"]) sheet["!ref"] = encoded;
		}

		if (!cell.c) cell.c = [];
		var o/*:Comment*/ = ({a: comment.author, t: comment.t, r: comment.r});
		if(comment.h) o.h = comment.h;
		cell.c.push(o);
	});
}

