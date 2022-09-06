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
	if(cmnttag && cmnttag[1]) cmnttag[1].split(/<\/\w*:?comment>/).forEach(function(x) {
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
		comment.t = (rt.t||"").replace(/\r\n/g,"\n").replace(/\r/g,"\n");
		if(opts.cellHTML) comment.h = rt.h;
		commentList.push(comment);
	});
	return commentList;
}

function write_comments_xml(data/*::, opts*/) {
	var o = [XML_HEADER, writextag('comments', null, { 'xmlns': XMLNS_main[0] })];

	var iauthor/*:Array<string>*/ = [];
	o.push("<authors>");
	data.forEach(function(x) { x[1].forEach(function(w) { var a = escapexml(w.a);
		if(iauthor.indexOf(a) == -1) {
			iauthor.push(a);
			o.push("<author>" + a + "</author>");
		}
		if(w.T && w.ID && iauthor.indexOf("tc=" + w.ID) == -1) {
			iauthor.push("tc=" + w.ID);
			o.push("<author>" + "tc=" + w.ID + "</author>");
		}
	}); });
	if(iauthor.length == 0) { iauthor.push("SheetJ5"); o.push("<author>SheetJ5</author>"); }
	o.push("</authors>");
	o.push("<commentList>");
	data.forEach(function(d) {
		/* 18.7.3 CT_Comment */
		var lastauthor = 0, ts = [], tcnt = 0;
		if(d[1][0] && d[1][0].T && d[1][0].ID) lastauthor = iauthor.indexOf("tc=" + d[1][0].ID);
		d[1].forEach(function(c) {
			if(c.a) lastauthor = iauthor.indexOf(escapexml(c.a));
			if(c.T) ++tcnt;
			ts.push(c.t == null ? "" : escapexml(c.t));
		});
		if(tcnt === 0) {
			d[1].forEach(function(c) {
				o.push('<comment ref="' + d[0] + '" authorId="' + iauthor.indexOf(escapexml(c.a)) + '"><text>');
				o.push(writetag("t", c.t == null ? "" : escapexml(c.t)));
				o.push('</text></comment>');
			});
		} else {
			/* based on Threaded Comments -> Comments projection */
			o.push('<comment ref="' + d[0] + '" authorId="' + lastauthor + '"><text>');
			var t = "Comment:\n    " + (ts[0]) + "\n";
			for(var i = 1; i < ts.length; ++i) t += "Reply:\n    " + ts[i] + "\n";
			o.push(writetag("t", escapexml(t)));
			o.push('</text></comment>');
		}
	});
	o.push("</commentList>");
	if(o.length>2) { o[o.length] = ('</comments>'); o[1]=o[1].replace("/>",">"); }
	return o.join("");
}

/* [MS-XLSX] 2.1.17 */
function parse_tcmnt_xml(data/*:string*/, opts)/*:Array<RawComment>*/ {
	var out = [];
	var pass = false, comment = {}, tidx = 0;
	data.replace(tagregex, function xml_tcmnt(x, idx) {
		var y/*:any*/ = parsexmltag(x);
		switch(strip_ns(y[0])) {
			case '<?xml': break;

			/* 2.6.207 ThreadedComments CT_ThreadedComments */
			case '<ThreadedComments': break;
			case '</ThreadedComments>': break;

			/* 2.6.205 threadedComment CT_ThreadedComment */
			case '<threadedComment': comment = {author: y.personId, guid: y.id, ref: y.ref, T: 1}; break;
			case '</threadedComment>': if(comment.t != null) out.push(comment); break;

			case '<text>': case '<text': tidx = idx + x.length; break;
			case '</text>': comment.t = data.slice(tidx, idx).replace(/\r\n/g, "\n").replace(/\r/g, "\n"); break;

			/* 2.6.206 mentions CT_ThreadedCommentMentions TODO */
			case '<mentions': case '<mentions>': pass = true; break;
			case '</mentions>': pass = false; break;

			/* 2.6.202 mention CT_Mention TODO */

			/* 18.2.10 extLst CT_ExtensionList ? */
			case '<extLst': case '<extLst>': case '</extLst>': case '<extLst/>': break;
			/* 18.2.7  ext CT_Extension + */
			case '<ext': pass=true; break;
			case '</ext>': pass=false; break;

			default: if(!pass && opts.WTF) throw new Error('unrecognized ' + y[0] + ' in threaded comments');
		}
		return x;
	});
	return out;
}

function write_tcmnt_xml(comments, people, opts) {
	var o = [XML_HEADER, writextag('ThreadedComments', null, { 'xmlns': XMLNS.TCMNT }).replace(/[\/]>/, ">")];
	comments.forEach(function(carr) {
		var rootid = "";
		(carr[1] || []).forEach(function(c, idx) {
			if(!c.T) { delete c.ID; return; }
			if(c.a && people.indexOf(c.a) == -1) people.push(c.a);
			var tcopts = {
				ref: carr[0],
				id: "{54EE7951-7262-4200-6969-" + ("000000000000" + opts.tcid++).slice(-12) + "}"
			};
			if(idx == 0) rootid = tcopts.id;
			else tcopts.parentId = rootid;
			c.ID = tcopts.id;
			if(c.a) tcopts.personId = "{54EE7950-7262-4200-6969-" + ("000000000000" + people.indexOf(c.a)).slice(-12) + "}";
			o.push(writextag('threadedComment', writetag('text', c.t||""), tcopts));
		});
	});
	o.push('</ThreadedComments>');
	return o.join("");
}

/* [MS-XLSX] 2.1.18 */
function parse_people_xml(data/*:string*/, opts) {
	var out = [];
	var pass = false;
	data.replace(tagregex, function xml_tcmnt(x) {
		var y/*:any*/ = parsexmltag(x);
		switch(strip_ns(y[0])) {
			case '<?xml': break;

			/* 2.4.85 personList CT_PersonList */
			case '<personList': break;
			case '</personList>': break;

			/* 2.6.203 person CT_Person TODO: providers */
			case '<person': out.push({name: y.displayname, id: y.id }); break;
			case '</person>': break;

			/* 18.2.10 extLst CT_ExtensionList ? */
			case '<extLst': case '<extLst>': case '</extLst>': case '<extLst/>': break;
			/* 18.2.7  ext CT_Extension + */
			case '<ext': pass=true; break;
			case '</ext>': pass=false; break;

			default: if(!pass && opts.WTF) throw new Error('unrecognized ' + y[0] + ' in threaded comments');
		}
		return x;
	});
	return out;
}
function write_people_xml(people/*, opts*/) {
	var o = [XML_HEADER, writextag('personList', null, {
		'xmlns': XMLNS.TCMNT,
		'xmlns:x': XMLNS_main[0]
	}).replace(/[\/]>/, ">")];
	people.forEach(function(person, idx) {
		o.push(writextag('person', null, {
			displayName: person,
			id: "{54EE7950-7262-4200-6969-" + ("000000000000" + idx).slice(-12) + "}",
			userId: person,
			providerId: "None"
		}));
	});
	o.push("</personList>");
	return o.join("");
}
