/* [MS-XLSB] 2.4.28 BrtBeginComment */
function parse_BrtBeginComment(data, length) {
	var out = {};
	out.iauthor = data.read_shift(4);
	var rfx = parse_UncheckedRfX(data, 16);
	out.rfx = rfx.s;
	out.ref = encode_cell(rfx.s);
	data.l += 16; /*var guid = parse_GUID(data); */
	return out;
}
function write_BrtBeginComment(data, o) {
	if(o == null) o = new_buf(36);
	o.write_shift(4, data[1].iauthor);
	write_UncheckedRfX((data[0]/*:any*/), o);
	o.write_shift(4, 0);
	o.write_shift(4, 0);
	o.write_shift(4, 0);
	o.write_shift(4, 0);
	return o;
}

/* [MS-XLSB] 2.4.324 BrtCommentAuthor */
var parse_BrtCommentAuthor = parse_XLWideString;

/* [MS-XLSB] 2.4.325 BrtCommentText */
var parse_BrtCommentText = parse_RichStr;

/* [MS-XLSB] 2.1.7.8 Comments */
function parse_comments_bin(data, opts) {
	var out = [];
	var authors = [];
	var c = {};
	var pass = false;
	recordhopper(data, function hopper_cmnt(val, R_n, RT) {
		switch(RT) {
			case 0x0278: /* 'BrtCommentAuthor' */
				authors.push(val); break;
			case 0x027B: /* 'BrtBeginComment' */
				c = val; break;
			case 0x027D: /* 'BrtCommentText' */
				c.t = val.t; c.h = val.h; c.r = val.r; break;
			case 0x027C: /* 'BrtEndComment' */
				c.author = authors[c.iauthor];
				delete c.iauthor;
				if(opts.sheetRows && opts.sheetRows <= c.rfx.r) break;
				if(!c.t) c.t = "";
				delete c.rfx; out.push(c); break;

			/* case 'BrtUid': */

			case 0x0023: /* 'BrtFRTBegin' */
				pass = true; break;
			case 0x0024: /* 'BrtFRTEnd' */
				pass = false; break;
			case 0x0025: /* 'BrtACBegin' */ break;
			case 0x0026: /* 'BrtACEnd' */ break;


			default:
				if((R_n||"").indexOf("Begin") > 0){}
				else if((R_n||"").indexOf("End") > 0){}
				else if(!pass || opts.WTF) throw new Error("Unexpected record " + RT + " " + R_n);
		}
	});
	return out;
}

function write_comments_bin(data, opts) {
	var ba = buf_array();
	var iauthor = [];
	write_record(ba, "BrtBeginComments");
	{ /* COMMENTAUTHORS */
		write_record(ba, "BrtBeginCommentAuthors");
		data.forEach(function(comment) {
			comment[1].forEach(function(c) {
				if(iauthor.indexOf(c.a) > -1) return;
				iauthor.push(c.a.substr(0,54));
				write_record(ba, "BrtCommentAuthor", write_XLWideString(c.a.substr(0, 54)));
			});
		});
		write_record(ba, "BrtEndCommentAuthors");
	}
	{ /* COMMENTLIST */
		write_record(ba, "BrtBeginCommentList");
		data.forEach(function(comment) {
			comment[1].forEach(function(c) {
				c.iauthor = iauthor.indexOf(c.a);
				var range = {s:decode_cell(comment[0]),e:decode_cell(comment[0])};
				write_record(ba, "BrtBeginComment", write_BrtBeginComment([range, c]));
				if(c.t && c.t.length > 0) write_record(ba, "BrtCommentText", write_RichStr(c));
				write_record(ba, "BrtEndComment");
				delete c.iauthor;
			});
		});
		write_record(ba, "BrtEndCommentList");
	}
	write_record(ba, "BrtEndComments");
	return ba.end();
}
