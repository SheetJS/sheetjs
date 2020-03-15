var XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
var attregexg=/([^"\s?>\/]+)\s*=\s*((?:")([^"]*)(?:")|(?:')([^']*)(?:')|([^'">\s]+))/g;
var tagregex=/<[\/\?]?[a-zA-Z0-9:_-]+(?:\s+[^"\s?>\/]+\s*=\s*(?:"[^"]*"|'[^']*'|[^'">\s=]+))*\s?[\/\?]?>/g;

if(!(XML_HEADER.match(tagregex))) tagregex = /<[^>]*>/g;
var nsregex=/<\w*:/, nsregex2 = /<(\/?)\w+:/;
function parsexmltag(tag/*:string*/, skip_root/*:?boolean*/, skip_LC/*:?boolean*/)/*:any*/ {
	var z = ({}/*:any*/);
	var eq = 0, c = 0;
	for(; eq !== tag.length; ++eq) if((c = tag.charCodeAt(eq)) === 32 || c === 10 || c === 13) break;
	if(!skip_root) z[0] = tag.slice(0, eq);
	if(eq === tag.length) return z;
	var m = tag.match(attregexg), j=0, v="", i=0, q="", cc="", quot = 1;
	if(m) for(i = 0; i != m.length; ++i) {
		cc = m[i];
		for(c=0; c != cc.length; ++c) if(cc.charCodeAt(c) === 61) break;
		q = cc.slice(0,c).trim();
		while(cc.charCodeAt(c+1) == 32) ++c;
		quot = ((eq=cc.charCodeAt(c+1)) == 34 || eq == 39) ? 1 : 0;
		v = cc.slice(c+1+quot, cc.length-quot);
		for(j=0;j!=q.length;++j) if(q.charCodeAt(j) === 58) break;
		if(j===q.length) {
			if(q.indexOf("_") > 0) q = q.slice(0, q.indexOf("_")); // from ods
			z[q] = v;
			if(!skip_LC) z[q.toLowerCase()] = v;
		}
		else {
			var k = (j===5 && q.slice(0,5)==="xmlns"?"xmlns":"")+q.slice(j+1);
			if(z[k] && q.slice(j-3,j) == "ext") continue; // from ods
			z[k] = v;
			if(!skip_LC) z[k.toLowerCase()] = v;
		}
	}
	return z;
}
function strip_ns(x/*:string*/)/*:string*/ { return x.replace(nsregex2, "<$1"); }

var encodings = {
	'&quot;': '"',
	'&apos;': "'",
	'&gt;': '>',
	'&lt;': '<',
	'&amp;': '&'
};
var rencoding = evert(encodings);
//var rencstr = "&<>'\"".split("");

// TODO: CP remap (need to read file version to determine OS)
var unescapexml/*:StringConv*/ = (function() {
	/* 22.4.2.4 bstr (Basic String) */
	var encregex = /&(?:quot|apos|gt|lt|amp|#x?([\da-fA-F]+));/ig, coderegex = /_x([\da-fA-F]{4})_/ig;
	return function unescapexml(text/*:string*/)/*:string*/ {
		var s = text + '', i = s.indexOf("<![CDATA[");
		if(i == -1) return s.replace(encregex, function($$, $1) { return encodings[$$]||String.fromCharCode(parseInt($1,$$.indexOf("x")>-1?16:10))||$$; }).replace(coderegex,function(m,c) {return String.fromCharCode(parseInt(c,16));});
		var j = s.indexOf("]]>");
		return unescapexml(s.slice(0, i)) + s.slice(i+9,j) + unescapexml(s.slice(j+3));
	};
})();

var decregex=/[&<>'"]/g, charegex = /[\u0000-\u0008\u000b-\u001f]/g;
function escapexml(text/*:string*/)/*:string*/{
	var s = text + '';
	return s.replace(decregex, function(y) { return rencoding[y]; }).replace(charegex,function(s) { return "_x" + ("000"+s.charCodeAt(0).toString(16)).slice(-4) + "_";});
}
function escapexmltag(text/*:string*/)/*:string*/{ return escapexml(text).replace(/ /g,"_x0020_"); }

var htmlcharegex = /[\u0000-\u001f]/g;
function escapehtml(text/*:string*/)/*:string*/{
	var s = text + '';
	return s.replace(decregex, function(y) { return rencoding[y]; }).replace(/\n/g, "<br/>").replace(htmlcharegex,function(s) { return "&#x" + ("000"+s.charCodeAt(0).toString(16)).slice(-4) + ";"; });
}

function escapexlml(text/*:string*/)/*:string*/{
	var s = text + '';
	return s.replace(decregex, function(y) { return rencoding[y]; }).replace(htmlcharegex,function(s) { return "&#x" + (s.charCodeAt(0).toString(16)).toUpperCase() + ";"; });
}

/* TODO: handle codepages */
var xlml_fixstr/*:StringConv*/ = (function() {
	var entregex = /&#(\d+);/g;
	function entrepl($$/*:string*/,$1/*:string*/)/*:string*/ { return String.fromCharCode(parseInt($1,10)); }
	return function xlml_fixstr(str/*:string*/)/*:string*/ { return str.replace(entregex,entrepl); };
})();
var xlml_unfixstr/*:StringConv*/ = (function() {
	return function xlml_unfixstr(str/*:string*/)/*:string*/ { return str.replace(/(\r\n|[\r\n])/g,"\&#10;"); };
})();

function parsexmlbool(value/*:any*/)/*:boolean*/ {
	switch(value) {
		case 1: case true: case '1': case 'true': case 'TRUE': return true;
		/* case '0': case 'false': case 'FALSE':*/
		default: return false;
	}
}

var utf8read/*:StringConv*/ = function utf8reada(orig/*:string*/)/*:string*/ {
	var out = "", i = 0, c = 0, d = 0, e = 0, f = 0, w = 0;
	while (i < orig.length) {
		c = orig.charCodeAt(i++);
		if (c < 128) { out += String.fromCharCode(c); continue; }
		d = orig.charCodeAt(i++);
		if (c>191 && c<224) { f = ((c & 31) << 6); f |= (d & 63); out += String.fromCharCode(f); continue; }
		e = orig.charCodeAt(i++);
		if (c < 240) { out += String.fromCharCode(((c & 15) << 12) | ((d & 63) << 6) | (e & 63)); continue; }
		f = orig.charCodeAt(i++);
		w = (((c & 7) << 18) | ((d & 63) << 12) | ((e & 63) << 6) | (f & 63))-65536;
		out += String.fromCharCode(0xD800 + ((w>>>10)&1023));
		out += String.fromCharCode(0xDC00 + (w&1023));
	}
	return out;
};

var utf8write/*:StringConv*/ = function(orig/*:string*/)/*:string*/ {
	var out/*:Array<string>*/ = [], i = 0, c = 0, d = 0;
	while(i < orig.length) {
		c = orig.charCodeAt(i++);
		switch(true) {
			case c < 128: out.push(String.fromCharCode(c)); break;
			case c < 2048:
				out.push(String.fromCharCode(192 + (c >> 6)));
				out.push(String.fromCharCode(128 + (c & 63)));
				break;
			case c >= 55296 && c < 57344:
				c -= 55296; d = orig.charCodeAt(i++) - 56320 + (c<<10);
				out.push(String.fromCharCode(240 + ((d >>18) & 7)));
				out.push(String.fromCharCode(144 + ((d >>12) & 63)));
				out.push(String.fromCharCode(128 + ((d >> 6) & 63)));
				out.push(String.fromCharCode(128 + (d & 63)));
				break;
			default:
				out.push(String.fromCharCode(224 + (c >> 12)));
				out.push(String.fromCharCode(128 + ((c >> 6) & 63)));
				out.push(String.fromCharCode(128 + (c & 63)));
		}
	}
	return out.join("");
};

if(has_buf) {
	var utf8readb = function utf8readb(data) {
		var out = Buffer.alloc(2*data.length), w, i, j = 1, k = 0, ww=0, c;
		for(i = 0; i < data.length; i+=j) {
			j = 1;
			if((c=data.charCodeAt(i)) < 128) w = c;
			else if(c < 224) { w = (c&31)*64+(data.charCodeAt(i+1)&63); j=2; }
			else if(c < 240) { w=(c&15)*4096+(data.charCodeAt(i+1)&63)*64+(data.charCodeAt(i+2)&63); j=3; }
			else { j = 4;
				w = (c & 7)*262144+(data.charCodeAt(i+1)&63)*4096+(data.charCodeAt(i+2)&63)*64+(data.charCodeAt(i+3)&63);
				w -= 65536; ww = 0xD800 + ((w>>>10)&1023); w = 0xDC00 + (w&1023);
			}
			if(ww !== 0) { out[k++] = ww&255; out[k++] = ww>>>8; ww = 0; }
			out[k++] = w%256; out[k++] = w>>>8;
		}
		return out.slice(0,k).toString('ucs2');
	};
	var corpus = "foo bar baz\u00e2\u0098\u0083\u00f0\u009f\u008d\u00a3";
	if(utf8read(corpus) == utf8readb(corpus)) utf8read = utf8readb;
	var utf8readc = function utf8readc(data) { return Buffer_from(data, 'binary').toString('utf8'); };
	if(utf8read(corpus) == utf8readc(corpus)) utf8read = utf8readc;

	utf8write = function(data) { return Buffer_from(data, 'utf8').toString("binary"); };
}

// matches <foo>...</foo> extracts content
var matchtag = (function() {
	var mtcache/*:{[k:string]:RegExp}*/ = ({}/*:any*/);
	return function matchtag(f/*:string*/,g/*:?string*/)/*:RegExp*/ {
		var t = f+"|"+(g||"");
		if(mtcache[t]) return mtcache[t];
		return (mtcache[t] = new RegExp('<(?:\\w+:)?'+f+'(?: xml:space="preserve")?(?:[^>]*)>([\\s\\S]*?)</(?:\\w+:)?'+f+'>',((g||"")/*:any*/)));
	};
})();

var htmldecode/*:{(s:string):string}*/ = (function() {
	var entities/*:Array<[RegExp, string]>*/ = [
		['nbsp', ' '], ['middot', '·'],
		['quot', '"'], ['apos', "'"], ['gt',   '>'], ['lt',   '<'], ['amp',  '&']
	].map(function(x/*:[string, string]*/) { return [new RegExp('&' + x[0] + ';', "ig"), x[1]]; });
	return function htmldecode(str/*:string*/)/*:string*/ {
		var o = str
				// Remove new lines and spaces from start of content
				.replace(/^[\t\n\r ]+/, "")
				// Remove new lines and spaces from end of content
				.replace(/[\t\n\r ]+$/,"")
				// Added line which removes any white space characters after and before html tags
				.replace(/>\s+/g,">").replace(/\s+</g,"<")
				// Replace remaining new lines and spaces with space
				.replace(/[\t\n\r ]+/g, " ")
				// Replace <br> tags with new lines
				.replace(/<\s*[bB][rR]\s*\/?>/g,"\n")
				// Strip HTML elements
				.replace(/<[^>]*>/g,"");
		for(var i = 0; i < entities.length; ++i) o = o.replace(entities[i][0], entities[i][1]);
		return o;
	};
})();

var vtregex = (function(){ var vt_cache = {};
	return function vt_regex(bt) {
		if(vt_cache[bt] !== undefined) return vt_cache[bt];
		return (vt_cache[bt] = new RegExp("<(?:vt:)?" + bt + ">([\\s\\S]*?)</(?:vt:)?" + bt + ">", 'g') );
};})();
var vtvregex = /<\/?(?:vt:)?variant>/g, vtmregex = /<(?:vt:)([^>]*)>([\s\S]*)</;
function parseVector(data/*:string*/, opts)/*:Array<{v:string,t:string}>*/ {
	var h = parsexmltag(data);

	var matches/*:Array<string>*/ = data.match(vtregex(h.baseType))||[];
	var res/*:Array<any>*/ = [];
	if(matches.length != h.size) {
		if(opts.WTF) throw new Error("unexpected vector length " + matches.length + " != " + h.size);
		return res;
	}
	matches.forEach(function(x/*:string*/) {
		var v = x.replace(vtvregex,"").match(vtmregex);
		if(v) res.push({v:utf8read(v[2]), t:v[1]});
	});
	return res;
}

var wtregex = /(^\s|\s$|\n)/;
function writetag(f/*:string*/,g/*:string*/)/*:string*/ { return '<' + f + (g.match(wtregex)?' xml:space="preserve"' : "") + '>' + g + '</' + f + '>'; }

function wxt_helper(h)/*:string*/ { return keys(h).map(function(k) { return " " + k + '="' + h[k] + '"';}).join(""); }
function writextag(f/*:string*/,g/*:?string*/,h) { return '<' + f + ((h != null) ? wxt_helper(h) : "") + ((g != null) ? (g.match(wtregex)?' xml:space="preserve"' : "") + '>' + g + '</' + f : "/") + '>';}

function write_w3cdtf(d/*:Date*/, t/*:?boolean*/)/*:string*/ { try { return d.toISOString().replace(/\.\d*/,""); } catch(e) { if(t) throw e; } return ""; }

function write_vt(s)/*:string*/ {
	switch(typeof s) {
		case 'string': return writextag('vt:lpwstr', escapexml(s));
		case 'number': return writextag((s|0)==s?'vt:i4':'vt:r8', escapexml(String(s)));
		case 'boolean': return writextag('vt:bool',s?'true':'false');
	}
	if(s instanceof Date) return writextag('vt:filetime', write_w3cdtf(s));
	throw new Error("Unable to serialize " + s);
}

var XMLNS = ({
	'dc': 'http://purl.org/dc/elements/1.1/',
	'dcterms': 'http://purl.org/dc/terms/',
	'dcmitype': 'http://purl.org/dc/dcmitype/',
	'mx': 'http://schemas.microsoft.com/office/mac/excel/2008/main',
	'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
	'sjs': 'http://schemas.openxmlformats.org/package/2006/sheetjs/core-properties',
	'vt': 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
	'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
	'xsd': 'http://www.w3.org/2001/XMLSchema'
}/*:any*/);

XMLNS.main = [
	'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
	'http://purl.oclc.org/ooxml/spreadsheetml/main',
	'http://schemas.microsoft.com/office/excel/2006/main',
	'http://schemas.microsoft.com/office/excel/2006/2'
];

var XLMLNS = ({
	'o':    'urn:schemas-microsoft-com:office:office',
	'x':    'urn:schemas-microsoft-com:office:excel',
	'ss':   'urn:schemas-microsoft-com:office:spreadsheet',
	'dt':   'uuid:C2F41010-65B3-11d1-A29F-00AA00C14882',
	'mv':   'http://macVmlSchemaUri',
	'v':    'urn:schemas-microsoft-com:vml',
	'html': 'http://www.w3.org/TR/REC-html40'
}/*:any*/);
