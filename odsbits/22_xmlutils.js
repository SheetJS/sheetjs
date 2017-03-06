var attregexg=/[^\s?>\/]+=["'][^"]*['"]/g;
var tagregex=/<[^>]*>/g;
var nsregex=/<\w*:/, nsregex2 = /<(\/?)\w+:/;
function parsexmltag(tag, skip_root) {
	var z/*:any*/ = [];
	var eq = 0, c = 0;
	for(; eq !== tag.length; ++eq) if((c = tag.charCodeAt(eq)) === 32 || c === 10 || c === 13) break;
	if(!skip_root) z[0] = tag.substr(0, eq);
	if(eq === tag.length) return z;
	var m = tag.match(attregexg), j=0, v="", i=0, q="", cc="";
	if(m) for(i = 0; i != m.length; ++i) {
		cc = m[i];
		for(c=0; c != cc.length; ++c) if(cc.charCodeAt(c) === 61) break;
		q = cc.substr(0,c); v = cc.substring(c+2, cc.length-1);
		for(j=0;j!=q.length;++j) if(q.charCodeAt(j) === 58) break;
		if(j===q.length) {
			if(q.indexOf("_") > 0) q = q.substr(0, q.indexOf("_"));
			z[q] = v;
		}
		else {
			var k = (j===5 && q.substr(0,5)==="xmlns"?"xmlns":"")+q.substr(j+1);
			if(z[k] && q.substr(j-3,3) == "ext") continue;
			z[k] = v;
		}
	}
	return z;
}
function strip_ns(x) { return x.replace(nsregex2, "<$1"); }

var encodings = {
	'&quot;': '"',
	'&apos;': "'",
	'&gt;': '>',
	'&lt;': '<',
	'&amp;': '&'
};
var rencoding = {
	'"': '&quot;',
	"'": '&apos;',
	'>': '&gt;',
	'<': '&lt;',
	'&': '&amp;'
};
var rencstr = "&<>'\"".split("");

// TODO: CP remap (need to read file version to determine OS)
/* 22.4.2.4 bstr (Basic String) */
var encregex = /&[a-z]*;/g, coderegex = /_x([\da-fA-F]{4})_/g;
function unescapexml(text){
	var s = text + '';
	return s.replace(encregex, function($$) { return encodings[$$]; }).replace(coderegex,function(m,c) {return String.fromCharCode(parseInt(c,16));});
}
var decregex=/[&<>'"]/g, charegex = /[\u0000-\u0008\u000b-\u001f]/g;
function escapexml(text){
	var s = text + '';
	return s.replace(decregex, function(y) { return rencoding[y]; }).replace(charegex,function(s) { return "_x" + ("000"+s.charCodeAt(0).toString(16)).substr(-4) + "_";});
}

function parsexmlbool(value) {
	switch(value) {
		case '1': case 'true': case 'TRUE': return true;
		/* case '0': case 'false': case 'FALSE':*/
		default: return false;
	}
}

function datenum(v) {
	var epoch = Date.parse(v);
	return (epoch + 2209161600000) / (24 * 60 * 60 * 1000);
}

/* ISO 8601 Duration */
function parse_isodur(s) {
	var sec = 0, mt = 0, time = false;
	var m = s.match(/P([0-9\.]+Y)?([0-9\.]+M)?([0-9\.]+D)?T([0-9\.]+H)?([0-9\.]+M)?([0-9\.]+S)?/);
	if(!m) throw new Error("|" + s + "| is not an ISO8601 Duration");
	for(var i = 1; i != m.length; ++i) {
		if(!m[i]) continue;
		mt = 1;
		if(i > 3) time = true;
		switch(m[i].substr(m[i].length-1)) {
			case 'Y':
				throw new Error("Unsupported ISO Duration Field: " + m[i].substr(m[i].length-1));
			case 'D': mt *= 24;
				/* falls through */
			case 'H': mt *= 60;
				/* falls through */
			case 'M':
				if(!time) throw new Error("Unsupported ISO Duration Field: M");
				else mt *= 60;
				/* falls through */
			case 'S': break;
		}
		sec += mt * parseInt(m[i], 10);
	}
	return sec;
}

var XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
