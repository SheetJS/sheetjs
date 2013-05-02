function parsexmltag(tag) {
	var words = tag.split(/\s+/);
	var z = {'0': words[0]};
	if(words.length === 1) return z;
	tag.match(/(\w+)="([^"]*)"/g).map(
		function(x){var y=x.match(/(\w+)="([^"]*)"/); z[y[1]] = y[2]; });
	return z;
}

var encodings = {
	'&quot;': '"',
	'&apos;': "'",
	'&gt;': '>',
	'&lt;': '<',
	'&amp;': '&'
};

// TODO: CP remap (need to read file version to determine OS)
function unescapexml(text){
	var s = text + '';
	for(var y in encodings) s = s.replace(new RegExp(y,'g'), encodings[y]);
	return s.replace(/_x([0-9a-fA-F]*)_/g,function(m,c) {return _chr(parseInt(c,16));});
}

function parsexmlbool(value, tag) {
	switch(value) {
		case '0': case 0: case 'false': case 'FALSE': return false;
		case '1': case 1: case 'true': case 'TRUE': return true;
		default: throw "bad boolean value " + value + " in "+(tag||"?");
	}
}

var utf8read = function(orig) {
	var out = "", i = 0, c = 0, c1 = 0, c2 = 0, c3 = 0;
	while (i < orig.length) {
		c = orig.charCodeAt(i++);
		if (c < 128) out += _chr(c);
		else {
			c2 = orig.charCodeAt(i++);
			if (c>191 && c<224) out += _chr((c & 31) << 6 | c2 & 63);
			else {
				c3 = orig.charCodeAt(i++);
				out += _chr((c & 15) << 12 | (c2 & 63) << 6 | c3 & 63);
			}
		}
	}
	return out;
};

// matches <foo>...</foo> extracts content
function matchtag(f,g) {return new RegExp('<'+f+'(?: xml:space="preserve")?>([^\u2603]*)</'+f+'>',(g||"")+"m");}

function parseVector(data) {
	var h = parsexmltag(data);

	var matches = data.match(new RegExp("<vt:" + h.baseType + ">(.*?)</vt:" + h.baseType + ">", 'g'));
	if(matches.length != h.size) throw "unexpected vector length " + matches.length + " != " + h.size;
	var res = [];
	matches.forEach(function(x) {
		var v = x.replace(/<[/]?vt:variant>/g,"").match(/<vt:([^>]*)>(.*)</);
		res.push({v:v[2], t:v[1]});
	});
	return res;
}

function isval(x) { return typeof x !== "undefined" && x !== null; }
