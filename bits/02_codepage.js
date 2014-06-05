var current_codepage = 1252, current_cptable;
if(typeof module !== "undefined" && typeof require !== 'undefined') {
	if(typeof cptable === 'undefined') cptable = require('./dist/cpexcel');
	current_cptable = cptable[current_codepage];
}
function reset_cp() { set_cp(1252); }
function set_cp(cp) { current_codepage = cp; if(typeof cptable !== 'undefined') current_cptable = cptable[cp]; }

var _getchar = function(x) { return String.fromCharCode(x); };
if(typeof cptable !== 'undefined') _getchar = function(x) {
	if (current_codepage === 1200) return String.fromCharCode(x);
	if (current_cptable) return current_cptable.dec[x];
	return cptable.utils.decode(current_codepage, [x%256,x>>8])[0];
};

function char_codes(data) { var o = []; for(var i = 0; i != data.length; ++i) o[i] = data.charCodeAt(i); return o; }
function debom_xml(data) {
	if(typeof cptable !== 'undefined') {
		if(data.charCodeAt(0) === 0xFF && data.charCodeAt(1) === 0xFE) { return cptable.utils.decode(1200, char_codes(data.substr(2))); }
	}
	return data;
}
