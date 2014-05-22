var current_codepage = 1252, current_cptable;
if(typeof module !== "undefined" && typeof require !== 'undefined') {
	if(typeof cptable === 'undefined') cptable = require('codepage');
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

function char_codes(data) { return data.split("").map(function(x) { return x.charCodeAt(0); }); }
function debom_xml(data) {
	if(typeof cptable !== 'undefined') {
		if(data.charCodeAt(0) === 0xFF && data.charCodeAt(1) === 0xFE) { return cptable.utils.decode(1200, char_codes(data.substr(2))); }
	}
	return data;
}
