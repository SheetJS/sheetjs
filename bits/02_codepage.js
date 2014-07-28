var current_codepage = 1252, current_cptable;
if(typeof module !== "undefined" && typeof require !== 'undefined') {
	if(typeof cptable === 'undefined') cptable = require('./dist/cpexcel');
	current_cptable = cptable[current_codepage];
}
function reset_cp() { set_cp(1252); }
var set_cp = function(cp) { current_codepage = cp; };

function char_codes(data) { var o = []; for(var i = 0, len = data.length; i < len; ++i) o[i] = data.charCodeAt(i); return o; }
var debom_xml = function(data) { return data; };

if(typeof cptable !== 'undefined') {
	set_cp = function(cp) { current_codepage = cp; current_cptable = cptable[cp]; };
	debom_xml = function(data) {
		if(data.charCodeAt(0) === 0xFF && data.charCodeAt(1) === 0xFE) { return cptable.utils.decode(1200, char_codes(data.substr(2))); }
		return data;
	};
}
