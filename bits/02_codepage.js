var current_codepage = 1200, current_cptable;
if(typeof module !== "undefined" && typeof require !== 'undefined') {
	if(typeof cptable === 'undefined') cptable = require('./dist/cpexcel');
	current_cptable = cptable[current_codepage];
}
function reset_cp() { set_cp(1200); }
var set_cp = function(cp) { current_codepage = cp; };

function char_codes(data) { var o = []; for(var i = 0, len = data.length; i < len; ++i) o[i] = data.charCodeAt(i); return o; }
var debom = function(data/*:string*/)/*:string*/ {
	var c1 = data.charCodeAt(0), c2 = data.charCodeAt(1);
	if(c1 == 0xFF && c2 == 0xFE) return data.substr(2);
	if(c1 == 0xFE && c2 == 0xFF) return data.substr(2);
	if(c1 == 0xFEFF) return data.substr(1);
	return data;
};

var _getchar = function _gc1(x) { return String.fromCharCode(x); };
if(typeof cptable !== 'undefined') {
	set_cp = function(cp) { current_codepage = cp; current_cptable = cptable[cp]; };
	debom = function(data) {
		if(data.charCodeAt(0) === 0xFF && data.charCodeAt(1) === 0xFE) { return cptable.utils.decode(1200, char_codes(data.substr(2))); }
		return data;
	};
	_getchar = function _gc2(x) {
		if(current_codepage === 1200) return String.fromCharCode(x);
		return cptable.utils.decode(current_codepage, [x&255,x>>8])[0];
	};
}
