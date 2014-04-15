var current_codepage, current_cptable;
if(typeof module !== "undefined" && typeof require !== 'undefined') {
	if(typeof cptable === 'undefined') cptable = require('codepage');
	current_codepage = 1252; current_cptable = cptable[1252];
}
function reset_cp() {
	current_codepage = 1252; if(typeof cptable !== 'undefined') current_cptable = cptable[1252];
}
function _getchar(x) { return String.fromCharCode(x); }

