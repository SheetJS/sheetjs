/*:: declare var XLSX: any; */
var get_utils = function() {
	if(typeof XLSX !== 'undefined') return XLSX.utils;
	if(typeof module !== "undefined" && typeof require !== 'undefined') try {
		return require('../' + 'xlsx').utils;
	} catch(e) {
		try { return require('./' + 'xlsx').utils; }
		catch(ee) { return require('xl' + 'sx').utils; }
	}
	throw new Error("Cannot find XLSX utils");
};
