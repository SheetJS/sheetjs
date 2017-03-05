/*:: declare var XLSX: any; */
var get_utils = function() {
	if(typeof XLSX !== 'undefined') return XLSX.utils;
	if(typeof module !== "undefined" && typeof require !== 'undefined') try {
		return require('../xlsx.js').utils;
	} catch(e) {
		return require('./xlsx.js').utils;
	}
	throw new Error("Cannot find XLSX utils");
};
