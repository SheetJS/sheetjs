var get_utils = function() {
	if(typeof XLSX !== 'undefined') return XLSX.utils;
	if(typeof module !== "undefined" && typeof require !== 'undefined') return require('xl' + 'sx').utils;
	throw new Error("Cannot find XLSX utils");
};
