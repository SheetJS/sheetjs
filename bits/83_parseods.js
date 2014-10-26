/* Helper function to call out to ODS parser */
function parse_ods(zip, opts) {
	if(typeof module !== "undefined" && typeof require !== 'undefined' && typeof ODS === 'undefined') ODS = require('./od' + 's');
	if(typeof ODS === 'undefined' || !ODS.parse_ods) throw new Error("Unsupported ODS");
	return ODS.parse_ods(zip, opts);
}
