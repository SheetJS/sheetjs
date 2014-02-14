function fixopts(opts) {
	var defaults = [
		['cellNF', false], /* emit cell number format string as .z */
		['cellHTML', true], /* emit html string as .h */
		['cellFormula', true], /* emit formulae as .h */

		['sheetStubs', false], /* emit empty cells */

		['bookSheets', false], /* only try to get sheet names (no Sheets) */
		['bookProps', false], /* only try to get properties (no Sheets) */

		['WTF', false] /* WTF mode (do not use) */
	];
	defaults.forEach(function(d) { if(typeof opts[d[0]] === 'undefined') opts[d[0]] = d[1]; });
}
