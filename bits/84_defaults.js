function fixopts(opts) {
	var defaults = [
		['cellNF', false], /* emit cell number format string as .z */

		['sheetStubs', true], /* emit empty cells */

		['WTF', false] /* WTF mode (do not use) */
	];
	defaults.forEach(function(d) { if(typeof opts[d[0]] === 'undefined') opts[d[0]] = d[1]; });
}
