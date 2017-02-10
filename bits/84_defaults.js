function fix_opts_func(defaults/*:Array<Array<any> >*/)/*:{(o:any):void}*/ {
	return function fix_opts(opts) {
		for(var i = 0; i != defaults.length; ++i) {
			var d = defaults[i];
			if(opts[d[0]] === undefined) opts[d[0]] = d[1];
			if(d[2] === 'n') opts[d[0]] = Number(opts[d[0]]);
		}
	};
}

var fix_read_opts = fix_opts_func([
	['cellNF', false], /* emit cell number format string as .z */
	['cellHTML', true], /* emit html string as .h */
	['cellFormula', true], /* emit formulae as .f */
	['cellStyles', false], /* emits style/theme as .s */
	['cellDates', false], /* emit date cells with type `d` */

	['sheetStubs', false], /* emit empty cells */
	['sheetRows', 0, 'n'], /* read n rows (0 = read all rows) */

	['bookDeps', false], /* parse calculation chains */
	['bookSheets', false], /* only try to get sheet names (no Sheets) */
	['bookProps', false], /* only try to get properties (no Sheets) */
	['bookFiles', false], /* include raw file structure (keys, files, cfb) */
	['bookVBA', false], /* include vba raw data (vbaraw) */

	['password',''], /* password */
	['WTF', false] /* WTF mode (throws errors) */
]);


var fix_write_opts = fix_opts_func([
	['cellDates', false], /* write date cells with type `d` */

	['bookSST', false], /* Generate Shared String Table */

	['bookType', 'xlsx'], /* Type of workbook (xlsx/m/b) */

	['compression', false], /* Use file compression */

	['WTF', false] /* WTF mode (throws errors) */
]);
