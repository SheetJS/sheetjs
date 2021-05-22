/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
var SheetJSFT = [
	"xlsx", "xlsb", "xlsm", "xls", "xml", "csv", "txt", "ods", "fods", "uos", "sylk", "dif", "dbf", "prn", "qpw", "123", "wb*", "wq*", "html", "htm"
].map(function(x) { return "." + x; }).join(",");

var SJSTemplate = [
	'<div>',
		'<input type="file" multiple="false" id="sheetjs-input" accept="' + SheetJSFT + '" @change="onchange" />',
		'<br/>',
		'<button type="button" id="export-table" style="visibility:hidden" @click="onexport">Export to XLSX</button>',
		'<br/>',
		'<div id="out-table"></div>',
	'</div>'
].join("");

Vue.component('html-preview', {
	template: SJSTemplate,
	methods: {
		async onchange (evt) {
			var file = evt.target.files[0];

			if (!file) return;

			var buffer = await file.arrayBuffer();
			/* read workbook */
			var wb = XLSX.read(buffer);

			/* grab first sheet */
			var wsname = wb.SheetNames[0];
			var ws = wb.Sheets[wsname];

			/* generate HTML */
			var HTML = XLSX.utils.sheet_to_html(ws);

			/* update table */
			document.getElementById('out-table').innerHTML = HTML;
			/* show export button */
			document.getElementById('export-table').style.visibility = "visible";
		},
		onexport (evt) {
			/* generate workbook object from table */
			var wb = XLSX.utils.table_to_book(document.getElementById('out-table'));
			/* generate file and force a download*/
			XLSX.writeFile(wb, "sheetjs.xlsx");
		}
	}
});
