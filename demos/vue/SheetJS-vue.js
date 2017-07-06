var SheetJSFT = [
	"xlsx", "xlsb", "xlsm", "xls", "xml", "csv", "txt", "ods", "fods", "uos", "sylk", "dif", "dbf", "prn", "qpw", "123", "wb*", "wq*", "html", "htm"
].map(function(x) { return "." + x; }).join(",");

var SJSTemplate = [
	'<div>',
		'<input type="file" multiple="false" id="sheetjs-input" accept="' + SheetJSFT + '" @change="onchange" />',
		'<br/>',
		'<button type="button" id="expor-table" style="visibility:hidden" @click="onexport">Export to XLSX</button>',
		'<br/>',
		'<div id="out-table"></div>',
	'</div>'
].join("");

function s2ab(s) {
	if(typeof ArrayBuffer !== 'undefined') {
		var buf = new ArrayBuffer(s.length);
		var view = new Uint8Array(buf);
		for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
		return buf;
	} else {
		var buf = new Array(s.length);
		for (var i=0; i!=s.length; ++i) buf[i] = s.charCodeAt(i) & 0xFF;
		return buf;
	}
}

Vue.component('html-preview', {
	template: SJSTemplate,
	methods: {
		// as per: https://github.com/SheetJS/js-xlsx/wiki/Reading-XLSX-from-FileReader.readAsArrayBuffer()
		// changing the readAsBinaryString (deprecated) to readAsArrayBuffer()
		onchange: function(evt) {
			var file;
			var files = evt.target.files;

			if (!files || files.length == 0) return;

			file = files[0];
			
			var reader = new FileReader();
			reader.onload = function (e) {
				// pre-process data
				var binary = "";
				var bytes = new Uint8Array(e.target.result);
				var length = bytes.byteLength;
				for (var i = 0; i < length; i++) {
					binary += String.fromCharCode(bytes[i]);
				}
				
				/* read workbook */
				var wb = XLSX.read(binary, {type: 'binary'});

				/* grab first sheet */
				var wsname = wb.SheetNames[0];
				var ws = wb.Sheets[wsname];

				/* generate HTML */
				var HTML = XLSX.utils.sheet_to_html(ws); 

				/* update table */
				document.getElementById('out-table').innerHTML = HTML;
				/* show export button */
				document.getElementById('expor-table').style.visibility = "visible";
			};

			reader.readAsArrayBuffer(file);
		},
		onexport: function(evt) {
			/* generate workbook object from table */
			var wb = XLSX.utils.table_to_book(document.getElementById('out-table'));
			/* get binary string as output */
			var wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });

			/* force a download */
			saveAs(new Blob([s2ab(wbout)], { type: 'application/octet-stream' }), "sheetjs.xlsx");
		}
	}
});
