var XLSX = require('xlsx');
var electron = require('electron').remote;

var process_wb = (function() {
	var HTMLOUT = document.getElementById('htmlout');
	var XPORT = document.getElementById('xport');

	return function process_wb(wb) {
		XPORT.disabled = false;
		HTMLOUT.innerHTML = "";
		wb.SheetNames.forEach(function(sheetName) {
			var htmlstr = XLSX.utils.sheet_to_html(wb.Sheets[sheetName],{editable:true});
			HTMLOUT.innerHTML += htmlstr;
		});
	};
})();

var _gaq = _gaq || [];
_gaq.push(['_setAccount', 'UA-36810333-1']);
_gaq.push(['_trackPageview']);

(function() {
	var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
	ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
	var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
})();

var do_file = (function() {
	return function do_file(files) {
		var f = files[0];
		var reader = new FileReader();
		reader.onload = function(e) {
			var data = e.target.result;
			data = new Uint8Array(data);
			process_wb(XLSX.read(data, {type: 'array'}));
		};
		reader.readAsArrayBuffer(f);
	};
})();

(function() {
	var drop = document.getElementById('drop');

	function handleDrop(e) {
		e.stopPropagation();
		e.preventDefault();
		do_file(e.dataTransfer.files);
	}

	function handleDragover(e) {
		e.stopPropagation();
		e.preventDefault();
		e.dataTransfer.dropEffect = 'copy';
	}

	drop.addEventListener('dragenter', handleDragover, false);
	drop.addEventListener('dragover', handleDragover, false);
	drop.addEventListener('drop', handleDrop, false);
})();

(function() {
	var readf = document.getElementById('readf');
	function handleF(e) {
		var o = electron.dialog.showOpenDialog({
			title: 'Select a file',
			filters: [{
				name: "Spreadsheets",
				extensions: "xls|xlsx|xlsm|xlsb|xml|xlw|xlc|csv|txt|dif|sylk|slk|prn|ods|fods|uos|dbf|wks|123|wq1|qpw|htm|html".split("|")
			}],
			properties: ['openFile']
		});
		if(o.length > 0) process_wb(XLSX.readFile(o[0]));
	}
	readf.addEventListener('click', handleF, false);
})();

(function() {
	var xlf = document.getElementById('xlf');
	function handleFile(e) { do_file(e.target.files); }
	xlf.addEventListener('change', handleFile, false);
})();

var export_xlsx = (function() {
	var HTMLOUT = document.getElementById('htmlout');
	var XTENSION = "xls|xlsx|xlsm|xlsb|xml|csv|txt|dif|sylk|slk|prn|ods|fods|htm|html".split("|")
	return function() {
		var wb = XLSX.utils.table_to_book(HTMLOUT);
		var o = electron.dialog.showSaveDialog({
			title: 'Save file as',
			filters: [{
				name: "Spreadsheets",
				extensions: XTENSION
			}]
		});
		console.log(o);
		XLSX.writeFile(wb, o);
		electron.dialog.showMessageBox({ message: "Exported data to " + o, buttons: ["OK"] });
	};
})();
