var X = require('xlsx');
var electron = require('electron').remote;

var process_wb = (function() {
	var HTMLOUT = document.getElementById('htmlout');

	return function process_wb(wb) {
		HTMLOUT.innerHTML = "";
		wb.SheetNames.forEach(function(sheetName) {
			var htmlstr = X.write(wb, {sheet:sheetName, type:'binary', bookType:'html'});
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
			process_wb(X.read(data, {type: 'array'}));
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
		if(o.length > 0) process_wb(X.readFile(o[0]));
	}
	readf.addEventListener('click', handleF, false);
})();

(function() {
	var xlf = document.getElementById('xlf');
	function handleFile(e) { do_file(e.target.files); }
	xlf.addEventListener('change', handleFile, false);
})();
