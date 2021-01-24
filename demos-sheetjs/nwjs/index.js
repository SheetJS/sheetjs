var fs = require('fs');

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
	var xlf = document.getElementById('xlf');
	function handleFile(e) { do_file(e.target.files); }
	xlf.addEventListener('change', handleFile, false);
})();

var export_xlsx = (function() {
	/* pre-build the nwsaveas input element */
	var HTMLOUT = document.getElementById('htmlout');
	var input = document.createElement('input');
	input.style.display = 'none';
	input.setAttribute('nwsaveas', 'sheetjs.xlsx');
	input.setAttribute('type', 'file');
	document.body.appendChild(input);
	input.addEventListener('cancel',function(){ alert("Save was canceled!"); });
	input.addEventListener('change',function(e){
		var filename=this.value, bookType=(filename.match(/[^\.]*$/)||["xlsx"])[0];
		var wb = XLSX.utils.table_to_book(HTMLOUT);
		var wbout = XLSX.write(wb, {type:'buffer', bookType:bookType});
		fs.writeFile(filename, wbout, function(err) {
			if(!err) return alert("Saved to " + filename);
			alert("Error: " + (err.message || err));
		});
	});

	return function() { input.click(); };
})();
