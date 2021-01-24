/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
importScripts('https://cdnjs.cloudflare.com/ajax/libs/systemjs/0.20.19/system.js');

SystemJS.config({
	meta: {
		'xlsx': {
			exports: 'XLSX' // <-- tell SystemJS to expose the XLSX variable
		}
	},
	map: {
		'xlsx': 'xlsx.full.min.js', // <-- make sure xlsx.full.min.js is in same dir
		'fs': '',     // <--|
		'crypto': '', // <--| suppress native node modules
		'stream': ''  // <--|
	}
});

onmessage = function(evt) {
	/* the real action is in the _cb function from xlsxworker.js */
	SystemJS.import('xlsxworker.js').then(function() { _cb(evt); });
};
