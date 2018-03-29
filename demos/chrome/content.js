/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* eslint-env browser */
/* global XLSX, chrome */
var coords = [0,0];
document.addEventListener('mousedown', function(mouse) {
	if(mouse && mouse.button == 2) coords = [mouse.clientX, mouse.clientY];
});

chrome.runtime.onMessage.addListener(function(msg, sender, cb) {
	if(!msg && !msg['Sheet']) return;
	if(msg.Sheet == "JS") {
		var elt = document.elementFromPoint(coords[0], coords[1]);
		while(elt != null) {
			if(elt.tagName.toLowerCase() == "table") return cb(XLSX.utils.table_to_book(elt));
			elt = elt.parentElement;
		}
	} else if(msg.Sheet == "J5") {
		var tables = document.getElementsByTagName("table");
		var wb = XLSX.utils.book_new();
		for(var i = 0; i < tables.length; ++i) {
			var ws = XLSX.utils.table_to_sheet(tables[i]);
			XLSX.utils.book_append_sheet(wb, ws, "Table" + i);
		}
		return cb(wb);
	}
	cb(coords);
});
