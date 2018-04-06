/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* eslint-env browser */
/* global XLSX, chrome */
chrome.runtime.onInstalled.addListener(function() {
	chrome.contextMenus.create({
		type: "normal",
		id: "sjsexport",
		title: "Export Table to XLSX",
		contexts: ["page", "selection"]
	});
	chrome.contextMenus.create({
		type: "normal",
		id: "sj5export",
		title: "Export All Tables in Page",
		contexts: ["page", "selection"]
	});
	chrome.contextMenus.onClicked.addListener(function(info/*, tab*/) {
		var mode = "";
		switch(info.menuItemId) {
			case 'sjsexport': mode = "JS"; break;
			case 'sj5export': mode = "J5"; break;
			default: return;
		}
		chrome.tabs.query({active: true, currentWindow: true}, function(tabs){
			chrome.tabs.sendMessage(tabs[0].id, {Sheet:mode}, sjsexport_cb);
		});
	});

	chrome.contextMenus.create({
		id: "sjsabout",
		title: "About",
		contexts: ["browser_action"]
	});
	chrome.contextMenus.onClicked.addListener(function(info/*, tab*/) {
		if(info.menuItemId !== "sjsabout") return;
		chrome.tabs.create({url: "https://sheetjs.com/"});
	});
});

function sjsexport_cb(wb) {
	if(!wb || !wb.SheetNames || !wb.Sheets) { return alert("Error in exporting table"); }
	XLSX.writeFile(wb, "export.xlsx");
}
