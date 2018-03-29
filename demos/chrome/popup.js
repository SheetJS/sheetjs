/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* eslint-env browser */
/* global XLSX, chrome */
document.getElementById('sjsversion').innerText = "SheetJS " + XLSX.version;

document.getElementById('sjsversion').addEventListener('click', function() {
  chrome.tabs.create({url: "https://sheetjs.com/"}); return false;
});

/* recursively walk the bookmark tree */
function recurse_bookmarks(data, tree, path) {
  if(tree.url) data.push({Name: tree.title, Location: tree.url, Path:path});
  var T = path ? (path + "::" + tree.title) : tree.title;
  (tree.children||[]).forEach(function(C) { recurse_bookmarks(data, C, T); });
}

/* export bookmark data */
document.getElementById('sjsdownload').addEventListener('click', function() {
  chrome.bookmarks.getTree(function(res) {
    var data = [];
    res.forEach(function(t) { recurse_bookmarks(data, t, ""); });

    /* create worksheet */
    var ws = XLSX.utils.json_to_sheet(data, { header: ['Name', 'Location', 'Path'] });

    /* create workbook and export */
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Bookmarks');
    XLSX.writeFile(wb, "bookmarks.xlsx");
  });
});