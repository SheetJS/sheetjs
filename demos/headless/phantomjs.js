/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
/* eslint-env phantomjs */
var XLSX = require('xlsx');

var page = require('webpage').create();
page.onConsoleMessage = function(msg) { console.log(msg); };

/* this code will be run in the page */
var code = [ "function(){",
  /* call table_to_book on first table */
  "var wb = XLSX.utils.table_to_book(document.body.getElementsByTagName('table')[0]);",

  /* generate XLSB file and return binary string */
  "return XLSX.write(wb, {type: 'binary', bookType: 'xlsb'});",
"}" ].join("");

page.open('https://sheetjs.com/demos/table', function() {
  console.log("Page Loaded");
  /* Load the browser script from the UNPKG CDN */
  page.includeJs("https://unpkg.com/xlsx/dist/xlsx.full.min.js", function() {
    /* Verify the page is loaded by logging the version number */
    var version = "function(){ console.log('Library Version:' + window.XLSX.version); }";
    page.evaluateJavaScript(version);

    /* The code will return a binary string */
    var bin = page.evaluateJavaScript(code);
    var workbook = XLSX.read(bin, {type: "binary"});
    console.log(XLSX.utils.sheet_to_csv(workbook.Sheets[workbook.SheetNames[0]]));

    /* XLSX.writeFile will not work here -- have to write manually */
    require("fs").write("phantomjs.xlsb", bin, "wb");
    phantom.exit();
  });
});

