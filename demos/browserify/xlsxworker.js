/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
var XLSX = require('../../'); // test against development version
//var XLSX = require('xlsx'); // use in production
postMessage({t:"ready"});

onmessage = function (evt) {
  var v;
  try {
    v = XLSX.read(evt.data.d, {type: evt.data.b});
postMessage({t:"xlsx", d:JSON.stringify(v)});
  } catch(e) { postMessage({t:"e",d:e.stack||e}); }
};
