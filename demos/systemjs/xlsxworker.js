/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
var XLSX = require('xlsx');
postMessage({t:"ready"});

/* expose a global that can be accessed from the worker script */
_cb = function (evt) {
  var v;
  try {
    v = XLSX.read(evt.data.d, {type: evt.data.b});
postMessage({t:"xlsx", d:JSON.stringify(v)});
  } catch(e) { postMessage({t:"e",d:e.stack||e}); }
};
