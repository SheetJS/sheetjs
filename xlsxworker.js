/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* uncomment the next line for encoding support */
importScripts('dist/cpexcel.js');
importScripts('jszip.js');
importScripts('xlsx.js');
postMessage({t:"ready"});

onmessage = function (evt) {
  var v;
  try {
    v = XLSX.read(evt.data.d, {type: evt.data.b});
postMessage({t:"xlsx", d:JSON.stringify(v)});
  } catch(e) { postMessage({t:"e",d:e.stack||e}); }
};
