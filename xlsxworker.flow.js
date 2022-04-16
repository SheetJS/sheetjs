/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/*:: declare var XLSX: XLSXModule; */
/*:: declare var self: DedicatedWorkerGlobalScope; */
importScripts('dist/shim.min.js');
/* uncomment the next line for encoding support */
importScripts('dist/cpexcel.js');
importScripts('xlsx.js');
/*::self.*/postMessage({t:"ready"});

onmessage = function (evt) {
  var v;
  try {
    v = XLSX.read(evt.data.d, {type: evt.data.b, codepage: evt.data.c});
    /*::self.*/postMessage({t:"xlsx", d:JSON.stringify(v)});
  } catch(e) { /*::self.*/postMessage({t:"e",d:e.stack||e}); }
};
