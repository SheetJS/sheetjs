/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/*:: declare var XLSX: XLSXModule; */
/*:: declare var self: DedicatedWorkerGlobalScope; */
/* uncomment the next line for encoding support */
importScripts('dist/cpexcel.js');
importScripts('jszip.js');
importScripts('xlsx.js');
/*::self.*/postMessage({t:"ready"});

onmessage = function (oEvent) {
  var v;
  try {
    v = XLSX.read(oEvent.data.d, {type: oEvent.data.b});
    /*::self.*/postMessage({t:"xlsx", d:JSON.stringify(v)});
  } catch(e) { /*::self.*/postMessage({t:"e",d:e.stack||e}); }
};
