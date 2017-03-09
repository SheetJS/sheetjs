/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* uncomment the next line for encoding support */
/*:: declare var XLSX: XLSXModule; */
/*:: declare var self: DedicatedWorkerGlobalScope; */
importScripts('dist/cpexcel.js');
importScripts('jszip.js');
importScripts('xlsx.js');
/* uncomment the next line for ODS support */
importScripts('dist/ods.js');
/*::self.*/postMessage({t:"ready"});

onmessage = function (oEvent) {
  var v;
  try {
    v = XLSX.read(oEvent.data.d, {type: oEvent.data.b ? 'binary' : 'base64'});
  } catch(e) { /*::self.*/postMessage({t:"e",d:e.stack||e}); }
  /*::self.*/
  postMessage({t:"xlsx", d:JSON.stringify(v)});
};
