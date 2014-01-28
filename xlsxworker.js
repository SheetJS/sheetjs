/* xlsx.js (C) 2013-2014 SheetJS -- http://sheetjs.com */
importScripts('jszip.js');
importScripts('xlsx.js');
postMessage({t:"ready"});

onmessage = function (oEvent) {
  var v;
  try {
    //v = XLSX.read(oEvent.data, {type: 'base64'});
    v = XLSX.read(oEvent.data.d, {type: oEvent.data.b ? 'binary': 'base64'});
  } catch(e) { postMessage({t:"e",d:e.stack}); }
  postMessage({t:"xlsx", d:JSON.stringify(v)});
};
