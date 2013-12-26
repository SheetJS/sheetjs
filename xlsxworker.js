/* xlsx.js (C) 2013 SheetJS -- http://sheetjs.com */
importScripts('jszip.js');
importScripts('xlsx.js');
postMessage({t:"ready"});

onmessage = function (oEvent) {
  var v;
  try {
    //postMessage({t:'e', d:oEvent.data.b});
    v = XLSX.read(oEvent.data.d, {type: oEvent.data.b ? 'binary': 'base64'});
  } catch(e) { postMessage({t:"e",d:e.stack}); }
  postMessage({t:"xlsx", d:JSON.stringify(v)});
};
