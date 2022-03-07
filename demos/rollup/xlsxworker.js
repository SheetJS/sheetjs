/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
import { read } from 'xlsx';
postMessage({t:"ready"});

onmessage = function (evt) {
  var v;
  try {
    v = read(evt.data.d, {type: evt.data.b});
postMessage({t:"xlsx", d:JSON.stringify(v)});
  } catch(e) { postMessage({t:"e",d:e.stack||e}); }
};
