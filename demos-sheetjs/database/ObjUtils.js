/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/*global XLSX, module, require */
var ObjUtils = (function() {

var X;
if(typeof XLSX !== "undefined") X = XLSX;
else if(typeof require !== 'undefined') X = require('xlsx');
else throw new Error("Could not find XLSX");

function walk(obj, key, arr) {
  if(Array.isArray(obj)) return;
  if(typeof obj != "object" || obj instanceof Date) { arr.push({path:key, value:obj}); return; }
  Object.keys(obj).forEach(function(k) {
    walk(obj[k], key ? key + "." + k : k, arr);
  });
}

function object_to_workbook(obj) {
  var wb = X.utils.book_new();

  var base = []; walk(obj, "", base);
  var ws = X.utils.json_to_sheet(base, {header:["path", "value"]});
  X.utils.book_append_sheet(wb, ws, "_keys");

  Object.keys(obj).forEach(function(k) {
    if(!Array.isArray(obj[k])) return;
    X.utils.book_append_sheet(wb, X.utils.json_to_sheet(obj[k]), k);
  });

  return wb;
}

function deepset(obj, path, value) {
  if(path.indexOf(".") == -1) return obj[path] = value;
  var parts = path.split(".");
  if(!obj[parts[0]]) obj[parts[0]] = {};
  return deepset(obj[parts[0]], parts.slice(1).join("."), value);
}
function workbook_set_object(obj, wb) {
  var ws = wb.Sheets["_keys"]; if(ws) {
    var data = X.utils.sheet_to_json(ws, {raw:true});
    data.forEach(function(r) { deepset(obj, r.path, r.value); });
  }
  wb.SheetNames.forEach(function(n) {
    if(n == "_keys") return;
    obj[n] = X.utils.sheet_to_json(wb.Sheets[n], {raw:true});
  });
}

function workbook_to_object(wb) { var obj = {}; workbook_set_object(obj, wb); return obj; }

return {
  workbook_to_object: workbook_to_object,
  object_to_workbook: object_to_workbook,
  workbook_set_object: workbook_set_object
};
})();

if(typeof module !== 'undefined') module.exports = ObjUtils;
