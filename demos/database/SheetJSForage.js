/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/*global ObjUtils, localforage */
localforage.load = async function foo(data) {
  var keys = Object.keys(data);
  for(var i = 0; i < keys.length; ++i) {
    var key = keys[i], val = JSON.stringify(data[keys[i]])
    await localforage.setItem(key, val);
  }
};

localforage.dump = async function() {
  var obj = {};
  var length = await localforage.length();
  for(var i = 0; i < length; ++i) {
    var key = await this.key(i);
    var val = await this.getItem(key);
    obj[key] = JSON.parse(val);
  }
  return ObjUtils.object_to_workbook(obj);
};
