/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* eslint-env browser */
/*global ObjUtils */
Storage.prototype.load = function(data) {
  var self = this;
  Object.keys(data).forEach(function(k) {
    self.setItem(k, JSON.stringify(data[k]));
  });
};

Storage.prototype.dump = function() {
  var obj = {};
  for(var i = 0; i < this.length; ++i) {
    var key = this.key(i);
    obj[key] = JSON.parse(this.getItem(key));
  }
  return ObjUtils.object_to_workbook(obj);
};
