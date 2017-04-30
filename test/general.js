/* vim: set ts=2: */
/*jshint loopfunc:true, mocha:true, node:true */
var SSF = require('../');
var fs = require('fs'), assert = require('assert');
var data = JSON.parse(fs.readFileSync('./test/general.json','utf8'));
var skip = [];
describe('General format', function() {
  data.forEach(function(d) {
    it(d[1]+" for "+d[0], skip.indexOf(d[1]) > -1 ? null : function(){
      assert.equal(SSF.format(d[1], d[0], {}), d[2]);
    });
  });
  it('should handle special values', function() {
    assert.equal(SSF.format("General", true), "TRUE");
    assert.equal(SSF.format("General", undefined), "");
    assert.equal(SSF.format("General", null), "");
  });
});
