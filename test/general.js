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
  it('should handle dates', function() {
    assert.equal(SSF.format("General", new Date(2017, 1, 19)), "2/19/17");
    assert.equal(SSF.format("General", new Date(2017, 1, 19), {date1904:true}), "2/19/17");
    assert.equal(SSF.format("General", new Date(1901, 0, 1)), "1/1/01");
    if(SSF.format("General", new Date(1901, 0, 1), {date1904:true}) == "1/1/01") throw new Error("date1904 invalid date");
    assert.equal(SSF.format("General", new Date(1904, 0, 1)), "1/1/04");
    assert.equal(SSF.format("General", new Date(1904, 0, 1), {date1904:true}), "1/1/04");
  });
});
