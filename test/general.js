/* vim: set ts=2: */
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
  it('should fail for undefined and null', function() {
    assert.throws(function() {
      SSF.format("General", undefined);
      SSF.format("General", null);
    });
  });
});
