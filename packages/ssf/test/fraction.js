/* vim: set ts=2: */
/*jshint loopfunc:true, mocha:true, node:true */
var SSF = require('../');
var fs = require('fs'), assert = require('assert');
var data = JSON.parse(fs.readFileSync('./test/fraction.json','utf8'));
var skip = [];
describe('fractional formats', function() {
  data.forEach(function(d) {
    it(d[1]+" for "+d[0], skip.indexOf(d[1]) > -1 ? null : function(){
      var expected = d[2], actual = SSF.format(d[1], d[0], {});
      //var r = actual.match(/(-?)\d* *\d+\/\d+/);
      assert.equal(actual, expected);
    });
  });
});
