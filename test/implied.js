/* vim: set ts=2: */
var SSF = require('../');
var fs = require('fs'), assert = require('assert');
var data = JSON.parse(fs.readFileSync('./test/implied.json','utf8'));
describe('implied formats', function() {
  data.forEach(function(d) {
    it(d[1]+" for "+d[0], (d[1]<14||d[1]>22)?null:function(){
      assert.equal(SSF.format(d[1], d[0], {}), d[2]);
    });
  });
});
