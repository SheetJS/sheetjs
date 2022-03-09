/* vim: set ts=2: */
/*jshint loopfunc:true, mocha:true, node:true */
var SSF = require('../');
var fs = require('fs'), assert = require('assert');
var data = JSON.parse(fs.readFileSync('./test/implied.json','utf8'));
var skip = [];
function doit(d) {
  d[1].forEach(function(r){if(r.length === 2)assert.equal(SSF.format(r[0],d[0]),r[1]);});
}
describe('implied formats', function() {
  data.forEach(function(d) {
    if(d.length == 2) it(String(d[0]), function() { doit(d); });
    else it(d[1]+" for "+d[0], skip.indexOf(d[1]) > -1 ? null : function(){
      assert.equal(SSF.format(d[1], d[0], {}), d[2]);
    });
  });
});
