/* vim: set ts=2: */
/*jshint loopfunc:true, mocha:true, node:true */
var SSF = require('../');
var fs = require('fs');
var data = fs.readFileSync('./test/valid.tsv','utf8').split("\n");
var _data = [0, 1, -2, 3.45, -67.89, "foo"];
function doit(d) {
  it(d[0], function() {
    for(var w = 0; w < _data.length; ++w) {
      SSF.format(d[0], _data[w]);
    }
  });
}
describe('valid formats', function() {
  for(var j=0;j<data.length;++j) {
    if(!data[j]) return;
    doit(data[j].replace(/#{255}/g,"").split("\t"));
  }
});
