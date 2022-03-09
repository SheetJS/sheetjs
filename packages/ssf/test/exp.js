/* vim: set ts=2: */
/*jshint loopfunc:true, mocha:true, node:true */
var SSF = require('../');
var fs = require('fs');
var data = fs.readFileSync('./test/exp.tsv','utf8').split("\n");
function doit(d, headers) {
  it(d[0], function() {
    for(var w = 1; w < headers.length; ++w) {
      var expected = d[w].replace("|", ""), actual;
      try { actual = SSF.format(headers[w], parseFloat(d[0]), {}); } catch(e) { }
      if(actual != expected && d[w].charAt(0) !== "|") throw new Error([actual, expected, w, headers[w],d[0],d].join("|"));
    }
  });
}
describe('exponential formats', function() {
  var headers = data[0].split("\t");
  for(var j=1;j<data.length;++j) {
    if(!data[j]) return;
    doit(data[j].replace(/#{255}/g,"").split("\t"), headers);
  }
});
