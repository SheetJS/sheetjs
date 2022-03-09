/* vim: set ts=2: */
/*jshint loopfunc:true, mocha:true, node:true */
var SSF = require('../');
var fs = require('fs');
var data = fs.readFileSync('./test/comma.tsv','utf8').split("\n");

function doit(w, headers) {
  it(headers[w], function() {
    for(var j=1;j<data.length;++j) {
      if(!data[j]) continue;
      var d = data[j].replace(/#{255}/g,"").split("\t");
      var expected = d[w].replace("|", ""), actual;
      try { actual = SSF.format(headers[w], Number(d[0]), {}); } catch(e) { }
      if(actual != expected && d[w][0] !== "|") throw new Error([actual, expected, w, headers[w],d[0],d].join("|"));
    }
  });
}
describe('comma formats', function() {
  var headers = data[0].split("\t");
  for(var w = 1; w < headers.length; ++w) doit(w, headers);
});
