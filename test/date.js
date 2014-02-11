/* vim: set ts=2: */
/*jshint loopfunc:true */
var SSF = require('../');
var fs = require('fs'), assert = require('assert');
var dates = fs.readFileSync('./test/dates.tsv','utf8').split("\n");
var times = fs.readFileSync('./test/times.tsv','utf8').split("\n");

function doit(data) {
  var step = Math.ceil(data.length/100), i = 1;
  var headers = data[0].split("\t");
  for(j=0;j<=100;++j) it(j, function() {
    for(var k = 0; k <= step; ++k,++i) {
      if(!data[i]) return;
      var d = data[i].replace(/#{255}/g,"").split("\t");
      for(var w = 1; w < headers.length; ++w) {
        var expected = d[w], actual = SSF.format(headers[w], Number(d[0]), {});
        if(actual != expected) throw [actual, expected, w, headers[w],d[0],d].join("|");
        actual = SSF.format(headers[w].toUpperCase(), Number(d[0]), {});
        if(actual != expected) throw [actual, expected, w, headers[w],d[0],d].join("|");
      }
    }
  });
}
describe('time formats', function() { doit(times.slice(0,1000)); });
describe('date formats', function() {
  doit(process.env.MINTEST ? dates.slice(0,1000) : dates);
  it('should fail for bad formats', function() {
    var bad = [];
    var chk = function(fmt){ return function(){ SSF.format(fmt,0); }; };
    bad.forEach(function(fmt){assert.throws(chk(fmt));});
  });
});
