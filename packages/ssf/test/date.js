/* vim: set ts=2: */
/*jshint -W041 */
/*jshint loopfunc:true, mocha:true, node:true */
var SSF = require('../');
var fs = require('fs'), assert = require('assert');
var dates = fs.readFileSync('./test/dates.tsv','utf8').split("\n");
var date2 = fs.readFileSync('./test/cal.tsv',  'utf8').split("\n");
var times = fs.readFileSync('./test/times.tsv','utf8').split("\n");
function doit(data) {
  var step = Math.ceil(data.length/100), i = 1;
  var headers = data[0].split("\t");
  for(var j = 0; j <= 100; ++j) it(String(j), function() {
    for(var k = 0; k <= step; ++k,++i) {
      if(data[i] == null || data[i].length < 3) return;
      var d = data[i].replace(/#{255}/g,"").split("\t");
      for(var w = 1; w < headers.length; ++w) {
        var expected = d[w], actual = SSF.format(headers[w], parseFloat(d[0]), {});
        if(actual != expected) throw new Error([actual, expected, w, headers[w],d[0],d,i].join("|"));
        actual = SSF.format(headers[w].toUpperCase(), parseFloat(d[0]), {});
        if(actual != expected) throw new Error([actual, expected, w, headers[w].toUpperCase(),d[0],d,i].join("|"));
      }
    }
  });
}

describe('time formats', function() {
  doit(process.env.MINTEST ? times.slice(0,4000) : times);
});

describe('date formats', function() {
  doit(process.env.MINTEST ? dates.slice(0,4000) : dates);
  if(0) doit(process.env.MINTEST ? date2.slice(0,1000) : date2);
  it('should fail for bad formats', function() {
    var bad = [];
    var chk = function(fmt){ return function(){ SSF.format(fmt,0); }; };
    bad.forEach(function(fmt){assert.throws(chk(fmt));});
  });
});
