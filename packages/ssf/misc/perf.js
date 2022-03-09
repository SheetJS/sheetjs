/* vim: set ts=2: */
/*jshint loopfunc:true */
var SSF = require('../');
var fs = require('fs')//, assert = require('assert');
var data = JSON.parse(fs.readFileSync('./test/oddities.json','utf8'));
var dates = fs.readFileSync('./test/dates.tsv','utf8').split("\n");
var date2 = fs.readFileSync('./test/cal.tsv',  'utf8').split("\n");
var times = fs.readFileSync('./test/times.tsv','utf8').split("\n");
function doit(data) {
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  var headers = data[0].split("\t");
  for(var k = 1; k <= data.length; ++k) {
    if(data[k] == null) return;
    var d = data[k].replace(/#{255}/g,"").split("\t");
    for(var w = 1; w < headers.length; ++w) {
      var expected = d[w], actual = SSF.format(headers[w], parseFloat(d[0]), {});
    }
  }
}

function testit() {
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  /* Prevent Optimization */
  doit(times.slice(0,4000));
  doit(dates.slice(0,4000));
    for(var i = 0; i != 1000; ++i) {
    for(var k = 0; k != data.length; ++k) {
      var d = data[k];
      for(var j=1;j<d.length;++j) {
        if(d[j].length == 2) {
          var expected = d[j][1], actual = SSF.format(d[0], d[j][0], {});
          //if(actual != expected) console.log(d[j]);
        }
      }
    }
  }
}

testit();
