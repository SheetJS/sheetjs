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
      var row = data[i].replace(/#{255}/g,"").split("\t");
      testRow(row, headers, {});
    }
  });
}

function testRow(row, headers, opts) {
  for(var w = 1; w < headers.length; ++w) {
    var expected = row[w], actual = SSF.format(headers[w], parseFloat(row[0]), opts);
    if(actual != expected) throw new Error([actual, expected, w, headers[w],row[0],row].join("|"));
    actual = SSF.format(headers[w].toUpperCase(), parseFloat(row[0]), opts);
    if(actual != expected) throw new Error([actual, expected, w, headers[w].toUpperCase(),row[0],row].join("|"));
  }
}

describe('time formats', function() {
  doit(process.env.MINTEST ? times.slice(0,4000) : times);
});

if(false) describe('time format rounding', function() {
  var headers=['value', 'yyyy mmm ddd dd hh:mm:ss'];
  var testCases = [
    {desc: "rounds up to 1 minute", value: "0.00069", date1904: {"false": "1900 Jan Sat 00 00:01:00", "true": "1904 Jan Fri 01 00:01:00"}},
    {desc: "rounds up to 2 munutes", value: "0.001388", date1904: {"false": "1900 Jan Sat 00 00:02:00", "true": "1904 Jan Fri 01 00:02:00"}},
    {desc: "rounds up to 10 minutes", value: "0.00694", date1904: {"false": "1900 Jan Sat 00 00:10:00", "true": "1904 Jan Fri 01 00:10:00"}},
    {desc: "rounds up to 2 hours", value: "0.08333", date1904: {"false": "1900 Jan Sat 00 02:00:00", "true": "1904 Jan Fri 01 02:00:00"}},
    {desc: "rounds up day", value: "0.999999", date1904: {"false": "1900 Jan Sun 01 00:00:00", "true": "1904 Jan Sat 02 00:00:00"}},
    {desc: "rounds up month", value: "31.999999", date1904: {"false": "1900 Feb Wed 01 00:00:00", "true": "1904 Feb Tue 02 00:00:00"}},
    {desc: "rounds up to 1900 leap day", value: "59.999999", date1904: {"false": "1900 Feb Wed 29 00:00:00", "true": "1904 Mar Tue 01 00:00:00"}},
    {desc: "rounds 1900 leap day up", value: "60.999999", date1904: {"false": "1900 Mar Thu 01 00:00:00", "true": "1904 Mar Wed 02 00:00:00"}},
    {desc: "rounds up day in March", value: "77.999999", date1904: {"false": "1900 Mar Sun 18 00:00:00", "true": "1904 Mar Sat 19 00:00:00"}},
    {desc: "rounds up leap year 1900", value: "366.999999", date1904: {"false": "1901 Jan Tue 01 00:00:00", "true": "1905 Jan Mon 02 00:00:00"}},
    {desc: "rounds up leap year 1904", value: "365.999999", date1904: {"false": "1900 Dec Mon 31 00:00:00", "true": "1905 Jan Sun 01 00:00:00"}}
  ];
  [{date1904: true}, {date1904: false}].forEach(function(opts) {
    testCases.forEach(function(testCase) {
      it(testCase.desc + " (1904: " + opts.date1904 + ")", function() {
        testRow([testCase.value, testCase.date1904[String(opts.date1904)]], headers, opts);
      });
    });
  });
});

describe('time format precision rounding', function() {
  var value = "4018.99999998843";
  var testCases = [
    {desc: "end-of-year thousandths rounding", format: "mm/dd/yyyy hh:mm:ss.000", expected: "12/31/1910 23:59:59.999"},
    //{desc: "end-of-year hundredths round up", format: "mm/dd/yyyy hh:mm:ss.00", expected: "01/01/1911 00:00:00.00"},
    //{desc: "end-of-year minutes round up", format: "mm/dd/yyyy hh:mm", expected: "01/01/1911 00:00"},
    {desc: "hour duration thousandths rounding", format: "[hh]:mm:ss.000", expected: "96455:59:59.999"},
    //{desc: "hour duration hundredths round up", format: "[hh]:mm:ss.00", expected: "96456:00:00.00"},
    //{desc: "hour duration minute round up (w/ ss)", format: "[hh]:mm:ss", expected: "96456:00:00"},
    //{desc: "hour duration minute round up", format: "[hh]:mm", expected: "96456:00"},
    {desc: "hour duration round up", format: "[hh]", expected: "96456"},
    {desc: "minute duration thousandths rounding", format: "[mm]:ss.000", expected: "5787359:59.999"},
    //{desc: "minute duration hundredths round up", format: "[mm]:ss.00", expected: "5787360:00.00"},
    //{desc: "minute duration round up", format: "[mm]:ss", expected: "5787360:00"},
    //{desc: "second duration thousandths rounding", format: "[ss].000", expected: "347241599.999"},
    {desc: "second duration hundredths round up", format: "[ss].00", expected: "347241600.00"},
    {desc: "second duration round up", format: "[ss]", expected: "347241600"}
  ];
  testCases.forEach(function(testCase) {
    var headers = ["value", testCase.format];
    it(testCase.desc, function() { testRow([value, testCase.expected], headers, {}); });
  });
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
