/* vim: set ts=2: */
/*jshint loopfunc:true, mocha:true, node:true */
var SSF = require('../');
var fs = require('fs'), assert = require('assert');
var is_date = JSON.parse(fs.readFileSync('./test/is_date.json','utf8'));
describe('utilities', function() {
  it('correctly determines if formats are dates', function() {
		is_date.forEach(function(d) { assert.equal(SSF.is_date(d[0]), d[1], d[0]); });
  });
});
