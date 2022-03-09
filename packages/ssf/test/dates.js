/* vim: set ts=2: */
/*jshint -W041 */
/*jshint loopfunc:true, mocha:true, node:true, evil:true */
var SSF = require('../');
var fs = require('fs'), assert = require('assert');
var data = JSON.parse(fs.readFileSync('./test/date.json','utf8'));

describe('date values', function() {
	it('should roundtrip dates', function() { data.forEach(function(d) {
		assert.equal(SSF.format("yyyy-mm-dd HH:MM:SS", eval(d[0]), {date1904:!!d[2]}), d[1]);
	}); });
});
