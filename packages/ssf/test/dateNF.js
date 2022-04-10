/* vim: set ts=2: */
/*jshint loopfunc:true, mocha:true, node:true */
/*eslint-env mocha, node */
var SSF = require('../');
var assert = require('assert');
describe('dateNF override', function() {
  it('should override format code 14', function() {
    assert.equal(SSF.format(14, 43880), "2/19/20");
    assert.equal(SSF.format(14, 43880, {dateNF:"yyyy-mm-dd"}), "2020-02-19");
    assert.equal(SSF.format(14, 43880, {dateNF:"dd/mm/yyyy"}), "19/02/2020");
  });
  it('should override format "m/d/yy"', function() {
    assert.equal(SSF.format('m/d/yy', 43880), "2/19/20");
    assert.equal(SSF.format('m/d/yy', 43880, {dateNF:"yyyy-mm-dd"}), "2020-02-19");
    assert.equal(SSF.format('m/d/yy', 43880, {dateNF:"dd/mm/yyyy"}), "19/02/2020");
  });
});
describe('asian formats', function() {
	it('上午/下午 (AM/PM)', function() {
		assert.equal(SSF.format('上午/下午', 0),    '上午');
		assert.equal(SSF.format('上午/下午', 0.25), '上午');
		assert.equal(SSF.format('上午/下午', 0.49), '上午');
		assert.equal(SSF.format('上午/下午', 0.5),  '下午');
		assert.equal(SSF.format('上午/下午', 0.51), '下午');
		assert.equal(SSF.format('上午/下午', 0.99), '下午');
		assert.equal(SSF.format('上午/下午', 1),    '上午');
	});
	it('bb (buddhist)', function() {
		[
			[12345,
				[ 'yyyy',   '1933'],
				[ 'eeee',   '1933'],
				[ 'bbbb',   '2476'],
				//[ 'ปปปป',   '๒๔๗๖'],
				[ 'b2yyyy', '1352'],
				[ 'b2eeee', '1352'],
				[ 'b2bbbb', '1895']
				//[ 'b2ปปปป', '๑๘๙๕']
			]
		].forEach(function(row) {
			row.slice(1).forEach(function(fmt) {
				assert.equal(SSF.format(fmt[0], row[0]), fmt[1]);
			});
		});
	});
	it.skip('thai fields', function() {
		SSF.format('\u0E27/\u0E14/\u0E1B\u0E1B\u0E1B\u0E1B \u0E0A\u0E0A:\u0E19\u0E19:\u0E17\u0E17', 12345.67);
		assert.equal(SSF.format('\u0E27/\u0E14/\u0E1B\u0E1B\u0E1B\u0E1B \u0E0A\u0E0A:\u0E19\u0E19:\u0E17\u0E17', 12345.67), "๑๘/๑๐/๒๔๗๖ ๑๖:๐๔:๔๘");
	});
});
describe('meridiem', function() {
	it('should force am/pm uppercase', function() {
		for(var i = 0; i < 16; ++i) {
			var a = ((i & 8) ? "A" : "a") + ((i & 4) ? "M" : "m");
			var p = ((i & 2) ? "P" : "p") + ((i & 1) ? "M" : "m");
			var fmt = a + "/" + p;
			assert.equal(SSF.format(fmt, 0.25), "AM");
			assert.equal(SSF.format(fmt, 0.75), "PM");
		}
	});
	it('should use actual case for a/p', function() {
		for(var i = 0; i < 12; ++i) {
			var a = ((i & 2) ? "A" : "a");
			var p = ((i & 1) ? "P" : "p");
			var w = ["", 'm', 'M'][i>>2];
			var fmt = a + "/" + p + ["", '"m"', '"M"'][i>>2];
			assert.equal(SSF.format(fmt, 0.25), a + w);
			assert.equal(SSF.format(fmt, 0.75), p + w);
		}
	});
});