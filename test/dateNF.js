/* vim: set ts=2: */
/*jshint loopfunc:true, mocha:true, node:true */
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
