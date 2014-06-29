var describe = function(m,cb){if(cb) cb();};
describe.skip = function(m,cb){};
var it = function(m,cb){if(cb) cb();};
it.skip = function(m,cb){};
var before = function(cb){if(cb) cb();};
