var has_buf = (typeof Buffer !== 'undefined');

function new_raw_buf(len) {
	/* jshint -W056 */
	return new (has_buf ? Buffer : Array)(len);
	/* jshint +W056 */
}

function s2a(s) {
	if(has_buf) return new Buffer(s, "binary");
	return s.split("").map(function(x){ return x.charCodeAt(0) & 0xff; });
}

var bconcat = function(bufs) { return [].concat.apply([], bufs); };

var chr0 = /\u0000/g, chr1 = /[\u0001-\u0006]/;
