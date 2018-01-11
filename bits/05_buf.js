var has_buf = (typeof Buffer !== 'undefined' && typeof process !== 'undefined' && typeof process.versions !== 'undefined' && process.versions.node);

function new_raw_buf(len/*:number*/) {
	/* jshint -W056 */
	// $FlowIgnore
	return new (has_buf ? Buffer : Array)(len);
	/* jshint +W056 */
}

function s2a(s/*:string*/) {
	if(has_buf) return new Buffer(s, "binary");
	return s.split("").map(function(x){ return x.charCodeAt(0) & 0xff; });
}

function s2ab(s/*:string*/) {
	if(typeof ArrayBuffer === 'undefined') return s2a(s);
	var buf = new ArrayBuffer(s.length), view = new Uint8Array(buf);
	for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
	return buf;
}

function arr2str(data/*:any*/)/*:string*/ {
	if(Array.isArray(data)) return data.map(_chr).join("");
	var o/*:Array<string>*/ = []; for(var i = 0; i < data.length; ++i) o[i] = _chr(data[i]); return o.join("");
}

function ab2a(data/*:ArrayBuffer|Uint8Array*/)/*:Array<number>*/ {
	if(typeof ArrayBuffer == 'undefined') throw new Error("Unsupported");
	if(data instanceof ArrayBuffer) return ab2a(new Uint8Array(data));
	/*:: if(data instanceof ArrayBuffer) throw new Error("unreachable"); */
	var o = new Array(data.length);
	for(var i = 0; i < data.length; ++i) o[i] = data[i];
	return o;
}

var bconcat = function(bufs) { return [].concat.apply([], bufs); };

var chr0 = /\u0000/g, chr1 = /[\u0001-\u0006]/g;
