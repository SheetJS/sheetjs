var has_buf = (typeof Buffer !== 'undefined' && typeof process !== 'undefined' && typeof process.versions !== 'undefined' && process.versions.node);

var Buffer_from = /*::(*/function(){}/*:: :any)*/;

if(typeof Buffer !== 'undefined') {
	var nbfs = !Buffer.from;
	if(!nbfs) try { Buffer.from("foo", "utf8"); } catch(e) { nbfs = true; }
	Buffer_from = nbfs ? function(buf, enc) { return (enc) ? new Buffer(buf, enc) : new Buffer(buf); } : Buffer.from.bind(Buffer);
	// $FlowIgnore
	if(!Buffer.alloc) Buffer.alloc = function(n) { return new Buffer(n); };
}

function new_raw_buf(len/*:number*/) {
	/* jshint -W056 */
	return has_buf ? Buffer.alloc(len) : new Array(len);
	/* jshint +W056 */
}

var s2a = function s2a(s/*:string*/)/*:any*/ {
	// $FlowIgnore
	if(has_buf) return Buffer_from(s, "binary");
	return s.split("").map(function(x/*:string*/)/*:number*/{ return x.charCodeAt(0) & 0xff; });
};

function s2ab(s/*:string*/)/*:any*/ {
	if(typeof ArrayBuffer === 'undefined') return s2a(s);
	var buf = new ArrayBuffer(s.length), view = new Uint8Array(buf);
	for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
	return buf;
}

function a2s(data/*:any*/)/*:string*/ {
	if(Array.isArray(data)) return data.map(_chr).join("");
	var o/*:Array<string>*/ = []; for(var i = 0; i < data.length; ++i) o[i] = _chr(data[i]); return o.join("");
}

function a2u(data/*:Array<number>*/)/*:Uint8Array*/ {
	if(typeof Uint8Array === 'undefined') throw new Error("Unsupported");
	return new Uint8Array(data);
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
