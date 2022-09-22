var has_buf = /*#__PURE__*/(function() { return typeof Buffer !== 'undefined' && typeof process !== 'undefined' && typeof process.versions !== 'undefined' && !!process.versions.node; })();

var Buffer_from = /*#__PURE__*/(function() {
	if(typeof Buffer !== 'undefined') {
		var nbfs = !Buffer.from;
		if(!nbfs) try { Buffer.from("foo", "utf8"); } catch(e) { nbfs = true; }
		return nbfs ? function(buf, enc) { return (enc) ? new Buffer(buf, enc) : new Buffer(buf); } : Buffer.from.bind(Buffer);
	}
	return function() {};
})();
var buf_utf16le = /*#__PURE__*/(function() {
	if(typeof Buffer === 'undefined') return false;
	var x = Buffer_from([65,0]);
	if(!x) return false;
	var o = x.toString("utf16le");
	return o.length == 1;
})();


function new_raw_buf(len/*:number*/) {
	/* jshint -W056 */
	if(has_buf) return Buffer.alloc ? Buffer.alloc(len) : new Buffer(len);
	return typeof Uint8Array != "undefined" ? new Uint8Array(len) : new Array(len);
	/* jshint +W056 */
}

function new_unsafe_buf(len/*:number*/) {
	/* jshint -W056 */
	if(has_buf) return Buffer.allocUnsafe ? Buffer.allocUnsafe(len) : new Buffer(len);
	return typeof Uint8Array != "undefined" ? new Uint8Array(len) : new Array(len);
	/* jshint +W056 */
}

var s2a = function s2a(s/*:string*/)/*:any*/ {
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
	if(Array.isArray(data)) return data.map(function(c) { return String.fromCharCode(c); }).join("");
	var o/*:Array<string>*/ = []; for(var i = 0; i < data.length; ++i) o[i] = String.fromCharCode(data[i]); return o.join("");
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

var bconcat = has_buf ? function(bufs) { return Buffer.concat(bufs.map(function(buf) { return Buffer.isBuffer(buf) ? buf : Buffer_from(buf); })); } : function(bufs) {
	if(typeof Uint8Array !== "undefined") {
		var i = 0, maxlen = 0;
		for(i = 0; i < bufs.length; ++i) maxlen += bufs[i].length;
		var o = new Uint8Array(maxlen);
		var len = 0;
		for(i = 0, maxlen = 0; i < bufs.length; maxlen += len, ++i) {
			len = bufs[i].length;
			if(bufs[i] instanceof Uint8Array) o.set(bufs[i], maxlen);
			else if(typeof bufs[i] == "string") o.set(new Uint8Array(s2a(bufs[i])), maxlen);
			else o.set(new Uint8Array(bufs[i]), maxlen);
		}
		return o;
	}
	return [].concat.apply([], bufs.map(function(buf) { return Array.isArray(buf) ? buf : [].slice.call(buf); }));
};

function utf8decode(content/*:string*/) {
	var out = [], widx = 0, L = content.length + 250;
	var o = new_raw_buf(content.length + 255);
	for(var ridx = 0; ridx < content.length; ++ridx) {
		var c = content.charCodeAt(ridx);
		if(c < 0x80) o[widx++] = c;
		else if(c < 0x800) {
			o[widx++] = (192|((c>>6)&31));
			o[widx++] = (128|(c&63));
		} else if(c >= 0xD800 && c < 0xE000) {
			c = (c&1023)+64;
			var d = content.charCodeAt(++ridx)&1023;
			o[widx++] = (240|((c>>8)&7));
			o[widx++] = (128|((c>>2)&63));
			o[widx++] = (128|((d>>6)&15)|((c&3)<<4));
			o[widx++] = (128|(d&63));
		} else {
			o[widx++] = (224|((c>>12)&15));
			o[widx++] = (128|((c>>6)&63));
			o[widx++] = (128|(c&63));
		}
		if(widx > L) {
			out.push(o.slice(0, widx));
			widx = 0;
			o = new_raw_buf(65535);
			L = 65530;
		}
	}
	out.push(o.slice(0, widx));
	return bconcat(out);
}

var chr0 = /\u0000/g, chr1 = /[\u0001-\u0006]/g;
