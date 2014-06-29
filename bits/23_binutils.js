function readIEEE754(buf, idx, isLE, nl, ml) {
	if(isLE === undefined) isLE = true;
	if(!nl) nl = 8;
	if(!ml && nl === 8) ml = 52;
	var e, m, el = nl * 8 - ml - 1, eMax = (1 << el) - 1, eBias = eMax >> 1;
	var bits = -7, d = isLE ? -1 : 1, i = isLE ? (nl - 1) : 0, s = buf[idx + i];

	i += d;
	e = s & ((1 << (-bits)) - 1); s >>>= (-bits); bits += el;
	for (; bits > 0; e = e * 256 + buf[idx + i], i += d, bits -= 8);
	m = e & ((1 << (-bits)) - 1); e >>>= (-bits); bits += ml;
	for (; bits > 0; m = m * 256 + buf[idx + i], i += d, bits -= 8);
	if (e === eMax) return m ? NaN : ((s ? -1 : 1) * Infinity);
	else if (e === 0) e = 1 - eBias;
	else { m = m + Math.pow(2, ml); e = e - eBias; }
	return (s ? -1 : 1) * m * Math.pow(2, e - ml);
}

var __toBuffer, ___toBuffer;
__toBuffer = ___toBuffer = function(bufs) {
	var x = [];
	for(var i = 0; i != bufs[0].length; ++i) { x = x.concat(bufs[0][i]); }
	return x;
};
if(typeof Buffer !== "undefined") {
	__toBuffer = function(bufs) { return (bufs[0].length > 0 && Buffer.isBuffer(bufs[0][0])) ? Buffer.concat(bufs[0]) : ___toBuffer(bufs);};
}

var ___readUInt32LE = function(b, idx) { return b.readUInt32LE ? b.readUInt32LE(idx) : b[idx+3]*(1<<24)+(b[idx+2]<<16)+(b[idx+1]<<8)+b[idx]; };
var ___readInt32LE = function(b, idx) { return (b[idx+3]<<24)+(b[idx+2]<<16)+(b[idx+1]<<8)+b[idx]; };

var __readUInt8 = function(b, idx) { return b.readUInt8 ? b.readUInt8(idx) : b[idx]; };
var __readUInt16LE = function(b, idx) { return b.readUInt16LE ? b.readUInt16LE(idx) : b[idx+1]*(1<<8)+b[idx]; };
var __readInt16LE = function(b, idx) { var u = __readUInt16LE(b,idx); if(!(u & 0x8000)) return u; return (0xffff - u + 1) * -1; };
var __readUInt32LE = typeof Buffer !== "undefined" ? function(b, i) { return Buffer.isBuffer(b) ? b.readUInt32LE(i) : ___readUInt32LE(b,i); } : ___readUInt32LE;
var __readInt32LE = typeof Buffer !== "undefined" ? function(b, i) { return Buffer.isBuffer(b) ? b.readInt32LE(i) : ___readInt32LE(b,i); } : ___readInt32LE;
var __readDoubleLE = function(b, idx) { return b.readDoubleLE ? b.readDoubleLE(idx) : readIEEE754(b, idx||0);};


function ReadShift(size, t) {
	var o="", oo=[], w, vv, i, loc;
	if(t === 'dbcs') {
		loc = this.l;
		if(typeof Buffer !== 'undefined' && this instanceof Buffer) o = this.slice(this.l, this.l+2*size).toString("utf16le");
		else for(i = 0; i != size; ++i) { o+=String.fromCharCode(__readUInt16LE(this, loc)); loc+=2; }
		size *= 2;
	} else switch(size) {
		case 1: o = __readUInt8(this, this.l); break;
		case 2: o = (t === 'i' ? __readInt16LE : __readUInt16LE)(this, this.l); break;
		case 4: o = __readUInt32LE(this, this.l); break;
		case 8: if(t === 'f') { o = __readDoubleLE(this, this.l); break; }
	}
	this.l+=size; return o;
}

function WriteShift(t, val, f) {
	var size, i;
	if(f === 'dbcs') {
		for(i = 0; i != val.length; ++i) this.writeUInt16LE(val.charCodeAt(i), this.l + 2 * i);
		size = 2 * val.length;
	} else switch(t) {
		case  1: size = 1; this.writeUInt8(val, this.l); break;
		case  4: size = 4; this.writeUInt32LE(val, this.l); break;
		case  8: size = 8; if(f === 'f') { this.writeDoubleLE(val, this.l); break; }
		/* falls through */
		case 16: break;
		case -4: size = 4; this.writeInt32LE(val, this.l); break;
	}
	this.l += size; return this;
}

function prep_blob(blob, pos) {
	blob.l = pos || 0;
	blob.read_shift = ReadShift;
	blob.write_shift = WriteShift;
}

function parsenoop(blob, length) { blob.l += length; }

function writenoop(blob, length) { blob.l += length; }

function new_buf(sz) {
	var o = typeof Buffer !== 'undefined' ? new Buffer(sz) : new Array(sz);
	prep_blob(o, 0);
	return o;
}

function is_buf(a) { return (typeof Buffer !== 'undefined' && a instanceof Buffer) || Array.isArray(a); }
