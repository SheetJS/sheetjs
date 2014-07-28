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
__toBuffer = ___toBuffer = function toBuffer_(bufs) { var x = []; for(var i = 0; i < bufs[0].length; ++i) { x.push.apply(x, bufs[0][i]); } return x; };
var __double, ___double;
__double = ___double = function(b, idx) { return readIEEE754(b, idx);};

var is_buf = function is_buf_a(a) { return Array.isArray(a); };
if(has_buf) {
	__toBuffer = function(bufs) { return (bufs[0].length > 0 && Buffer.isBuffer(bufs[0][0])) ? Buffer.concat(bufs[0]) : ___toBuffer(bufs);};
	__double = function double_(b,i) { if(Buffer.isBuffer(b)) return b.readDoubleLE(i); return ___double(b,i); };
	is_buf = function is_buf_b(a) { return Buffer.isBuffer(a) || Array.isArray(a); };
}


var __readUInt8 = function(b, idx) { return b[idx]; };
var __readUInt16LE = function(b, idx) { return b[idx+1]*(1<<8)+b[idx]; };
var __readInt16LE = function(b, idx) { var u = b[idx+1]*(1<<8)+b[idx]; return (u < 0x8000) ? u : (0xffff - u + 1) * -1; };
var __readUInt32LE = function(b, idx) { return b[idx+3]*(1<<24)+(b[idx+2]<<16)+(b[idx+1]<<8)+b[idx]; };
var __readInt32LE = function(b, idx) { return (b[idx+3]<<24)|(b[idx+2]<<16)|(b[idx+1]<<8)|b[idx]; };


function ReadShift(size, t) {
	var o="", oo=[], w, vv, i, loc;
	if(t === 'dbcs') {
		loc = this.l;
		if(has_buf && Buffer.isBuffer(this)) o = this.slice(this.l, this.l+2*size).toString("utf16le");
		else for(i = 0; i != size; ++i) { o+=String.fromCharCode(__readUInt16LE(this, loc)); loc+=2; }
		size *= 2;
	} else switch(size) {
		case 1: o = __readUInt8(this, this.l); break;
		case 2: o = (t === 'i' ? __readInt16LE : __readUInt16LE)(this, this.l); break;
		case 4: o = __readUInt32LE(this, this.l); break;
		case 8: if(t === 'f') { o = __double(this, this.l); break; }
	}
	this.l+=size; return o;
}

function WriteShift(t, val, f) {
	var size, i;
	if(f === 'dbcs') {
		for(i = 0; i != val.length; ++i) this.writeUInt16LE(val.charCodeAt(i), this.l + 2 * i);
		size = 2 * val.length;
	} else switch(t) {
		case  1: size = 1; this[this.l] = val&255; break;
		case  3: size = 3; this[this.l+2] = val & 255; val >>>= 8; this[this.l+1] = val&255; val >>>= 8; this[this.l] = val&255; break;
		case  4: size = 4; this.writeUInt32LE(val, this.l); break;
		case  8: size = 8; if(f === 'f') { this.writeDoubleLE(val, this.l); break; }
		/* falls through */
		case 16: break;
		case -4: size = 4; this.writeInt32LE(val, this.l); break;
	}
	this.l += size; return this;
}

function prep_blob(blob, pos) {
	blob.l = pos;
	blob.read_shift = ReadShift;
	blob.write_shift = WriteShift;
}

function parsenoop(blob, length) { blob.l += length; }

function writenoop(blob, length) { blob.l += length; }

function new_buf(sz) {
	var o = has_buf ? new Buffer(sz) : new Array(sz);
	prep_blob(o, 0);
	return o;
}

