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

var __toBuffer;
if(typeof Buffer !== "undefined") {
	Buffer.prototype.hexlify= function() { return this.toString('hex'); };
	__toBuffer = function(bufs) { return Buffer.concat(bufs[0]); };
} else {
	__toBuffer = function(bufs) {
		var x = [];
		for(var i = 0; i != bufs[0].length; ++i) { x = x.concat(bufs[0][i]); }
		return x;
	};
}

var __readUInt8 = function(b, idx) { return b.readUInt8 ? b.readUInt8(idx) : b[idx]; };
var __readUInt16LE = function(b, idx) { return b.readUInt16LE ? b.readUInt16LE(idx) : b[idx+1]*(1<<8)+b[idx]; };
var __readInt16LE = function(b, idx) { var u = __readUInt16LE(b,idx); if(!(u & 0x8000)) return u; return (0xffff - u + 1) * -1; };
var __readUInt32LE = function(b, idx) { return b.readUInt32LE ? b.readUInt32LE(idx) : b[idx+3]*(1<<24)+b[idx+2]*(1<<16)+b[idx+1]*(1<<8)+b[idx]; };
var __readInt32LE = function(b, idx) { if(b.readInt32LE) return b.readInt32LE(idx); var u = __readUInt32LE(b,idx); if(!(u & 0x80000000)) return u; return (0xffffffff - u + 1) * -1; };
var __readDoubleLE = function(b, idx) { return b.readDoubleLE ? b.readDoubleLE(idx) : readIEEE754(b, idx||0);};


function ReadShift(size, t) {
	var o = "", oo = [], w, vv, i, loc; t = t || 'u';
	if(size === 'ieee754') { size = 8; t = 'f'; }
	switch(size) {
		case 1: o = __readUInt8(this, this.l); break;
		case 2: o=(t==='u' ? __readUInt16LE : __readInt16LE)(this, this.l); break;
		case 4: o = __readUInt32LE(this, this.l); break;
		case 8: if(t === 'f') { o = __readDoubleLE(this, this.l); break; }
		/* falls through */
		case 16: o = this.toString('hex', this.l,this.l+size); break;

		/* sbcs and dbcs support continue records in the SST way TODO codepages */
		/* TODO: DBCS http://msdn.microsoft.com/en-us/library/cc194788.aspx */
		case 'dbcs': size = 2*t; loc = this.l;
			for(i = 0; i != t; ++i) {
				oo.push(_getchar(__readUInt16LE(this, loc)));
				loc+=2;
			} o = oo.join(""); break;

		case 'sbcs': size = t; o = ""; loc = this.l;
			for(i = 0; i != t; ++i) {
				o += _getchar(__readUInt8(this, loc));
				loc+=1;
			} break;
	}
	this.l+=size; return o;
}

function prep_blob(blob, pos) {
	blob.read_shift = ReadShift.bind(blob);
	blob.l = pos || 0;
	var read = ReadShift.bind(blob);
	return [read];
}

function parsenoop(blob, length) { blob.l += length; }
