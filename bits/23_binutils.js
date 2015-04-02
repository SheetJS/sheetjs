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
var __utf16le, ___utf16le;
__utf16le = ___utf16le = function utf16le_(b,s,e) { var ss=[]; for(var i=s; i<e; i+=2) ss.push(String.fromCharCode(__readUInt16LE(b,i))); return ss.join(""); };
var __hexlify, ___hexlify;
__hexlify = ___hexlify = function hexlify_(b,s,l) { return b.slice(s,(s+l)).map(function(x){return (x<16?"0":"") + x.toString(16);}).join(""); };
var __utf8, ___utf8;
__utf8 = ___utf8 = function(b,s,e) { var ss=[]; for(var i=s; i<e; i++) ss.push(String.fromCharCode(__readUInt8(b,i))); return ss.join(""); };
var __lpstr, ___lpstr;
__lpstr = ___lpstr = function lpstr_(b,i) { var len = __readUInt32LE(b,i); return len > 0 ? __utf8(b, i+4,i+4+len-1) : "";};
var __lpwstr, ___lpwstr;
__lpwstr = ___lpwstr = function lpwstr_(b,i) { var len = 2*__readUInt32LE(b,i); return len > 0 ? __utf8(b, i+4,i+4+len-1) : "";};
var __double, ___double;
__double = ___double = function(b, idx) { return readIEEE754(b, idx);};

var is_buf = function is_buf_a(a) { return Array.isArray(a); };
if(has_buf) {
	__utf16le = function utf16le_b(b,s,e) { if(!Buffer.isBuffer(b)) return ___utf16le(b,s,e); return b.toString('utf16le',s,e); };
	__hexlify = function(b,s,l) { return Buffer.isBuffer(b) ? b.toString('hex',s,s+l) : ___hexlify(b,s,l); };
	__lpstr = function lpstr_b(b,i) { if(!Buffer.isBuffer(b)) return ___lpstr(b, i); var len = b.readUInt32LE(i); return len > 0 ? b.toString('utf8',i+4,i+4+len-1) : "";};
	__lpwstr = function lpwstr_b(b,i) { if(!Buffer.isBuffer(b)) return ___lpwstr(b, i); var len = 2*b.readUInt32LE(i); return b.toString('utf16le',i+4,i+4+len-1);};
	__utf8 = function utf8_b(s,e) { return this.toString('utf8',s,e); };
	__toBuffer = function(bufs) { return (bufs[0].length > 0 && Buffer.isBuffer(bufs[0][0])) ? Buffer.concat(bufs[0]) : ___toBuffer(bufs);};
	bconcat = function(bufs) { return Buffer.isBuffer(bufs[0]) ? Buffer.concat(bufs) : [].concat.apply([], bufs); };
	__double = function double_(b,i) { if(Buffer.isBuffer(b)) return b.readDoubleLE(i); return ___double(b,i); };
	is_buf = function is_buf_b(a) { return Buffer.isBuffer(a) || Array.isArray(a); };
}

/* from js-xls */
if(typeof cptable !== 'undefined') {
	__utf16le = function(b,s,e) { return cptable.utils.decode(1200, b.slice(s,e)); };
	__utf8 = function(b,s,e) { return cptable.utils.decode(65001, b.slice(s,e)); };
	__lpstr = function(b,i) { var len = __readUInt32LE(b,i); return len > 0 ? cptable.utils.decode(current_codepage, b.slice(i+4, i+4+len-1)) : "";};
	__lpwstr = function(b,i) { var len = 2*__readUInt32LE(b,i); return len > 0 ? cptable.utils.decode(1200, b.slice(i+4,i+4+len-1)) : "";};
}

var __readUInt8 = function(b, idx) { return b[idx]; };
var __readUInt16LE = function(b, idx) { return b[idx+1]*(1<<8)+b[idx]; };
var __readInt16LE = function(b, idx) { var u = b[idx+1]*(1<<8)+b[idx]; return (u < 0x8000) ? u : (0xffff - u + 1) * -1; };
var __readUInt32LE = function(b, idx) { return b[idx+3]*(1<<24)+(b[idx+2]<<16)+(b[idx+1]<<8)+b[idx]; };
var __readInt32LE = function(b, idx) { return (b[idx+3]<<24)|(b[idx+2]<<16)|(b[idx+1]<<8)|b[idx]; };

var ___unhexlify = function(s) { return s.match(/../g).map(function(x) { return parseInt(x,16);}); };
var __unhexlify = typeof Buffer !== "undefined" ? function(s) { return Buffer.isBuffer(s) ? new Buffer(s, 'hex') : ___unhexlify(s); } : ___unhexlify;

function ReadShift(size, t) {
	var o="", oI, oR, oo=[], w, vv, i, loc;
	switch(t) {
		case 'dbcs':
			loc = this.l;
			if(has_buf && Buffer.isBuffer(this)) o = this.slice(this.l, this.l+2*size).toString("utf16le");
			else for(i = 0; i != size; ++i) { o+=String.fromCharCode(__readUInt16LE(this, loc)); loc+=2; }
			size *= 2;
			break;

		case 'utf8': o = __utf8(this, this.l, this.l + size); break;
		case 'utf16le': size *= 2; o = __utf16le(this, this.l, this.l + size); break;

		/* [MS-OLEDS] 2.1.4 LengthPrefixedAnsiString */
		case 'lpstr': o = __lpstr(this, this.l); size = 5 + o.length; break;
		/* [MS-OLEDS] 2.1.5 LengthPrefixedUnicodeString */
		case 'lpwstr': o = __lpwstr(this, this.l); size = 5 + o.length; if(o[o.length-1] == '\u0000') size += 2; break;

		case 'cstr': size = 0; o = "";
			while((w=__readUInt8(this, this.l + size++))!==0) oo.push(_getchar(w));
			o = oo.join(""); break;
		case 'wstr': size = 0; o = "";
			while((w=__readUInt16LE(this,this.l +size))!==0){oo.push(_getchar(w));size+=2;}
			size+=2; o = oo.join(""); break;

		/* sbcs and dbcs support continue records in the SST way TODO codepages */
		case 'dbcs-cont': o = ""; loc = this.l;
			for(i = 0; i != size; ++i) {
				if(this.lens && this.lens.indexOf(loc) !== -1) {
					w = __readUInt8(this, loc);
					this.l = loc + 1;
					vv = ReadShift.call(this, size-i, w ? 'dbcs-cont' : 'sbcs-cont');
					return oo.join("") + vv;
				}
				oo.push(_getchar(__readUInt16LE(this, loc)));
				loc+=2;
			} o = oo.join(""); size *= 2; break;

		case 'sbcs-cont': o = ""; loc = this.l;
			for(i = 0; i != size; ++i) {
				if(this.lens && this.lens.indexOf(loc) !== -1) {
					w = __readUInt8(this, loc);
					this.l = loc + 1;
					vv = ReadShift.call(this, size-i, w ? 'dbcs-cont' : 'sbcs-cont');
					return oo.join("") + vv;
				}
				oo.push(_getchar(__readUInt8(this, loc)));
				loc+=1;
			} o = oo.join(""); break;

		default:
	switch(size) {
		case 1: oI = __readUInt8(this, this.l); this.l++; return oI;
		case 2: oI = (t === 'i' ? __readInt16LE : __readUInt16LE)(this, this.l); this.l += 2; return oI;
		case 4:
			if(t === 'i' || (this[this.l+3] & 0x80)===0) { oI = __readInt32LE(this, this.l); this.l += 4; return oI; }
			else { oR = __readUInt32LE(this, this.l); this.l += 4; return oR; } break;
		case 8: if(t === 'f') { oR = __double(this, this.l); this.l += 8; return oR; }
		/* falls through */
		case 16: o = __hexlify(this, this.l, size); break;
	}}
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

function CheckField(hexstr, fld) {
	var m = __hexlify(this,this.l,hexstr.length>>1);
	if(m !== hexstr) throw fld + 'Expected ' + hexstr + ' saw ' + m;
	this.l += hexstr.length>>1;
}

function prep_blob(blob, pos) {
	blob.l = pos;
	blob.read_shift = ReadShift;
	blob.chk = CheckField;
	blob.write_shift = WriteShift;
}

function parsenoop(blob, length) { blob.l += length; }

function writenoop(blob, length) { blob.l += length; }

function new_buf(sz) {
	var o = new_raw_buf(sz);
	prep_blob(o, 0);
	return o;
}

