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

function s2a(s) {
	if(typeof Buffer !== 'undefined') return new Buffer(s, "binary");
	var w = s.split("").map(function(x){ return x.charCodeAt(0) & 0xff; });
	return w;
}

var __toBuffer;
if(typeof Buffer !== "undefined") {
	Buffer.prototype.hexlify= function() { return this.toString('hex'); };
	Buffer.prototype.utf16le= function(s,e){return this.toString('utf16le',s,e).replace(/\u0000/,'').replace(/[\u0001-\u0006]/,'!');};
	Buffer.prototype.utf8 = function(s,e) { return this.toString('utf8',s,e); };
	Buffer.prototype.lpstr = function(i) { var len = this.readUInt32LE(i); return len > 0 ? this.utf8(i+4,i+4+len-1) : "";};
	Buffer.prototype.lpwstr = function(i) { var len = 2*this.readUInt32LE(i); return this.utf8(i+4,i+4+len-1);};
	if(typeof cptable !== "undefined") Buffer.prototype.lpstr = function(i) {
		var len = this.readUInt32LE(i);
		if(len === 0) return "";
		if(typeof current_cptable === "undefined") return this.utf8(i+4,i+4+len-1);
		var t = Array(this.slice(i+4,i+4+len-1));
		//1console.log("start", this.l, len, t);
		var c, j = i+4, o = "", cc;
		for(;j!=i+4+len;++j) {
			c = this.readUInt8(j);
			cc = current_cptable.dec[c];
			if(typeof cc === 'undefined') {
				c = c*256 + this.readUInt8(++j);
				cc = current_cptable.dec[c];
			}
			if(typeof cc === 'undefined') throw "Unrecognized character " + c.toString(16);
			if(c === 0) break;
			o += cc;
		//1console.log(cc, cc.charCodeAt(0), o, this.l);
		}
		return o;
	};
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

var __hexlify = function(b) { return b.map(function(x){return (x<16?"0":"") + x.toString(16);}).join(""); };

var __utf16le = function(b,s,e) { if(b.utf16le) return b.utf16le(s,e); var str = ""; for(var i=s; i<e; i+=2) str += String.fromCharCode(__readUInt16LE(b,i)); return str.replace(/\u0000/,'').replace(/[\u0001-\u0006]/,'!'); };

var __utf8 = function(b,s,e) { if(b.utf8) return b.utf8(s,e); var str = ""; for(var i=s; i<e; i++) str += String.fromCharCode(__readUInt8(b,i)); return str; };

var __lpstr = function(b,i) { if(b.lpstr) return b.lpstr(i); var len = __readUInt32LE(b,i); return len > 0 ? __utf8(b, i+4,i+4+len-1) : "";};
var __lpwstr = function(b,i) { if(b.lpwstr) return b.lpwstr(i); var len = 2*__readUInt32LE(b,i); return __utf8(b, i+4,i+4+len-1);};

function bconcat(bufs) { return (typeof Buffer !== 'undefined') ? Buffer.concat(bufs) : [].concat.apply([], bufs); }

function ReadShift(size, t) {
	var o, w, vv, i, loc; t = t || 'u';
	if(size === 'ieee754') { size = 8; t = 'f'; }
	switch(size) {
		case 1: o = __readUInt8(this, this.l); break;
		case 2: o=(t==='u' ? __readUInt16LE : __readInt16LE)(this, this.l); break;
		case 4: o = __readUInt32LE(this, this.l); break;
		case 8: if(t === 'f') { o = __readDoubleLE(this, this.l); break; }
		/* falls through */
		case 16: o = this.toString('hex', this.l,this.l+size); break;

		case 'utf8': size = t; o = __utf8(this, this.l, this.l + size); break;
		case 'utf16le': size=2*t; o = __utf16le(this, this.l, this.l + size); break;

		/* [MS-OLEDS] 2.1.4 LengthPrefixedAnsiString */
		case 'lpstr': o = __lpstr(this, this.l); size = 5 + o.length; break;

		case 'lpwstr': o = __lpwstr(this, this.l); size = 5 + o.length; if(o[o.length-1] == '\u0000') size += 2; break;

		/* sbcs and dbcs support continue records in the SST way TODO codepages */
		/* TODO: DBCS http://msdn.microsoft.com/en-us/library/cc194788.aspx */
		case 'dbcs': size = 2*t; o = ""; loc = this.l;
			for(i = 0; i != t; ++i) {
				if(this.lens && this.lens.indexOf(loc) !== -1) {
					w = __readUInt8(this, loc);
					this.l = loc + 1;
					vv = ReadShift.call(this, w ? 'dbcs' : 'sbcs', t-i);
					return o + vv;
				}
				o += _getchar(__readUInt16LE(this, loc));
				loc+=2;
			} break;

		case 'sbcs': size = t; o = ""; loc = this.l;
			for(i = 0; i != t; ++i) {
				if(this.lens && this.lens.indexOf(loc) !== -1) {
					w = __readUInt8(this, loc);
					this.l = loc + 1;
					vv = ReadShift.call(this, w ? 'dbcs' : 'sbcs', t-i);
					return o + vv;
				}
				o += _getchar(__readUInt8(this, loc));
				loc+=1;
			} break;

		case 'cstr': size = 0; o = "";
			while((w=__readUInt8(this, this.l + size++))!==0) o+= _getchar(w);
			break;
		case 'wstr': size = 0; o = "";
			while((w=__readUInt16LE(this,this.l +size))!==0){o+= _getchar(w);size+=2;}
			size+=2; break;
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
