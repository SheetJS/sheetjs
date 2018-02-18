function read_double_le(b/*:RawBytes|CFBlob*/, idx/*:number*/)/*:number*/ {
	var s = 1 - 2 * (b[idx + 7] >>> 7);
	var e = ((b[idx + 7] & 0x7f) << 4) + ((b[idx + 6] >>> 4) & 0x0f);
	var m = (b[idx+6]&0x0f);
	for(var i = 5; i >= 0; --i) m = m * 256 + b[idx + i];
	if(e == 0x7ff) return m == 0 ? (s * Infinity) : NaN;
	if(e == 0) e = -1022;
	else { e -= 1023; m += Math.pow(2,52); }
	return s * Math.pow(2, e - 52) * m;
}

function write_double_le(b/*:RawBytes|CFBlob*/, v/*:number*/, idx/*:number*/) {
	var bs = ((((v < 0) || (1/v == -Infinity)) ? 1 : 0) << 7), e = 0, m = 0;
	var av = bs ? (-v) : v;
	if(!isFinite(av)) { e = 0x7ff; m = isNaN(v) ? 0x6969 : 0; }
	else if(av == 0) e = m = 0;
	else {
		e = Math.floor(Math.log(av) / Math.LN2);
		m = av * Math.pow(2, 52 - e);
		if((e <= -1023) && (!isFinite(m) || (m < Math.pow(2,52)))) { e = -1022; }
		else { m -= Math.pow(2,52); e+=1023; }
	}
	for(var i = 0; i <= 5; ++i, m/=256) b[idx + i] = m & 0xff;
	b[idx + 6] = ((e & 0x0f) << 4) | (m & 0xf);
	b[idx + 7] = (e >> 4) | bs;
}

var __toBuffer = function(bufs/*:Array<any>*/) { var x=[],w=10240; for(var i=0;i<bufs[0].length;++i) for(var j=0,L=bufs[0][i].length;j<L;j+=w) x.push.apply(x, bufs[0][i].slice(j,j+w)); return x; };
var ___toBuffer = __toBuffer;
var __utf16le = function(b/*:RawBytes|CFBlob*/,s/*:number*/,e/*:number*/)/*:string*/ { var ss/*:Array<string>*/=[]; for(var i=s; i<e; i+=2) ss.push(String.fromCharCode(__readUInt16LE(b,i))); return ss.join("").replace(chr0,''); };
var ___utf16le = __utf16le;
var __hexlify = function(b/*:RawBytes|CFBlob*/,s/*:number*/,l/*:number*/)/*:string*/ { var ss/*:Array<string>*/=[]; for(var i=s; i<s+l; ++i) ss.push(("0" + b[i].toString(16)).slice(-2)); return ss.join(""); };
var ___hexlify = __hexlify;
var __utf8 = function(b/*:RawBytes|CFBlob*/,s/*:number*/,e/*:number*/) { var ss=[]; for(var i=s; i<e; i++) ss.push(String.fromCharCode(__readUInt8(b,i))); return ss.join(""); };
var ___utf8 = __utf8;
var __lpstr = function(b/*:RawBytes|CFBlob*/,i/*:number*/) { var len = __readUInt32LE(b,i); return len > 0 ? __utf8(b, i+4,i+4+len-1) : "";};
var ___lpstr = __lpstr;
var __cpstr = function(b/*:RawBytes|CFBlob*/,i/*:number*/) { var len = __readUInt32LE(b,i); return len > 0 ? __utf8(b, i+4,i+4+len-1) : "";};
var ___cpstr = __cpstr;
var __lpwstr = function(b/*:RawBytes|CFBlob*/,i/*:number*/) { var len = 2*__readUInt32LE(b,i); return len > 0 ? __utf8(b, i+4,i+4+len-1) : "";};
var ___lpwstr = __lpwstr;
var __lpp4, ___lpp4;
__lpp4 = ___lpp4 = function lpp4_(b/*:RawBytes|CFBlob*/,i/*:number*/) { var len = __readUInt32LE(b,i); return len > 0 ? __utf16le(b, i+4,i+4+len) : "";};
var __8lpp4 = function(b/*:RawBytes|CFBlob*/,i/*:number*/) { var len = __readUInt32LE(b,i); return len > 0 ? __utf8(b, i+4,i+4+len) : "";};
var ___8lpp4 = __8lpp4;
var __double, ___double;
__double = ___double = function(b/*:RawBytes|CFBlob*/, idx/*:number*/) { return read_double_le(b, idx);};
var is_buf = function is_buf_a(a) { return Array.isArray(a); };

if(has_buf/*:: && typeof Buffer !== 'undefined'*/) {
	__utf16le = function(b/*:RawBytes|CFBlob*/,s/*:number*/,e/*:number*/)/*:string*/ { if(!Buffer.isBuffer(b)/*:: || !(b instanceof Buffer)*/) return ___utf16le(b,s,e); return b.toString('utf16le',s,e).replace(chr0,'')/*.replace(chr1,'!')*/; };
	__hexlify = function(b/*:RawBytes|CFBlob*/,s/*:number*/,l/*:number*/)/*:string*/ { return Buffer.isBuffer(b)/*:: && b instanceof Buffer*/ ? b.toString('hex',s,s+l) : ___hexlify(b,s,l); };
	__lpstr = function lpstr_b(b/*:RawBytes|CFBlob*/, i/*:number*/) { if(!Buffer.isBuffer(b)/*:: || !(b instanceof Buffer)*/) return ___lpstr(b, i); var len = b.readUInt32LE(i); return len > 0 ? b.toString('utf8',i+4,i+4+len-1) : "";};
	__cpstr = function cpstr_b(b/*:RawBytes|CFBlob*/, i/*:number*/) { if(!Buffer.isBuffer(b)/*:: || !(b instanceof Buffer)*/) return ___cpstr(b, i); var len = b.readUInt32LE(i); return len > 0 ? b.toString('utf8',i+4,i+4+len-1) : "";};
	__lpwstr = function lpwstr_b(b/*:RawBytes|CFBlob*/, i/*:number*/) { if(!Buffer.isBuffer(b)/*:: || !(b instanceof Buffer)*/) return ___lpwstr(b, i); var len = 2*b.readUInt32LE(i); return b.toString('utf16le',i+4,i+4+len-1);};
	__lpp4 = function lpp4_b(b/*:RawBytes|CFBlob*/, i/*:number*/) { if(!Buffer.isBuffer(b)/*:: || !(b instanceof Buffer)*/) return ___lpp4(b, i); var len = b.readUInt32LE(i); return b.toString('utf16le',i+4,i+4+len);};
	__8lpp4 = function lpp4_8b(b/*:RawBytes|CFBlob*/, i/*:number*/) { if(!Buffer.isBuffer(b)/*:: || !(b instanceof Buffer)*/) return ___8lpp4(b, i); var len = b.readUInt32LE(i); return b.toString('utf8',i+4,i+4+len);};
	__utf8 = function utf8_b(b/*:RawBytes|CFBlob*/, s/*:number*/, e/*:number*/) { return (Buffer.isBuffer(b)/*:: && (b instanceof Buffer)*/) ? b.toString('utf8',s,e) : ___utf8(b,s,e); };
	__toBuffer = function(bufs) { return (bufs[0].length > 0 && Buffer.isBuffer(bufs[0][0])) ? Buffer.concat(bufs[0]) : ___toBuffer(bufs);};
	bconcat = function(bufs) { return Buffer.isBuffer(bufs[0]) ? Buffer.concat(bufs) : [].concat.apply([], bufs); };
	__double = function double_(b/*:RawBytes|CFBlob*/, i/*:number*/) { if(Buffer.isBuffer(b)/*::&& b instanceof Buffer*/) return b.readDoubleLE(i); return ___double(b,i); };
	is_buf = function is_buf_b(a) { return Buffer.isBuffer(a) || Array.isArray(a); };
}

/* from js-xls */
if(typeof cptable !== 'undefined') {
	__utf16le = function(b/*:RawBytes|CFBlob*/,s/*:number*/,e/*:number*/) { return cptable.utils.decode(1200, b.slice(s,e)).replace(chr0, ''); };
	__utf8 = function(b/*:RawBytes|CFBlob*/,s/*:number*/,e/*:number*/) { return cptable.utils.decode(65001, b.slice(s,e)); };
	__lpstr = function(b/*:RawBytes|CFBlob*/,i/*:number*/) { var len = __readUInt32LE(b,i); return len > 0 ? cptable.utils.decode(current_ansi, b.slice(i+4, i+4+len-1)) : "";};
	__cpstr = function(b/*:RawBytes|CFBlob*/,i/*:number*/) { var len = __readUInt32LE(b,i); return len > 0 ? cptable.utils.decode(current_codepage, b.slice(i+4, i+4+len-1)) : "";};
	__lpwstr = function(b/*:RawBytes|CFBlob*/,i/*:number*/) { var len = 2*__readUInt32LE(b,i); return len > 0 ? cptable.utils.decode(1200, b.slice(i+4,i+4+len-1)) : "";};
	__lpp4 = function(b/*:RawBytes|CFBlob*/,i/*:number*/) { var len = __readUInt32LE(b,i); return len > 0 ? cptable.utils.decode(1200, b.slice(i+4,i+4+len)) : "";};
	__8lpp4 = function(b/*:RawBytes|CFBlob*/,i/*:number*/) { var len = __readUInt32LE(b,i); return len > 0 ? cptable.utils.decode(65001, b.slice(i+4,i+4+len)) : "";};
}

var __readUInt8 = function(b/*:RawBytes|CFBlob*/, idx/*:number*/)/*:number*/ { return b[idx]; };
var __readUInt16LE = function(b/*:RawBytes|CFBlob*/, idx/*:number*/)/*:number*/ { return (b[idx+1]*(1<<8))+b[idx]; };
var __readInt16LE = function(b/*:RawBytes|CFBlob*/, idx/*:number*/)/*:number*/ { var u = (b[idx+1]*(1<<8))+b[idx]; return (u < 0x8000) ? u : ((0xffff - u + 1) * -1); };
var __readUInt32LE = function(b/*:RawBytes|CFBlob*/, idx/*:number*/)/*:number*/ { return b[idx+3]*(1<<24)+(b[idx+2]<<16)+(b[idx+1]<<8)+b[idx]; };
var __readInt32LE = function(b/*:RawBytes|CFBlob*/, idx/*:number*/)/*:number*/ { return (b[idx+3]<<24)|(b[idx+2]<<16)|(b[idx+1]<<8)|b[idx]; };
var __readInt32BE = function(b/*:RawBytes|CFBlob*/, idx/*:number*/)/*:number*/ { return (b[idx]<<24)|(b[idx+1]<<16)|(b[idx+2]<<8)|b[idx+3]; };

function ReadShift(size/*:number*/, t/*:?string*/)/*:number|string*/ {
	var o="", oI/*:: :number = 0*/, oR, oo=[], w, vv, i, loc;
	switch(t) {
		case 'dbcs':
			loc = this.l;
			if(has_buf && Buffer.isBuffer(this)) o = this.slice(this.l, this.l+2*size).toString("utf16le");
			else for(i = 0; i < size; ++i) { o+=String.fromCharCode(__readUInt16LE(this, loc)); loc+=2; }
			size *= 2;
			break;

		case 'utf8': o = __utf8(this, this.l, this.l + size); break;
		case 'utf16le': size *= 2; o = __utf16le(this, this.l, this.l + size); break;

		case 'wstr':
			if(typeof cptable !== 'undefined') o = cptable.utils.decode(current_codepage, this.slice(this.l, this.l+2*size));
			else return ReadShift.call(this, size, 'dbcs');
			size = 2 * size; break;

		/* [MS-OLEDS] 2.1.4 LengthPrefixedAnsiString */
		case 'lpstr-ansi': o = __lpstr(this, this.l); size = 4 + __readUInt32LE(this, this.l); break;
		case 'lpstr-cp': o = __cpstr(this, this.l); size = 4 + __readUInt32LE(this, this.l); break;
		/* [MS-OLEDS] 2.1.5 LengthPrefixedUnicodeString */
		case 'lpwstr': o = __lpwstr(this, this.l); size = 4 + 2 * __readUInt32LE(this, this.l); break;
		/* [MS-OFFCRYPTO] 2.1.2 Length-Prefixed Padded Unicode String (UNICODE-LP-P4) */
		case 'lpp4': size = 4 +  __readUInt32LE(this, this.l); o = __lpp4(this, this.l); if(size & 0x02) size += 2; break;
		/* [MS-OFFCRYPTO] 2.1.3 Length-Prefixed UTF-8 String (UTF-8-LP-P4) */
		case '8lpp4': size = 4 +  __readUInt32LE(this, this.l); o = __8lpp4(this, this.l); if(size & 0x03) size += 4 - (size & 0x03); break;

		case 'cstr': size = 0; o = "";
			while((w=__readUInt8(this, this.l + size++))!==0) oo.push(_getchar(w));
			o = oo.join(""); break;
		case '_wstr': size = 0; o = "";
			while((w=__readUInt16LE(this,this.l +size))!==0){oo.push(_getchar(w));size+=2;}
			size+=2; o = oo.join(""); break;

		/* sbcs and dbcs support continue records in the SST way TODO codepages */
		case 'dbcs-cont': o = ""; loc = this.l;
			for(i = 0; i < size; ++i) {
				if(this.lens && this.lens.indexOf(loc) !== -1) {
					w = __readUInt8(this, loc);
					this.l = loc + 1;
					vv = ReadShift.call(this, size-i, w ? 'dbcs-cont' : 'sbcs-cont');
					return oo.join("") + vv;
				}
				oo.push(_getchar(__readUInt16LE(this, loc)));
				loc+=2;
			} o = oo.join(""); size *= 2; break;

		case 'cpstr':
			if(typeof cptable !== 'undefined') {
				o = cptable.utils.decode(current_codepage, this.slice(this.l, this.l + size));
				break;
			}
		/* falls through */
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
		case 4: case -4:
			if(t === 'i' || ((this[this.l+3] & 0x80)===0)) { oI = ((size > 0) ? __readInt32LE : __readInt32BE)(this, this.l); this.l += 4; return oI; }
			else { oR = __readUInt32LE(this, this.l); this.l += 4; } return oR;
		case 8: case -8:
			if(t === 'f') {
				if(size == 8) oR = __double(this, this.l);
				else oR = __double([this[this.l+7],this[this.l+6],this[this.l+5],this[this.l+4],this[this.l+3],this[this.l+2],this[this.l+1],this[this.l+0]], 0);
				this.l += 8; return oR;
			} else size = 8;
		/* falls through */
		case 16: o = __hexlify(this, this.l, size); break;
	}}
	this.l+=size; return o;
}

var __writeUInt32LE = function(b/*:RawBytes|CFBlob*/, val/*:number*/, idx/*:number*/)/*:void*/ { b[idx] = (val & 0xFF); b[idx+1] = ((val >>> 8) & 0xFF); b[idx+2] = ((val >>> 16) & 0xFF); b[idx+3] = ((val >>> 24) & 0xFF); };
var __writeInt32LE  = function(b/*:RawBytes|CFBlob*/, val/*:number*/, idx/*:number*/)/*:void*/ { b[idx] = (val & 0xFF); b[idx+1] = ((val >> 8) & 0xFF); b[idx+2] = ((val >> 16) & 0xFF); b[idx+3] = ((val >> 24) & 0xFF); };
var __writeUInt16LE = function(b/*:RawBytes|CFBlob*/, val/*:number*/, idx/*:number*/)/*:void*/ { b[idx] = (val & 0xFF); b[idx+1] = ((val >>> 8) & 0xFF); };

function WriteShift(t/*:number*/, val/*:string|number*/, f/*:?string*/)/*:any*/ {
	var size = 0, i = 0;
	if(f === 'dbcs') {
		/*:: if(typeof val !== 'string') throw new Error("unreachable"); */
		for(i = 0; i != val.length; ++i) __writeUInt16LE(this, val.charCodeAt(i), this.l + 2 * i);
		size = 2 * val.length;
	} else if(f === 'sbcs') {
		/* TODO: codepage */
		/*:: if(typeof val !== 'string') throw new Error("unreachable"); */
		val = val.replace(/[^\x00-\x7F]/g, "_");
		/*:: if(typeof val !== 'string') throw new Error("unreachable"); */
		for(i = 0; i != val.length; ++i) this[this.l + i] = (val.charCodeAt(i) & 0xFF);
		size = val.length;
	} else if(f === 'hex') {
		for(; i < t; ++i) {
			/*:: if(typeof val !== "string") throw new Error("unreachable"); */
			this[this.l++] = (parseInt(val.slice(2*i, 2*i+2), 16)||0);
		} return this;
	} else if(f === 'utf16le') {
			/*:: if(typeof val !== "string") throw new Error("unreachable"); */
			var end/*:number*/ = Math.min(this.l + t, this.length);
			for(i = 0; i < Math.min(val.length, t); ++i) {
				var cc = val.charCodeAt(i);
				this[this.l++] = (cc & 0xff);
				this[this.l++] = (cc >> 8);
			}
			while(this.l < end) this[this.l++] = 0;
			return this;
	} else /*:: if(typeof val === 'number') */ switch(t) {
		case  1: size = 1; this[this.l] = val&0xFF; break;
		case  2: size = 2; this[this.l] = val&0xFF; val >>>= 8; this[this.l+1] = val&0xFF; break;
		case  3: size = 3; this[this.l] = val&0xFF; val >>>= 8; this[this.l+1] = val&0xFF; val >>>= 8; this[this.l+2] = val&0xFF; break;
		case  4: size = 4; __writeUInt32LE(this, val, this.l); break;
		case  8: size = 8; if(f === 'f') { write_double_le(this, val, this.l); break; }
		/* falls through */
		case 16: break;
		case -4: size = 4; __writeInt32LE(this, val, this.l); break;
	}
	this.l += size; return this;
}

function CheckField(hexstr/*:string*/, fld/*:string*/)/*:void*/ {
	var m = __hexlify(this,this.l,hexstr.length>>1);
	if(m !== hexstr) throw new Error(fld + 'Expected ' + hexstr + ' saw ' + m);
	this.l += hexstr.length>>1;
}

function prep_blob(blob, pos/*:number*/)/*:void*/ {
	blob.l = pos;
	blob.read_shift = /*::(*/ReadShift/*:: :any)*/;
	blob.chk = CheckField;
	blob.write_shift = WriteShift;
}

function parsenoop(blob, length/*:: :number, opts?:any */) { blob.l += length; }

function new_buf(sz/*:number*/)/*:Block*/ {
	var o = new_raw_buf(sz);
	prep_blob(o, 0);
	return o;
}

