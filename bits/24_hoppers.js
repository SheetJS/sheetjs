/* [MS-XLSB] 2.1.4 Record */
function recordhopper(data, cb, opts) {
	var tmpbyte, cntbyte, length;
	prep_blob(data, data.l || 0);
	while(data.l < data.length) {
		var RT = data.read_shift(1);
		if(RT & 0x80) RT = (RT & 0x7F) + ((data.read_shift(1) & 0x7F)<<7);
		var R = XLSBRecordEnum[RT] || XLSBRecordEnum[0xFFFF];
		tmpbyte = data.read_shift(1);
		length = tmpbyte & 0x7F;
		for(cntbyte = 1; cntbyte <4 && (tmpbyte & 0x80); ++cntbyte) length += ((tmpbyte = data.read_shift(1)) & 0x7F)<<(7*cntbyte);
		var d = R.f(data, length, opts);
		if(cb(d, R, RT)) return;
	}
}

/* control buffer usage for fixed-length buffers */
function buf_array() {
	var bufs = [], blksz = 2048;
	var newblk = function ba_newblk(sz) {
		var o = new_buf(sz);
		prep_blob(o, 0);
		return o;
	};

	var curbuf = newblk(blksz);

	var endbuf = function ba_endbuf() {
		if(curbuf.length > curbuf.l) curbuf = curbuf.slice(0, curbuf.l);
		if(curbuf.length > 0) bufs.push(curbuf);
		curbuf = null;
	};

	var next = function ba_next(sz) {
		if(sz < curbuf.length - curbuf.l) return curbuf;
		endbuf();
		return (curbuf = newblk(Math.max(sz+1, blksz)));
	};

	var end = function ba_end() {
		endbuf();
		return __toBuffer([bufs]);
	};

	var push = function ba_push(buf) { endbuf(); curbuf = buf; next(blksz); };

	return { next:next, push:push, end:end, _bufs:bufs };
}

function write_record(ba/*:BufArray*/, type/*:string*/, payload, length/*:?number*/) {
	var t = evert_RE[type], l;
	if(!length) length = XLSBRecordEnum[t].p || (payload||[]).length || 0;
	l = 1 + (t >= 0x80 ? 1 : 0) + 1 + length;
	if(length >= 0x80) ++l; if(length >= 0x4000) ++l; if(length >= 0x200000) ++l;
	var o = ba.next(l);
	if(t <= 0x7F) o.write_shift(1, t);
	else {
		o.write_shift(1, (t & 0x7F) + 0x80);
		o.write_shift(1, (t >> 7));
	}
	for(var i = 0; i != 4; ++i) {
		if(length >= 0x80) { o.write_shift(1, (length & 0x7F)+0x80); length >>= 7; }
		else { o.write_shift(1, length); break; }
	}
	if(/*:: length != null &&*/length > 0 && is_buf(payload)) ba.push(payload);
}
