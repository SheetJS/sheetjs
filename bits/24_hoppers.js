/* [MS-XLSB] 2.1.4 Record */
var recordhopper = function(data, cb, opts) {
	var tmpbyte, cntbyte, length;
	prep_blob(data, data.l || 0);
	while(data.l < data.length) {
		var RT = data.read_shift(1);
		if(RT & 0x80) RT = (RT & 0x7F) + ((data.read_shift(1) & 0x7F)<<7);
		var R = RecordEnum[RT] || RecordEnum[0xFFFF];
		tmpbyte = data.read_shift(1);
		length = tmpbyte & 0x7F;
		for(cntbyte = 1; cntbyte <4 && (tmpbyte & 0x80); ++cntbyte) length += ((tmpbyte = data.read_shift(1)) & 0x7F)<<(7*cntbyte);
		var d = R.f(data, length, opts);
		if(cb(d, R, RT)) return;
	}
};

/* control buffer usage for fixed-length buffers */
var buf_array = function() {
	var bufs = [], blksz = 2048;
	var newblk = function(sz) {
		var o = new_buf(sz || blksz);
		prep_blob(o, 0, true);
		return o;
	};

	var curbuf = newblk();

	var endbuf = function() {
		curbuf.length = curbuf.l;
		if(curbuf.length > 0) bufs.push(curbuf);
		curbuf = null;
	};

	var next = function(sz) {
		if(sz < curbuf.length - curbuf.l) return curbuf;
		endbuf();
		return (curbuf = newblk(Math.max(sz+1, blksz)));
	};

	var end = function() {
		endbuf();
		return __toBuffer([bufs]);
	};

	var push = function(buf) { endbuf(); curbuf = buf; next(); };

	return { next:next, push:push, end:end, _bufs:bufs };
};

var write_record = function(ba, type, payload, length) {
	var t = evert_RE[type], l;
	if(!length) length = RecordEnum[t].p || (payload||[]).length || 0;
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
	if(length > 0 && is_buf(payload)) ba.push(payload);
};
