/* [MS-XLSB] 2.4.219 BrtBeginSst */
function parse_BrtBeginSst(data, length) {
	return [data.read_shift(4), data.read_shift(4)];
}

/* [MS-XLSB] 2.1.7.45 Shared Strings */
function parse_sst_bin(data, opts) {
	var s = [];
	var pass = false;
	recordhopper(data, function hopper_sst(val, R, RT) {
		switch(R.n) {
			case 'BrtBeginSst': s.Count = val[0]; s.Unique = val[1]; break;
			case 'BrtSSTItem': s.push(val); break;
			case 'BrtEndSst': return true;
			/* TODO: produce a test case with a future record */
			case 'BrtFRTBegin': pass = true; break;
			case 'BrtFRTEnd': pass = false; break;
			default: if(!pass || opts.WTF) throw new Error("Unexpected record " + RT + " " + R.n);
		}
	});
	return s;
}

function write_BrtBeginSst(sst, o) {
	if(!o) o = new_buf(8);
	o.write_shift(4, sst.Count);
	o.write_shift(4, sst.Unique);
	return o;
}

var write_BrtSSTItem = write_RichStr;

function write_sst_bin(sst, opts) {
	var ba = buf_array();
	write_record(ba, "BrtBeginSst", write_BrtBeginSst(sst));
	for(var i = 0; i < sst.length; ++i) write_record(ba, "BrtSSTItem", write_BrtSSTItem(sst[i]));
	write_record(ba, "BrtEndSst");
	return ba.end();
}
