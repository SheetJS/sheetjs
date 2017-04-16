/* [MS-XLSB] 2.4.219 BrtBeginSst */
function parse_BrtBeginSst(data, length) {
	return [data.read_shift(4), data.read_shift(4)];
}

/* [MS-XLSB] 2.1.7.45 Shared Strings */
function parse_sst_bin(data, opts)/*:SST*/ {
	var s/*:SST*/ = ([]/*:any*/);
	var pass = false;
	recordhopper(data, function hopper_sst(val, R_n, RT) {
		switch(RT) {
			case 0x009F: /* 'BrtBeginSst' */
				s.Count = val[0]; s.Unique = val[1]; break;
			case 0x0013: /* 'BrtSSTItem' */
				s.push(val); break;
			case 0x00A0: /* 'BrtEndSst' */
				return true;

			case 0x0023: /* 'BrtFRTBegin' */
				pass = true; break;
			case 0x0024: /* 'BrtFRTEnd' */
				pass = false; break;

			default:
				if(R_n.indexOf("Begin") > 0){}
				else if(R_n.indexOf("End") > 0){}
				if(!pass || opts.WTF) throw new Error("Unexpected record " + RT + " " + R_n);
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
	/* FRTSST */
	write_record(ba, "BrtEndSst");
	return ba.end();
}
