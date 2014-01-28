/* [MS-XLSB] 2.1.4 Record */
var recordhopper = function(data, cb) {
	var tmpbyte, cntbyte;
	prep_blob(data, data.l || 0);
	while(data.l < data.length) {
		var RT = data.read_shift(1);
		if(RT & 0x80) RT = (RT & 0x7F) + ((data.read_shift(1) & 0x7F)<<7);
		var R = RecordEnum[RT] || RecordEnum[0xFFFF];

		var length = tmpbyte = data.read_shift(1);
		for(cntbyte = 1; cntbyte <4 && (tmpbyte & 0x80); ++cntbyte) length += ((tmpbyte = data.read_shift(1)) & 0x7F)<<(7*cntbyte);
		var d = R.f(data, length);
		if(cb(d, R)) return;
	}
};
