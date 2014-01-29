/* [MS-XLSB] 2.4.219 BrtBeginSst */
var parse_BrtBeginSst = function(data, length) {
	return [data.read_shift(4), data.read_shift(4)];
};

/* [MS-XLSB] 2.1.7.45 Shared Strings */
var parse_sst_bin = function(data) {
	var s = [];
	recordhopper(data, function(val, R) {
		switch(R.n) {
			case 'BrtBeginSst': s.Count = val[0]; s.Unique = val[1]; break;
			case 'BrtSSTItem': s.push(val); break;
			case 'BrtEndSst': return true;
			default: throw new Error("Unexpected record " + R.n);
		}
	});
	return s;
};
