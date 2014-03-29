/* [MS-XLSB] 2.6.4.1 */
var parse_BrtCalcChainItem$ = function(data, length) {
	var out = {};
	out.i = data.read_shift(4);
	var cell = {};
	cell.r = data.read_shift(4);
	cell.c = data.read_shift(4);
	out.r = encode_cell(cell);
	var flags = data.read_shift(1);
	if(flags & 0x2) out.l = '1';
	if(flags & 0x8) out.a = '1';
	return out;
};

/* 18.6 Calculation Chain */
function parse_cc_bin(data, opts) {
	var out = [];
	var pass = false;
	recordhopper(data, function(val, R, RT) {
		switch(R.n) {
			case 'BrtCalcChainItem$': out.push(val); break;
			case 'BrtBeginCalcChain$': break;
			case 'BrtEndCalcChain$': break;
			default: if(!pass || opts.WTF) throw new Error("Unexpected record " + RT + " " + R.n);
		}
	});
	return out;
}
