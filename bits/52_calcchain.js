/* 18.6 Calculation Chain */
function parse_cc_xml(data/*::, name, opts*/)/*:Array<any>*/ {
	var d = [];
	if(!data) return d;
	var i = 1;
	(data.match(tagregex)||[]).forEach(function(x) {
		var y = parsexmltag(x);
		switch(y[0]) {
			case '<?xml': break;
			/* 18.6.2  calcChain CT_CalcChain 1 */
			case '<calcChain': case '<calcChain>': case '</calcChain>': break;
			/* 18.6.1  c CT_CalcCell 1 */
			case '<c': delete y[0]; if(y.i) i = y.i; else y.i = i; d.push(y); break;
		}
	});
	return d;
}

//function write_cc_xml(data, opts) { }

/* [MS-XLSB] 2.6.4.1 */
function parse_BrtCalcChainItem$(data) {
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
}

/* 18.6 Calculation Chain */
function parse_cc_bin(data, name, opts) {
	var out = [];
	var pass = false;
	recordhopper(data, function hopper_cc(val, R, RT) {
		switch(RT) {
			case 0x003F: /* 'BrtCalcChainItem$' */
				out.push(val); break;

			default:
				if(R.T){/* empty */}
				else if(!pass || opts.WTF) throw new Error("Unexpected record 0x" + RT.toString(16));
		}
	});
	return out;
}

//function write_cc_bin(data, opts) { }
