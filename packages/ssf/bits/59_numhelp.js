var pct1 = /%/g;
function write_num_pct(type/*:string*/, fmt/*:string*/, val/*:number*/)/*:string*/{
	var sfmt = fmt.replace(pct1,""), mul = fmt.length - sfmt.length;
	return write_num(type, sfmt, val * Math.pow(10,2*mul)) + fill("%",mul);
}
