function write_num_cm2(type/*:string*/, fmt/*:string*/, val/*:number*/)/*:string*/{
	var idx = fmt.length - 1;
	while(fmt.charCodeAt(idx-1) === 44) --idx;
	return write_num(type, fmt.substr(0,idx), val / Math.pow(10,3*(fmt.length-idx)));
}
function write_num_pct2(type/*:string*/, fmt/*:string*/, val/*:number*/)/*:string*/{
	var sfmt = fmt.replace(pct1,""), mul = fmt.length - sfmt.length;
	return write_num(type, sfmt, val * Math.pow(10,2*mul)) + fill("%",mul);
}
function write_num_exp2(fmt/*:string*/, val/*:number*/)/*:string*/{
	var o/*:string*/;
	var idx = fmt.indexOf("E") - fmt.indexOf(".") - 1;
	if(fmt.match(/^#+0.0E\+0$/)) {
		if(val == 0) return "0.0E+0";
		else if(val < 0) return "-" + write_num_exp2(fmt, -val);
		var period = fmt.indexOf("."); if(period === -1) period=fmt.indexOf('E');
		var ee = Math.floor(Math.log(val)*Math.LOG10E)%period;
		if(ee < 0) ee += period;
		o = (val/Math.pow(10,ee)).toPrecision(idx+1+(period+ee)%period);
		if(!o.match(/[Ee]/)) {
			var fakee = Math.floor(Math.log(val)*Math.LOG10E);
			if(o.indexOf(".") === -1) o = o.charAt(0) + "." + o.substr(1) + "E+" + (fakee - o.length+ee);
			else o += "E+" + (fakee - ee);
			o = o.replace(/\+-/,"-");
		}
		o = o.replace(/^([+-]?)(\d*)\.(\d*)[Ee]/,function($$,$1,$2,$3) { return $1 + $2 + $3.substr(0,(period+ee)%period) + "." + $3.substr(ee) + "E"; });
	} else o = val.toExponential(idx);
	if(fmt.match(/E\+00$/) && o.match(/e[+-]\d$/)) o = o.substr(0,o.length-1) + "0" + o.charAt(o.length-1);
	if(fmt.match(/E\-/) && o.match(/e\+/)) o = o.replace(/e\+/,"e");
	return o.replace("e","E");
}
