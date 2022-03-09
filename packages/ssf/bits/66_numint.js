function write_num_int(type/*:string*/, fmt/*:string*/, val/*:number*/)/*:string*/ {
	if(type.charCodeAt(0) === 40 && !fmt.match(closeparen)) {
		var ffmt = fmt.replace(/\( */,"").replace(/ \)/,"").replace(/\)/,"");
		if(val >= 0) return write_num_int('n', ffmt, val);
		return '(' + write_num_int('n', ffmt, -val) + ')';
	}
	if(fmt.charCodeAt(fmt.length - 1) === 44) return write_num_cm2(type, fmt, val);
	if(fmt.indexOf('%') !== -1) return write_num_pct2(type, fmt, val);
	if(fmt.indexOf('E') !== -1) return write_num_exp2(fmt, val);
	if(fmt.charCodeAt(0) === 36) return "$"+write_num_int(type,fmt.substr(fmt.charAt(1)==' '?2:1),val);
	var o;
	var r/*:?Array<string>*/, ri, ff, aval = Math.abs(val), sign = val < 0 ? "-" : "";
	if(fmt.match(/^00+$/)) return sign + pad0(aval,fmt.length);
	if(fmt.match(/^[#?]+$/)) {
		o = (""+val); if(val === 0) o = "";
		return o.length > fmt.length ? o : hashq(fmt.substr(0,fmt.length-o.length)) + o;
	}
	if((r = fmt.match(frac1))) return write_num_f2(r, aval, sign);
	if(fmt.match(/^#+0+$/)) return sign + pad0(aval,fmt.length - fmt.indexOf("0"));
	if((r = fmt.match(dec1))) {
		/*:: if(!Array.isArray(r)) throw new Error("unreachable"); */
		o = (""+val).replace(/^([^\.]+)$/,"$1."+hashq(r[1])).replace(/\.$/,"."+hashq(r[1]));
		o = o.replace(/\.(\d*)$/,function($$, $1) {
		/*:: if(!Array.isArray(r)) throw new Error("unreachable"); */
			return "." + $1 + fill("0", hashq(r[1]).length-$1.length); });
		return fmt.indexOf("0.") !== -1 ? o : o.replace(/^0\./,".");
	}
	fmt = fmt.replace(/^#+([0.])/, "$1");
	if((r = fmt.match(/^(0*)\.(#*)$/))) {
		return sign + (""+aval).replace(/\.(\d*[1-9])0*$/,".$1").replace(/^(-?\d*)$/,"$1.").replace(/^0\./,r[1].length?"0.":".");
	}
	if((r = fmt.match(/^#{1,3},##0(\.?)$/))) return sign + commaify((""+aval));
	if((r = fmt.match(/^#,##0\.([#0]*0)$/))) {
		return val < 0 ? "-" + write_num_int(type, fmt, -val) : commaify((""+val)) + "." + fill('0',r[1].length);
	}
	if((r = fmt.match(/^#,#*,#0/))) return write_num_int(type,fmt.replace(/^#,#*,/,""),val);
	if((r = fmt.match(/^([0#]+)(\\?-([0#]+))+$/))) {
		o = _strrev(write_num_int(type, fmt.replace(/[\\-]/g,""), val));
		ri = 0;
		return _strrev(_strrev(fmt.replace(/\\/g,"")).replace(/[0#]/g,function(x){return ri<o.length?o.charAt(ri++):x==='0'?'0':"";}));
	}
	if(fmt.match(phone)) {
		o = write_num_int(type, "##########", val);
		return "(" + o.substr(0,3) + ") " + o.substr(3, 3) + "-" + o.substr(6);
	}
	var oa = "";
	if((r = fmt.match(/^([#0?]+)( ?)\/( ?)([#0?]+)/))) {
		ri = Math.min(/*::String(*/r[4]/*::)*/.length,7);
		ff = frac(aval, Math.pow(10,ri)-1, false);
		o = "" + sign;
		oa = write_num("n", /*::String(*/r[1]/*::)*/, ff[1]);
		if(oa.charAt(oa.length-1) == " ") oa = oa.substr(0,oa.length-1) + "0";
		o += oa + /*::String(*/r[2]/*::)*/ + "/" + /*::String(*/r[3]/*::)*/;
		oa = rpad_(ff[2],ri);
		if(oa.length < r[4].length) oa = hashq(r[4].substr(r[4].length-oa.length)) + oa;
		o += oa;
		return o;
	}
	if((r = fmt.match(/^# ([#0?]+)( ?)\/( ?)([#0?]+)/))) {
		ri = Math.min(Math.max(r[1].length, r[4].length),7);
		ff = frac(aval, Math.pow(10,ri)-1, true);
		return sign + (ff[0]||(ff[1] ? "" : "0")) + " " + (ff[1] ? pad_(ff[1],ri) + r[2] + "/" + r[3] + rpad_(ff[2],ri): fill(" ", 2*ri+1 + r[2].length + r[3].length));
	}
	if((r = fmt.match(/^[#0?]+$/))) {
		o = "" + val;
		if(fmt.length <= o.length) return o;
		return hashq(fmt.substr(0,fmt.length-o.length)) + o;
	}
	if((r = fmt.match(/^([#0]+)\.([#0]+)$/))) {
		o = "" + val.toFixed(Math.min(r[2].length,10)).replace(/([^0])0+$/,"$1");
		ri = o.indexOf(".");
		var lres = fmt.indexOf(".") - ri, rres = fmt.length - o.length - lres;
		return hashq(fmt.substr(0,lres) + o + fmt.substr(fmt.length-rres));
	}
	if((r = fmt.match(/^00,000\.([#0]*0)$/))) {
		return val < 0 ? "-" + write_num_int(type, fmt, -val) : commaify(""+val).replace(/^\d,\d{3}$/,"0$&").replace(/^\d*$/,function($$) { return "00," + ($$.length < 3 ? pad0(0,3-$$.length) : "") + $$; }) + "." + pad0(0,r[1].length);
	}
	switch(fmt) {
		case "###,###":
		case "##,###":
		case "#,###": var x = commaify(""+aval); return x !== "0" ? sign + x : "";
		default:
			if(fmt.match(/\.[0#?]*$/)) return write_num_int(type, fmt.slice(0,fmt.lastIndexOf(".")), val) + hashq(fmt.slice(fmt.lastIndexOf(".")));
	}
	throw new Error("unsupported format |" + fmt + "|");
}
