function general_fmt_int(v/*:number*/)/*:string*/ { return v.toString(10); }
SSF._general_int = general_fmt_int;
var general_fmt_num = (function make_general_fmt_num() {
var gnr1 = /\.(\d*[1-9])0+$/, gnr2 = /\.0*$/, gnr4 = /\.(\d*[1-9])0+/, gnr5 = /\.0*[Ee]/, gnr6 = /(E[+-])(\d)$/;
function gfn2(v) {
	var w = (v<0?12:11);
	var o = gfn5(v.toFixed(12)); if(o.length <= w) return o;
	o = v.toPrecision(10); if(o.length <= w) return o;
	return v.toExponential(5);
}
function gfn3(v) {
	var o = v.toFixed(11).replace(gnr1,".$1");
	if(o.length > (v<0?12:11)) o = v.toPrecision(6);
	return o;
}
function gfn4(o) {
	for(var i = 0; i != o.length; ++i) if((o.charCodeAt(i) | 0x20) === 101) return o.replace(gnr4,".$1").replace(gnr5,"E").replace("e","E").replace(gnr6,"$10$2");
	return o;
}
function gfn5(o) {
	return o.indexOf(".") > -1 ? o.replace(gnr2,"").replace(gnr1,".$1") : o;
}
return function general_fmt_num(v/*:number*/)/*:string*/ {
	var V = Math.floor(Math.log(Math.abs(v))*Math.LOG10E), o;
	if(V >= -4 && V <= -1) o = v.toPrecision(10+V);
	else if(Math.abs(V) <= 9) o = gfn2(v);
	else if(V === 10) o = v.toFixed(10).substr(0,12);
	else o = gfn3(v);
	return gfn5(gfn4(o));
};})();
SSF._general_num = general_fmt_num;
function general_fmt(v/*:any*/, opts/*:any*/) {
	switch(typeof v) {
		case 'string': return v;
		case 'boolean': return v ? "TRUE" : "FALSE";
		case 'number': return (v|0) === v ? general_fmt_int(v) : general_fmt_num(v);
		case 'undefined': return "";
		case 'object':
			if(v == null) return "";
			if(v instanceof Date) return format(14, datenum_local(v, opts && opts.date1904), opts);
	}
	throw new Error("unsupported value in General format: " + v);
}
SSF._general = general_fmt;
