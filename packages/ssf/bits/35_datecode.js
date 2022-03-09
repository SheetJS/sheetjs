function parse_date_code(v/*:number*/,opts/*:?any*/,b2/*:?boolean*/) {
	if(v > 2958465 || v < 0) return null;
	var date = (v|0), time = Math.floor(86400 * (v - date)), dow=0;
	var dout=[];
	var out={D:date, T:time, u:86400*(v-date)-time,y:0,m:0,d:0,H:0,M:0,S:0,q:0};
	if(Math.abs(out.u) < 1e-6) out.u = 0;
	if(opts && opts.date1904) date += 1462;
	if(out.u > 0.9999) {
		out.u = 0;
		if(++time == 86400) { out.T = time = 0; ++date; ++out.D; }
	}
	if(date === 60) {dout = b2 ? [1317,10,29] : [1900,2,29]; dow=3;}
	else if(date === 0) {dout = b2 ? [1317,8,29] : [1900,1,0]; dow=6;}
	else {
		if(date > 60) --date;
		/* 1 = Jan 1 1900 in Gregorian */
		var d = new Date(1900, 0, 1);
		d.setDate(d.getDate() + date - 1);
		dout = [d.getFullYear(), d.getMonth()+1,d.getDate()];
		dow = d.getDay();
		if(date < 60) dow = (dow + 6) % 7;
		if(b2) dow = fix_hijri(d, dout);
	}
	out.y = dout[0]; out.m = dout[1]; out.d = dout[2];
	out.S = time % 60; time = Math.floor(time / 60);
	out.M = time % 60; time = Math.floor(time / 60);
	out.H = time;
	out.q = dow;
	return out;
}
SSF.parse_date_code = parse_date_code;
var basedate = new Date(1899, 11, 31, 0, 0, 0);
var dnthresh = basedate.getTime();
var base1904 = new Date(1900, 2, 1, 0, 0, 0);
function datenum_local(v/*:Date*/, date1904/*:?boolean*/)/*:number*/ {
	var epoch = v.getTime();
	if(date1904) epoch -= 1461*24*60*60*1000;
	else if(v >= base1904) epoch += 24*60*60*1000;
	return (epoch - (dnthresh + (v.getTimezoneOffset() - basedate.getTimezoneOffset()) * 60000)) / (24 * 60 * 60 * 1000);
}
