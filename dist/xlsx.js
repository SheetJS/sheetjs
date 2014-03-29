/* xlsx.js (C) 2013-2014 SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
/*jshint eqnull:true */
/* Spreadsheet Format -- jump to XLSX for the XLSX code */
/* ssf.js (C) 2013-2014 SheetJS -- http://sheetjs.com */
var SSF = {};
var make_ssf = function(SSF){
var _strrev = function(x) { return String(x).split("").reverse().join("");};
function fill(c,l) { return new Array(l+1).join(c); }
function pad(v,d,c){var t=String(v);return t.length>=d?t:(fill(c||0,d-t.length)+t);}
function rpad(v,d,c){var t=String(v);return t.length>=d?t:(t+fill(c||0,d-t.length));}
SSF.version = '0.6.1';
/* Options */
var opts_fmt = {};
function fixopts(o){for(var y in opts_fmt) if(o[y]===undefined) o[y]=opts_fmt[y];}
SSF.opts = opts_fmt;
opts_fmt.date1904 = 0;
opts_fmt.output = "";
opts_fmt.mode = "";
var table_fmt = {
	0:  'General',
	1:  '0',
	2:  '0.00',
	3:  '#,##0',
	4:  '#,##0.00',
	9:  '0%',
	10: '0.00%',
	11: '0.00E+00',
	12: '# ?/?',
	13: '# ??/??',
	14: 'm/d/yy',
	15: 'd-mmm-yy',
	16: 'd-mmm',
	17: 'mmm-yy',
	18: 'h:mm AM/PM',
	19: 'h:mm:ss AM/PM',
	20: 'h:mm',
	21: 'h:mm:ss',
	22: 'm/d/yy h:mm',
	37: '#,##0 ;(#,##0)',
	38: '#,##0 ;[Red](#,##0)',
	39: '#,##0.00;(#,##0.00)',
	40: '#,##0.00;[Red](#,##0.00)',
	45: 'mm:ss',
	46: '[h]:mm:ss',
	47: 'mmss.0',
	48: '##0.0E+0',
	49: '@',
	56: '"上午/下午 "hh"時"mm"分"ss"秒 "',
	65535: 'General'
};
var days = [
	['Sun', 'Sunday'],
	['Mon', 'Monday'],
	['Tue', 'Tuesday'],
	['Wed', 'Wednesday'],
	['Thu', 'Thursday'],
	['Fri', 'Friday'],
	['Sat', 'Saturday']
];
var months = [
	['J', 'Jan', 'January'],
	['F', 'Feb', 'February'],
	['M', 'Mar', 'March'],
	['A', 'Apr', 'April'],
	['M', 'May', 'May'],
	['J', 'Jun', 'June'],
	['J', 'Jul', 'July'],
	['A', 'Aug', 'August'],
	['S', 'Sep', 'September'],
	['O', 'Oct', 'October'],
	['N', 'Nov', 'November'],
	['D', 'Dec', 'December']
];
var frac = function frac(x, D, mixed) {
	var sgn = x < 0 ? -1 : 1;
	var B = x * sgn;
	var P_2 = 0, P_1 = 1, P = 0;
	var Q_2 = 1, Q_1 = 0, Q = 0;
	var A = Math.floor(B);
	while(Q_1 < D) {
		A = Math.floor(B);
		P = A * P_1 + P_2;
		Q = A * Q_1 + Q_2;
		if((B - A) < 0.0000000005) break;
		B = 1 / (B - A);
		P_2 = P_1; P_1 = P;
		Q_2 = Q_1; Q_1 = Q;
	}
	if(Q > D) { Q = Q_1; P = P_1; }
	if(Q > D) { Q = Q_2; P = P_2; }
	if(!mixed) return [0, sgn * P, Q];
	if(Q===0) throw "Unexpected state: "+P+" "+P_1+" "+P_2+" "+Q+" "+Q_1+" "+Q_2;
	var q = Math.floor(sgn * P/Q);
	return [q, sgn*P - q*Q, Q];
};
var general_fmt = function(v) {
	if(typeof v === 'boolean') return v ? "TRUE" : "FALSE";
	if(typeof v === 'number') {
		var o, V = v < 0 ? -v : v;
		if(V >= 0.1 && V < 1) o = v.toPrecision(9);
		else if(V >= 0.01 && V < 0.1) o = v.toPrecision(8);
		else if(V >= 0.001 && V < 0.01) o = v.toPrecision(7);
		else if(V >= 0.0001 && V < 0.001) o = v.toPrecision(6);
		else if(V >= Math.pow(10,10) && V < Math.pow(10,11)) o = v.toFixed(10).substr(0,12);
		else if(V > Math.pow(10,-9) && V < Math.pow(10,11)) {
			o = v.toFixed(12).replace(/(\.[0-9]*[1-9])0*$/,"$1").replace(/\.$/,"");
			if(o.length > 11+(v<0?1:0)) o = v.toPrecision(10);
			if(o.length > 11+(v<0?1:0)) o = v.toExponential(5);
		}
		else {
			o = v.toFixed(11).replace(/(\.[0-9]*[1-9])0*$/,"$1");
				if(o.length > 11 + (v<0?1:0)) o = v.toPrecision(6);
		}
		o = o.replace(/(\.[0-9]*[1-9])0+e/,"$1e").replace(/\.0*e/,"e");
		return o.replace("e","E").replace(/\.0*$/,"").replace(/\.([0-9]*[^0])0*$/,".$1").replace(/(E[+-])([0-9])$/,"$1"+"0"+"$2");
	}
	if(typeof v === 'string') return v;
	throw new Error("unsupported value in General format: " + v);
};
SSF._general = general_fmt;
var parse_date_code = function parse_date_code(v,opts) {
	var date = Math.floor(v), time = Math.floor(86400 * (v - date)+1e-6), dow=0;
	var dout=[], out={D:date, T:time, u:86400*(v-date)-time}; fixopts(opts = (opts||{}));
	if(opts.date1904) date += 1462;
	if(date > 2958465) return null;
	if(out.u > .999) {
		out.u = 0;
		if(++time == 86400) { time = 0; ++date; }
	}
	if(date === 60) {dout = [1900,2,29]; dow=3;}
	else if(date === 0) {dout = [1900,1,0]; dow=6;}
	else {
		if(date > 60) --date;
		/* 1 = Jan 1 1900 */
		var d = new Date(1900,0,1);
		d.setDate(d.getDate() + date - 1);
		dout = [d.getFullYear(), d.getMonth()+1,d.getDate()];
		dow = d.getDay();
		if(/* opts.mode === 'excel' && */ date < 60) dow = (dow + 6) % 7;
	}
	out.y = dout[0]; out.m = dout[1]; out.d = dout[2];
	out.S = time % 60; time = Math.floor(time / 60);
	out.M = time % 60; time = Math.floor(time / 60);
	out.H = time;
	out.q = dow;
	return out;
};
SSF.parse_date_code = parse_date_code;
/*jshint -W086 */
var write_date = function(type, fmt, val) {
	if(val < 0) return "";
	var o, ss;
	switch(type) {
		case 'y': switch(fmt) { /* year */
			case 'y': case 'yy': return pad(val.y % 100,2);
			default: return pad(val.y % 10000,4);
		}
		case 'm': switch(fmt) { /* month */
			case 'm': return val.m;
			case 'mm': return pad(val.m,2);
			case 'mmm': return months[val.m-1][1];
			case 'mmmmm': return months[val.m-1][0];
			default: return months[val.m-1][2];
		}
		case 'd': switch(fmt) { /* day */
			case 'd': return val.d;
			case 'dd': return pad(val.d,2);
			case 'ddd': return days[val.q][0];
			default: return days[val.q][1];
		}
		case 'h': switch(fmt) { /* 12-hour */
			case 'h': return 1+(val.H+11)%12;
			case 'hh': return pad(1+(val.H+11)%12, 2);
			default: throw 'bad hour format: ' + fmt;
		}
		case 'H': switch(fmt) { /* 24-hour */
			case 'h': return val.H;
			case 'hh': return pad(val.H, 2);
			default: throw 'bad hour format: ' + fmt;
		}
		case 'M': switch(fmt) { /* minutes */
			case 'm': return val.M;
			case 'mm': return pad(val.M, 2);
			default: throw 'bad minute format: ' + fmt;
		}
		case 's': switch(fmt) { /* seconds */
			case 's': ss=Math.round(val.S+val.u); return ss >= 60 ? 0 : ss;
			case 'ss': ss=Math.round(val.S+val.u); if(ss>=60) ss=0; return pad(ss,2);
			case 'ss.0': ss=Math.round(10*(val.S+val.u)); if(ss>=600) ss = 0; o = pad(ss,3); return o.substr(0,2)+"." + o.substr(2);
			case 'ss.00': ss=Math.round(100*(val.S+val.u)); if(ss>=6000) ss = 0; o = pad(ss,4); return o.substr(0,2)+"." + o.substr(2);
			case 'ss.000': ss=Math.round(1000*(val.S+val.u)); if(ss>=60000) ss = 0; o = pad(ss,5); return o.substr(0,2)+"." + o.substr(2);
			default: throw 'bad second format: ' + fmt;
		}
		case 'Z': switch(fmt) {
			case '[h]': case '[hh]': o = val.D*24+val.H; break;
			case '[m]': case '[mm]': o = (val.D*24+val.H)*60+val.M; break;
			case '[s]': case '[ss]': o = ((val.D*24+val.H)*60+val.M)*60+Math.round(val.S+val.u); break;
			default: throw 'bad abstime format: ' + fmt;
		} return fmt.length === 3 ? o : pad(o, 2);
		/* TODO: handle the ECMA spec format ee -> yy */
		case 'e': { return val.y; } break;
	}
};
/*jshint +W086 */
var commaify = function(s) { return _strrev(_strrev(s).replace(/.../g,"$&,")).replace(/^,/,""); };
var write_num = function(type, fmt, val) {
	if(type === '(') {
		var ffmt = fmt.replace(/\( */,"").replace(/ \)/,"").replace(/\)/,"");
		if(val >= 0) return write_num('n', ffmt, val);
		return '(' + write_num('n', ffmt, -val) + ')';
	}
	var mul = 0, o;
	fmt = fmt.replace(/%/g,function(x) { mul++; return ""; });
	if(mul !== 0) return write_num(type, fmt, val * Math.pow(10,2*mul)) + fill("%",mul);
	fmt = fmt.replace(/(\.0+)(,+)$/g,function($$,$1,$2) { mul=$2.length; return $1; });
	if(mul !== 0) return write_num(type, fmt, val / Math.pow(10,3*mul));
	if(fmt.indexOf("E") > -1) {
		var idx = fmt.indexOf("E") - fmt.indexOf(".") - 1;
		if(fmt.match(/^#+0.0E\+0$/)) {
		var period = fmt.indexOf("."); if(period === -1) period=fmt.indexOf('E');
			var ee = (Number(val.toExponential(0).substr(2+(val<0))))%period;
			if(ee < 0) ee += period;
			o = (val/Math.pow(10,ee)).toPrecision(idx+1+(period+ee)%period);
			if(!o.match(/[Ee]/)) {
				var fakee = (Number(val.toExponential(0).substr(2+(val<0))));
				if(o.indexOf(".") === -1) o = o[0] + "." + o.substr(1) + "E+" + (fakee - o.length+ee);
				else o += "E+" + (fakee - ee);
				while(o.substr(0,2) === "0.") {
					o = o[0] + o.substr(2,period) + "." + o.substr(2+period);
					o = o.replace(/^0+([1-9])/,"$1").replace(/^0+\./,"0.");
				}
				o = o.replace(/\+-/,"-");
			}
			o = o.replace(/^([+-]?)([0-9]*)\.([0-9]*)[Ee]/,function($$,$1,$2,$3) { return $1 + $2 + $3.substr(0,(period+ee)%period) + "." + $3.substr(ee) + "E"; });
		} else o = val.toExponential(idx);
		if(fmt.match(/E\+00$/) && o.match(/e[+-][0-9]$/)) o = o.substr(0,o.length-1) + "0" + o[o.length-1];
		if(fmt.match(/E\-/) && o.match(/e\+/)) o = o.replace(/e\+/,"e");
		return o.replace("e","E");
	}
	if(fmt[0] === "$") return "$"+write_num(type,fmt.substr(fmt[1]==' '?2:1),val);
	var r, rr, ff, aval = val < 0 ? -val : val, sign = val < 0 ? "-" : "";
	if((r = fmt.match(/# (\?+)([ ]?)\/([ ]?)(\d+)/))) {
		var den = Number(r[4]), rnd = Math.round(aval * den), base = Math.floor(rnd/den);
		var myn = (rnd - base*den), myd = den;
		return sign + (base?base:"") + " " + (myn === 0 ? fill(" ", r[1].length + 1 + r[4].length) : pad(myn,r[1].length," ") + r[2] + "/" + r[3] + pad(myd,r[4].length));
	}
	if(fmt.match(/^#+0+$/)) fmt = fmt.replace(/#/g,"");
	if(fmt.match(/^00+$/)) return (val<0?"-":"")+pad(Math.round(aval),fmt.length);
	if(fmt.match(/^[#?]+$/)) return String(Math.round(val)).replace(/^0$/,"");
	if((r = fmt.match(/^#*0*\.(0+)/))) {
		o = Math.round(val * Math.pow(10,r[1].length));
		rr = String(o/Math.pow(10,r[1].length)).replace(/^([^\.]+)$/,"$1."+r[1]).replace(/\.$/,"."+r[1]).replace(/\.([0-9]*)$/,function($$, $1) { return "." + $1 + fill("0", r[1].length-$1.length); });
		return fmt.match(/0\./) ? rr : rr.replace(/^0\./,".");
	}
	fmt = fmt.replace(/^#+([0.])/, "$1");
	if((r = fmt.match(/^(0*)\.(#*)$/))) {
		o = Math.round(aval*Math.pow(10,r[2].length));
		return sign + String(o / Math.pow(10,r[2].length)).replace(/\.(\d*[1-9])0*$/,".$1").replace(/^([-]?\d*)$/,"$1.").replace(/^0\./,r[1].length?"0.":".");
	}
	if((r = fmt.match(/^#,##0([.]?)$/))) return sign + commaify(String(Math.round(aval)));
	if((r = fmt.match(/^#,##0\.([#0]*0)$/))) {
		rr = Math.round((val-Math.floor(val))*Math.pow(10,r[1].length));
		return val < 0 ? "-" + write_num(type, fmt, -val) : commaify(String(Math.floor(val))) + "." + pad(rr,r[1].length,0);
	}
	if((r = fmt.match(/^#,#*,#0/))) return write_num(type,fmt.replace(/^#,#*,/,""),val);
	if((r = fmt.match(/^([?]+)([ ]?)\/([ ]?)([?]+)/))) {
		rr = Math.min(Math.max(r[1].length, r[4].length),7);
		ff = frac(aval, Math.pow(10,rr)-1, false);
		return sign + (ff[0]||(ff[1] ? "" : "0")) + (ff[1] ? pad(ff[1],rr," ") + r[2] + "/" + r[3] + rpad(ff[2],rr," "): fill(" ", 2*rr+1 + r[2].length + r[3].length));
	}
	if((r = fmt.match(/^# ([?]+)([ ]?)\/([ ]?)([?]+)/))) {
		rr = Math.min(Math.max(r[1].length, r[4].length),7);
		ff = frac(aval, Math.pow(10,rr)-1, true);
		return sign + (ff[0]||(ff[1] ? "" : "0")) + " " + (ff[1] ? pad(ff[1],rr," ") + r[2] + "/" + r[3] + rpad(ff[2],rr," "): fill(" ", 2*rr+1 + r[2].length + r[3].length));
	}
	if((r = fmt.match(/^00,000\.([#0]*0)$/))) {
		rr = val == Math.floor(val) ? 0 : Math.round((val-Math.floor(val))*Math.pow(10,r[1].length));
		return val < 0 ? "-" + write_num(type, fmt, -val) : commaify(String(Math.floor(val))).replace(/^\d,\d{3}$/,"0$&").replace(/^\d*$/,function($$) { return "00," + ($$.length < 3 ? pad(0,3-$$.length) : "") + $$; }) + "." + pad(rr,r[1].length,0);
	}
	switch(fmt) {
		case "0": case "#0": return Math.round(val);
		case "#,###": var x = commaify(String(Math.round(aval))); return x !== "0" ? sign + x : "";
		default:
	}
	throw new Error("unsupported format |" + fmt + "|");
};
function split_fmt(fmt) {
	var out = [];
	var in_str = -1;
	for(var i = 0, j = 0; i < fmt.length; ++i) {
		if(in_str != -1) { if(fmt[i] == '"') in_str = -1; continue; }
		if(fmt[i] == "_" || fmt[i] == "*" || fmt[i] == "\\") { ++i; continue; }
		if(fmt[i] == '"') { in_str = i; continue; }
		if(fmt[i] != ";") continue;
		out.push(fmt.slice(j,i));
		j = i+1;
	}
	out.push(fmt.slice(j));
	if(in_str !=-1) throw new Error("Format |" + fmt + "| unterminated string at " + in_str);
	return out;
}
SSF._split = split_fmt;
function eval_fmt(fmt, v, opts, flen) {
	var out = [], o = "", i = 0, c = "", lst='t', q = {}, dt;
	fixopts(opts = (opts || {}));
	var hr='H';
	/* Tokenize */
	while(i < fmt.length) {
		switch((c = fmt[i])) {
			case 'G': /* General */
				if(fmt.substr(i, i+6).toLowerCase() !== "general")
					throw new Error('unrecognized character ' + fmt[i] + ' in ' +fmt);
				out.push({t:'G',v:'General'}); i+=7; break;
			case '"': /* Literal text */
				for(o="";fmt[++i] !== '"' && i < fmt.length;) o += fmt[i];
				out.push({t:'t', v:o}); ++i; break;
			case '\\': var w = fmt[++i], t = "()".indexOf(w) === -1 ? 't' : w;
				out.push({t:t, v:w}); ++i; break;
			case '_': out.push({t:'t', v:" "}); i+=2; break;
			case '@': /* Text Placeholder */
				out.push({t:'T', v:v}); ++i; break;
			/* Dates */
			case 'M': case 'D': case 'Y': case 'H': case 'S': case 'E':
				c = c.toLowerCase();
				/* falls through */
			case 'm': case 'd': case 'y': case 'h': case 's': case 'e': case 'g':
				if(v < 0) return "";
				if(!dt) dt = parse_date_code(v, opts);
				if(!dt) return "";
				o = fmt[i]; while((fmt[++i]||"").toLowerCase() === c) o+=c;
				if(c === 's' && fmt[i] === '.' && fmt[i+1] === '0') { o+='.'; while(fmt[++i] === '0') o+= '0'; }
				if(c === 'm' && lst.toLowerCase() === 'h') c = 'M'; /* m = minute */
				if(c === 'h') c = hr;
				o = o.toLowerCase();
				q={t:c, v:o}; out.push(q); lst = c; break;
			case 'A':
				if(!dt) dt = parse_date_code(v, opts);
				if(!dt) return "";
				q={t:c,v:"A"};
				if(fmt.substr(i, 3) === "A/P") {q.v = dt.H >= 12 ? "P" : "A"; q.t = 'T'; hr='h';i+=3;}
				else if(fmt.substr(i,5) === "AM/PM") { q.v = dt.H >= 12 ? "PM" : "AM"; q.t = 'T'; i+=5; hr='h'; }
				else { q.t = "t"; i++; }
				out.push(q); lst = c; break;
			case '[': /* TODO: Fix this -- ignore all conditionals and formatting */
				o = c;
				while(fmt[i++] !== ']' && i < fmt.length) o += fmt[i];
				if(o.substr(-1) !== ']') throw 'unterminated "[" block: |' + o + '|';
				if(o.match(/\[[HhMmSs]*\]/)) {
					if(!dt) dt = parse_date_code(v, opts);
					if(!dt) return "";
					out.push({t:'Z', v:o.toLowerCase()});
				} else { o=""; }
				break;
			/* Numbers */
			case '0': case '#': case '.':
				o = c; while("0#?.,E+-%".indexOf(c=fmt[++i]) > -1) o += c;
				out.push({t:'n', v:o}); break;
			case '?':
				o = fmt[i]; while(fmt[++i] === c) o+=c;
				q={t:c, v:o}; out.push(q); lst = c; break;
			case '*': ++i; if(fmt[i] == ' ' || fmt[i] == '*') ++i; break; // **
			case '(': case ')': out.push({t:(flen===1?'t':c),v:c}); ++i; break;
			case '1': case '2': case '3': case '4': case '5': case '6': case '7': case '8': case '9':
				o = fmt[i]; while("0123456789".indexOf(fmt[++i]) > -1) o+=fmt[i];
				out.push({t:'D', v:o}); break;
			case ' ': out.push({t:c,v:c}); ++i; break;
			default:
				if(",$-+/():!^&'~{}<>=€".indexOf(c) === -1)
					throw 'unrecognized character ' + fmt[i] + ' in ' + fmt;
				out.push({t:'t', v:c}); ++i; break;
		}
	}
	var bt = 0;
	for(i=out.length-1, lst='t'; i >= 0; --i) {
		switch(out[i].t) {
			case 'h': case 'H': out[i].t = hr; lst='h'; if(bt < 1) bt = 1; break;
			case 's': if(bt < 3) bt = 3;
			/* falls through */
			case 'd': case 'y': case 'M': case 'e': lst=out[i].t; break;
			case 'm': if(lst === 's') { out[i].t = 'M'; if(bt < 2) bt = 2; } break;
			case 'Z':
				if(bt < 1 && out[i].v.match(/[Hh]/)) bt = 1;
				if(bt < 2 && out[i].v.match(/[Mm]/)) bt = 2;
				if(bt < 3 && out[i].v.match(/[Ss]/)) bt = 3;
		}
	}
	switch(bt) {
		case 0: break;
		case 1:
			if(dt.u >= .5) { dt.u = 0; ++dt.S; }
			if(dt.S >= 60) { dt.S = 0; ++dt.M; }
			if(dt.M >= 60) { dt.M = 0; ++dt.H; }
			break;
		case 2:
			if(dt.u >= .5) { dt.u = 0; ++dt.S; }
			if(dt.S >= 60) { dt.S = 0; ++dt.M; }
			break;
	}
	/* replace fields */
	for(i=0; i < out.length; ++i) {
		switch(out[i].t) {
			case 't': case 'T': case ' ': case 'D': break;
			case 'd': case 'm': case 'y': case 'h': case 'H': case 'M': case 's': case 'e': case 'Z':
				out[i].v = write_date(out[i].t, out[i].v, dt, bt);
				out[i].t = 't'; break;
			case 'n': case '(': case '?':
				var jj = i+1;
				while(out[jj] && ("?D".indexOf(out[jj].t) > -1 || (" t".indexOf(out[jj].t) > -1 && "?t".indexOf((out[jj+1]||{}).t)>-1 && (out[jj+1].t == '?' || out[jj+1].v == '/')) || out[i].t == '(' && (out[jj].t == ')' || out[jj].t == 'n') || out[jj].t == 't' && (out[jj].v == '/' || '$€'.indexOf(out[jj].v) > -1 || (out[jj].v == ' ' && (out[jj+1]||{}).t == '?')))) {
					out[i].v += out[jj].v;
					delete out[jj]; ++jj;
				}
				out[i].v = write_num(out[i].t, out[i].v, (flen >1 && v < 0 && i>0 && out[i-1].v == "-" ? -v:v));
				out[i].t = 't';
				i = jj-1; break;
			case 'G': out[i].t = 't'; out[i].v = general_fmt(v,opts); break;
		}
	}
	return out.map(function(x){return x.v;}).join("");
}
SSF._eval = eval_fmt;
function choose_fmt(fmt, v, o) {
	if(typeof fmt === 'number') fmt = ((o&&o.table) ? o.table : table_fmt)[fmt];
	if(typeof fmt === "string") fmt = split_fmt(fmt);
	var l = fmt.length;
	if(l<4 && fmt[l-1].indexOf("@")>-1) --l;
	switch(fmt.length) {
		case 1: fmt = fmt[0].indexOf("@")>-1 ? ["General", "General", "General", fmt[0]] : [fmt[0], fmt[0], fmt[0], "@"]; break;
		case 2: fmt = fmt[1].indexOf("@")>-1 ? [fmt[0], fmt[0], fmt[0], fmt[1]] : [fmt[0], fmt[1], fmt[0], "@"]; break;
		case 3: fmt = fmt[2].indexOf("@")>-1 ? [fmt[0], fmt[1], fmt[0], fmt[2]] : [fmt[0], fmt[1], fmt[2], "@"]; break;
		case 4: break;
		default: throw "cannot find right format for |" + fmt + "|";
	}
	if(typeof v !== "number") return [fmt.length, fmt[3]];
	return [l, v > 0 ? fmt[0] : v < 0 ? fmt[1] : fmt[2]];
}
var format = function format(fmt,v,o) {
	fixopts(o = (o||{}));
	if(typeof fmt === "string" && fmt.toLowerCase() === "general") return general_fmt(v, o);
	if(typeof fmt === 'number') fmt = (o.table || table_fmt)[fmt];
	var f = choose_fmt(fmt, v, o);
	if(f[1].toLowerCase() === "general") return general_fmt(v,o);
	if(v === true) v = "TRUE"; if(v === false) v = "FALSE";
	if(v === "" || typeof v === "undefined") return "";
	return eval_fmt(f[1], v, o, f[0]);
};

SSF._choose = choose_fmt;
SSF._table = table_fmt;
SSF.load = function(fmt, idx) { table_fmt[idx] = fmt; };
SSF.format = format;
SSF.get_table = function() { return table_fmt; };
SSF.load_table = function(tbl) { for(var i=0; i!=0x0188; ++i) if(tbl[i]) SSF.load(tbl[i], i); };
};
make_ssf(SSF);
var XLSX = {};
(function(XLSX){
XLSX.version = '0.5.14';
var current_codepage, current_cptable, cptable;
if(typeof module !== "undefined" && typeof require !== 'undefined') {
	if(typeof cptable === 'undefined') cptable = require('codepage');
	current_codepage = 1252; current_cptable = cptable[1252];
}
function reset_cp() {
	current_codepage = 1252; if(typeof cptable !== 'undefined') current_cptable = cptable[1252];
}
function _getchar(x) { return String.fromCharCode(x); }

function getdata(data) {
	if(!data) return null;
	if(data.data) return data.name.substr(-4) !== ".bin" ? data.data : data.data.split("").map(function(x) { return x.charCodeAt(0); });
	if(data.asNodeBuffer && typeof Buffer !== 'undefined' && data.name.substr(-4)===".bin") return data.asNodeBuffer();
	if(data.asBinary && data.name.substr(-4) !== ".bin") return data.asBinary();
	if(data._data && data._data.getContent) {
		/* TODO: something far more intelligent */
		if(data.name.substr(-4) === ".bin") return Array.prototype.slice.call(data._data.getContent());
		return Array.prototype.slice.call(data._data.getContent(),0).map(function(x) { return String.fromCharCode(x); }).join("");
	}
	return null;
}

function getzipfile(zip, file) {
	var f = file; if(zip.files[f]) return zip.files[f];
	f = file.toLowerCase(); if(zip.files[f]) return zip.files[f];
	f = f.replace(/\//g,'\\'); if(zip.files[f]) return zip.files[f];
	throw new Error("Cannot find file " + file + " in zip");
}

function getzipdata(zip, file, safe) {
	if(!safe) return getdata(getzipfile(zip, file));
	if(!file) return null;
	try { return getzipdata(zip, file); } catch(e) { return null; }
}

var _fs, jszip;
if(typeof JSZip !== 'undefined') jszip = JSZip;
if (typeof exports !== 'undefined') {
	if (typeof module !== 'undefined' && module.exports) {
		if(typeof Buffer !== 'undefined' && typeof jszip === 'undefined') jszip = require('jszip');
		if(typeof jszip === 'undefined') jszip = require('./jszip').JSZip;
		_fs = require('fs');
	}
}
var attregexg=/(\w+)=((?:")([^"]*)(?:")|(?:')([^']*)(?:'))/g;
var attregex=/(\w+)=((?:")(?:[^"]*)(?:")|(?:')(?:[^']*)(?:'))/;
function parsexmltag(tag) {
	var words = tag.split(/\s+/);
	var z = {'0': words[0]};
	if(words.length === 1) return z;
	(tag.match(attregexg) || []).map(
		function(x){var y=x.match(attregex); z[y[1].replace(/^[a-zA-Z]*:/,"")] = y[2].substr(1,y[2].length-2); });
	return z;
}

function evert(obj) {
	var o = {};
	Object.keys(obj).forEach(function(k) { if(obj.hasOwnProperty(k)) o[obj[k]] = k; });
	return o;
}

var encodings = {
	'&quot;': '"',
	'&apos;': "'",
	'&gt;': '>',
	'&lt;': '<',
	'&amp;': '&'
};
var rencoding = evert(encodings);
var rencstr = "&<>'\"".split("");

// TODO: CP remap (need to read file version to determine OS)
function unescapexml(text){
	var s = text + '';
	for(var y in encodings) s = s.replace(new RegExp(y,'g'), encodings[y]);
	return s.replace(/_x([0-9a-fA-F]*)_/g,function(m,c) {return _chr(parseInt(c,16));});
}
function escapexml(text){
	var s = text + '';
	rencstr.forEach(function(y){s=s.replace(new RegExp(y,'g'), rencoding[y]);});
	return s;
}


function parsexmlbool(value, tag) {
	switch(value) {
		case '0': case 0: case 'false': case 'FALSE': return false;
		case '1': case 1: case 'true': case 'TRUE': return true;
		default: throw "bad boolean value " + value + " in "+(tag||"?");
	}
}

var utf8read = function(orig) {
	var out = [], i = 0, c = 0, c1 = 0, c2 = 0, c3 = 0;
	while (i < orig.length) {
		c = orig.charCodeAt(i++);
		if (c < 128) out.push(_chr(c));
		else {
			c2 = orig.charCodeAt(i++);
			if (c>191 && c<224) out.push(_chr((c & 31) << 6 | c2 & 63));
			else {
				c3 = orig.charCodeAt(i++);
				out.push(_chr((c & 15) << 12 | (c2 & 63) << 6 | c3 & 63));
			}
		}
	}
	return out.join("");
};

// matches <foo>...</foo> extracts content
function matchtag(f,g) {return new RegExp('<'+f+'(?: xml:space="preserve")?>([^\u2603]*)</'+f+'>',(g||"")+"m");}

function parseVector(data) {
	var h = parsexmltag(data);

	var matches = data.match(new RegExp("<vt:" + h.baseType + ">(.*?)</vt:" + h.baseType + ">", 'g'))||[];
	if(matches.length != h.size) throw "unexpected vector length " + matches.length + " != " + h.size;
	var res = [];
	matches.forEach(function(x) {
		var v = x.replace(/<[/]?vt:variant>/g,"").match(/<vt:([^>]*)>(.*)</);
		res.push({v:v[2], t:v[1]});
	});
	return res;
}

function isval(x) { return typeof x !== "undefined" && x !== null; }
function readIEEE754(buf, idx, isLE, nl, ml) {
	if(isLE === undefined) isLE = true;
	if(!nl) nl = 8;
	if(!ml && nl === 8) ml = 52;
	var e, m, el = nl * 8 - ml - 1, eMax = (1 << el) - 1, eBias = eMax >> 1;
	var bits = -7, d = isLE ? -1 : 1, i = isLE ? (nl - 1) : 0, s = buf[idx + i];

	i += d;
	e = s & ((1 << (-bits)) - 1); s >>>= (-bits); bits += el;
	for (; bits > 0; e = e * 256 + buf[idx + i], i += d, bits -= 8);
	m = e & ((1 << (-bits)) - 1); e >>>= (-bits); bits += ml;
	for (; bits > 0; m = m * 256 + buf[idx + i], i += d, bits -= 8);
	if (e === eMax) return m ? NaN : ((s ? -1 : 1) * Infinity);
	else if (e === 0) e = 1 - eBias;
	else { m = m + Math.pow(2, ml); e = e - eBias; }
	return (s ? -1 : 1) * m * Math.pow(2, e - ml);
}

var __toBuffer;
if(typeof Buffer !== "undefined") {
	Buffer.prototype.hexlify= function() { return this.toString('hex'); };
	__toBuffer = function(bufs) { return Buffer.concat(bufs[0]); };
} else {
	__toBuffer = function(bufs) {
		var x = [];
		for(var i = 0; i != bufs[0].length; ++i) { x = x.concat(bufs[0][i]); }
		return x;
	};
}

var __readUInt8 = function(b, idx) { return b.readUInt8 ? b.readUInt8(idx) : b[idx]; };
var __readUInt16LE = function(b, idx) { return b.readUInt16LE ? b.readUInt16LE(idx) : b[idx+1]*(1<<8)+b[idx]; };
var __readInt16LE = function(b, idx) { var u = __readUInt16LE(b,idx); if(!(u & 0x8000)) return u; return (0xffff - u + 1) * -1; };
var __readUInt32LE = function(b, idx) { return b.readUInt32LE ? b.readUInt32LE(idx) : b[idx+3]*(1<<24)+b[idx+2]*(1<<16)+b[idx+1]*(1<<8)+b[idx]; };
var __readInt32LE = function(b, idx) { if(b.readInt32LE) return b.readInt32LE(idx); var u = __readUInt32LE(b,idx); if(!(u & 0x80000000)) return u; return (0xffffffff - u + 1) * -1; };
var __readDoubleLE = function(b, idx) { return b.readDoubleLE ? b.readDoubleLE(idx) : readIEEE754(b, idx||0);};


function ReadShift(size, t) {
	var o = "", oo = [], w, vv, i, loc; t = t || 'u';
	if(size === 'ieee754') { size = 8; t = 'f'; }
	switch(size) {
		case 1: o = __readUInt8(this, this.l); break;
		case 2: o=(t==='u' ? __readUInt16LE : __readInt16LE)(this, this.l); break;
		case 4: o = __readUInt32LE(this, this.l); break;
		case 8: if(t === 'f') { o = __readDoubleLE(this, this.l); break; }
		/* falls through */
		case 16: o = this.toString('hex', this.l,this.l+size); break;

		/* sbcs and dbcs support continue records in the SST way TODO codepages */
		/* TODO: DBCS http://msdn.microsoft.com/en-us/library/cc194788.aspx */
		case 'dbcs': size = 2*t; loc = this.l;
			for(i = 0; i != t; ++i) {
				oo.push(_getchar(__readUInt16LE(this, loc)));
				loc+=2;
			} o = oo.join(""); break;

		case 'sbcs': size = t; o = ""; loc = this.l;
			for(i = 0; i != t; ++i) {
				o += _getchar(__readUInt8(this, loc));
				loc+=1;
			} break;
	}
	this.l+=size; return o;
}

function prep_blob(blob, pos) {
	blob.read_shift = ReadShift.bind(blob);
	blob.l = pos || 0;
	var read = ReadShift.bind(blob);
	return [read];
}

function parsenoop(blob, length) { blob.l += length; }
/* [MS-XLSB] 2.1.4 Record */
var recordhopper = function(data, cb, opts) {
	var tmpbyte, cntbyte, length;
	prep_blob(data, data.l || 0);
	while(data.l < data.length) {
		var RT = data.read_shift(1);
		if(RT & 0x80) RT = (RT & 0x7F) + ((data.read_shift(1) & 0x7F)<<7);
		var R = RecordEnum[RT] || RecordEnum[0xFFFF];
		tmpbyte = data.read_shift(1);
		length = tmpbyte & 0x7F;
		for(cntbyte = 1; cntbyte <4 && (tmpbyte & 0x80); ++cntbyte) length += ((tmpbyte = data.read_shift(1)) & 0x7F)<<(7*cntbyte);
		var d = R.f(data, length, opts);
		if(cb(d, R, RT)) return;
	}
};

/* [MS-XLSB] 2.5.143 */
var parse_StrRun = function(data, length) {
	return { ich: data.read_shift(2), ifnt: data.read_shift(2) };
};

/* [MS-XLSB] 2.1.7.121 */
var parse_RichStr = function(data, length) {
	var start = data.l;
	var flags = data.read_shift(1);
	var fRichStr = flags & 1, fExtStr = flags & 2;
	var str = parse_XLWideString(data);
	var rgsStrRun = [];
	var z = {
		t: str,
		r:"<t>" + escapexml(str) + "</t>",
		h: str
	};
	if(fRichStr) {
		/* TODO: formatted string */
		var dwSizeStrRun = data.read_shift(4);
		for(var i = 0; i != dwSizeStrRun; ++i) rgsStrRun.push(parse_StrRun(data));
		z.r = JSON.stringify(rgsStrRun);
	}
	if(fExtStr) {
		/* TODO: phonetic string */
	}
	data.l = start + length;
	return z;
};

/* [MS-XLSB] 2.5.9 */
function parse_Cell(data) {
	var col = data.read_shift(4);
	var iStyleRef = data.read_shift(2);
	iStyleRef += data.read_shift(1) <<16;
	var fPhShow = data.read_shift(1);
	return { c:col, iStyleRef: iStyleRef };
}

/* [MS-XLSB] 2.5.21 */
var parse_CodeName = function(data, length) { return parse_XLWideString(data, length); };

/* [MS-XLSB] 2.5.114 */
var parse_RelID = function(data, length) { return parse_XLNullableWideString(data, length); };

/* [MS-XLSB] 2.5.122 */
function parse_RkNumber(data) {
	var b = data.slice(data.l, data.l+4);
	var fX100 = b[0] & 1, fInt = b[0] & 2;
	data.l+=4;
	b[0] &= ~3;
	var RK = fInt === 0 ? __readDoubleLE([0,0,0,0,b[0],b[1],b[2],b[3]],0) : __readInt32LE(b,0)>>2;
	return fX100 ? RK/100 : RK;
}

/* [MS-XLSB] 2.5.153 */
var parse_UncheckedRfX = function(data) {
	var cell = {s: {}, e: {}};
	cell.s.r = data.read_shift(4);
	cell.e.r = data.read_shift(4);
	cell.s.c = data.read_shift(4);
	cell.e.c = data.read_shift(4);
	return cell;
};

/* [MS-XLSB] 2.5.166 */
var parse_XLNullableWideString = function(data) {
	var cchCharacters = data.read_shift(4);
	return cchCharacters === 0 || cchCharacters === 0xFFFFFFFF ? "" : data.read_shift('dbcs', cchCharacters);
};

/* [MS-XLSB] 2.5.168 */
var parse_XLWideString = function(data) {
	var cchCharacters = data.read_shift(4);
	return cchCharacters === 0 ? "" : data.read_shift('dbcs', cchCharacters);
};

/* [MS-XLSB] 2.5.171 */
function parse_Xnum(data, length) { return data.read_shift('ieee754'); }

/* [MS-XLSB] 2.5.198.2 */
var BErr = {
	0x00: "#NULL!",
	0x07: "#DIV/0!",
	0x0F: "#VALUE!",
	0x17: "#REF!",
	0x1D: "#NAME?",
	0x24: "#NUM!",
	0x2A: "#N/A",
	0x2B: "#GETTING_DATA",
	0xFF: "#WTF?"
};
var RBErr = evert(BErr);

/* [MS-XLSB] 2.4.321 BrtColor */
function parse_BrtColor(data, length) {
	var read = data.read_shift.bind(data);
	var out = {};
	var d = read(1);
	out.fValidRGB = d & 1;
	out.xColorType = d >>> 1;
	out.index = read(1)
	out.nTintAndShade = read(2, 'i');
	out.bRed   = read(1);
	out.bGreen = read(1);
	out.bBlue  = read(1);
	out.bAlpha = read(1);
}

/* [MS-XLSB 2.5.52 */
function parse_FontFlags(data, length) {
	var d = data.read_shift(1);
	data.l++;
	var out = {
		fItalic: d & 0x2,
		fStrikeout: d & 0x8,
		fOutline: d & 0x10,
		fShadow: d & 0x20,
		fCondense: d & 0x40,
		fExtend: d & 0x80,
	};
	return out;
}
/* Parse a list of <r> tags */
var parse_rs = (function() {
	var tregex = matchtag("t"), rpregex = matchtag("rPr");
	/* 18.4.7 rPr CT_RPrElt */
	var parse_rpr = function(rpr, intro, outro) {
		var font = {};
		(rpr.match(/<[^>]*>/g)||[]).forEach(function(x) {
			var y = parsexmltag(x);
			switch(y[0]) {
				/* 18.8.12 condense CT_BooleanProperty */
				/* ** not required . */
				case '<condense': break;
				/* 18.8.17 extend CT_BooleanProperty */
				/* ** not required . */
				case '<extend': break;
				/* 18.8.36 shadow CT_BooleanProperty */
				/* ** not required . */
				case '<shadow':
					/* falls through */
				case '<shadow/>': break;

				/* 18.4.1 charset CT_IntProperty TODO */
				case '<charset': break;

				/* 18.4.2 outline CT_BooleanProperty TODO */
				case '<outline':
					/* falls through */
				case '<outline/>': break;

				/* 18.4.5 rFont CT_FontName */
				case '<rFont': font.name = y.val; break;

				/* 18.4.11 sz CT_FontSize */
				case '<sz': font.sz = y.val; break;

				/* 18.4.10 strike CT_BooleanProperty */
				case '<strike':
					if(!y.val) break;
					/* falls through */
				case '<strike/>': font.strike = 1; break;
				case '</strike>': break;

				/* 18.4.13 u CT_UnderlineProperty */
				case '<u':
					if(!y.val) break;
					/* falls through */
				case '<u/>': font.u = 1; break;
				case '</u>': break;

				/* 18.8.2 b */
				case '<b':
					if(!y.val) break;
					/* falls through */
				case '<b/>': font.b = 1; break;
				case '</b>': break;

				/* 18.8.26 i */
				case '<i':
					if(!y.val) break;
					/* falls through */
				case '<i/>': font.i = 1; break;
				case '</i>': break;

				/* 18.3.1.15 color CT_Color TODO: tint, theme, auto, indexed */
				case '<color':
					if(y.rgb) font.color = y.rgb.substr(2,6);
					break;

				/* 18.8.18 family ST_FontFamily */
				case '<family': font.family = y.val; break;

				/* 18.4.14 vertAlign CT_VerticalAlignFontProperty TODO */
				case '<vertAlign': break;

				/* 18.8.35 scheme CT_FontScheme TODO */
				case '<scheme': break;

				default:
					if(y[0][2] !== '/') throw 'Unrecognized rich format ' + y[0];
			}
		});
		/* TODO: These should be generated styles, not inline */
		var style = [];
		if(font.b) style.push("font-weight: bold;");
		if(font.i) style.push("font-style: italic;");
		intro.push('<span style="' + style.join("") + '">');
		outro.push("</span>");
	};

	/* 18.4.4 r CT_RElt */
	function parse_r(r) {
		var terms = [[],"",[]];
		/* 18.4.12 t ST_Xstring */
		var t = r.match(tregex);
		if(!isval(t)) return "";
		terms[1] = t[1];

		var rpr = r.match(rpregex);
		if(isval(rpr)) parse_rpr(rpr[1], terms[0], terms[2]);
		return terms[0].join("") + terms[1].replace(/\r\n/g,'<br/>') + terms[2].join("");
	}
	return function(rs) {
		return rs.replace(/<r>/g,"").split(/<\/r>/).map(parse_r).join("");
	};
})();

/* 18.4.8 si CT_Rst */
var parse_si = function(x, opts) {
	var html = opts ? opts.cellHTML : true;
	var z = {};
	if(!x) return null;
	var y;
	/* 18.4.12 t ST_Xstring (Plaintext String) */
	if(x[1] === 't') {
		z.t = utf8read(unescapexml(x.substr(x.indexOf(">")+1).split(/<\/t>/)[0]));
		z.r = x;
		if(html) z.h = z.t;
	}
	/* 18.4.4 r CT_RElt (Rich Text Run) */
	else if((y = x.match(/<r>/))) {
		z.r = x;
		/* TODO: properly parse (note: no other valid child can have body text) */
		z.t = utf8read(unescapexml(x.replace(/<[^>]*>/gm,"")));
		if(html) z.h = parse_rs(x);
	}
	/* 18.4.3 phoneticPr CT_PhoneticPr (TODO: needed for Asian support) */
	/* 18.4.6 rPh CT_PhoneticRun (TODO: needed for Asian support) */
	return z;
};

/* 18.4 Shared String Table */
var parse_sst_xml = function(data, opts) {
	var s = [];
	/* 18.4.9 sst CT_Sst */
	var sst = data.match(new RegExp("<sst([^>]*)>([\\s\\S]*)<\/sst>","m"));
	if(isval(sst)) {
		s = sst[2].replace(/<(?:si|sstItem)>/g,"").split(/<\/(?:si|sstItem)>/).map(function(x) { return parse_si(x, opts); }).filter(function(x) { return x; });
		sst = parsexmltag(sst[1]); s.Count = sst.count; s.Unique = sst.uniqueCount;
	}
	return s;
};

/* [MS-XLSB] 2.4.219 BrtBeginSst */
var parse_BrtBeginSst = function(data, length) {
	return [data.read_shift(4), data.read_shift(4)];
};

/* [MS-XLSB] 2.1.7.45 Shared Strings */
var parse_sst_bin = function(data, opts) {
	var s = [];
	recordhopper(data, function(val, R) {
		switch(R.n) {
			case 'BrtBeginSst': s.Count = val[0]; s.Unique = val[1]; break;
			case 'BrtSSTItem': s.push(val); break;
			case 'BrtEndSst': return true;
			/* TODO: produce a test case with a future record */
			case 'BrtFRTBegin': pass = true; break;
			case 'BrtFRTEnd': pass = false; break;
			default: if(!pass || opts.WTF) throw new Error("Unexpected record " + RT + " " + R.n);
		}
	});
	return s;
};
var styles = {}; // shared styles

/* 18.8.31 numFmts CT_NumFmts */
function parseNumFmts(t) {
	styles.NumberFmt = [];
	for(var y in SSF._table) styles.NumberFmt[y] = SSF._table[y];
	t[0].match(/<[^>]*>/g).forEach(function(x) {
		var y = parsexmltag(x);
		switch(y[0]) {
			case '<numFmts': case '</numFmts>': case '<numFmts/>': break;
			case '<numFmt': {
				var f=unescapexml(y.formatCode), i=parseInt(y.numFmtId,10);
				styles.NumberFmt[i] = f; if(i>0) SSF.load(f,i);
			} break;
			default: throw 'unrecognized ' + y[0] + ' in numFmts';
		}
	});
}

/* 18.8.10 cellXfs CT_CellXfs */
function parseCXfs(t) {
	styles.CellXf = [];
	t[0].match(/<[^>]*>/g).forEach(function(x) {
		var y = parsexmltag(x);
		switch(y[0]) {
			case '<cellXfs': case '<cellXfs/>': case '</cellXfs>': break;

			/* 18.8.45 xf CT_Xf */
			case '<xf': delete y[0];
				if(y.numFmtId) y.numFmtId = parseInt(y.numFmtId, 10);
				styles.CellXf.push(y); break;
			case '</xf>': break;

			/* 18.8.1 alignment CT_CellAlignment */
			case '<alignment': break;

			/* 18.8.33 protection CT_CellProtection */
			case '<protection': case '</protection>': case '<protection/>': break;

			case '<extLst': case '</extLst>': break;
			case '<ext': break;
			default: throw 'unrecognized ' + y[0] + ' in cellXfs';
		}
	});
}

/* 18.8 Styles CT_Stylesheet*/
function parse_sty_xml(data) {
	/* 18.8.39 styleSheet CT_Stylesheet */
	var t;

	/* numFmts CT_NumFmts ? */
	if((t=data.match(/<numFmts([^>]*)>.*<\/numFmts>/))) parseNumFmts(t);

	/* fonts CT_Fonts ? */
	/* fills CT_Fills ? */
	/* borders CT_Borders ? */
	/* cellStyleXfs CT_CellStyleXfs ? */

	/* cellXfs CT_CellXfs ? */
	if((t=data.match(/<cellXfs([^>]*)>.*<\/cellXfs>/))) parseCXfs(t);

	/* dxfs CT_Dxfs ? */
	/* tableStyles CT_TableStyles ? */
	/* colors CT_Colors ? */
	/* extLst CT_ExtensionList ? */

	return styles;
}
/* [MS-XLSB] 2.4.651 BrtFmt */
function parse_BrtFmt(data, length) {
	var ifmt = data.read_shift(2);
	var stFmtCode = parse_XLWideString(data,length-2);
	return [ifmt, stFmtCode];
}

/* [MS-XLSB] 2.4.653 BrtFont TODO */
function parse_BrtFont(data, length) {
	var read = data.read_shift.bind(data);
	var out = {flags:{}};
	out.dyHeight = read(2);
	out.grbit = parse_FontFlags(data, 2);
	out.bls = read(2);
	out.sss = read(2);
	out.uls = read(1);
	out.bFamily = read(1);
	out.bCharSet = read(1);
	data.l++;
	out.brtColor = parse_BrtColor(data, 8);
	out.bFontScheme = read(1);
	out.name = parse_XLWideString(data, length - 21);

	out.flags.Bold = out.bls === 0x02BC;
	out.flags.Italic = out.grbit.fItalic;
	out.flags.Strikeout = out.grbit.fStrikeout;
	out.flags.Outline = out.grbit.fOutline;
	out.flags.Shadow = out.grbit.fShadow;
	out.flags.Condense = out.grbit.fCondense;
	out.flags.Extend = out.grbit.fExtend;
	out.flags.Sub = out.sss & 0x2;
	out.flags.Sup = out.sss & 0x1;
	return out;
}

/* [MS-XLSB] 2.4.816 BrtXF */
function parse_BrtXF(data, length) {
	var ixfeParent = data.read_shift(2);
	var ifmt = data.read_shift(2);
	parsenoop(data, length-4);
	return {ixfe:ixfeParent, ifmt:ifmt };
}

/* [MS-XLSB] 2.1.7.50 Styles */
function parse_sty_bin(data, opts) {
	styles.NumberFmt = [];
	for(var y in SSF._table) styles.NumberFmt[y] = SSF._table[y];

	styles.CellXf = [];
	var state = ""; /* TODO: this should be a stack */
	var pass = false;
	recordhopper(data, function(val, R, RT) {
		switch(R.n) {
			case 'BrtFmt':
				styles.NumberFmt[val[0]] = val[1]; SSF.load(val[1], val[0]);
				break;
			case 'BrtFont': break; /* TODO */
			case 'BrtKnownFonts': break; /* TODO */
			case 'BrtFill': break; /* TODO */
			case 'BrtBorder': break; /* TODO */
			case 'BrtXF':
				if(state === "CELLXFS") {
					styles.CellXf.push(val);
				}
				break; /* TODO */
			case 'BrtStyle': break; /* TODO */
			case 'BrtRowHdr': break; /* TODO */
			case 'BrtCellMeta': break; /* ?? */
			case 'BrtDXF': break; /* TODO */
			case 'BrtMRUColor': break; /* TODO */
			case 'BrtIndexedColor': break; /* TODO */
			case 'BrtBeginStyleSheet': break;
			case 'BrtEndStyleSheet': break;
			case 'BrtBeginTableStyle': break;
			case 'BrtTableStyleElement': break;
			case 'BrtEndTableStyle': break;
			case 'BrtBeginFmts': state = "FMTS"; break;
			case 'BrtEndFmts': state = ""; break;
			case 'BrtBeginFonts': state = "FONTS"; break;
			case 'BrtEndFonts': state = ""; break;
			case 'BrtACBegin': state = "ACFONTS"; break;
			case 'BrtACEnd': state = ""; break;
			case 'BrtBeginFills': state = "FILLS"; break;
			case 'BrtEndFills': state = ""; break;
			case 'BrtBeginBorders': state = "BORDERS"; break;
			case 'BrtEndBorders': state = ""; break;
			case 'BrtBeginCellStyleXFs': state = "CELLSTYLEXFS"; break;
			case 'BrtEndCellStyleXFs': state = ""; break;
			case 'BrtBeginCellXFs': state = "CELLXFS"; break;
			case 'BrtEndCellXFs': state = ""; break;
			case 'BrtBeginStyles': state = "STYLES"; break;
			case 'BrtEndStyles': state = ""; break;
			case 'BrtBeginDXFs': state = "DXFS"; break;
			case 'BrtEndDXFs': state = ""; break;
			case 'BrtBeginTableStyles': state = "TABLESTYLES"; break;
			case 'BrtEndTableStyles': state = ""; break;
			case 'BrtBeginColorPalette': state = "COLORPALETTE"; break;
			case 'BrtEndColorPalette': state = ""; break;
			case 'BrtBeginIndexedColors': state = "INDEXEDCOLORS"; break;
			case 'BrtEndIndexedColors': state = ""; break;
			case 'BrtBeginMRUColors': state = "MRUCOLORS"; break;
			case 'BrtEndMRUColors': state = ""; break;
			case 'BrtFRTBegin': pass = true; break;
			case 'BrtFRTEnd': pass = false; break;
			case 'BrtBeginStyleSheetExt14': break;
			case 'BrtBeginSlicerStyles': break;
			case 'BrtEndSlicerStyles': break;
			case 'BrtBeginTimelineStylesheetExt15': break;
			case 'BrtEndTimelineStylesheetExt15': break;
			case 'BrtBeginTimelineStyles': break;
			case 'BrtEndTimelineStyles': break;
			case 'BrtEndStyleSheetExt14': break;
			default: if(!pass || opts.WTF) throw new Error("Unexpected record " + RT + " " + R.n);
		}
	});
	return styles;
}

var ct2type = {
	"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml": "workbooks", /*XLSX*/
	"application/vnd.ms-excel.sheet.macroEnabled.main+xml":"workbooks", /*XLSM*/
	"application/vnd.ms-excel.sheet.binary.macroEnabled.main":"workbooks", /*XLSB*/

	"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml":"sheets", /*XLS[XM]*/
	"application/vnd.ms-excel.worksheet":"sheets", /*XLSB*/

	"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml":"styles", /*XLS[XM]*/
	"application/vnd.ms-excel.styles":"styles", /*XLSB*/

	"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml": "strs", /*XLS[XM]*/
	"application/vnd.ms-excel.sharedStrings": "strs", /*XLSB*/

	"application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml": "calcchains", /*XLS[XM]*/
	//"application/vnd.ms-excel.calcChain": "calcchains", /*XLSB*/

	"application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml": "comments", /*XLS[XM]*/
	"application/vnd.ms-excel.comments": "comments", /*XLSB*/

	"application/vnd.openxmlformats-package.core-properties+xml": "coreprops",
	"application/vnd.openxmlformats-officedocument.extended-properties+xml": "extprops",
	"application/vnd.openxmlformats-officedocument.custom-properties+xml": "custprops",
	"application/vnd.openxmlformats-officedocument.theme+xml":"themes",
	"foo": "bar"
};

var XMLNS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';

function parseProps(data) {
	var p = { Company:'' }, q = {};
	var strings = ["Application", "DocSecurity", "Company", "AppVersion"];
	var bools = ["HyperlinksChanged","SharedDoc","LinksUpToDate","ScaleCrop"];
	var xtra = ["HeadingPairs", "TitlesOfParts"];
	var xtracp = ["category", "contentStatus", "lastModifiedBy", "lastPrinted", "revision", "version"];
	var xtradc = ["creator", "description", "identifier", "language", "subject", "title"];
	var xtradcterms = ["created", "modified"];
	xtra = xtra.concat(xtracp.map(function(x) { return "cp:" + x; }));
	xtra = xtra.concat(xtradc.map(function(x) { return "dc:" + x; }));
	xtra = xtra.concat(xtradcterms.map(function(x) { return "dcterms:" + x; }));


	strings.forEach(function(f){p[f] = (data.match(matchtag(f))||[])[1];});
	bools.forEach(function(f){p[f] = (data.match(matchtag(f))||[])[1] == "true";});
	xtra.forEach(function(f) {
		var cur = data.match(new RegExp("<" + f + "[^>]*>(.*)<\/" + f + ">"));
		if(cur && cur.length > 0) q[f] = cur[1];
	});

	if(q.HeadingPairs && q.TitlesOfParts) {
		var v = parseVector(q.HeadingPairs);
		var j = 0, widx = 0;
		for(var i = 0; i !== v.length; ++i) {
			switch(v[i].v) {
				case "Worksheets": widx = j; p.Worksheets = +(v[++i].v); break;
				case "Named Ranges": ++i; break; // TODO: Handle Named Ranges
			}
		}
		var parts = parseVector(q.TitlesOfParts).map(function(x) { return utf8read(x.v); });
		p.SheetNames = parts.slice(widx, widx + p.Worksheets);
	}
	p.Creator = q["dc:creator"];
	p.LastModifiedBy = q["cp:lastModifiedBy"];
	p.CreatedDate = new Date(q["dcterms:created"]);
	p.ModifiedDate = new Date(q["dcterms:modified"]);
	return p;
}

/* 15.2.12.2 Custom File Properties Part */
function parseCustomProps(data) {
	var p = {}, name;
	data.match(/<[^>]+>([^<]*)/g).forEach(function(x) {
		var y = parsexmltag(x);
		switch(y[0]) {
			case '<property': name = y.name; break;
			case '</property>': name = null; break;
			default: if (x.indexOf('<vt:') === 0) {
				var toks = x.split('>');
				var type = toks[0].substring(4), text = toks[1];
				/* 22.4.2.32 (CT_Variant). Omit the binary types from 22.4 (Variant Types) */
				switch(type) {
					case 'lpstr': case 'lpwstr': case 'bstr': case 'lpwstr':
						p[name] = unescapexml(text);
						break;
					case 'bool':
						p[name] = parsexmlbool(text, '<vt:bool>');
						break;
					case 'i1': case 'i2': case 'i4': case 'i8': case 'int': case 'uint':
						p[name] = parseInt(text, 10);
						break;
					case 'r4': case 'r8': case 'decimal':
						p[name] = parseFloat(text);
						break;
					case 'filetime': case 'date':
						p[name] = text; // should we make this into a date?
						break;
					case 'cy': case 'error':
						p[name] = unescapexml(text);
						break;
					default:
						console.warn('Unexpected', x, type, toks);
				}
			}
		}
	});
	return p;
}

/* 18.6 Calculation Chain */
function parseDeps(data) {
	var d = [];
	var l = 0, i = 1;
	(data.match(/<[^>]*>/g)||[]).forEach(function(x) {
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

var ctext = {};
function parseCT(data) {
	if(!data || !data.match) return data;
	var ct = { workbooks: [], sheets: [], calcchains: [], themes: [], styles: [],
		coreprops: [], extprops: [], custprops: [], strs:[], comments: [], xmlns: "" };
	(data.match(/<[^>]*>/g)||[]).forEach(function(x) {
		var y = parsexmltag(x);
		switch(y[0]) {
			case '<?xml': break;
			case '<Types': ct.xmlns = y.xmlns; break;
			case '<Default': ctext[y.Extension] = y.ContentType; break;
			case '<Override':
				if(y.ContentType in ct2type)ct[ct2type[y.ContentType]].push(y.PartName);
				break;
		}
	});
	if(ct.xmlns !== XMLNS_CT) throw new Error("Unknown Namespace: " + ct.xmlns);
	ct.calcchain = ct.calcchains.length > 0 ? ct.calcchains[0] : "";
	ct.sst = ct.strs.length > 0 ? ct.strs[0] : "";
	ct.style = ct.styles.length > 0 ? ct.styles[0] : "";
	delete ct.calcchains;
	return ct;
}



/* 9.3.2 OPC Relationships Markup */
function parseRels(data, currentFilePath) {
	if (!data) return data;
	if (currentFilePath.charAt(0) !== '/') {
		currentFilePath = '/'+currentFilePath;
	}
	var rels = {};

	var resolveRelativePathIntoAbsolute = function (to) {
		var toksFrom = currentFilePath.split('/');
		toksFrom.pop(); // folder path
		var toksTo = to.split('/');
		var reversed = [];
		while (toksTo.length !== 0) {
			var tokTo = toksTo.shift();
			if (tokTo === '..') {
				toksFrom.pop();
			} else if (tokTo !== '.') {
				toksFrom.push(tokTo);
			}
		}
		return toksFrom.join('/');
	};

	data.match(/<[^>]*>/g).forEach(function(x) {
		var y = parsexmltag(x);
		/* 9.3.2.2 OPC_Relationships */
		if (y[0] === '<Relationship') {
			var rel = {}; rel.Type = y.Type; rel.Target = y.Target; rel.Id = y.Id; rel.TargetMode = y.TargetMode;
			var canonictarget = resolveRelativePathIntoAbsolute(y.Target);
			rels[canonictarget] = rel;
		}
	});

	return rels;
}


/* 18.7.3 CT_Comment */
function parse_comments_xml(data, opts) {
	if(data.match(/<comments *\/>/)) return [];
	var authors = [];
	var commentList = [];
	data.match(/<authors>([^\u2603]*)<\/authors>/m)[1].split('</author>').forEach(function(x) {
		if(x === "" || x.trim() === "") return;
		authors.push(x.match(/<author[^>]*>(.*)/)[1]);
	});
	data.match(/<commentList>([^\u2603]*)<\/commentList>/m)[1].split('</comment>').forEach(function(x, index) {
		if(x === "" || x.trim() === "") return;
		var y = parsexmltag(x.match(/<comment[^>]*>/)[0]);
		var comment = { author: y.authorId && authors[y.authorId] ? authors[y.authorId] : undefined, ref: y.ref, guid: y.guid };
		var cell = decode_cell(y.ref);
		if(opts.sheetRows && opts.sheetRows <= cell.r) return;
		var textMatch = x.match(/<text>([^\u2603]*)<\/text>/m);
		if (!textMatch || !textMatch[1]) return; // a comment may contain an empty text tag.
		var rt = parse_si(textMatch[1]);
		comment.r = rt.r;
		comment.t = rt.t;
		if(opts.cellHTML) comment.h = rt.h;
		commentList.push(comment);
	});
	return commentList;
}
/* [MS-XLSB] 2.4.28 BrtBeginComment */
var parse_BrtBeginComment = function(data, length) {
	var out = {};
	out.iauthor = data.read_shift(4);
	var rfx = parse_UncheckedRfX(data, 16);
	out.rfx = rfx.s;
	out.ref = encode_cell(rfx.s);
	data.l += 16; /*var guid = parse_GUID(data); */
	return out;
};

/* [MS-XLSB] 2.4.324 BrtCommentAuthor */
var parse_BrtCommentAuthor = parse_XLWideString;

/* [MS-XLSB] 2.4.325 BrtCommentText */
var parse_BrtCommentText = parse_RichStr;

/* [MS-XLSB] 2.1.7.8 Comments */
function parse_comments_bin(data, opts) {
	var out = [];
	var authors = [];
	var c = {};
	var pass = false;
	recordhopper(data, function(val, R, RT) {
		switch(R.n) {
			case 'BrtCommentAuthor': authors.push(val); break;
			case 'BrtBeginComment': c = val; break;
			case 'BrtCommentText': c.t = val.t; c.h = val.h; c.r = val.r; break;
			case 'BrtEndComment':
				c.author = authors[c.iauthor];
				delete c.iauthor;
				if(opts.sheetRows && opts.sheetRows <= c.rfx.r) break;
				delete c.rfx; out.push(c); break;
			case 'BrtBeginComments': break;
			case 'BrtEndComments': break;
			case 'BrtBeginCommentAuthors': break;
			case 'BrtEndCommentAuthors': break;
			case 'BrtBeginCommentList': break;
			case 'BrtEndCommentList': break;
			default: if(!pass || opts.WTF) throw new Error("Unexpected record " + RT + " " + R.n);
		}
	});
	return out;
}

function parse_comments(zip, dirComments, sheets, sheetRels, opts) {
	for(var i = 0; i != dirComments.length; ++i) {
		var canonicalpath=dirComments[i];
		var comments=parse_cmnt(getzipdata(zip, canonicalpath.replace(/^\//,''), true), canonicalpath, opts);
		if(!comments || !comments.length) continue;
		// find the sheets targeted by these comments
		var sheetNames = Object.keys(sheets);
		for(var j = 0; j != sheetNames.length; ++j) {
			var sheetName = sheetNames[j];
			var rels = sheetRels[sheetName];
			if(rels) {
				var rel = rels[canonicalpath];
				if(rel) insertCommentsIntoSheet(sheetName, sheets[sheetName], comments);
			}
		}
	}
}

function insertCommentsIntoSheet(sheetName, sheet, comments) {
	comments.forEach(function(comment) {
		var cell = sheet[comment.ref];
		if (!cell) {
			cell = {};
			sheet[comment.ref] = cell;
			var range = decode_range(sheet["!ref"]);
			var thisCell = decode_cell(comment.ref);
			if(range.s.r > thisCell.r) range.s.r = thisCell.r;
			if(range.e.r < thisCell.r) range.e.r = thisCell.r;
			if(range.s.c > thisCell.c) range.s.c = thisCell.c;
			if(range.e.c < thisCell.c) range.e.c = thisCell.c;
			var encoded = encode_range(range);
			if (encoded !== sheet["!ref"]) sheet["!ref"] = encoded;
		}

		if (!cell.c) cell.c = [];
		var o = {a: comment.author, t: comment.t, r: comment.r};
		if(comment.h) o.h = comment.h;
		cell.c.push(o);
	});
}

/* [MS-XLSB] 2.5.97.4 CellParsedFormula TODO: use similar logic to js-xls */
var parse_CellParsedFormula = function(data, length) {
	var cce = data.read_shift(4);
	return parsenoop(data, length-4);
};
var strs = {}; // shared strings
var _ssfopts = {}; // spreadsheet formatting options

/* 18.3 Worksheets */
function parse_ws_xml(data, opts) {
	if(!data) return data;
	/* 18.3.1.99 worksheet CT_Worksheet */
	var s = {};

	/* 18.3.1.35 dimension CT_SheetDimension ? */
	var ref = data.match(/<dimension ref="([^"]*)"\s*\/>/);
	if(ref && ref.length == 2 && ref[1].indexOf(":") !== -1) s["!ref"] = ref[1];

	/* 18.3.1.55 mergeCells CT_MergeCells */
	var mergecells = [];
	if(data.match(/<\/mergeCells>/)) {
		var merges = data.match(/<mergeCell ref="([A-Z0-9:]+)"\s*\/>/g);
		mergecells = merges.map(function(range) { 
			return decode_range(/<mergeCell ref="([A-Z0-9:]+)"\s*\/>/.exec(range)[1]);
		});
	}

	var refguess = {s: {r:1000000, c:1000000}, e: {r:0, c:0} };
	var q = ["v","f"];
	var sidx = 0;

	/* 18.3.1.80 sheetData CT_SheetData ? */
	if(!data.match(/<sheetData *\/>/))
	data.match(/<sheetData>([^\u2603]*)<\/sheetData>/m)[1].split("</row>").forEach(function(x) {
		if(x === "" || x.trim() === "") return;

		/* 18.3.1.73 row CT_Row */
		var row = parsexmltag(x.match(/<row[^>]*>/)[0]);
		if(opts.sheetRows && opts.sheetRows < +row.r) return;
		if(refguess.s.r > row.r - 1) refguess.s.r = row.r - 1;
		if(refguess.e.r < row.r - 1) refguess.e.r = row.r - 1;
		/* 18.3.1.4 c CT_Cell */
		var cells = x.substr(x.indexOf('>')+1).split(/<c /);
		cells.forEach(function(c, idx) { if(c === "" || c.trim() === "") return;
			var cref = c.match(/r=["']([^"']*)["']/);
			c = "<c " + c;
			if(cref && cref.length == 2) idx = decode_cell(cref[1]).c;
			var cell = parsexmltag((c.match(/<c[^>]*>/)||[c])[0]); delete cell[0];
			var d = c.substr(c.indexOf('>')+1);
			var p = {};
			q.forEach(function(f){var x=d.match(matchtag(f));if(x)p[f]=unescapexml(x[1]);});

			/* SCHEMA IS ACTUALLY INCORRECT HERE.  IF A CELL HAS NO T, EMIT "" */
			if(cell.t === undefined && p.v === undefined) {
				if(!opts.sheetStubs) return;
				p.t = "str"; p.v = undefined;
			}
			else p.t = (cell.t ? cell.t : "n"); // default is "n" in schema
			if(refguess.s.c > idx) refguess.s.c = idx;
			if(refguess.e.c < idx) refguess.e.c = idx;
			/* 18.18.11 t ST_CellType */
			switch(p.t) {
				case 'n': p.v = parseFloat(p.v); break;
				case 's': {
					sidx = parseInt(p.v, 10);
					p.v = strs[sidx].t;
					p.r = strs[sidx].r;
					if(opts.cellHTML) p.h = strs[sidx].h;
				} break;
				case 'str': if(p.v) p.v = utf8read(p.v); break;
				case 'inlineStr':
					var is = d.match(/<is>(.*)<\/is>/);
					is = is ? parse_si(is[1]) : {t:"",r:""};
					p.t = 'str'; p.v = is.t;
					break; // inline string
				case 'b': if(typeof p.v !== 'boolean') p.v = parsexmlbool(p.v); break;
				case 'd': /* TODO: date1904 logic */
					var epoch = Date.parse(p.v);
					p.v = (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
					p.t = 'n';
					break;
				/* in case of error, stick value in .raw */
				case 'e': p.raw = RBErr[p.v]; break;
			}

			/* formatting */
			var fmtid = 0;
			if(cell.s && styles.CellXf) {
				var cf = styles.CellXf[cell.s];
				if(cf && cf.numFmtId) fmtid = cf.numFmtId;
			}
			try {
				p.w = SSF.format(fmtid,p.v,_ssfopts);
				if(opts.cellNF) p.z = SSF._table[fmtid];
			} catch(e) { if(opts.WTF) throw e; }
			s[cell.r] = p;
		});
	});
	if(!s["!ref"]) s["!ref"] = encode_range(refguess);
	if(opts.sheetRows) {
		var tmpref = decode_range(s["!ref"]);
		if(opts.sheetRows < +tmpref.e.r) {
			tmpref.e.r = opts.sheetRows - 1;
			if(tmpref.e.r > refguess.e.r) tmpref.e.r = refguess.e.r;
			if(tmpref.e.r < tmpref.s.r) tmpref.s.r = tmpref.e.r;
			if(tmpref.e.c > refguess.e.c) tmpref.e.c = refguess.e.c;
			if(tmpref.e.c < tmpref.s.c) tmpref.s.c = tmpref.e.c;
			s["!fullref"] = s["!ref"];
			s["!ref"] = encode_range(tmpref);
		}
	}
	if(mergecells.length > 0) s["!merges"] = mergecells;
	return s;
}


/* [MS-XLSB] 2.4.718 BrtRowHdr */
var parse_BrtRowHdr = function(data, length) {
	var z = {};
	z.r = data.read_shift(4);
	data.l += length-4;
	return z;
};

/* [MS-XLSB] 2.4.812 BrtWsDim */
var parse_BrtWsDim = parse_UncheckedRfX;

/* [MS-XLSB] 2.4.815 BrtWsProp */
var parse_BrtWsProp = function(data, length) {
	var z = {};
	/* TODO: pull flags */
	data.l += 19;
	z.name = parse_CodeName(data, length - 19);
	return z;
};

/* [MS-XLSB] 2.4.303 BrtCellBlank */
var parse_BrtCellBlank = parsenoop;

/* [MS-XLSB] 2.4.304 BrtCellBool */
var parse_BrtCellBool = function(data, length) {
	var cell = parse_Cell(data);
	var fBool = data.read_shift(1);
	return [cell, fBool, 'b'];
};

/* [MS-XLSB] 2.4.305 BrtCellError */
var parse_BrtCellError = function(data, length) {
	var cell = parse_Cell(data);
	var fBool = data.read_shift(1);
	return [cell, fBool, 'e'];
};

/* [MS-XLSB] 2.4.308 BrtCellIsst */
var parse_BrtCellIsst = function(data, length) {
	var cell = parse_Cell(data);
	var isst = data.read_shift(4);
	return [cell, isst, 's'];
};

/* [MS-XLSB] 2.4.310 BrtCellReal */
var parse_BrtCellReal = function(data, length) {
	var cell = parse_Cell(data);
	var value = parse_Xnum(data);
	return [cell, value, 'n'];
};

/* [MS-XLSB] 2.4.311 BrtCellRk */
var parse_BrtCellRk = function(data, length) {
	var cell = parse_Cell(data);
	var value = parse_RkNumber(data);
	return [cell, value, 'n'];
};

/* [MS-XLSB] 2.4.314 BrtCellSt */
var parse_BrtCellSt = function(data, length) {
	var cell = parse_Cell(data);
	var value = parse_XLWideString(data);
	return [cell, value, 'str'];
};

/* [MS-XLSB] 2.4.647 BrtFmlaBool */
var parse_BrtFmlaBool = function(data, length, opts) {
	var cell = parse_Cell(data);
	var value = data.read_shift(1);
	var o = [cell, value, 'b'];
	if(opts.cellFormula) {
		var formula = parse_CellParsedFormula(data, length-9);
		o[3] = ""; /* TODO */
	}
	else data.l += length-9;
	return o;
};

/* [MS-XLSB] 2.4.648 BrtFmlaError */
var parse_BrtFmlaError = function(data, length, opts) {
	var cell = parse_Cell(data);
	var value = data.read_shift(1);
	var o = [cell, value, 'e'];
	if(opts.cellFormula) {
		var formula = parse_CellParsedFormula(data, length-9);
		o[3] = ""; /* TODO */
	}
	else data.l += length-9;
	return o;
};

/* [MS-XLSB] 2.4.649 BrtFmlaNum */
var parse_BrtFmlaNum = function(data, length, opts) {
	var cell = parse_Cell(data);
	var value = parse_Xnum(data);
	var o = [cell, value, 'n'];
	if(opts.cellFormula) {
		var formula = parse_CellParsedFormula(data, length - 16);
		o[3] = ""; /* TODO */
	}
	else data.l += length-16;
	return o;
};

/* [MS-XLSB] 2.4.650 BrtFmlaString */
var parse_BrtFmlaString = function(data, length, opts) {
	var start = data.l;
	var cell = parse_Cell(data);
	var value = parse_XLWideString(data);
	var o = [cell, value, 'str'];
	if(opts.cellFormula) {
		var formula = parse_CellParsedFormula(data, start + length - data.l);
		o[3] = ""; /* TODO */
	}
	else data.l = start + length;
	return o;
};

/* [MS-XLSB] 2.4.676 BrtMergeCell */
var parse_BrtMergeCell = parse_UncheckedRfX;

/* [MS-XLSB] 2.1.7.61 Worksheet */
var parse_ws_bin = function(data, opts) {
	if(!data) return data;
	var s = {};

	var ref;
	var refguess = {s: {r:1000000, c:1000000}, e: {r:0, c:0} };

	var pass = false, end = false;
	var row, p, cf;
	var mergecells = [];
	recordhopper(data, function(val, R) {
		if(end) return;
		switch(R.n) {
			case 'BrtWsDim': ref = val; break;
			case 'BrtRowHdr':
				row = val;
				if(opts.sheetRows && opts.sheetRows <= row.r) end=true;
				break;

			case 'BrtFmlaBool':
			case 'BrtFmlaError':
			case 'BrtFmlaNum':
			case 'BrtFmlaString':
			case 'BrtCellBool':
			case 'BrtCellError':
			case 'BrtCellIsst':
			case 'BrtCellReal':
			case 'BrtCellRk':
			case 'BrtCellSt':
				p = {t:val[2]};
				switch(val[2]) {
					case 'n': p.v = val[1]; break;
					case 's': p.v = strs[val[1]].t; p.r = strs[val[1]].r; break;
					case 'b': p.v = val[1] ? true : false; break;
					case 'e': p.raw = val[1]; p.v = BErr[p.raw]; break;
					case 'str': p.v = utf8read(val[1]); break;
				}
				if(opts.cellFormula && val.length > 3) p.f = val[3];
				if((cf = styles.CellXf[val[0].iStyleRef])) try {
					p.w = SSF.format(cf.ifmt,p.v,_ssfopts);
					if(opts.cellNF) p.z = SSF._table[cf.ifmt];
				} catch(e) { if(opts.WTF) throw e; }
				s[encode_cell({c:val[0].c,r:row.r})] = p;
				if(refguess.s.r > row.r) refguess.s.r = row.r;
				if(refguess.s.c > val[0].c) refguess.s.c = val[0].c;
				if(refguess.e.r < row.r) refguess.e.r = row.r;
				if(refguess.e.c < val[0].c) refguess.e.c = val[0].c;
				break;

			case 'BrtCellBlank': break; // (blank cell)

			/* Merge Cells */
			case 'BrtBeginMergeCells': break;
			case 'BrtEndMergeCells': break;
			case 'BrtMergeCell': mergecells.push(val); break;
			
			case 'BrtArrFmla': break; // TODO
			case 'BrtShrFmla': break; // TODO
			case 'BrtBeginSheet': break;
			case 'BrtWsProp': break; // TODO
			case 'BrtSheetCalcProp': break; // TODO
			case 'BrtBeginWsViews': break; // TODO
			case 'BrtBeginWsView': break; // TODO
			case 'BrtPane': break; // TODO
			case 'BrtSel': break; // TODO
			case 'BrtEndWsView': break; // TODO
			case 'BrtEndWsViews': break; // TODO
			case 'BrtACBegin': break; // TODO
			case 'BrtRwDescent': break; // TODO
			case 'BrtACEnd': break; // TODO
			case 'BrtWsFmtInfoEx14': break; // TODO
			case 'BrtWsFmtInfo': break; // TODO
			case 'BrtBeginColInfos': break; // TODO
			case 'BrtColInfo': break; // TODO
			case 'BrtEndColInfos': break; // TODO
			case 'BrtBeginSheetData': break; // TODO
			case 'BrtEndSheetData': break; // TODO
			case 'BrtSheetProtection': break; // TODO
			case 'BrtPrintOptions': break; // TODO
			case 'BrtMargins': break; // TODO
			case 'BrtPageSetup': break; // TODO
			case 'BrtFRTBegin': pass = true; break;
			case 'BrtFRTEnd': pass = false; break;
			case 'BrtEndSheet': break; // TODO
			case 'BrtHLink': break; // TODO
			case 'BrtDrawing': break; // TODO
			case 'BrtLegacyDrawing': break; // TODO
			case 'BrtLegacyDrawingHF': break; // TODO
			case 'BrtPhoneticInfo': break; // TODO
			case 'BrtBeginHeaderFooter': break; // TODO
			case 'BrtEndHeaderFooter': break; // TODO
			case 'BrtBrk': break; // TODO
			case 'BrtBeginRwBrk': break; // TODO
			case 'BrtEndRwBrk': break; // TODO
			case 'BrtBeginColBrk': break; // TODO
			case 'BrtEndColBrk': break; // TODO
			case 'BrtBeginUserShViews': break; // TODO
			case 'BrtBeginUserShView': break; // TODO
			case 'BrtEndUserShView': break; // TODO
			case 'BrtEndUserShViews': break; // TODO
			case 'BrtBkHim': break; // TODO
			case 'BrtBeginOleObjects': break; // TODO
			case 'BrtOleObject': break; // TODO
			case 'BrtEndOleObjects': break; // TODO
			case 'BrtBeginListParts': break; // TODO
			case 'BrtListPart': break; // TODO
			case 'BrtEndListParts': break; // TODO
			case 'BrtBeginSortState': break; // TODO
			case 'BrtBeginSortCond': break; // TODO
			case 'BrtEndSortCond': break; // TODO
			case 'BrtEndSortState': break; // TODO
			case 'BrtBeginConditionalFormatting': break; // TODO
			case 'BrtEndConditionalFormatting': break; // TODO
			case 'BrtBeginCFRule': break; // TODO
			case 'BrtEndCFRule': break; // TODO
			case 'BrtBeginDVals': break; // TODO
			case 'BrtDVal': break; // TODO
			case 'BrtEndDVals': break; // TODO
			case 'BrtRangeProtection': break; // TODO
			case 'BrtBeginDCon': break; // TODO
			case 'BrtEndDCon': break; // TODO
			case 'BrtBeginDRefs': break;
			case 'BrtDRef': break;
			case 'BrtEndDRefs': break;

			/* ActiveX */
			case 'BrtBeginActiveXControls': break;
			case 'BrtActiveX': break;
			case 'BrtEndActiveXControls': break;

			/* AutoFilter */
			case 'BrtBeginAFilter': break;
			case 'BrtEndAFilter': break;
			case 'BrtBeginFilterColumn': break;
			case 'BrtBeginFilters': break;
			case 'BrtFilter': break; 
			case 'BrtEndFilters': break;
			case 'BrtEndFilterColumn': break;
			case 'BrtDynamicFilter': break;
			case 'BrtTop10Filter': break;
			case 'BrtBeginCustomFilters': break;
			case 'BrtCustomFilter': break;
			case 'BrtEndCustomFilters': break;

			/* Cell Watch */
			case 'BrtBeginCellWatches': break;
			case 'BrtCellWatch': break;
			case 'BrtEndCellWatches': break;

			/* Table */
			case 'BrtTable': break;

			/* Ignore Cell Errors */
			case 'BrtBeginCellIgnoreECs': break;
			case 'BrtCellIgnoreEC': break;
			case 'BrtEndCellIgnoreECs': break;

			default: if(!pass || opts.WTF) throw new Error("Unexpected record " + R.n);
		}
	}, opts);
	s["!ref"] = encode_range(ref);
	if(opts.sheetRows) {
		var tmpref = decode_range(s["!ref"]);
		if(opts.sheetRows < +tmpref.e.r) {
			tmpref.e.r = opts.sheetRows - 1;
			if(tmpref.e.r > refguess.e.r) tmpref.e.r = refguess.e.r;
			if(tmpref.e.r < tmpref.s.r) tmpref.s.r = tmpref.e.r;
			if(tmpref.e.c > refguess.e.c) tmpref.e.c = refguess.e.c;
			if(tmpref.e.c < tmpref.s.c) tmpref.s.c = tmpref.e.c;
			s["!fullref"] = s["!ref"];
			s["!ref"] = encode_range(tmpref);
		}
	}
	if(mergecells.length > 0) s["!merges"] = mergecells;
	return s;
};

/* 18.2.28 (CT_WorkbookProtection) Defaults */
var WBPropsDef = {
	allowRefreshQuery: '0',
	autoCompressPictures: '1',
	backupFile: '0',
	checkCompatibility: '0',
	codeName: '',
	date1904: '0',
	dateCompatibility: '1',
	//defaultThemeVersion: '0',
	filterPrivacy: '0',
	hidePivotFieldList: '0',
	promptedSolutions: '0',
	publishItems: '0',
	refreshAllConnections: false,
	saveExternalLinkValues: '1',
	showBorderUnselectedTables: '1',
	showInkAnnotation: '1',
	showObjects: 'all',
	showPivotChartFilter: '0'
	//updateLinks: 'userSet'
};

/* 18.2.30 (CT_BookView) Defaults */
var WBViewDef = {
	activeTab: '0',
	autoFilterDateGrouping: '1',
	firstSheet: '0',
	minimized: '0',
	showHorizontalScroll: '1',
	showSheetTabs: '1',
	showVerticalScroll: '1',
	tabRatio: '600',
	visibility: 'visible'
	//window{Height,Width}, {x,y}Window
};

/* 18.2.19 (CT_Sheet) Defaults */
var SheetDef = {
	state: 'visible'
};

/* 18.2.2  (CT_CalcPr) Defaults */
var CalcPrDef = {
	calcCompleted: 'true',
	calcMode: 'auto',
	calcOnSave: 'true',
	concurrentCalc: 'true',
	fullCalcOnLoad: 'false',
	fullPrecision: 'true',
	iterate: 'false',
	iterateCount: '100',
	iterateDelta: '0.001',
	refMode: 'A1'
};

/* 18.2.3 (CT_CustomWorkbookView) Defaults */
var CustomWBViewDef = {
	autoUpdate: 'false',
	changesSavedWin: 'false',
	includeHiddenRowCol: 'true',
	includePrintSettings: 'true',
	maximized: 'false',
	minimized: 'false',
	onlySync: 'false',
	personalView: 'false',
	showComments: 'commIndicator',
	showFormulaBar: 'true',
	showHorizontalScroll: 'true',
	showObjects: 'all',
	showSheetTabs: 'true',
	showStatusbar: 'true',
	showVerticalScroll: 'true',
	tabRatio: '600',
	xWindow: '0',
	yWindow: '0'
};
var XMLNS_WB = [
	'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
	'http://schemas.microsoft.com/office/excel/2006/main',
	'http://schemas.microsoft.com/office/excel/2006/2'
];

/* 18.2 Workbook */
function parse_wb_xml(data) {
	var wb = { AppVersion:{}, WBProps:{}, WBView:[], Sheets:[], CalcPr:{}, xmlns: "" };
	var pass = false;
	data.match(/<[^>]*>/g).forEach(function(x) {
		var y = parsexmltag(x);
		switch(y[0]) {
			case '<?xml': break;

			/* 18.2.27 workbook CT_Workbook 1 */
			case '<workbook': wb.xmlns = y.xmlns; break;
			case '</workbook>': break;

			/* 18.2.13 fileVersion CT_FileVersion ? */
			case '<fileVersion': delete y[0]; wb.AppVersion = y; break;
			case '<fileVersion/>': break;

			/* 18.2.12 fileSharing CT_FileSharing ? */
			case '<fileSharing': case '<fileSharing/>': break;

			/* 18.2.28 workbookPr CT_WorkbookPr ? */
			case '<workbookPr': delete y[0]; wb.WBProps = y; break;
			case '<workbookPr/>': delete y[0]; wb.WBProps = y; break;

			/* 18.2.29 workbookProtection CT_WorkbookProtection ? */
			case '<workbookProtection/>': break;

			/* 18.2.1  bookViews CT_BookViews ? */
			case '<bookViews>': case '</bookViews>': break;
			/* 18.2.30   workbookView CT_BookView + */
			case '<workbookView': delete y[0]; wb.WBView.push(y); break;

			/* 18.2.20 sheets CT_Sheets 1 */
			case '<sheets>': case '</sheets>': break; // aggregate sheet
			/* 18.2.19   sheet CT_Sheet + */
			case '<sheet': delete y[0]; y.name = utf8read(y.name); wb.Sheets.push(y); break;

			/* 18.2.15 functionGroups CT_FunctionGroups ? */
			case '<functionGroups': case '<functionGroups/>': break;
			/* 18.2.14   functionGroup CT_FunctionGroup + */
			case '<functionGroup': break;

			/* 18.2.9  externalReferences CT_ExternalReferences ? */
			case '<externalReferences': case '</externalReferences>': break;
			/* 18.2.8    externalReference CT_ExternalReference + */
			case '<externalReference': break;

			/* 18.2.6  definedNames CT_DefinedNames ? */
			case '<definedNames/>': break;
			case '<definedNames>': pass=true; break;
			case '</definedNames>': pass=false; break;
			/* 18.2.5    definedName CT_DefinedName + */
			case '<definedName': case '<definedName/>': case '</definedName>': break;

			/* 18.2.2  calcPr CT_CalcPr ? */
			case '<calcPr': delete y[0]; wb.CalcPr = y; break;
			case '<calcPr/>': delete y[0]; wb.CalcPr = y; break;

			/* 18.2.16 oleSize CT_OleSize ? (ref required) */
			case '<oleSize': break;

			/* 18.2.4  customWorkbookViews CT_CustomWorkbookViews ? */
			case '<customWorkbookViews>': case '</customWorkbookViews>': case '<customWorkbookViews': break;
			/* 18.2.3    customWorkbookView CT_CustomWorkbookView + */
			case '<customWorkbookView': case '</customWorkbookView>': break;

			/* 18.2.18 pivotCaches CT_PivotCaches ? */
			case '<pivotCaches>': case '</pivotCaches>': case '<pivotCaches': break;
			/* 18.2.17 pivotCache CT_PivotCache ? */
			case '<pivotCache': break;

			/* 18.2.21 smartTagPr CT_SmartTagPr ? */
			case '<smartTagPr': case '<smartTagPr/>': break;

			/* 18.2.23 smartTagTypes CT_SmartTagTypes ? */
			case '<smartTagTypes': case '<smartTagTypes>': case '</smartTagTypes>': break;
			/* 18.2.22   smartTagType CT_SmartTagType ? */
			case '<smartTagType': break;

			/* 18.2.24 webPublishing CT_WebPublishing ? */
			case '<webPublishing': case '<webPublishing/>': break;

			/* 18.2.11 fileRecoveryPr CT_FileRecoveryPr ? */
			case '<fileRecoveryPr': case '<fileRecoveryPr/>': break;

			/* 18.2.26 webPublishObjects CT_WebPublishObjects ? */
			case '<webPublishObjects>': case '<webPublishObjects': case '</webPublishObjects>': break;
			/* 18.2.25 webPublishObject CT_WebPublishObject ? */
			case '<webPublishObject': break;

			/* 18.2.10 extLst CT_ExtensionList ? */
			case '<extLst>': case '</extLst>': case '<extLst/>': break;
			/* 18.2.7    ext CT_Extension + */
			case '<ext': pass=true; break; //TODO: check with versions of excel
			case '</ext>': pass=false; break;

			/* Others */
			case '<mx:ArchID': break;
			case '<mc:AlternateContent': pass=true; break;
			case '</mc:AlternateContent>': pass=false; break;
		}
	});
	if(XMLNS_WB.indexOf(wb.xmlns) === -1) throw new Error("Unknown Namespace: " + wb.xmlns);

	var z;
	/* defaults */
	for(z in WBPropsDef) if(typeof wb.WBProps[z] === 'undefined') wb.WBProps[z] = WBPropsDef[z];
	for(z in CalcPrDef) if(typeof wb.CalcPr[z] === 'undefined') wb.CalcPr[z] = CalcPrDef[z];

	wb.WBView.forEach(function(w){for(var z in WBViewDef) if(typeof w[z] === 'undefined') w[z]=WBViewDef[z]; });
	wb.Sheets.forEach(function(w){for(var z in SheetDef) if(typeof w[z] === 'undefined') w[z]=SheetDef[z]; });

	_ssfopts.date1904 = parsexmlbool(wb.WBProps.date1904, 'date1904');

	return wb;
}

/* [MS-XLSB] 2.4.301 BrtBundleSh */
var parse_BrtBundleSh = function(data, length) {
	var z = {};
	z.hsState = data.read_shift(4); //ST_SheetState
	z.iTabID = data.read_shift(4);
	z.strRelID = parse_RelID(data,length-8);
	z.name = parse_XLWideString(data);
	return z;
};

/* [MS-XLSB] 2.1.7.60 Workbook */
var parse_wb_bin = function(data, opts) {
	var wb = { AppVersion:{}, WBProps:{}, WBView:[], Sheets:[], CalcPr:{}, xmlns: "" };
	var pass = false, z;

	recordhopper(data, function(val, R) {
		switch(R.n) {
			case 'BrtBundleSh': wb.Sheets.push(val); break;

			case 'BrtBeginBook': break;
			case 'BrtFileVersion': break;
			case 'BrtWbProp': break;
			case 'BrtACBegin': break;
			case 'BrtAbsPath15': break;
			case 'BrtACEnd': break;
			/*case 'BrtBookProtectionIso': break;*/
			case 'BrtBookProtection': break;
			case 'BrtBeginBookViews': break;
			case 'BrtBookView': break;
			case 'BrtEndBookViews': break;
			case 'BrtBeginBundleShs': break;
			case 'BrtEndBundleShs': break;
			case 'BrtBeginFnGroup': break;
			case 'BrtEndFnGroup': break;
			case 'BrtBeginExternals': break;
			case 'BrtSupSelf': break;
			case 'BrtSupBookSrc': break;
			case 'BrtExternSheet': break;
			case 'BrtEndExternals': break;
			case 'BrtName': break;
			case 'BrtCalcProp': break;
			case 'BrtUserBookView': break;
			case 'BrtBeginPivotCacheIDs': break;
			case 'BrtBeginPivotCacheID': break;
			case 'BrtEndPivotCacheID': break;
			case 'BrtEndPivotCacheIDs': break;
			case 'BrtWebOpt': break;
			case 'BrtFileRecover': break;
			case 'BrtFileSharing': break;
			/*case 'BrtBeginWebPubItems': break;
			case 'BrtBeginWebPubItem': break;
			case 'BrtEndWebPubItem': break;
			case 'BrtEndWebPubItems': break;*/
			case 'BrtFRTBegin': pass = true; break;
			case 'BrtFRTEnd': pass = false; break;
			case 'BrtEndBook': break;
			default: if(!pass) throw new Error("Unexpected record " + R.n);
		}
	});

	/* defaults */
	for(z in WBPropsDef) if(typeof wb.WBProps[z] === 'undefined') wb.WBProps[z] = WBPropsDef[z];
	for(z in CalcPrDef) if(typeof wb.CalcPr[z] === 'undefined') wb.CalcPr[z] = CalcPrDef[z];

	wb.WBView.forEach(function(w){for(var z in WBViewDef) if(typeof w[z] === 'undefined') w[z]=WBViewDef[z]; });
	wb.Sheets.forEach(function(w){for(var z in SheetDef) if(typeof w[z] === 'undefined') w[z]=SheetDef[z]; });

	_ssfopts.date1904 = parsexmlbool(wb.WBProps.date1904, 'date1904');

	return wb;
};
function parse_wb(data, name, opts) {
	return name.substr(-4)===".bin" ? parse_wb_bin(data, opts) : parse_wb_xml(data, opts);
}

function parse_ws(data, name, opts) {
	return name.substr(-4)===".bin" ? parse_ws_bin(data, opts) : parse_ws_xml(data, opts);
}

function parse_sty(data, name, opts) {
	return name.substr(-4)===".bin" ? parse_sty_bin(data, opts) : parse_sty_xml(data, opts);
}

function parse_sst(data, name, opts) {
	return name.substr(-4)===".bin" ? parse_sst_bin(data, opts) : parse_sst_xml(data, opts);
}

function parse_cmnt(data, name, opts) {
	return name.substr(-4)===".bin" ? parse_comments_bin(data, opts) : parse_comments_xml(data, opts);
}
/* [MS-XLSB] 2.3 Record Enumeration */
var RecordEnum = {
	0x0000: { n:"BrtRowHdr", f:parse_BrtRowHdr },
	0x0001: { n:"BrtCellBlank", f:parse_BrtCellBlank },
	0x0002: { n:"BrtCellRk", f:parse_BrtCellRk },
	0x0003: { n:"BrtCellError", f:parse_BrtCellError },
	0x0004: { n:"BrtCellBool", f:parse_BrtCellBool },
	0x0005: { n:"BrtCellReal", f:parse_BrtCellReal },
	0x0006: { n:"BrtCellSt", f:parse_BrtCellSt },
	0x0007: { n:"BrtCellIsst", f:parse_BrtCellIsst },
	0x0008: { n:"BrtFmlaString", f:parse_BrtFmlaString },
	0x0009: { n:"BrtFmlaNum", f:parse_BrtFmlaNum },
	0x000A: { n:"BrtFmlaBool", f:parse_BrtFmlaBool },
	0x000B: { n:"BrtFmlaError", f:parse_BrtFmlaError },
	0x0013: { n:"BrtSSTItem", f:parse_RichStr },
	0x0014: { n:"BrtPCDIMissing", f:parsenoop },
	0x0015: { n:"BrtPCDINumber", f:parsenoop },
	0x0016: { n:"BrtPCDIBoolean", f:parsenoop },
	0x0017: { n:"BrtPCDIError", f:parsenoop },
	0x0018: { n:"BrtPCDIString", f:parsenoop },
	0x0019: { n:"BrtPCDIDatetime", f:parsenoop },
	0x001A: { n:"BrtPCDIIndex", f:parsenoop },
	0x001B: { n:"BrtPCDIAMissing", f:parsenoop },
	0x001C: { n:"BrtPCDIANumber", f:parsenoop },
	0x001D: { n:"BrtPCDIABoolean", f:parsenoop },
	0x001E: { n:"BrtPCDIAError", f:parsenoop },
	0x001F: { n:"BrtPCDIAString", f:parsenoop },
	0x0020: { n:"BrtPCDIADatetime", f:parsenoop },
	0x0021: { n:"BrtPCRRecord", f:parsenoop },
	0x0022: { n:"BrtPCRRecordDt", f:parsenoop },
	0x0023: { n:"BrtFRTBegin", f:parsenoop },
	0x0024: { n:"BrtFRTEnd", f:parsenoop },
	0x0025: { n:"BrtACBegin", f:parsenoop },
	0x0026: { n:"BrtACEnd", f:parsenoop },
	0x0027: { n:"BrtName", f:parsenoop },
	0x0028: { n:"BrtIndexRowBlock", f:parsenoop },
	0x002A: { n:"BrtIndexBlock", f:parsenoop },
	0x002B: { n:"BrtFont", f:parse_BrtFont },
	0x002C: { n:"BrtFmt", f:parse_BrtFmt },
	0x002D: { n:"BrtFill", f:parsenoop },
	0x002E: { n:"BrtBorder", f:parsenoop },
	0x002F: { n:"BrtXF", f:parse_BrtXF },
	0x0030: { n:"BrtStyle", f:parsenoop },
	0x0031: { n:"BrtCellMeta", f:parsenoop },
	0x0032: { n:"BrtValueMeta", f:parsenoop },
	0x0033: { n:"BrtMdb", f:parsenoop },
	0x0034: { n:"BrtBeginFmd", f:parsenoop },
	0x0035: { n:"BrtEndFmd", f:parsenoop },
	0x0036: { n:"BrtBeginMdx", f:parsenoop },
	0x0037: { n:"BrtEndMdx", f:parsenoop },
	0x0038: { n:"BrtBeginMdxTuple", f:parsenoop },
	0x0039: { n:"BrtEndMdxTuple", f:parsenoop },
	0x003A: { n:"BrtMdxMbrIstr", f:parsenoop },
	0x003B: { n:"BrtStr", f:parsenoop },
	0x003C: { n:"BrtColInfo", f:parsenoop },
	0x003E: { n:"BrtCellRString", f:parsenoop },
	0x0040: { n:"BrtDVal", f:parsenoop },
	0x0041: { n:"BrtSxvcellNum", f:parsenoop },
	0x0042: { n:"BrtSxvcellStr", f:parsenoop },
	0x0043: { n:"BrtSxvcellBool", f:parsenoop },
	0x0044: { n:"BrtSxvcellErr", f:parsenoop },
	0x0045: { n:"BrtSxvcellDate", f:parsenoop },
	0x0046: { n:"BrtSxvcellNil", f:parsenoop },
	0x0080: { n:"BrtFileVersion", f:parsenoop },
	0x0081: { n:"BrtBeginSheet", f:parsenoop },
	0x0082: { n:"BrtEndSheet", f:parsenoop },
	0x0083: { n:"BrtBeginBook", f:parsenoop },
	0x0084: { n:"BrtEndBook", f:parsenoop },
	0x0085: { n:"BrtBeginWsViews", f:parsenoop },
	0x0086: { n:"BrtEndWsViews", f:parsenoop },
	0x0087: { n:"BrtBeginBookViews", f:parsenoop },
	0x0088: { n:"BrtEndBookViews", f:parsenoop },
	0x0089: { n:"BrtBeginWsView", f:parsenoop },
	0x008A: { n:"BrtEndWsView", f:parsenoop },
	0x008B: { n:"BrtBeginCsViews", f:parsenoop },
	0x008C: { n:"BrtEndCsViews", f:parsenoop },
	0x008D: { n:"BrtBeginCsView", f:parsenoop },
	0x008E: { n:"BrtEndCsView", f:parsenoop },
	0x008F: { n:"BrtBeginBundleShs", f:parsenoop },
	0x0090: { n:"BrtEndBundleShs", f:parsenoop },
	0x0091: { n:"BrtBeginSheetData", f:parsenoop },
	0x0092: { n:"BrtEndSheetData", f:parsenoop },
	0x0093: { n:"BrtWsProp", f:parse_BrtWsProp },
	0x0094: { n:"BrtWsDim", f:parse_BrtWsDim },
	0x0097: { n:"BrtPane", f:parsenoop },
	0x0098: { n:"BrtSel", f:parsenoop },
	0x0099: { n:"BrtWbProp", f:parsenoop },
	0x009A: { n:"BrtWbFactoid", f:parsenoop },
	0x009B: { n:"BrtFileRecover", f:parsenoop },
	0x009C: { n:"BrtBundleSh", f:parse_BrtBundleSh },
	0x009D: { n:"BrtCalcProp", f:parsenoop },
	0x009E: { n:"BrtBookView", f:parsenoop },
	0x009F: { n:"BrtBeginSst", f:parse_BrtBeginSst },
	0x00A0: { n:"BrtEndSst", f:parsenoop },
	0x00A1: { n:"BrtBeginAFilter", f:parsenoop },
	0x00A2: { n:"BrtEndAFilter", f:parsenoop },
	0x00A3: { n:"BrtBeginFilterColumn", f:parsenoop },
	0x00A4: { n:"BrtEndFilterColumn", f:parsenoop },
	0x00A5: { n:"BrtBeginFilters", f:parsenoop },
	0x00A6: { n:"BrtEndFilters", f:parsenoop },
	0x00A7: { n:"BrtFilter", f:parsenoop },
	0x00A8: { n:"BrtColorFilter", f:parsenoop },
	0x00A9: { n:"BrtIconFilter", f:parsenoop },
	0x00AA: { n:"BrtTop10Filter", f:parsenoop },
	0x00AB: { n:"BrtDynamicFilter", f:parsenoop },
	0x00AC: { n:"BrtBeginCustomFilters", f:parsenoop },
	0x00AD: { n:"BrtEndCustomFilters", f:parsenoop },
	0x00AE: { n:"BrtCustomFilter", f:parsenoop },
	0x00AF: { n:"BrtAFilterDateGroupItem", f:parsenoop },
	0x00B0: { n:"BrtMergeCell", f:parse_BrtMergeCell },
	0x00B1: { n:"BrtBeginMergeCells", f:parsenoop },
	0x00B2: { n:"BrtEndMergeCells", f:parsenoop },
	0x00B3: { n:"BrtBeginPivotCacheDef", f:parsenoop },
	0x00B4: { n:"BrtEndPivotCacheDef", f:parsenoop },
	0x00B5: { n:"BrtBeginPCDFields", f:parsenoop },
	0x00B6: { n:"BrtEndPCDFields", f:parsenoop },
	0x00B7: { n:"BrtBeginPCDField", f:parsenoop },
	0x00B8: { n:"BrtEndPCDField", f:parsenoop },
	0x00B9: { n:"BrtBeginPCDSource", f:parsenoop },
	0x00BA: { n:"BrtEndPCDSource", f:parsenoop },
	0x00BB: { n:"BrtBeginPCDSRange", f:parsenoop },
	0x00BC: { n:"BrtEndPCDSRange", f:parsenoop },
	0x00BD: { n:"BrtBeginPCDFAtbl", f:parsenoop },
	0x00BE: { n:"BrtEndPCDFAtbl", f:parsenoop },
	0x00BF: { n:"BrtBeginPCDIRun", f:parsenoop },
	0x00C0: { n:"BrtEndPCDIRun", f:parsenoop },
	0x00C1: { n:"BrtBeginPivotCacheRecords", f:parsenoop },
	0x00C2: { n:"BrtEndPivotCacheRecords", f:parsenoop },
	0x00C3: { n:"BrtBeginPCDHierarchies", f:parsenoop },
	0x00C4: { n:"BrtEndPCDHierarchies", f:parsenoop },
	0x00C5: { n:"BrtBeginPCDHierarchy", f:parsenoop },
	0x00C6: { n:"BrtEndPCDHierarchy", f:parsenoop },
	0x00C7: { n:"BrtBeginPCDHFieldsUsage", f:parsenoop },
	0x00C8: { n:"BrtEndPCDHFieldsUsage", f:parsenoop },
	0x00C9: { n:"BrtBeginExtConnection", f:parsenoop },
	0x00CA: { n:"BrtEndExtConnection", f:parsenoop },
	0x00CB: { n:"BrtBeginECDbProps", f:parsenoop },
	0x00CC: { n:"BrtEndECDbProps", f:parsenoop },
	0x00CD: { n:"BrtBeginECOlapProps", f:parsenoop },
	0x00CE: { n:"BrtEndECOlapProps", f:parsenoop },
	0x00CF: { n:"BrtBeginPCDSConsol", f:parsenoop },
	0x00D0: { n:"BrtEndPCDSConsol", f:parsenoop },
	0x00D1: { n:"BrtBeginPCDSCPages", f:parsenoop },
	0x00D2: { n:"BrtEndPCDSCPages", f:parsenoop },
	0x00D3: { n:"BrtBeginPCDSCPage", f:parsenoop },
	0x00D4: { n:"BrtEndPCDSCPage", f:parsenoop },
	0x00D5: { n:"BrtBeginPCDSCPItem", f:parsenoop },
	0x00D6: { n:"BrtEndPCDSCPItem", f:parsenoop },
	0x00D7: { n:"BrtBeginPCDSCSets", f:parsenoop },
	0x00D8: { n:"BrtEndPCDSCSets", f:parsenoop },
	0x00D9: { n:"BrtBeginPCDSCSet", f:parsenoop },
	0x00DA: { n:"BrtEndPCDSCSet", f:parsenoop },
	0x00DB: { n:"BrtBeginPCDFGroup", f:parsenoop },
	0x00DC: { n:"BrtEndPCDFGroup", f:parsenoop },
	0x00DD: { n:"BrtBeginPCDFGItems", f:parsenoop },
	0x00DE: { n:"BrtEndPCDFGItems", f:parsenoop },
	0x00DF: { n:"BrtBeginPCDFGRange", f:parsenoop },
	0x00E0: { n:"BrtEndPCDFGRange", f:parsenoop },
	0x00E1: { n:"BrtBeginPCDFGDiscrete", f:parsenoop },
	0x00E2: { n:"BrtEndPCDFGDiscrete", f:parsenoop },
	0x00E3: { n:"BrtBeginPCDSDTupleCache", f:parsenoop },
	0x00E4: { n:"BrtEndPCDSDTupleCache", f:parsenoop },
	0x00E5: { n:"BrtBeginPCDSDTCEntries", f:parsenoop },
	0x00E6: { n:"BrtEndPCDSDTCEntries", f:parsenoop },
	0x00E7: { n:"BrtBeginPCDSDTCEMembers", f:parsenoop },
	0x00E8: { n:"BrtEndPCDSDTCEMembers", f:parsenoop },
	0x00E9: { n:"BrtBeginPCDSDTCEMember", f:parsenoop },
	0x00EA: { n:"BrtEndPCDSDTCEMember", f:parsenoop },
	0x00EB: { n:"BrtBeginPCDSDTCQueries", f:parsenoop },
	0x00EC: { n:"BrtEndPCDSDTCQueries", f:parsenoop },
	0x00ED: { n:"BrtBeginPCDSDTCQuery", f:parsenoop },
	0x00EE: { n:"BrtEndPCDSDTCQuery", f:parsenoop },
	0x00EF: { n:"BrtBeginPCDSDTCSets", f:parsenoop },
	0x00F0: { n:"BrtEndPCDSDTCSets", f:parsenoop },
	0x00F1: { n:"BrtBeginPCDSDTCSet", f:parsenoop },
	0x00F2: { n:"BrtEndPCDSDTCSet", f:parsenoop },
	0x00F3: { n:"BrtBeginPCDCalcItems", f:parsenoop },
	0x00F4: { n:"BrtEndPCDCalcItems", f:parsenoop },
	0x00F5: { n:"BrtBeginPCDCalcItem", f:parsenoop },
	0x00F6: { n:"BrtEndPCDCalcItem", f:parsenoop },
	0x00F7: { n:"BrtBeginPRule", f:parsenoop },
	0x00F8: { n:"BrtEndPRule", f:parsenoop },
	0x00F9: { n:"BrtBeginPRFilters", f:parsenoop },
	0x00FA: { n:"BrtEndPRFilters", f:parsenoop },
	0x00FB: { n:"BrtBeginPRFilter", f:parsenoop },
	0x00FC: { n:"BrtEndPRFilter", f:parsenoop },
	0x00FD: { n:"BrtBeginPNames", f:parsenoop },
	0x00FE: { n:"BrtEndPNames", f:parsenoop },
	0x00FF: { n:"BrtBeginPName", f:parsenoop },
	0x0100: { n:"BrtEndPName", f:parsenoop },
	0x0101: { n:"BrtBeginPNPairs", f:parsenoop },
	0x0102: { n:"BrtEndPNPairs", f:parsenoop },
	0x0103: { n:"BrtBeginPNPair", f:parsenoop },
	0x0104: { n:"BrtEndPNPair", f:parsenoop },
	0x0105: { n:"BrtBeginECWebProps", f:parsenoop },
	0x0106: { n:"BrtEndECWebProps", f:parsenoop },
	0x0107: { n:"BrtBeginEcWpTables", f:parsenoop },
	0x0108: { n:"BrtEndECWPTables", f:parsenoop },
	0x0109: { n:"BrtBeginECParams", f:parsenoop },
	0x010A: { n:"BrtEndECParams", f:parsenoop },
	0x010B: { n:"BrtBeginECParam", f:parsenoop },
	0x010C: { n:"BrtEndECParam", f:parsenoop },
	0x010D: { n:"BrtBeginPCDKPIs", f:parsenoop },
	0x010E: { n:"BrtEndPCDKPIs", f:parsenoop },
	0x010F: { n:"BrtBeginPCDKPI", f:parsenoop },
	0x0110: { n:"BrtEndPCDKPI", f:parsenoop },
	0x0111: { n:"BrtBeginDims", f:parsenoop },
	0x0112: { n:"BrtEndDims", f:parsenoop },
	0x0113: { n:"BrtBeginDim", f:parsenoop },
	0x0114: { n:"BrtEndDim", f:parsenoop },
	0x0115: { n:"BrtIndexPartEnd", f:parsenoop },
	0x0116: { n:"BrtBeginStyleSheet", f:parsenoop },
	0x0117: { n:"BrtEndStyleSheet", f:parsenoop },
	0x0118: { n:"BrtBeginSXView", f:parsenoop },
	0x0119: { n:"BrtEndSXVI", f:parsenoop },
	0x011A: { n:"BrtBeginSXVI", f:parsenoop },
	0x011B: { n:"BrtBeginSXVIs", f:parsenoop },
	0x011C: { n:"BrtEndSXVIs", f:parsenoop },
	0x011D: { n:"BrtBeginSXVD", f:parsenoop },
	0x011E: { n:"BrtEndSXVD", f:parsenoop },
	0x011F: { n:"BrtBeginSXVDs", f:parsenoop },
	0x0120: { n:"BrtEndSXVDs", f:parsenoop },
	0x0121: { n:"BrtBeginSXPI", f:parsenoop },
	0x0122: { n:"BrtEndSXPI", f:parsenoop },
	0x0123: { n:"BrtBeginSXPIs", f:parsenoop },
	0x0124: { n:"BrtEndSXPIs", f:parsenoop },
	0x0125: { n:"BrtBeginSXDI", f:parsenoop },
	0x0126: { n:"BrtEndSXDI", f:parsenoop },
	0x0127: { n:"BrtBeginSXDIs", f:parsenoop },
	0x0128: { n:"BrtEndSXDIs", f:parsenoop },
	0x0129: { n:"BrtBeginSXLI", f:parsenoop },
	0x012A: { n:"BrtEndSXLI", f:parsenoop },
	0x012B: { n:"BrtBeginSXLIRws", f:parsenoop },
	0x012C: { n:"BrtEndSXLIRws", f:parsenoop },
	0x012D: { n:"BrtBeginSXLICols", f:parsenoop },
	0x012E: { n:"BrtEndSXLICols", f:parsenoop },
	0x012F: { n:"BrtBeginSXFormat", f:parsenoop },
	0x0130: { n:"BrtEndSXFormat", f:parsenoop },
	0x0131: { n:"BrtBeginSXFormats", f:parsenoop },
	0x0132: { n:"BrtEndSxFormats", f:parsenoop },
	0x0133: { n:"BrtBeginSxSelect", f:parsenoop },
	0x0134: { n:"BrtEndSxSelect", f:parsenoop },
	0x0135: { n:"BrtBeginISXVDRws", f:parsenoop },
	0x0136: { n:"BrtEndISXVDRws", f:parsenoop },
	0x0137: { n:"BrtBeginISXVDCols", f:parsenoop },
	0x0138: { n:"BrtEndISXVDCols", f:parsenoop },
	0x0139: { n:"BrtEndSXLocation", f:parsenoop },
	0x013A: { n:"BrtBeginSXLocation", f:parsenoop },
	0x013B: { n:"BrtEndSXView", f:parsenoop },
	0x013C: { n:"BrtBeginSXTHs", f:parsenoop },
	0x013D: { n:"BrtEndSXTHs", f:parsenoop },
	0x013E: { n:"BrtBeginSXTH", f:parsenoop },
	0x013F: { n:"BrtEndSXTH", f:parsenoop },
	0x0140: { n:"BrtBeginISXTHRws", f:parsenoop },
	0x0141: { n:"BrtEndISXTHRws", f:parsenoop },
	0x0142: { n:"BrtBeginISXTHCols", f:parsenoop },
	0x0143: { n:"BrtEndISXTHCols", f:parsenoop },
	0x0144: { n:"BrtBeginSXTDMPS", f:parsenoop },
	0x0145: { n:"BrtEndSXTDMPs", f:parsenoop },
	0x0146: { n:"BrtBeginSXTDMP", f:parsenoop },
	0x0147: { n:"BrtEndSXTDMP", f:parsenoop },
	0x0148: { n:"BrtBeginSXTHItems", f:parsenoop },
	0x0149: { n:"BrtEndSXTHItems", f:parsenoop },
	0x014A: { n:"BrtBeginSXTHItem", f:parsenoop },
	0x014B: { n:"BrtEndSXTHItem", f:parsenoop },
	0x014C: { n:"BrtBeginMetadata", f:parsenoop },
	0x014D: { n:"BrtEndMetadata", f:parsenoop },
	0x014E: { n:"BrtBeginEsmdtinfo", f:parsenoop },
	0x014F: { n:"BrtMdtinfo", f:parsenoop },
	0x0150: { n:"BrtEndEsmdtinfo", f:parsenoop },
	0x0151: { n:"BrtBeginEsmdb", f:parsenoop },
	0x0152: { n:"BrtEndEsmdb", f:parsenoop },
	0x0153: { n:"BrtBeginEsfmd", f:parsenoop },
	0x0154: { n:"BrtEndEsfmd", f:parsenoop },
	0x0155: { n:"BrtBeginSingleCells", f:parsenoop },
	0x0156: { n:"BrtEndSingleCells", f:parsenoop },
	0x0157: { n:"BrtBeginList", f:parsenoop },
	0x0158: { n:"BrtEndList", f:parsenoop },
	0x0159: { n:"BrtBeginListCols", f:parsenoop },
	0x015A: { n:"BrtEndListCols", f:parsenoop },
	0x015B: { n:"BrtBeginListCol", f:parsenoop },
	0x015C: { n:"BrtEndListCol", f:parsenoop },
	0x015D: { n:"BrtBeginListXmlCPr", f:parsenoop },
	0x015E: { n:"BrtEndListXmlCPr", f:parsenoop },
	0x015F: { n:"BrtListCCFmla", f:parsenoop },
	0x0160: { n:"BrtListTrFmla", f:parsenoop },
	0x0161: { n:"BrtBeginExternals", f:parsenoop },
	0x0162: { n:"BrtEndExternals", f:parsenoop },
	0x0163: { n:"BrtSupBookSrc", f:parsenoop },
	0x0165: { n:"BrtSupSelf", f:parsenoop },
	0x0166: { n:"BrtSupSame", f:parsenoop },
	0x0167: { n:"BrtSupTabs", f:parsenoop },
	0x0168: { n:"BrtBeginSupBook", f:parsenoop },
	0x0169: { n:"BrtPlaceholderName", f:parsenoop },
	0x016A: { n:"BrtExternSheet", f:parsenoop },
	0x016B: { n:"BrtExternTableStart", f:parsenoop },
	0x016C: { n:"BrtExternTableEnd", f:parsenoop },
	0x016E: { n:"BrtExternRowHdr", f:parsenoop },
	0x016F: { n:"BrtExternCellBlank", f:parsenoop },
	0x0170: { n:"BrtExternCellReal", f:parsenoop },
	0x0171: { n:"BrtExternCellBool", f:parsenoop },
	0x0172: { n:"BrtExternCellError", f:parsenoop },
	0x0173: { n:"BrtExternCellString", f:parsenoop },
	0x0174: { n:"BrtBeginEsmdx", f:parsenoop },
	0x0175: { n:"BrtEndEsmdx", f:parsenoop },
	0x0176: { n:"BrtBeginMdxSet", f:parsenoop },
	0x0177: { n:"BrtEndMdxSet", f:parsenoop },
	0x0178: { n:"BrtBeginMdxMbrProp", f:parsenoop },
	0x0179: { n:"BrtEndMdxMbrProp", f:parsenoop },
	0x017A: { n:"BrtBeginMdxKPI", f:parsenoop },
	0x017B: { n:"BrtEndMdxKPI", f:parsenoop },
	0x017C: { n:"BrtBeginEsstr", f:parsenoop },
	0x017D: { n:"BrtEndEsstr", f:parsenoop },
	0x017E: { n:"BrtBeginPRFItem", f:parsenoop },
	0x017F: { n:"BrtEndPRFItem", f:parsenoop },
	0x0180: { n:"BrtBeginPivotCacheIDs", f:parsenoop },
	0x0181: { n:"BrtEndPivotCacheIDs", f:parsenoop },
	0x0182: { n:"BrtBeginPivotCacheID", f:parsenoop },
	0x0183: { n:"BrtEndPivotCacheID", f:parsenoop },
	0x0184: { n:"BrtBeginISXVIs", f:parsenoop },
	0x0185: { n:"BrtEndISXVIs", f:parsenoop },
	0x0186: { n:"BrtBeginColInfos", f:parsenoop },
	0x0187: { n:"BrtEndColInfos", f:parsenoop },
	0x0188: { n:"BrtBeginRwBrk", f:parsenoop },
	0x0189: { n:"BrtEndRwBrk", f:parsenoop },
	0x018A: { n:"BrtBeginColBrk", f:parsenoop },
	0x018B: { n:"BrtEndColBrk", f:parsenoop },
	0x018C: { n:"BrtBrk", f:parsenoop },
	0x018D: { n:"BrtUserBookView", f:parsenoop },
	0x018E: { n:"BrtInfo", f:parsenoop },
	0x018F: { n:"BrtCUsr", f:parsenoop },
	0x0190: { n:"BrtUsr", f:parsenoop },
	0x0191: { n:"BrtBeginUsers", f:parsenoop },
	0x0193: { n:"BrtEOF", f:parsenoop },
	0x0194: { n:"BrtUCR", f:parsenoop },
	0x0195: { n:"BrtRRInsDel", f:parsenoop },
	0x0196: { n:"BrtRREndInsDel", f:parsenoop },
	0x0197: { n:"BrtRRMove", f:parsenoop },
	0x0198: { n:"BrtRREndMove", f:parsenoop },
	0x0199: { n:"BrtRRChgCell", f:parsenoop },
	0x019A: { n:"BrtRREndChgCell", f:parsenoop },
	0x019B: { n:"BrtRRHeader", f:parsenoop },
	0x019C: { n:"BrtRRUserView", f:parsenoop },
	0x019D: { n:"BrtRRRenSheet", f:parsenoop },
	0x019E: { n:"BrtRRInsertSh", f:parsenoop },
	0x019F: { n:"BrtRRDefName", f:parsenoop },
	0x01A0: { n:"BrtRRNote", f:parsenoop },
	0x01A1: { n:"BrtRRConflict", f:parsenoop },
	0x01A2: { n:"BrtRRTQSIF", f:parsenoop },
	0x01A3: { n:"BrtRRFormat", f:parsenoop },
	0x01A4: { n:"BrtRREndFormat", f:parsenoop },
	0x01A5: { n:"BrtRRAutoFmt", f:parsenoop },
	0x01A6: { n:"BrtBeginUserShViews", f:parsenoop },
	0x01A7: { n:"BrtBeginUserShView", f:parsenoop },
	0x01A8: { n:"BrtEndUserShView", f:parsenoop },
	0x01A9: { n:"BrtEndUserShViews", f:parsenoop },
	0x01AA: { n:"BrtArrFmla", f:parsenoop },
	0x01AB: { n:"BrtShrFmla", f:parsenoop },
	0x01AC: { n:"BrtTable", f:parsenoop },
	0x01AD: { n:"BrtBeginExtConnections", f:parsenoop },
	0x01AE: { n:"BrtEndExtConnections", f:parsenoop },
	0x01AF: { n:"BrtBeginPCDCalcMems", f:parsenoop },
	0x01B0: { n:"BrtEndPCDCalcMems", f:parsenoop },
	0x01B1: { n:"BrtBeginPCDCalcMem", f:parsenoop },
	0x01B2: { n:"BrtEndPCDCalcMem", f:parsenoop },
	0x01B3: { n:"BrtBeginPCDHGLevels", f:parsenoop },
	0x01B4: { n:"BrtEndPCDHGLevels", f:parsenoop },
	0x01B5: { n:"BrtBeginPCDHGLevel", f:parsenoop },
	0x01B6: { n:"BrtEndPCDHGLevel", f:parsenoop },
	0x01B7: { n:"BrtBeginPCDHGLGroups", f:parsenoop },
	0x01B8: { n:"BrtEndPCDHGLGroups", f:parsenoop },
	0x01B9: { n:"BrtBeginPCDHGLGroup", f:parsenoop },
	0x01BA: { n:"BrtEndPCDHGLGroup", f:parsenoop },
	0x01BB: { n:"BrtBeginPCDHGLGMembers", f:parsenoop },
	0x01BC: { n:"BrtEndPCDHGLGMembers", f:parsenoop },
	0x01BD: { n:"BrtBeginPCDHGLGMember", f:parsenoop },
	0x01BE: { n:"BrtEndPCDHGLGMember", f:parsenoop },
	0x01BF: { n:"BrtBeginQSI", f:parsenoop },
	0x01C0: { n:"BrtEndQSI", f:parsenoop },
	0x01C1: { n:"BrtBeginQSIR", f:parsenoop },
	0x01C2: { n:"BrtEndQSIR", f:parsenoop },
	0x01C3: { n:"BrtBeginDeletedNames", f:parsenoop },
	0x01C4: { n:"BrtEndDeletedNames", f:parsenoop },
	0x01C5: { n:"BrtBeginDeletedName", f:parsenoop },
	0x01C6: { n:"BrtEndDeletedName", f:parsenoop },
	0x01C7: { n:"BrtBeginQSIFs", f:parsenoop },
	0x01C8: { n:"BrtEndQSIFs", f:parsenoop },
	0x01C9: { n:"BrtBeginQSIF", f:parsenoop },
	0x01CA: { n:"BrtEndQSIF", f:parsenoop },
	0x01CB: { n:"BrtBeginAutoSortScope", f:parsenoop },
	0x01CC: { n:"BrtEndAutoSortScope", f:parsenoop },
	0x01CD: { n:"BrtBeginConditionalFormatting", f:parsenoop },
	0x01CE: { n:"BrtEndConditionalFormatting", f:parsenoop },
	0x01CF: { n:"BrtBeginCFRule", f:parsenoop },
	0x01D0: { n:"BrtEndCFRule", f:parsenoop },
	0x01D1: { n:"BrtBeginIconSet", f:parsenoop },
	0x01D2: { n:"BrtEndIconSet", f:parsenoop },
	0x01D3: { n:"BrtBeginDatabar", f:parsenoop },
	0x01D4: { n:"BrtEndDatabar", f:parsenoop },
	0x01D5: { n:"BrtBeginColorScale", f:parsenoop },
	0x01D6: { n:"BrtEndColorScale", f:parsenoop },
	0x01D7: { n:"BrtCFVO", f:parsenoop },
	0x01D8: { n:"BrtExternValueMeta", f:parsenoop },
	0x01D9: { n:"BrtBeginColorPalette", f:parsenoop },
	0x01DA: { n:"BrtEndColorPalette", f:parsenoop },
	0x01DB: { n:"BrtIndexedColor", f:parsenoop },
	0x01DC: { n:"BrtMargins", f:parsenoop },
	0x01DD: { n:"BrtPrintOptions", f:parsenoop },
	0x01DE: { n:"BrtPageSetup", f:parsenoop },
	0x01DF: { n:"BrtBeginHeaderFooter", f:parsenoop },
	0x01E0: { n:"BrtEndHeaderFooter", f:parsenoop },
	0x01E1: { n:"BrtBeginSXCrtFormat", f:parsenoop },
	0x01E2: { n:"BrtEndSXCrtFormat", f:parsenoop },
	0x01E3: { n:"BrtBeginSXCrtFormats", f:parsenoop },
	0x01E4: { n:"BrtEndSXCrtFormats", f:parsenoop },
	0x01E5: { n:"BrtWsFmtInfo", f:parsenoop },
	0x01E6: { n:"BrtBeginMgs", f:parsenoop },
	0x01E7: { n:"BrtEndMGs", f:parsenoop },
	0x01E8: { n:"BrtBeginMGMaps", f:parsenoop },
	0x01E9: { n:"BrtEndMGMaps", f:parsenoop },
	0x01EA: { n:"BrtBeginMG", f:parsenoop },
	0x01EB: { n:"BrtEndMG", f:parsenoop },
	0x01EC: { n:"BrtBeginMap", f:parsenoop },
	0x01ED: { n:"BrtEndMap", f:parsenoop },
	0x01EE: { n:"BrtHLink", f:parsenoop },
	0x01EF: { n:"BrtBeginDCon", f:parsenoop },
	0x01F0: { n:"BrtEndDCon", f:parsenoop },
	0x01F1: { n:"BrtBeginDRefs", f:parsenoop },
	0x01F2: { n:"BrtEndDRefs", f:parsenoop },
	0x01F3: { n:"BrtDRef", f:parsenoop },
	0x01F4: { n:"BrtBeginScenMan", f:parsenoop },
	0x01F5: { n:"BrtEndScenMan", f:parsenoop },
	0x01F6: { n:"BrtBeginSct", f:parsenoop },
	0x01F7: { n:"BrtEndSct", f:parsenoop },
	0x01F8: { n:"BrtSlc", f:parsenoop },
	0x01F9: { n:"BrtBeginDXFs", f:parsenoop },
	0x01FA: { n:"BrtEndDXFs", f:parsenoop },
	0x01FB: { n:"BrtDXF", f:parsenoop },
	0x01FC: { n:"BrtBeginTableStyles", f:parsenoop },
	0x01FD: { n:"BrtEndTableStyles", f:parsenoop },
	0x01FE: { n:"BrtBeginTableStyle", f:parsenoop },
	0x01FF: { n:"BrtEndTableStyle", f:parsenoop },
	0x0200: { n:"BrtTableStyleElement", f:parsenoop },
	0x0201: { n:"BrtTableStyleClient", f:parsenoop },
	0x0202: { n:"BrtBeginVolDeps", f:parsenoop },
	0x0203: { n:"BrtEndVolDeps", f:parsenoop },
	0x0204: { n:"BrtBeginVolType", f:parsenoop },
	0x0205: { n:"BrtEndVolType", f:parsenoop },
	0x0206: { n:"BrtBeginVolMain", f:parsenoop },
	0x0207: { n:"BrtEndVolMain", f:parsenoop },
	0x0208: { n:"BrtBeginVolTopic", f:parsenoop },
	0x0209: { n:"BrtEndVolTopic", f:parsenoop },
	0x020A: { n:"BrtVolSubtopic", f:parsenoop },
	0x020B: { n:"BrtVolRef", f:parsenoop },
	0x020C: { n:"BrtVolNum", f:parsenoop },
	0x020D: { n:"BrtVolErr", f:parsenoop },
	0x020E: { n:"BrtVolStr", f:parsenoop },
	0x020F: { n:"BrtVolBool", f:parsenoop },
	0x0212: { n:"BrtBeginSortState", f:parsenoop },
	0x0213: { n:"BrtEndSortState", f:parsenoop },
	0x0214: { n:"BrtBeginSortCond", f:parsenoop },
	0x0215: { n:"BrtEndSortCond", f:parsenoop },
	0x0216: { n:"BrtBookProtection", f:parsenoop },
	0x0217: { n:"BrtSheetProtection", f:parsenoop },
	0x0218: { n:"BrtRangeProtection", f:parsenoop },
	0x0219: { n:"BrtPhoneticInfo", f:parsenoop },
	0x021A: { n:"BrtBeginECTxtWiz", f:parsenoop },
	0x021B: { n:"BrtEndECTxtWiz", f:parsenoop },
	0x021C: { n:"BrtBeginECTWFldInfoLst", f:parsenoop },
	0x021D: { n:"BrtEndECTWFldInfoLst", f:parsenoop },
	0x021E: { n:"BrtBeginECTwFldInfo", f:parsenoop },
	0x0224: { n:"BrtFileSharing", f:parsenoop },
	0x0225: { n:"BrtOleSize", f:parsenoop },
	0x0226: { n:"BrtDrawing", f:parsenoop },
	0x0227: { n:"BrtLegacyDrawing", f:parsenoop },
	0x0228: { n:"BrtLegacyDrawingHF", f:parsenoop },
	0x0229: { n:"BrtWebOpt", f:parsenoop },
	0x022A: { n:"BrtBeginWebPubItems", f:parsenoop },
	0x022B: { n:"BrtEndWebPubItems", f:parsenoop },
	0x022C: { n:"BrtBeginWebPubItem", f:parsenoop },
	0x022D: { n:"BrtEndWebPubItem", f:parsenoop },
	0x022E: { n:"BrtBeginSXCondFmt", f:parsenoop },
	0x022F: { n:"BrtEndSXCondFmt", f:parsenoop },
	0x0230: { n:"BrtBeginSXCondFmts", f:parsenoop },
	0x0231: { n:"BrtEndSXCondFmts", f:parsenoop },
	0x0232: { n:"BrtBkHim", f:parsenoop },
	0x0234: { n:"BrtColor", f:parsenoop },
	0x0235: { n:"BrtBeginIndexedColors", f:parsenoop },
	0x0236: { n:"BrtEndIndexedColors", f:parsenoop },
	0x0239: { n:"BrtBeginMRUColors", f:parsenoop },
	0x023A: { n:"BrtEndMRUColors", f:parsenoop },
	0x023C: { n:"BrtMRUColor", f:parsenoop },
	0x023D: { n:"BrtBeginDVals", f:parsenoop },
	0x023E: { n:"BrtEndDVals", f:parsenoop },
	0x0241: { n:"BrtSupNameStart", f:parsenoop },
	0x0242: { n:"BrtSupNameValueStart", f:parsenoop },
	0x0243: { n:"BrtSupNameValueEnd", f:parsenoop },
	0x0244: { n:"BrtSupNameNum", f:parsenoop },
	0x0245: { n:"BrtSupNameErr", f:parsenoop },
	0x0246: { n:"BrtSupNameSt", f:parsenoop },
	0x0247: { n:"BrtSupNameNil", f:parsenoop },
	0x0248: { n:"BrtSupNameBool", f:parsenoop },
	0x0249: { n:"BrtSupNameFmla", f:parsenoop },
	0x024A: { n:"BrtSupNameBits", f:parsenoop },
	0x024B: { n:"BrtSupNameEnd", f:parsenoop },
	0x024C: { n:"BrtEndSupBook", f:parsenoop },
	0x024D: { n:"BrtCellSmartTagProperty", f:parsenoop },
	0x024E: { n:"BrtBeginCellSmartTag", f:parsenoop },
	0x024F: { n:"BrtEndCellSmartTag", f:parsenoop },
	0x0250: { n:"BrtBeginCellSmartTags", f:parsenoop },
	0x0251: { n:"BrtEndCellSmartTags", f:parsenoop },
	0x0252: { n:"BrtBeginSmartTags", f:parsenoop },
	0x0253: { n:"BrtEndSmartTags", f:parsenoop },
	0x0254: { n:"BrtSmartTagType", f:parsenoop },
	0x0255: { n:"BrtBeginSmartTagTypes", f:parsenoop },
	0x0256: { n:"BrtEndSmartTagTypes", f:parsenoop },
	0x0257: { n:"BrtBeginSXFilters", f:parsenoop },
	0x0258: { n:"BrtEndSXFilters", f:parsenoop },
	0x0259: { n:"BrtBeginSXFILTER", f:parsenoop },
	0x025A: { n:"BrtEndSXFilter", f:parsenoop },
	0x025B: { n:"BrtBeginFills", f:parsenoop },
	0x025C: { n:"BrtEndFills", f:parsenoop },
	0x025D: { n:"BrtBeginCellWatches", f:parsenoop },
	0x025E: { n:"BrtEndCellWatches", f:parsenoop },
	0x025F: { n:"BrtCellWatch", f:parsenoop },
	0x0260: { n:"BrtBeginCRErrs", f:parsenoop },
	0x0261: { n:"BrtEndCRErrs", f:parsenoop },
	0x0262: { n:"BrtCrashRecErr", f:parsenoop },
	0x0263: { n:"BrtBeginFonts", f:parsenoop },
	0x0264: { n:"BrtEndFonts", f:parsenoop },
	0x0265: { n:"BrtBeginBorders", f:parsenoop },
	0x0266: { n:"BrtEndBorders", f:parsenoop },
	0x0267: { n:"BrtBeginFmts", f:parsenoop },
	0x0268: { n:"BrtEndFmts", f:parsenoop },
	0x0269: { n:"BrtBeginCellXFs", f:parsenoop },
	0x026A: { n:"BrtEndCellXFs", f:parsenoop },
	0x026B: { n:"BrtBeginStyles", f:parsenoop },
	0x026C: { n:"BrtEndStyles", f:parsenoop },
	0x0271: { n:"BrtBigName", f:parsenoop },
	0x0272: { n:"BrtBeginCellStyleXFs", f:parsenoop },
	0x0273: { n:"BrtEndCellStyleXFs", f:parsenoop },
	0x0274: { n:"BrtBeginComments", f:parsenoop },
	0x0275: { n:"BrtEndComments", f:parsenoop },
	0x0276: { n:"BrtBeginCommentAuthors", f:parsenoop },
	0x0277: { n:"BrtEndCommentAuthors", f:parsenoop },
	0x0278: { n:"BrtCommentAuthor", f:parse_BrtCommentAuthor },
	0x0279: { n:"BrtBeginCommentList", f:parsenoop },
	0x027A: { n:"BrtEndCommentList", f:parsenoop },
	0x027B: { n:"BrtBeginComment", f:parse_BrtBeginComment},
	0x027C: { n:"BrtEndComment", f:parsenoop },
	0x027D: { n:"BrtCommentText", f:parse_BrtCommentText },
	0x027E: { n:"BrtBeginOleObjects", f:parsenoop },
	0x027F: { n:"BrtOleObject", f:parsenoop },
	0x0280: { n:"BrtEndOleObjects", f:parsenoop },
	0x0281: { n:"BrtBeginSxrules", f:parsenoop },
	0x0282: { n:"BrtEndSxRules", f:parsenoop },
	0x0283: { n:"BrtBeginActiveXControls", f:parsenoop },
	0x0284: { n:"BrtActiveX", f:parsenoop },
	0x0285: { n:"BrtEndActiveXControls", f:parsenoop },
	0x0286: { n:"BrtBeginPCDSDTCEMembersSortBy", f:parsenoop },
	0x0288: { n:"BrtBeginCellIgnoreECs", f:parsenoop },
	0x0289: { n:"BrtCellIgnoreEC", f:parsenoop },
	0x028A: { n:"BrtEndCellIgnoreECs", f:parsenoop },
	0x028B: { n:"BrtCsProp", f:parsenoop },
	0x028C: { n:"BrtCsPageSetup", f:parsenoop },
	0x028D: { n:"BrtBeginUserCsViews", f:parsenoop },
	0x028E: { n:"BrtEndUserCsViews", f:parsenoop },
	0x028F: { n:"BrtBeginUserCsView", f:parsenoop },
	0x0290: { n:"BrtEndUserCsView", f:parsenoop },
	0x0291: { n:"BrtBeginPcdSFCIEntries", f:parsenoop },
	0x0292: { n:"BrtEndPCDSFCIEntries", f:parsenoop },
	0x0293: { n:"BrtPCDSFCIEntry", f:parsenoop },
	0x0294: { n:"BrtBeginListParts", f:parsenoop },
	0x0295: { n:"BrtListPart", f:parsenoop },
	0x0296: { n:"BrtEndListParts", f:parsenoop },
	0x0297: { n:"BrtSheetCalcProp", f:parsenoop },
	0x0298: { n:"BrtBeginFnGroup", f:parsenoop },
	0x0299: { n:"BrtFnGroup", f:parsenoop },
	0x029A: { n:"BrtEndFnGroup", f:parsenoop },
	0x029B: { n:"BrtSupAddin", f:parsenoop },
	0x029C: { n:"BrtSXTDMPOrder", f:parsenoop },
	0x029D: { n:"BrtCsProtection", f:parsenoop },
	0x029F: { n:"BrtBeginWsSortMap", f:parsenoop },
	0x02A0: { n:"BrtEndWsSortMap", f:parsenoop },
	0x02A1: { n:"BrtBeginRRSort", f:parsenoop },
	0x02A2: { n:"BrtEndRRSort", f:parsenoop },
	0x02A3: { n:"BrtRRSortItem", f:parsenoop },
	0x02A4: { n:"BrtFileSharingIso", f:parsenoop },
	0x02A5: { n:"BrtBookProtectionIso", f:parsenoop },
	0x02A6: { n:"BrtSheetProtectionIso", f:parsenoop },
	0x02A7: { n:"BrtCsProtectionIso", f:parsenoop },
	0x02A8: { n:"BrtRangeProtectionIso", f:parsenoop },
	0x0400: { n:"BrtRwDescent", f:parsenoop },
	0x0401: { n:"BrtKnownFonts", f:parsenoop },
	0x0402: { n:"BrtBeginSXTupleSet", f:parsenoop },
	0x0403: { n:"BrtEndSXTupleSet", f:parsenoop },
	0x0404: { n:"BrtBeginSXTupleSetHeader", f:parsenoop },
	0x0405: { n:"BrtEndSXTupleSetHeader", f:parsenoop },
	0x0406: { n:"BrtSXTupleSetHeaderItem", f:parsenoop },
	0x0407: { n:"BrtBeginSXTupleSetData", f:parsenoop },
	0x0408: { n:"BrtEndSXTupleSetData", f:parsenoop },
	0x0409: { n:"BrtBeginSXTupleSetRow", f:parsenoop },
	0x040A: { n:"BrtEndSXTupleSetRow", f:parsenoop },
	0x040B: { n:"BrtSXTupleSetRowItem", f:parsenoop },
	0x040C: { n:"BrtNameExt", f:parsenoop },
	0x040D: { n:"BrtPCDH14", f:parsenoop },
	0x040E: { n:"BrtBeginPCDCalcMem14", f:parsenoop },
	0x040F: { n:"BrtEndPCDCalcMem14", f:parsenoop },
	0x0410: { n:"BrtSXTH14", f:parsenoop },
	0x0411: { n:"BrtBeginSparklineGroup", f:parsenoop },
	0x0412: { n:"BrtEndSparklineGroup", f:parsenoop },
	0x0413: { n:"BrtSparkline", f:parsenoop },
	0x0414: { n:"BrtSXDI14", f:parsenoop },
	0x0415: { n:"BrtWsFmtInfoEx14", f:parsenoop },
	0x0416: { n:"BrtBeginConditionalFormatting14", f:parsenoop },
	0x0417: { n:"BrtEndConditionalFormatting14", f:parsenoop },
	0x0418: { n:"BrtBeginCFRule14", f:parsenoop },
	0x0419: { n:"BrtEndCFRule14", f:parsenoop },
	0x041A: { n:"BrtCFVO14", f:parsenoop },
	0x041B: { n:"BrtBeginDatabar14", f:parsenoop },
	0x041C: { n:"BrtBeginIconSet14", f:parsenoop },
	0x041D: { n:"BrtDVal14", f:parsenoop },
	0x041E: { n:"BrtBeginDVals14", f:parsenoop },
	0x041F: { n:"BrtColor14", f:parsenoop },
	0x0420: { n:"BrtBeginSparklines", f:parsenoop },
	0x0421: { n:"BrtEndSparklines", f:parsenoop },
	0x0422: { n:"BrtBeginSparklineGroups", f:parsenoop },
	0x0423: { n:"BrtEndSparklineGroups", f:parsenoop },
	0x0425: { n:"BrtSXVD14", f:parsenoop },
	0x0426: { n:"BrtBeginSxview14", f:parsenoop },
	0x0427: { n:"BrtEndSxview14", f:parsenoop },
	0x042A: { n:"BrtBeginPCD14", f:parsenoop },
	0x042B: { n:"BrtEndPCD14", f:parsenoop },
	0x042C: { n:"BrtBeginExtConn14", f:parsenoop },
	0x042D: { n:"BrtEndExtConn14", f:parsenoop },
	0x042E: { n:"BrtBeginSlicerCacheIDs", f:parsenoop },
	0x042F: { n:"BrtEndSlicerCacheIDs", f:parsenoop },
	0x0430: { n:"BrtBeginSlicerCacheID", f:parsenoop },
	0x0431: { n:"BrtEndSlicerCacheID", f:parsenoop },
	0x0433: { n:"BrtBeginSlicerCache", f:parsenoop },
	0x0434: { n:"BrtEndSlicerCache", f:parsenoop },
	0x0435: { n:"BrtBeginSlicerCacheDef", f:parsenoop },
	0x0436: { n:"BrtEndSlicerCacheDef", f:parsenoop },
	0x0437: { n:"BrtBeginSlicersEx", f:parsenoop },
	0x0438: { n:"BrtEndSlicersEx", f:parsenoop },
	0x0439: { n:"BrtBeginSlicerEx", f:parsenoop },
	0x043A: { n:"BrtEndSlicerEx", f:parsenoop },
	0x043B: { n:"BrtBeginSlicer", f:parsenoop },
	0x043C: { n:"BrtEndSlicer", f:parsenoop },
	0x043D: { n:"BrtSlicerCachePivotTables", f:parsenoop },
	0x043E: { n:"BrtBeginSlicerCacheOlapImpl", f:parsenoop },
	0x043F: { n:"BrtEndSlicerCacheOlapImpl", f:parsenoop },
	0x0440: { n:"BrtBeginSlicerCacheLevelsData", f:parsenoop },
	0x0441: { n:"BrtEndSlicerCacheLevelsData", f:parsenoop },
	0x0442: { n:"BrtBeginSlicerCacheLevelData", f:parsenoop },
	0x0443: { n:"BrtEndSlicerCacheLevelData", f:parsenoop },
	0x0444: { n:"BrtBeginSlicerCacheSiRanges", f:parsenoop },
	0x0445: { n:"BrtEndSlicerCacheSiRanges", f:parsenoop },
	0x0446: { n:"BrtBeginSlicerCacheSiRange", f:parsenoop },
	0x0447: { n:"BrtEndSlicerCacheSiRange", f:parsenoop },
	0x0448: { n:"BrtSlicerCacheOlapItem", f:parsenoop },
	0x0449: { n:"BrtBeginSlicerCacheSelections", f:parsenoop },
	0x044A: { n:"BrtSlicerCacheSelection", f:parsenoop },
	0x044B: { n:"BrtEndSlicerCacheSelections", f:parsenoop },
	0x044C: { n:"BrtBeginSlicerCacheNative", f:parsenoop },
	0x044D: { n:"BrtEndSlicerCacheNative", f:parsenoop },
	0x044E: { n:"BrtSlicerCacheNativeItem", f:parsenoop },
	0x044F: { n:"BrtRangeProtection14", f:parsenoop },
	0x0450: { n:"BrtRangeProtectionIso14", f:parsenoop },
	0x0451: { n:"BrtCellIgnoreEC14", f:parsenoop },
	0x0457: { n:"BrtList14", f:parsenoop },
	0x0458: { n:"BrtCFIcon", f:parsenoop },
	0x0459: { n:"BrtBeginSlicerCachesPivotCacheIDs", f:parsenoop },
	0x045A: { n:"BrtEndSlicerCachesPivotCacheIDs", f:parsenoop },
	0x045B: { n:"BrtBeginSlicers", f:parsenoop },
	0x045C: { n:"BrtEndSlicers", f:parsenoop },
	0x045D: { n:"BrtWbProp14", f:parsenoop },
	0x045E: { n:"BrtBeginSXEdit", f:parsenoop },
	0x045F: { n:"BrtEndSXEdit", f:parsenoop },
	0x0460: { n:"BrtBeginSXEdits", f:parsenoop },
	0x0461: { n:"BrtEndSXEdits", f:parsenoop },
	0x0462: { n:"BrtBeginSXChange", f:parsenoop },
	0x0463: { n:"BrtEndSXChange", f:parsenoop },
	0x0464: { n:"BrtBeginSXChanges", f:parsenoop },
	0x0465: { n:"BrtEndSXChanges", f:parsenoop },
	0x0466: { n:"BrtSXTupleItems", f:parsenoop },
	0x0468: { n:"BrtBeginSlicerStyle", f:parsenoop },
	0x0469: { n:"BrtEndSlicerStyle", f:parsenoop },
	0x046A: { n:"BrtSlicerStyleElement", f:parsenoop },
	0x046B: { n:"BrtBeginStyleSheetExt14", f:parsenoop },
	0x046C: { n:"BrtEndStyleSheetExt14", f:parsenoop },
	0x046D: { n:"BrtBeginSlicerCachesPivotCacheID", f:parsenoop },
	0x046E: { n:"BrtEndSlicerCachesPivotCacheID", f:parsenoop },
	0x046F: { n:"BrtBeginConditionalFormattings", f:parsenoop },
	0x0470: { n:"BrtEndConditionalFormattings", f:parsenoop },
	0x0471: { n:"BrtBeginPCDCalcMemExt", f:parsenoop },
	0x0472: { n:"BrtEndPCDCalcMemExt", f:parsenoop },
	0x0473: { n:"BrtBeginPCDCalcMemsExt", f:parsenoop },
	0x0474: { n:"BrtEndPCDCalcMemsExt", f:parsenoop },
	0x0475: { n:"BrtPCDField14", f:parsenoop },
	0x0476: { n:"BrtBeginSlicerStyles", f:parsenoop },
	0x0477: { n:"BrtEndSlicerStyles", f:parsenoop },
	0x0478: { n:"BrtBeginSlicerStyleElements", f:parsenoop },
	0x0479: { n:"BrtEndSlicerStyleElements", f:parsenoop },
	0x047A: { n:"BrtCFRuleExt", f:parsenoop },
	0x047B: { n:"BrtBeginSXCondFmt14", f:parsenoop },
	0x047C: { n:"BrtEndSXCondFmt14", f:parsenoop },
	0x047D: { n:"BrtBeginSXCondFmts14", f:parsenoop },
	0x047E: { n:"BrtEndSXCondFmts14", f:parsenoop },
	0x0480: { n:"BrtBeginSortCond14", f:parsenoop },
	0x0481: { n:"BrtEndSortCond14", f:parsenoop },
	0x0482: { n:"BrtEndDVals14", f:parsenoop },
	0x0483: { n:"BrtEndIconSet14", f:parsenoop },
	0x0484: { n:"BrtEndDatabar14", f:parsenoop },
	0x0485: { n:"BrtBeginColorScale14", f:parsenoop },
	0x0486: { n:"BrtEndColorScale14", f:parsenoop },
	0x0487: { n:"BrtBeginSxrules14", f:parsenoop },
	0x0488: { n:"BrtEndSxrules14", f:parsenoop },
	0x0489: { n:"BrtBeginPRule14", f:parsenoop },
	0x048A: { n:"BrtEndPRule14", f:parsenoop },
	0x048B: { n:"BrtBeginPRFilters14", f:parsenoop },
	0x048C: { n:"BrtEndPRFilters14", f:parsenoop },
	0x048D: { n:"BrtBeginPRFilter14", f:parsenoop },
	0x048E: { n:"BrtEndPRFilter14", f:parsenoop },
	0x048F: { n:"BrtBeginPRFItem14", f:parsenoop },
	0x0490: { n:"BrtEndPRFItem14", f:parsenoop },
	0x0491: { n:"BrtBeginCellIgnoreECs14", f:parsenoop },
	0x0492: { n:"BrtEndCellIgnoreECs14", f:parsenoop },
	0x0493: { n:"BrtDxf14", f:parsenoop },
	0x0494: { n:"BrtBeginDxF14s", f:parsenoop },
	0x0495: { n:"BrtEndDxf14s", f:parsenoop },
	0x0499: { n:"BrtFilter14", f:parsenoop },
	0x049A: { n:"BrtBeginCustomFilters14", f:parsenoop },
	0x049C: { n:"BrtCustomFilter14", f:parsenoop },
	0x049D: { n:"BrtIconFilter14", f:parsenoop },
	0x049E: { n:"BrtPivotCacheConnectionName", f:parsenoop },
	0x0800: { n:"BrtBeginDecoupledPivotCacheIDs", f:parsenoop },
	0x0801: { n:"BrtEndDecoupledPivotCacheIDs", f:parsenoop },
	0x0802: { n:"BrtDecoupledPivotCacheID", f:parsenoop },
	0x0803: { n:"BrtBeginPivotTableRefs", f:parsenoop },
	0x0804: { n:"BrtEndPivotTableRefs", f:parsenoop },
	0x0805: { n:"BrtPivotTableRef", f:parsenoop },
	0x0806: { n:"BrtSlicerCacheBookPivotTables", f:parsenoop },
	0x0807: { n:"BrtBeginSxvcells", f:parsenoop },
	0x0808: { n:"BrtEndSxvcells", f:parsenoop },
	0x0809: { n:"BrtBeginSxRow", f:parsenoop },
	0x080A: { n:"BrtEndSxRow", f:parsenoop },
	0x080C: { n:"BrtPcdCalcMem15", f:parsenoop },
	0x0813: { n:"BrtQsi15", f:parsenoop },
	0x0814: { n:"BrtBeginWebExtensions", f:parsenoop },
	0x0815: { n:"BrtEndWebExtensions", f:parsenoop },
	0x0816: { n:"BrtWebExtension", f:parsenoop },
	0x0817: { n:"BrtAbsPath15", f:parsenoop },
	0x0818: { n:"BrtBeginPivotTableUISettings", f:parsenoop },
	0x0819: { n:"BrtEndPivotTableUISettings", f:parsenoop },
	0x081B: { n:"BrtTableSlicerCacheIDs", f:parsenoop },
	0x081C: { n:"BrtTableSlicerCacheID", f:parsenoop },
	0x081D: { n:"BrtBeginTableSlicerCache", f:parsenoop },
	0x081E: { n:"BrtEndTableSlicerCache", f:parsenoop },
	0x081F: { n:"BrtSxFilter15", f:parsenoop },
	0x0820: { n:"BrtBeginTimelineCachePivotCacheIDs", f:parsenoop },
	0x0821: { n:"BrtEndTimelineCachePivotCacheIDs", f:parsenoop },
	0x0822: { n:"BrtTimelineCachePivotCacheID", f:parsenoop },
	0x0823: { n:"BrtBeginTimelineCacheIDs", f:parsenoop },
	0x0824: { n:"BrtEndTimelineCacheIDs", f:parsenoop },
	0x0825: { n:"BrtBeginTimelineCacheID", f:parsenoop },
	0x0826: { n:"BrtEndTimelineCacheID", f:parsenoop },
	0x0827: { n:"BrtBeginTimelinesEx", f:parsenoop },
	0x0828: { n:"BrtEndTimelinesEx", f:parsenoop },
	0x0829: { n:"BrtBeginTimelineEx", f:parsenoop },
	0x082A: { n:"BrtEndTimelineEx", f:parsenoop },
	0x082B: { n:"BrtWorkBookPr15", f:parsenoop },
	0x082C: { n:"BrtPCDH15", f:parsenoop },
	0x082D: { n:"BrtBeginTimelineStyle", f:parsenoop },
	0x082E: { n:"BrtEndTimelineStyle", f:parsenoop },
	0x082F: { n:"BrtTimelineStyleElement", f:parsenoop },
	0x0830: { n:"BrtBeginTimelineStylesheetExt15", f:parsenoop },
	0x0831: { n:"BrtEndTimelineStylesheetExt15", f:parsenoop },
	0x0832: { n:"BrtBeginTimelineStyles", f:parsenoop },
	0x0833: { n:"BrtEndTimelineStyles", f:parsenoop },
	0x0834: { n:"BrtBeginTimelineStyleElements", f:parsenoop },
	0x0835: { n:"BrtEndTimelineStyleElements", f:parsenoop },
	0x0836: { n:"BrtDxf15", f:parsenoop },
	0x0837: { n:"BrtBeginDxfs15", f:parsenoop },
	0x0838: { n:"brtEndDxfs15", f:parsenoop },
	0x0839: { n:"BrtSlicerCacheHideItemsWithNoData", f:parsenoop },
	0x083A: { n:"BrtBeginItemUniqueNames", f:parsenoop },
	0x083B: { n:"BrtEndItemUniqueNames", f:parsenoop },
	0x083C: { n:"BrtItemUniqueName", f:parsenoop },
	0x083D: { n:"BrtBeginExtConn15", f:parsenoop },
	0x083E: { n:"BrtEndExtConn15", f:parsenoop },
	0x083F: { n:"BrtBeginOledbPr15", f:parsenoop },
	0x0840: { n:"BrtEndOledbPr15", f:parsenoop },
	0x0841: { n:"BrtBeginDataFeedPr15", f:parsenoop },
	0x0842: { n:"BrtEndDataFeedPr15", f:parsenoop },
	0x0843: { n:"BrtTextPr15", f:parsenoop },
	0x0844: { n:"BrtRangePr15", f:parsenoop },
	0x0845: { n:"BrtDbCommand15", f:parsenoop },
	0x0846: { n:"BrtBeginDbTables15", f:parsenoop },
	0x0847: { n:"BrtEndDbTables15", f:parsenoop },
	0x0848: { n:"BrtDbTable15", f:parsenoop },
	0x0849: { n:"BrtBeginDataModel", f:parsenoop },
	0x084A: { n:"BrtEndDataModel", f:parsenoop },
	0x084B: { n:"BrtBeginModelTables", f:parsenoop },
	0x084C: { n:"BrtEndModelTables", f:parsenoop },
	0x084D: { n:"BrtModelTable", f:parsenoop },
	0x084E: { n:"BrtBeginModelRelationships", f:parsenoop },
	0x084F: { n:"BrtEndModelRelationships", f:parsenoop },
	0x0850: { n:"BrtModelRelationship", f:parsenoop },
	0x0851: { n:"BrtBeginECTxtWiz15", f:parsenoop },
	0x0852: { n:"BrtEndECTxtWiz15", f:parsenoop },
	0x0853: { n:"BrtBeginECTWFldInfoLst15", f:parsenoop },
	0x0854: { n:"BrtEndECTWFldInfoLst15", f:parsenoop },
	0x0855: { n:"BrtBeginECTWFldInfo15", f:parsenoop },
	0x0856: { n:"BrtFieldListActiveItem", f:parsenoop },
	0x0857: { n:"BrtPivotCacheIdVersion", f:parsenoop },
	0x0858: { n:"BrtSXDI15", f:parsenoop },
	0xFFFF: { n:"", f:parsenoop }
};

function fixopts(opts) {
	var defaults = [
		['cellNF', false], /* emit cell number format string as .z */
		['cellHTML', true], /* emit html string as .h */
		['cellFormula', true], /* emit formulae as .f */

		['sheetStubs', false], /* emit empty cells */
		['sheetRows', 0, 'n'], /* read n rows (0 = read all rows) */
		['bookDeps', false], /* parse calculation chains */
		['bookSheets', false], /* only try to get sheet names (no Sheets) */
		['bookProps', false], /* only try to get properties (no Sheets) */
		['bookFiles', false], /* include raw file structure (keys, files) */

		['WTF', false] /* WTF mode (throws errors) */
	];
	defaults.forEach(function(d) {
		if(typeof opts[d[0]] === 'undefined') opts[d[0]] = d[1];
		if(d[2] === 'n') opts[d[0]] = Number(opts[d[0]]);
	});
}
function parseZip(zip, opts) {
	opts = opts || {};
	fixopts(opts);
	reset_cp();
	var entries = Object.keys(zip.files);
	var keys = entries.filter(function(x){return x.substr(-1) != '/';}).sort();
	var dir = parseCT(getzipdata(zip, '[Content_Types].xml'));
	var xlsb = false;
	var sheets;
	if(dir.workbooks.length === 0) {
		var binname = "xl/workbook.bin";
		if(!getzipfile(zip,binname)) throw new Error("Could not find workbook entry");
		dir.workbooks.push(binname);
		xlsb = true;
	}

	if(!opts.bookSheets && !opts.bookProps) {
		strs = {};
		if(dir.sst) strs=parse_sst(getzipdata(zip, dir.sst.replace(/^\//,'')), dir.sst, opts);

		styles = {};
		if(dir.style) styles = parse_sty(getzipdata(zip, dir.style.replace(/^\//,'')),dir.style, opts);
	}

	var wb = parse_wb(getzipdata(zip, dir.workbooks[0].replace(/^\//,'')), dir.workbooks[0], opts);

	var props = {}, propdata = "";
	try {
		propdata = dir.coreprops.length !== 0 ? getzipdata(zip, dir.coreprops[0].replace(/^\//,'')) : "";
		propdata += dir.extprops.length !== 0 ? getzipdata(zip, dir.extprops[0].replace(/^\//,'')) : "";
		props = propdata !== "" ? parseProps(propdata) : {};
	} catch(e) { }

	var custprops = {};
	if(!opts.bookSheets || opts.bookProps) {
		if (dir.custprops.length !== 0) {
			propdata = getzipdata(zip, dir.custprops[0].replace(/^\//,''), true);
			if(propdata) custprops = parseCustomProps(propdata);
		}
	}

	var out = {};
	if(opts.bookSheets || opts.bookProps) {
		if(props.Worksheets && props.SheetNames.length > 0) sheets=props.SheetNames;
		else if(wb.Sheets) sheets = wb.Sheets.map(function(x){ return x.name; });
		if(opts.bookProps) { out.Props = props; out.Custprops = custprops; }
		if(typeof sheets !== 'undefined') out.SheetNames = sheets;
		if(opts.bookSheets ? out.SheetNames : opts.bookProps) return out;
	}
	sheets = {};

	var deps = {};
	if(opts.bookDeps && dir.calcchain) deps=parseDeps(getzipdata(zip, dir.calcchain.replace(/^\//,'')));

	var i=0;
	var sheetRels = {};
	var path, relsPath;
	if(!props.Worksheets) {
		/* Google Docs doesn't generate the appropriate metadata, so we impute: */
		var wbsheets = wb.Sheets;
		props.Worksheets = wbsheets.length;
		props.SheetNames = [];
		for(var j = 0; j != wbsheets.length; ++j) {
			props.SheetNames[j] = wbsheets[j].name;
		}
	}
	/* Numbers iOS hack TODO: parse workbook rels to get names */
	var nmode = (getzipdata(zip,"xl/worksheets/sheet.xml",true))?1:0;
	for(i = 0; i != props.Worksheets; ++i) {
		try {
			//path = dir.sheets[i].replace(/^\//,'');
			path = 'xl/worksheets/sheet'+(i+1-nmode)+(xlsb?'.bin':'.xml');
			path = path.replace(/sheet0\./,"sheet.");
			relsPath = path.replace(/^(.*)(\/)([^\/]*)$/, "$1/_rels/$3.rels");
			sheets[props.SheetNames[i]]=parse_ws(getzipdata(zip, path),path,opts);
			sheetRels[props.SheetNames[i]]=parseRels(getzipdata(zip, relsPath, true), path);
		} catch(e) { if(opts.WTF) throw e; }
	}

	if(dir.comments) parse_comments(zip, dir.comments, sheets, sheetRels, opts);

	out = {
		Directory: dir,
		Workbook: wb,
		Props: props,
		Custprops: custprops,
		Deps: deps,
		Sheets: sheets,
		SheetNames: props.SheetNames,
		Strings: strs,
		Styles: styles,
	};
	if(opts.bookFiles) {
		out.keys = keys;
		out.files = zip.files;
	}
	return out;
}
function readSync(data, options) {
	var zip, d = data;
	var o = options||{};
	switch((o.type||"base64")){
		case "file":
			if(typeof Buffer !== 'undefined') { zip=new jszip(d=_fs.readFileSync(data)); break; }
			d = _fs.readFileSync(data).toString('base64');
			/* falls through */
		case "base64": zip = new jszip(d, { base64:true }); break;
		case "binary": zip = new jszip(d, { base64:false }); break;
	}
	return parseZip(zip, o);
}

function readFileSync(data, options) {
	var o = options||{}; o.type = 'file';
	return readSync(data, o);
}

XLSX.read = readSync;
XLSX.readFile = readFileSync;
XLSX.parseZip = parseZip;
return this;

})(XLSX);

var _chr = function(c) { return String.fromCharCode(c); };

function encode_col(col) { var s=""; for(++col; col; col=Math.floor((col-1)/26)) s = _chr(((col-1)%26) + 65) + s; return s; }
function encode_row(row) { return "" + (row + 1); }
function encode_cell(cell) { return encode_col(cell.c) + encode_row(cell.r); }

function decode_col(c) { var d = 0, i = 0; for(; i !== c.length; ++i) d = 26*d + c.charCodeAt(i) - 64; return d - 1; }
function decode_row(rowstr) { return Number(rowstr) - 1; }
function split_cell(cstr) { return cstr.replace(/(\$?[A-Z]*)(\$?[0-9]*)/,"$1,$2").split(","); }
function decode_cell(cstr) { var splt = split_cell(cstr); return { c:decode_col(splt[0]), r:decode_row(splt[1]) }; }
function decode_range(range) { var x =range.split(":").map(decode_cell); return {s:x[0],e:x[x.length-1]}; }
function encode_range(range) { return encode_cell(range.s) + ":" + encode_cell(range.e); }

function sheet_to_row_object_array(sheet, opts){
	var val, row, r, hdr = {}, isempty, R, C;
	var out = [];
	opts = opts || {};
	if(!sheet || !sheet["!ref"]) return out;
	r = XLSX.utils.decode_range(sheet["!ref"]);
	for(R=r.s.r, C = r.s.c; C <= r.e.c; ++C) {
		val = sheet[encode_cell({c:C,r:R})];
		if(!val) continue;
		if(val.w) hdr[C] = val.w;
		else switch(val.t) {
			case 's': case 'str': hdr[C] = val.v; break;
			case 'n': hdr[C] = val.v; break;
		}
	}

	for (R = r.s.r + 1; R <= r.e.r; ++R) {
		isempty = true;
		/* row index available as __rowNum__ */
		row = Object.create({ __rowNum__ : R });
		for (C = r.s.c; C <= r.e.c; ++C) {
			val = sheet[encode_cell({c: C,r: R})];
			if(!val || !val.t) continue;
			if(typeof val.w !== 'undefined' && !opts.raw) { row[hdr[C]] = val.w; isempty = false; }
			else switch(val.t){
				case 's': case 'str': case 'b': case 'n':
					if(typeof val.v !== 'undefined') {
						row[hdr[C]] = val.v;
						isempty = false;
					}
					break;
				case 'e': break; /* throw */
				default: throw 'unrecognized type ' + val.t;
			}
		}
		if(!isempty) out.push(row);
	}
	return out;
}

function sheet_to_csv(sheet, opts) {
	var stringify = function stringify(val) {
		if(!val.t) return "";
		if(typeof val.w !== 'undefined') return val.w;
		switch(val.t){
			case 'n': return String(val.v);
			case 's': case 'str': return typeof val.v !== 'undefined' ? val.v : "";
			case 'b': return val.v ? "TRUE" : "FALSE";
			case 'e': return val.v; /* throw out value in case of error */
			default: throw 'unrecognized type ' + val.t;
		}
	};
	var out = [], txt = "";
	opts = opts || {};
	if(!sheet || !sheet["!ref"]) return "";
	var r = XLSX.utils.decode_range(sheet["!ref"]);
	var fs = opts.FS||",", rs = opts.RS||"\n";

	for(var R = r.s.r; R <= r.e.r; ++R) {
		var row = [];
		for(var C = r.s.c; C <= r.e.c; ++C) {
			var val = sheet[XLSX.utils.encode_cell({c:C,r:R})];
			if(!val) { row.push(""); continue; }
			txt = String(stringify(val));
			if(txt.indexOf(fs)!==-1 || txt.indexOf(rs)!==-1 || txt.indexOf('"')!==-1)
				txt = "\"" + txt.replace(/"/g, '""') + "\"";
			row.push(txt);
		}
		out.push(row.join(fs));
	}
	return out.join(rs) + (out.length ? rs : "");
}
var make_csv = sheet_to_csv;

function get_formulae(ws) {
	var cmds = [];
	for(var y in ws) if(y[0] !=='!' && ws.hasOwnProperty(y)) {
		var x = ws[y];
		var val = "";
		if(x.f) val = x.f;
		else if(typeof x.w !== 'undefined') val = "'" + x.w;
		else if(typeof x.v === 'undefined') continue;
		else val = x.v;
		cmds.push(y + "=" + val);
	}
	return cmds;
}

XLSX.utils = {
	encode_col: encode_col,
	encode_row: encode_row,
	encode_cell: encode_cell,
	encode_range: encode_range,
	decode_col: decode_col,
	decode_row: decode_row,
	split_cell: split_cell,
	decode_cell: decode_cell,
	decode_range: decode_range,
	sheet_to_csv: sheet_to_csv,
	make_csv: sheet_to_csv,
	get_formulae: get_formulae,
	sheet_to_row_object_array: sheet_to_row_object_array
};

if(typeof require !== 'undefined' && typeof exports !== 'undefined') {
	exports.read = XLSX.read;
	exports.readFile = XLSX.readFile;
	exports.utils = XLSX.utils;
	exports.version = XLSX.version;
}
