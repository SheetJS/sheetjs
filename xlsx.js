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
SSF.version = '0.5.3';
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
	throw "unsupported value in General format: " + v;
};
SSF._general = general_fmt;
var parse_date_code = function parse_date_code(v,opts) {
	var date = Math.floor(v), time = Math.floor(86400 * (v - date)+1e-6), dow=0;
	var dout=[], out={D:date, T:time, u:86400*(v-date)-time}; fixopts(opts = (opts||{}));
	if(opts.date1904) date += 1462;
	if(date > 2958465) return null;
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
	var o;
	switch(type) {
		case 'y': switch(fmt) { /* year */
			case 'y': case 'yy': return pad(val.y % 100,2);
			case 'yyy': case 'yyyy': return pad(val.y % 10000,4);
			default: throw 'bad year format: ' + fmt;
		}
		case 'm': switch(fmt) { /* month */
			case 'm': return val.m;
			case 'mm': return pad(val.m,2);
			case 'mmm': return months[val.m-1][1];
			case 'mmmm': return months[val.m-1][2];
			case 'mmmmm': return months[val.m-1][0];
			default: throw 'bad month format: ' + fmt;
		}
		case 'd': switch(fmt) { /* day */
			case 'd': return val.d;
			case 'dd': return pad(val.d,2);
			case 'ddd': return days[val.q][0];
			case 'dddd': return days[val.q][1];
			default: throw 'bad day format: ' + fmt;
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
			case 's': return Math.round(val.S+val.u);
			case 'ss': return pad(Math.round(val.S+val.u), 2);
			case 'ss.0': o = pad(Math.round(10*(val.S+val.u)),3); return o.substr(0,2)+"." + o.substr(2);
			case 'ss.00': o = pad(Math.round(100*(val.S+val.u)),4); return o.substr(0,2)+"." + o.substr(2);
			case 'ss.000': o = pad(Math.round(1000*(val.S+val.u)),5); return o.substr(0,2)+"." + o.substr(2);
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
		default: throw 'bad format type ' + type + ' in ' + fmt;
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
	if(fmt.indexOf("E") > -1) {
		var idx = fmt.indexOf("E") - fmt.indexOf(".") - 1;
		if(fmt == '##0.0E+0') {
			var period = fmt.length - 5;
			var ee = (Number(val.toExponential(0).substr(2+(val<0))))%period;
			o = (val/Math.pow(10,ee)).toPrecision(idx+1+(period+ee)%period);
			if(!o.match(/[Ee]/)) {
				var fakee = (Number(val.toExponential(0).substr(2+(val<0))));
				if(o.indexOf(".") === -1) o = o[0] + "." + o.substr(1) + "E+" + (fakee - o.length+ee);
				else throw "missing E |" + o;
			}
			o = o.replace(/^([+-]?)([0-9]*)\.([0-9]*)[Ee]/,function($$,$1,$2,$3) { return $1 + $2 + $3.substr(0,(period+ee)%period) + "." + $3.substr(ee) + "E"; });
		} else o = val.toExponential(idx);
		if(fmt.match(/E\+00$/) && o.match(/e[+-][0-9]$/)) o = o.substr(0,o.length-1) + "0" + o[o.length-1];
		if(fmt.match(/E\-/) && o.match(/e\+/)) o = o.replace(/e\+/,"e");
		return o.replace("e","E");
	}
	if(fmt[0] === "$") return "$"+write_num(type,fmt.substr(fmt[1]==' '?2:1),val);
	var r, ff, aval = val < 0 ? -val : val, sign = val < 0 ? "-" : "";
	if((r = fmt.match(/# (\?+)([ ]?)\/([ ]?)(\d+)/))) {
		var den = Number(r[4]), rnd = Math.round(aval * den), base = Math.floor(rnd/den);
		var myn = (rnd - base*den), myd = den;
		return sign + (base?base:"") + " " + (myn === 0 ? fill(" ", r[1].length + 1 + r[4].length) : pad(myn,r[1].length," ") + r[2] + "/" + r[3] + pad(myd,r[4].length));
	}
	if(fmt.match(/^00+$/)) return (val<0?"-":"")+pad(Math.round(aval),fmt.length);
	if(fmt.match(/^[#?]+$/)) return String(Math.round(val)).replace(/^0$/,"");
	if(r = fmt.match(/^#*0+\.(0+)/)) {
		o = Math.round(val * Math.pow(10,r[1].length));
		return String(o/Math.pow(10,r[1].length)).replace(/^([^\.]+)$/,"$1."+r[1]).replace(/\.$/,"."+r[1]).replace(/\.([0-9]*)$/,function($$, $1) { return "." + $1 + fill("0", r[1].length-$1.length); });
	}
	if(r = fmt.match(/^# ([?]+)([ ]?)\/([ ]?)([?]+)/)) {
		var rr = Math.min(Math.max(r[1].length, r[4].length),7);
		ff = frac(aval, Math.pow(10,rr)-1, true);
		return sign + (ff[0]||(ff[1] ? "" : "0")) + " " + (ff[1] ? pad(ff[1],rr," ") + r[2] + "/" + r[3] + rpad(ff[2],rr," "): fill(" ", 2*rr+1 + r[2].length + r[3].length));
	}
	switch(fmt) {
		case "0": case "#0": return Math.round(val);
		case "#.##": o = Math.round(val*100);
			return String(o/100).replace(/^([^\.]+)$/,"$1.").replace(/^0\.$/,".");
		case "#,###": var x = commaify(String(Math.round(aval))); return x !== "0" ? sign + x : "";
		case "#,##0": return sign + commaify(String(Math.round(aval)));
		case "#,##0.0": r = Math.round((val-Math.floor(val))*10); return val < 0 ? "-" + write_num(type, fmt, -val) : commaify(String(Math.floor(val))) + "." + pad(r,1,"0");
		case "#,##0.00": r = Math.round((val-Math.floor(val))*100); return val < 0 ? "-" + write_num(type, fmt, -val) : commaify(String(Math.floor(val))) + "." + pad(r,2,"0");
		case "#,##0.000": r = Math.round((val-Math.floor(val))*1000); return val < 0 ? "-" + write_num(type, fmt, -val) : commaify(String(Math.floor(val))) + "." + pad(r,3,0);
		case "#,##0.0000": r = Math.round((val-Math.floor(val))*10000); return val < 0 ? "-" + write_num(type, fmt, -val) : commaify(String(Math.floor(val))) + "." + pad(r,4,0);
		case "#,##0.00000": r = Math.round((val-Math.floor(val))*Math.pow(10,5)); return val < 0 ? "-" + write_num(type, fmt, -val) : commaify(String(Math.floor(val))) + "." + pad(r,5,0);
		case "#,##0.000000": r = Math.round((val-Math.floor(val))*Math.pow(10,6)); return val < 0 ? "-" + write_num(type, fmt, -val) : commaify(String(Math.floor(val))) + "." + pad(r,6,0);
		case "#,##0.0000000": r = Math.round((val-Math.floor(val))*Math.pow(10,7)); return val < 0 ? "-" + write_num(type, fmt, -val) : commaify(String(Math.floor(val))) + "." + pad(r,7,0);
		case "#,##0.00000000": r = Math.round((val-Math.floor(val))*Math.pow(10,8)); return val < 0 ? "-" + write_num(type, fmt, -val) : commaify(String(Math.floor(val))) + "." + pad(r,8,0);
		case "#,##0.000000000": r = Math.round((val-Math.floor(val))*Math.pow(10,9)); return val < 0 ? "-" + write_num(type, fmt, -val) : commaify(String(Math.floor(val))) + "." + pad(r,9,0);
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
	if(in_str !=-1) throw "Format |" + fmt + "| unterminated string at " + in_str;
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
					throw 'unrecognized character ' + fmt[i] + ' in ' + fmt;
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
			case 'm': case 'd': case 'y': case 'h': case 's': case 'e':
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
				while(fmt[i++] !== ']') o += fmt[i];
				if(o.match(/\[[HhMmSs]*\]/)) {
					if(!dt) dt = parse_date_code(v, opts);
					if(!dt) return "";
					out.push({t:'Z', v:o.toLowerCase()});
				} else { o=""; }
				break;
			/* Numbers */
			case '0': case '#':
				o = c; while("0#?.,E+-%".indexOf(c=fmt[++i]) > -1) o += c;
				out.push({t:'n', v:o}); break;
			case '?':
				o = fmt[i]; while(fmt[++i] === c) o+=c;
				q={t:c, v:o}; out.push(q); lst = c; break;
			case '*': ++i; if(fmt[i] == ' ') ++i; break; // **
			case '(': case ')': out.push({t:(flen===1?'t':c),v:c}); ++i; break;
			case '1': case '2': case '3': case '4': case '5': case '6': case '7': case '8': case '9':
				o = fmt[i]; while("0123456789".indexOf(fmt[++i]) > -1) o+=fmt[i];
				out.push({t:'D', v:o}); break;
			case ' ': out.push({t:c,v:c}); ++i; break;
			default:
				if(",$-+/():!^&'~{}<>=".indexOf(c) === -1)
					throw 'unrecognized character ' + fmt[i] + ' in ' + fmt;
				out.push({t:'t', v:c}); ++i; break;
		}
	}

	/* walk backwards */
	for(i=out.length-1, lst='t'; i >= 0; --i) {
		switch(out[i].t) {
			case 'h': case 'H': out[i].t = hr; lst='h'; break;
			case 'd': case 'y': case 's': case 'M': case 'e': lst=out[i].t; break;
			case 'm': if(lst === 's') out[i].t = 'M'; break;
		}
	}

	/* replace fields */
	for(i=0; i < out.length; ++i) {
		switch(out[i].t) {
			case 't': case 'T': case ' ': break;
			case 'd': case 'm': case 'y': case 'h': case 'H': case 'M': case 's': case 'A': case 'e': case 'Z':
				out[i].v = write_date(out[i].t, out[i].v, dt);
				out[i].t = 't'; break;
			case 'n': case '(': case '?':
				var jj = i+1;
				while(out[jj] && ("?D".indexOf(out[jj].t) > -1 || (" t".indexOf(out[jj].t) > -1 && "?t".indexOf((out[jj+1]||{}).t)>-1 && (out[jj+1].t == '?' || out[jj+1].v == '/')) || out[i].t == '(' && (out[jj].t == ')' || out[jj].t == 'n') || out[jj].t == 't' && (out[jj].v == '/' || out[jj].v == '$' || (out[jj].v == ' ' && (out[jj+1]||{}).t == '?')))) {
					out[i].v += out[jj].v;
					delete out[jj]; ++jj;
				}
				out[i].v = write_num(out[i].t, out[i].v, v);
				out[i].t = 't';
				i = jj-1; break;
			case 'G': out[i].t = 't'; out[i].v = general_fmt(v,opts); break;
			default: console.error(out); throw "unrecognized type " + out[i].t;
		}
	}
	return out.map(function(x){return x.v;}).join("");
}
SSF._eval = eval_fmt;
function choose_fmt(fmt, v, o) {
	if(typeof fmt === 'number') fmt = ((o&&o.table) ? o.table : table_fmt)[fmt];
	if(typeof fmt === "string") fmt = split_fmt(fmt);
	var l = fmt.length;
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
XLSX.version = '0.3.10';
function parsexmltag(tag) {
	var words = tag.split(/\s+/);
	var z = {'0': words[0]};
	if(words.length === 1) return z;
	(tag.match(/(\w+)="([^"]*)"/g) || []).map(
		function(x){var y=x.match(/(\w+)="([^"]*)"/); z[y[1]] = y[2]; });
	return z;
}

var encodings = {
	'&quot;': '"',
	'&apos;': "'",
	'&gt;': '>',
	'&lt;': '<',
	'&amp;': '&'
};

// TODO: CP remap (need to read file version to determine OS)
function unescapexml(text){
	var s = text + '';
	for(var y in encodings) s = s.replace(new RegExp(y,'g'), encodings[y]);
	return s.replace(/_x([0-9a-fA-F]*)_/g,function(m,c) {return _chr(parseInt(c,16));});
}

function parsexmlbool(value, tag) {
	switch(value) {
		case '0': case 0: case 'false': case 'FALSE': return false;
		case '1': case 1: case 'true': case 'TRUE': return true;
		default: throw "bad boolean value " + value + " in "+(tag||"?");
	}
}

var utf8read = function(orig) {
	var out = "", i = 0, c = 0, c1 = 0, c2 = 0, c3 = 0;
	while (i < orig.length) {
		c = orig.charCodeAt(i++);
		if (c < 128) out += _chr(c);
		else {
			c2 = orig.charCodeAt(i++);
			if (c>191 && c<224) out += _chr((c & 31) << 6 | c2 & 63);
			else {
				c3 = orig.charCodeAt(i++);
				out += _chr((c & 15) << 12 | (c2 & 63) << 6 | c3 & 63);
			}
		}
	}
	return out;
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
				case '<shadow': break;

				/* 18.4.1 charset CT_IntProperty TODO */
				case '<charset': break;

				/* 18.4.2 outline CT_BooleanProperty TODO */
				case '<outline': break;

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
var parse_si = function(x) {
	var z = {};
	if(!x) return z;
	var y;
	/* 18.4.12 t ST_Xstring (Plaintext String) */
	if(x[1] === 't') {
		z.t = utf8read(unescapexml(x.replace(/<[^>]*>/g,"")));
		z.raw = x;
		z.r = z.t;
	}
	/* 18.4.4 r CT_RElt (Rich Text Run) */
	else if((y = x.match(/<r>/))) {
		z.raw = x;
		/* TODO: properly parse (note: no other valid child can have body text) */
		z.t = utf8read(unescapexml(x.replace(/<[^>]*>/gm,"")));
		z.r = parse_rs(x);
	}
	/* 18.4.3 phoneticPr CT_PhoneticPr (TODO: needed for Asian support) */
	/* 18.4.6 rPh CT_PhoneticRun (TODO: needed for Asian support) */
	return z;
};

/* 18.4 Shared String Table */
var parse_sst = (function(){
	return function(data) {
		var s = [];
		/* 18.4.9 sst CT_Sst */
		var sst = data.match(new RegExp("<sst([^>]*)>([\\s\\S]*)<\/sst>","m"));
		if(isval(sst)) {
			s = sst[2].replace(/<si>/g,"").split(/<\/si>/).map(parse_si);
			sst = parsexmltag(sst[1]); s.Count = sst.count; s.Unique = sst.uniqueCount;
		}
		return s;
	};
})();

var ct2type = {
	"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml": "workbooks",
	"application/vnd.ms-excel.sheet.macroEnabled.main+xml":"workbooks", /*XLSM*/
	"application/vnd.openxmlformats-package.core-properties+xml": "coreprops",
	"application/vnd.openxmlformats-officedocument.extended-properties+xml": "extprops",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml": "calcchains",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml":"sheets",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml": "strs",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml":"styles",
	"application/vnd.openxmlformats-officedocument.theme+xml":"themes",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml": "comments",
	"foo": "bar"
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

var XMLNS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
var XMLNS_WB = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

var strs = {}; // shared strings
var styles = {}; // shared styles
var _ssfopts = {}; // spreadsheet formatting options

/* 18.3 Worksheets */
function parseSheet(data) {
	if(!data) return data;
	/* 18.3.1.99 worksheet CT_Worksheet */
	var s = {};

	/* 18.3.1.35 dimension CT_SheetDimension ? */
	var ref = data.match(/<dimension ref="([^"]*)"\s*\/>/);
	if(ref && ref.length == 2 && ref[1].indexOf(":") !== -1) s["!ref"] = ref[1];

	var refguess = {s: {r:1000000, c:1000000}, e: {r:0, c:0} };
	var q = ["v","f"];
	var sidx = 0;
	/* 18.3.1.80 sheetData CT_SheetData ? */
	if(!data.match(/<sheetData *\/>/))
	data.match(/<sheetData>([^\u2603]*)<\/sheetData>/m)[1].split("</row>").forEach(function(x) {
		if(x === "" || x.trim() === "") return;

		/* 18.3.1.73 row CT_Row */
		var row = parsexmltag(x.match(/<row[^>]*>/)[0]);
		if(refguess.s.r > row.r - 1) refguess.s.r = row.r - 1;
		if(refguess.e.r < row.r - 1) refguess.e.r = row.r - 1;

		/* 18.3.1.4 c CT_Cell */
		var cells = x.substr(x.indexOf('>')+1).split(/<c/);
		cells.forEach(function(c, idx) { if(c === "" || c.trim() === "") return;
			var cref = c.match(/r="([^"]*)"/);
			c = "<c" + c;
			if(cref && cref.length == 2) {
				var cref_cell = decode_cell(cref[1]);
				idx = cref_cell.c;
			}
			if(refguess.s.c > idx) refguess.s.c = idx;
			if(refguess.e.c < idx) refguess.e.c = idx;
			var cell = parsexmltag((c.match(/<c[^>]*>/)||[c])[0]); delete cell[0];
			var d = c.substr(c.indexOf('>')+1);
			var p = {};
			q.forEach(function(f){var x=d.match(matchtag(f));if(x)p[f]=unescapexml(x[1]);});

			/* SCHEMA IS ACTUALLY INCORRECT HERE.  IF A CELL HAS NO T, EMIT "" */
			if(cell.t === undefined && p.v === undefined) { p.t = "str"; p.v = undefined; }
			else p.t = (cell.t ? cell.t : "n"); // default is "n" in schema
			switch(p.t) {
				case 'n': p.v = parseFloat(p.v); break;
				case 's': {
					sidx = parseInt(p.v, 10);
					p.v = strs[sidx].t;
					p.r = strs[sidx].r;
				} break;
				case 'str': if(p.v) p.v = utf8read(p.v); break; // normal string
				case 'inlineStr':
					p.t = 'str'; p.v = unescapexml((d.match(matchtag('t'))||["",""])[1]);
					break; // inline string
				case 'b':
					switch(p.v) {
						case '0': case 'FALSE': case "false": case false: p.v=false; break;
						case '1': case 'TRUE':  case "true":  case true:  p.v=true;  break;
						default: throw "Unrecognized boolean: " + p.v;
					} break;
				/* in case of error, stick value in .raw */
				case 'e': p.raw = p.v; p.v = undefined; break;
				default: throw "Unrecognized cell type: " + p.t;
			}

			/* formatting */
			var fmtid = 0;
			if(cell.s && styles.CellXf) {
				var cf = styles.CellXf[cell.s];
				if(cf && cf.numFmtId) fmtid = cf.numFmtId;
			}
					p.raw = p.v;
					p.rawt = p.t;
					try {
				p.v = SSF.format(fmtid,p.v,_ssfopts);
						p.t = 'str';
			} catch(e) { p.v = p.raw; p.t = p.rawt; }

			s[cell.r] = p;
		});
	});
	if(!s["!ref"]) s["!ref"] = encode_range(refguess);
	return s;
}

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
				case "Worksheets": widx = j; p.Worksheets = +v[++i]; break;
				case "Named Ranges": ++i; break; // TODO: Handle Named Ranges
			}
		}
		var parts = parseVector(q.TitlesOfParts).map(utf8read);
		p.SheetNames = parts.slice(widx, widx + p.Worksheets);
	}
	p.Creator = q["dc:creator"];
	p.LastModifiedBy = q["cp:lastModifiedBy"];
	p.CreatedDate = new Date(q["dcterms:created"]);
	p.ModifiedDate = new Date(q["dcterms:modified"]);
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
		coreprops: [], extprops: [], strs:[], comments: [], xmlns: "" };
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


/* 18.2 Workbook */
function parseWB(data) {
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
	if(wb.xmlns !== XMLNS_WB) throw new Error("Unknown Namespace: " + wb.xmlns);

	var z;
	/* defaults */
	for(z in WBPropsDef) if(typeof wb.WBProps[z] === 'undefined') wb.WBProps[z] = WBPropsDef[z];
	for(z in CalcPrDef) if(typeof wb.CalcPr[z] === 'undefined') wb.CalcPr[z] = CalcPrDef[z];

	wb.WBView.forEach(function(w){for(var z in WBViewDef) if(typeof w[z] === 'undefined') w[z]=WBViewDef[z]; });
	wb.Sheets.forEach(function(w){for(var z in SheetDef) if(typeof w[z] === 'undefined') w[z]=SheetDef[z]; });

	_ssfopts.date1904 = parsexmlbool(wb.WBProps.date1904, 'date1904');

	return wb;
}

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
				styles.NumberFmt[i] = f; SSF.load(f,i);
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
			case '<xf': if(y.numFmtId) y.numFmtId = parseInt(y.numFmtId, 10);
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
function parseStyles(data) {
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
	}

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
function parseComments(data) {
	if(data.match(/<comments *\/>/)) {
		throw new Error('Not a valid comments xml');
	}
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
		var textMatch = x.match(/<text>([^\u2603]*)<\/text>/m);
		if (!textMatch || !textMatch[1]) return; // a comment may contain an empty text tag.
		var rt = parse_si(textMatch[1]);
		comment.raw = rt.raw;
		comment.t = rt.t;
		comment.r = rt.r;
		commentList.push(comment);
	});
	return commentList;
}

function parseCommentsAddToSheets(zip, dirComments, sheets, sheetRels) {
	for(var i = 0; i != dirComments.length; ++i) {
		var canonicalpath=dirComments[i];
		var comments=parseComments(getdata(getzipfile(zip, canonicalpath.replace(/^\//,''))));
		// find the sheets targeted by these comments
		var sheetNames = Object.keys(sheets);
		for(var j = 0; j != sheetNames.length; ++j) {
			var sheetName = sheetNames[j];
			var rels = sheetRels[sheetName];
			if (rels) {
				var rel = rels[canonicalpath];
				if (rel) {
					insertCommentsIntoSheet(sheetName, sheets[sheetName], comments);
				}
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

		if (!cell.c) {
			cell.c = [];
		}
		cell.c.push({a: comment.author, t: comment.t, raw: comment.raw, r: comment.r});
	});
}

function getdata(data) {
	if(!data) return null;
	if(data.data) return data.data;
	if(data._data && data._data.getContent) return Array.prototype.slice.call(data._data.getContent(),0).map(function(x) { return String.fromCharCode(x); }).join("");
	return null;
}

function getzipfile(zip, file) {
	var f = file; if(zip.files[f]) return zip.files[f];
	f = file.toLowerCase(); if(zip.files[f]) return zip.files[f];
	f = f.replace(/\//g,'\\'); if(zip.files[f]) return zip.files[f];
	throw new Error("Cannot find file " + file + " in zip");
}

function parseZip(zip) {
	var entries = Object.keys(zip.files);
	var keys = entries.filter(function(x){return x.substr(-1) != '/';}).sort();
	var dir = parseCT(getdata(getzipfile(zip, '[Content_Types].xml')));
	if(dir.workbooks.length === 0) throw new Error("Could not find workbook entry");
	strs = {};
	if(dir.sst) strs=parse_sst(getdata(getzipfile(zip, dir.sst.replace(/^\//,''))));

	styles = {};
	if(dir.style) styles = parseStyles(getdata(getzipfile(zip, dir.style.replace(/^\//,''))));

	var wb = parseWB(getdata(getzipfile(zip, dir.workbooks[0].replace(/^\//,''))));
	var propdata = dir.coreprops.length !== 0 ? getdata(getzipfile(zip, dir.coreprops[0].replace(/^\//,''))) : "";
	propdata += dir.extprops.length !== 0 ? getdata(getzipfile(zip, dir.extprops[0].replace(/^\//,''))) : "";
	var props = propdata !== "" ? parseProps(propdata) : {};
	var deps = {};
	if(dir.calcchain) deps=parseDeps(getdata(getzipfile(zip, dir.calcchain.replace(/^\//,''))));
	var sheets = {}, i=0;
	var sheetRels = {};	
	if(!props.Worksheets) {
		/* Google Docs doesn't generate the appropriate metadata, so we impute: */
		var wbsheets = wb.Sheets;
		props.Worksheets = wbsheets.length;
		props.SheetNames = [];
		for(var j = 0; j != wbsheets.length; ++j) {
			props.SheetNames[j] = wbsheets[j].name;
		}
		for(i = 0; i != props.Worksheets; ++i) {
			try { /* TODO: remove these guards */
				var path = 'xl/worksheets/sheet' + (i+1) + '.xml';
				var relsPath = path.replace(/^(.*)(\/)([^\/]*)$/, "$1/_rels/$3.rels");
				sheets[props.SheetNames[i]]=parseSheet(getdata(getzipfile(zip, path)));
				sheetRels[props.SheetNames[i]]=parseRels(getdata(getzipfile(zip, relsPath)), path);
			} catch(e) {}
		}
	} else {
		for(i = 0; i != props.Worksheets; ++i) {
			try {
				var path = dir.sheets[i].replace(/^\//,'');
				var relsPath = path.replace(/^(.*)(\/)([^\/]*)$/, "$1/_rels/$3.rels");
				sheets[props.SheetNames[i]]=parseSheet(getdata(getzipfile(zip, path)));
				sheetRels[props.SheetNames[i]]=parseRels(getdata(getzipfile(zip, relsPath)), path);
			} catch(e) {}
		}
	}

	if(dir.comments) parseCommentsAddToSheets(zip, dir.comments, sheets, sheetRels);

	return {
		Directory: dir,
		Workbook: wb,
		Props: props,
		Deps: deps,
		Sheets: sheets,
		SheetNames: props.SheetNames,
		Strings: strs,
		Styles: styles,
		keys: keys,
		files: zip.files
	};
}

var _fs, jszip;
if(typeof JSZip !== 'undefined') jszip = JSZip;
if (typeof exports !== 'undefined') {
	if (typeof module !== 'undefined' && module.exports) {
		if(typeof jszip === 'undefined') jszip = require('./jszip').JSZip;
		_fs = require('fs');
	}
}

function readSync(data, options) {
	var zip, d = data;
	var o = options||{};
	switch((o.type||"base64")){
		case "file": d = _fs.readFileSync(data).toString('base64');
			/* falls through */
		case "base64": zip = new jszip(d, { base64:true }); break;
		case "binary": zip = new jszip(d, { base64:false }); break;
	}
	return parseZip(zip);
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
	var val, row, r, hdr = {}, isempty, R, C, v;
	var out = [];
	opts = opts || {};
	if(!sheet || !sheet["!ref"]) return out;
	r = XLSX.utils.decode_range(sheet["!ref"]);
	for(R=r.s.r, C = r.s.c; C <= r.e.c; ++C) {
		val = sheet[encode_cell({c:C,r:R})];
		if(!val) continue;
		switch(val.t) {
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
			v = (val || {}).v;
			switch(val.t){
				case 's': case 'str': case 'b': case 'n':
					if(val.v !== undefined) {
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
		switch(val.t){
			case 'n': return String(val.v);
			case 's': case 'str':
				if(typeof val.v === 'undefined') return "";
				return JSON.stringify(val.v);
			case 'b': return val.v ? "TRUE" : "FALSE";
			case 'e': return ""; /* throw out value in case of error */
			default: throw 'unrecognized type ' + val.t;
		}
	};
	var out = "", txt = "";
	opts = opts || {};
	if(!sheet || !sheet["!ref"]) return out;
	var r = XLSX.utils.decode_range(sheet["!ref"]);
	for(var R = r.s.r; R <= r.e.r; ++R) {
		var row = [];
		for(var C = r.s.c; C <= r.e.c; ++C) {
			var val = sheet[XLSX.utils.encode_cell({c:C,r:R})];
			if(!val) { row.push(""); continue; }
			txt = stringify(val);
			row.push(String(txt).replace(/\\r\\n/g,"\n").replace(/\\t/g,"\t").replace(/\\\\/g,"\\").replace("\\\"","\"\""));
		}
		out += row.join(opts.FS||",") + (opts.RS||"\n");
	}
	return out;
}
var make_csv = sheet_to_csv;

function get_formulae(ws) {
	var cmds = [];
	for(var y in ws) if(y[0] !=='!' && ws.hasOwnProperty(y)) {
		var x = ws[y];
		var val = "";
		if(x.f) val = x.f;
		else if(typeof x.v === 'number') val = x.v;
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
	exports.main = function(args) {
		var zip = XLSX.read(args[0], {type:'file'});
		console.log(zip.Sheets);
	};
if(typeof module !== 'undefined' && require.main === module)
	exports.main(process.argv.slice(2));
}
