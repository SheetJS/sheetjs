/* Spreadsheet Format -- jump to XLSX for the XLSX code */
/* ssf.js (C) 2013 SheetJS -- http://sheetjs.com */
var SSF = {};
var make_ssf = function(SSF){
String.prototype.reverse=function(){return this.split("").reverse().join("");};
var _strrev = function(x) { return String(x).reverse(); };
function fill(c,l) { return new Array(l+1).join(c); }
function pad(v,d,c){var t=String(v);return t.length>=d?t:(fill(c||0,d-t.length)+t);}
function rpad(v,d,c){var t=String(v);return t.length>=d?t:(t+fill(c||0,d-t.length));}
/* Options */
var opts_fmt = {};
function fixopts(o){for(var y in opts_fmt) if(o[y]===undefined) o[y]=opts_fmt[y];}
SSF.opts = opts_fmt;
opts_fmt.date1904 = 0;
opts_fmt.output = "";
opts_fmt.mode = "";
var table_fmt = {
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
	49: '@'
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
	var A = B|0;
	while(Q_1 < D) {
		A = B|0;
		P = A * P_1 + P_2;
		Q = A * Q_1 + Q_2;
		if((B - A) < 0.0000000001) break;
		B = 1 / (B - A);
		P_2 = P_1; P_1 = P;
		Q_2 = Q_1; Q_1 = Q;
	}
	if(Q > D) { Q = Q_1; P = P_1; }
	if(Q > D) { Q = Q_2; P = P_2; }
	if(!mixed) return [0, sgn * P, Q];
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
	var date = Math.floor(v), time = Math.round(86400 * (v - date)), dow=0;
	var dout=[], out={D:date, T:time, u:86400*(v-date)-time}; fixopts(opts = (opts||{}));
	if(opts.date1904) date += 1462;
	if(date === 60) {dout = [1900,2,29]; dow=3;}
	else if(date === 0) {dout = [1900,1,0]; dow=6;}
	else {
		if(date > 60) --date;
		/* 1 = Jan 1 1900 */
		var d = new Date(1900,0,1);
		d.setDate(d.getDate() + date - 1);
		dout = [d.getFullYear(), d.getMonth()+1,d.getDate()];
		dow = d.getDay();
		if(opts.mode === 'excel' && date < 60) dow = (dow + 6) % 7;
	}
	out.y = dout[0]; out.m = dout[1]; out.d = dout[2];
	out.S = time % 60; time = Math.floor(time / 60);
	out.M = time % 60; time = Math.floor(time / 60);
	out.H = time;
	out.q = dow;
	return out;
};
SSF.parse_date_code = parse_date_code;
var write_date = function(type, fmt, val) {
	if(val < 0) return "";
	switch(type) {
		case 'y': switch(fmt) { /* year */
			case 'y': case 'yy': return pad(val.y % 100,2);
			default: return val.y;
		} break;
		case 'm': switch(fmt) { /* month */
			case 'm': return val.m;
			case 'mm': return pad(val.m,2);
			case 'mmm': return months[val.m-1][1];
			case 'mmmm': return months[val.m-1][2];
			case 'mmmmm': return months[val.m-1][0];
			default: throw 'bad month format: ' + fmt;
		} break;
		case 'd': switch(fmt) { /* day */
			case 'd': return val.d;
			case 'dd': return pad(val.d,2);
			case 'ddd': return days[val.q][0];
			case 'dddd': return days[val.q][1];
			default: throw 'bad day format: ' + fmt;
		} break;
		case 'h': switch(fmt) { /* 12-hour */
			case 'h': return 1+(val.H+11)%12;
			case 'hh': return pad(1+(val.H+11)%12, 2);
			default: throw 'bad hour format: ' + fmt;
		} break;
		case 'H': switch(fmt) { /* 24-hour */
			case 'h': return val.H;
			case 'hh': return pad(val.H, 2);
			default: throw 'bad hour format: ' + fmt;
		} break;
		case 'M': switch(fmt) { /* minutes */
			case 'm': return val.M;
			case 'mm': return pad(val.M, 2);
			default: throw 'bad minute format: ' + fmt;
		} break;
		case 's': switch(fmt) { /* seconds */
			case 's': return val.S;
			case 'ss': return pad(val.S, 2);
			case 'ss.0': return pad(val.S,2) + "." + Math.round(10*val.u);
			default: throw 'bad second format: ' + fmt;
		} break;
		case 'Z': switch(fmt) {
			case '[h]': return val.D*24+val.H;
			default: throw 'bad abstime format: ' + fmt;
		} break;
		/* TODO: handle the ECMA spec format ee -> yy */
		case 'e': { return val.y; } break;
		case 'A': return (val.h>=12 ? 'P' : 'A') + fmt.substr(1);
		default: throw 'bad format type ' + type + ' in ' + fmt;
	}
};
String.prototype.reverse = function() { return this.split("").reverse().join(""); };
var commaify = function(s) { return s.reverse().replace(/.../g,"$&,").reverse().replace(/^,/,""); };
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
			var ee = Number(val.toExponential(0).substr(3))%3;
			o = (val/Math.pow(10,ee%3)).toPrecision(idx+1+(ee%3)).replace(/^([+-]?)([0-9]*)\.([0-9]*)[Ee]/,function($$,$1,$2,$3) { return $1 + $2 + $3.substr(0,ee) + "." + $3.substr(ee) + "E"; });
		} else o = val.toExponential(idx);
		if(fmt.match(/E\+00$/) && o.match(/e[+-][0-9]$/)) o = o.substr(0,o.length-1) + "0" + o[o.length-1];
		if(fmt.match(/E\-/) && o.match(/e\+/)) o = o.replace(/e\+/,"e");
		return o.replace("e","E");
	}
  if(fmt[0] === "$") return "$"+write_num(type,fmt.substr(fmt[1]==' '?2:1),val);
	var r, ff, aval = val < 0 ? -val : val, sign = val < 0 ? "-" : "";
	if((r = fmt.match(/# (\?+) \/ (\d+)/))) {
		var den = Number(r[2]), rnd = Math.round(aval * den), base = Math.floor(rnd/den);
		var myn = (rnd - base*den), myd = den;
		return sign + (base?base:"") + " " + (myn === 0 ? fill(" ", r[1].length + 1 + r[2].length) : pad(myn,r[1].length," ") + "/" + pad(myd,r[2].length));
	}
	if(fmt.match(/^00*$/)) return (val<0?"-":"")+pad(Math.round(Math.abs(val)), fmt.length);
	if(fmt.match(/^####*$/)) return "dafuq";
	switch(fmt) {
		case "0": return Math.round(val);
		case "0.0": o = Math.round(val*10);
			return String(o/10).replace(/^([^\.]+)$/,"$1.0").replace(/\.$/,".0");
		case "0.00": o = Math.round(val*100);
			return String(o/100).replace(/^([^\.]+)$/,"$1.00").replace(/\.$/,".00").replace(/\.([0-9])$/,".$1"+"0");
		case "0.000": o = Math.round(val*1000);
			return String(o/1000).replace(/^([^\.]+)$/,"$1.000").replace(/\.$/,".000").replace(/\.([0-9])$/,".$1"+"00").replace(/\.([0-9][0-9])$/,".$1"+"0");
		case "#,##0": return sign + commaify(String(Math.round(aval)));
		case "#,##0.0": r = Math.round((val-Math.floor(val))*10); return val < 0 ? "-" + write_num(type, fmt, -val) : commaify(String(Math.floor(val))) + "." + r;
		case "#,##0.00": r = Math.round((val-Math.floor(val))*100); return val < 0 ? "-" + write_num(type, fmt, -val) : commaify(String(Math.floor(val))) + "." + (r < 10 ? "0"+r:r);
		case "# ? / ?": ff = frac(aval, 9, true); return sign + (ff[0]||"") + " " + (ff[1] === 0 ? "   " : ff[1] + "/" + ff[2]);
		case "# ?? / ??": ff = frac(aval, 99, true); return sign + (ff[0]||"") + " " + (ff[1] ? pad(ff[1],2," ") + "/" + rpad(ff[2],2," ") : "     ");
		case "# ??? / ???": ff = frac(aval, 999, true); return sign + (ff[0]||"") + " " + (ff[1] ? pad(ff[1],3," ") + "/" + rpad(ff[2],3," ") : "       ");
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
			case '"': /* Literal text */
				for(o="";fmt[++i] !== '"';) o += fmt[i];
				out.push({t:'t', v:o}); ++i; break;
			case '\\': var w = fmt[++i], t = "()".indexOf(w) === -1 ? 't' : w;
				out.push({t:t, v:w}); ++i; break;
			case '_': out.push({t:'t', v:" "}); i+=2; break;
			case '@': /* Text Placeholder */
				out.push({t:'T', v:v}); ++i; break;
			/* Dates */
			case 'm': case 'd': case 'y': case 'h': case 's': case 'e':
				if(v < 0) return "";
				if(!dt) dt = parse_date_code(v, opts);
				o = fmt[i]; while(fmt[++i] === c) o+=c;
				if(c === 's' && fmt[i] === '.' && fmt[i+1] === '0') { o+='.'; while(fmt[++i] === '0') o+= '0'; }
				if(c === 'm' && lst.toLowerCase() === 'h') c = 'M'; /* m = minute */
				if(c === 'h') c = hr;
				q={t:c, v:o}; out.push(q); lst = c; break;
			case 'A':
				if(!dt) dt = parse_date_code(v, opts);
				q={t:c,v:"A"};
				if(fmt.substr(i, 3) === "A/P") {q.v = dt.H >= 12 ? "P" : "A"; q.t = 'T'; hr='h';i+=3;}
				else if(fmt.substr(i,5) === "AM/PM") { q.v = dt.H >= 12 ? "PM" : "AM"; q.t = 'T'; i+=5; hr='h'; }
				else q.t = "t";
				out.push(q); lst = c; break;
			case '[': /* TODO: Fix this -- ignore all conditionals and formatting */
				o = c;
				while(fmt[i++] !== ']') o += fmt[i];
				if(o == "[h]") out.push({t:'Z', v:o});
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
				if("$-+/():!^&'~{}<>=".indexOf(c) === -1)
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
			case 'n': case '(':
				var jj = i+1;
				while(out[jj] && ("? D".indexOf(out[jj].t) > -1 || out[i].t == '(' && (out[jj].t == ')' || out[jj].t == 'n') || out[jj].t == 't' && (out[jj].v == '/' || out[jj].v == '$' || (out[jj].v == ' ' && (out[jj+1]||{}).t == '?')))) {
					if(out[jj].v!==' ') out[i].v += ' ' + out[jj].v;
					delete out[jj]; ++jj;
				}
				out[i].v = write_num(out[i].t, out[i].v, v);
				out[i].t = 't';
				i = jj; break;
			default: throw "unrecognized type " + out[i].t;
		}
	}

	return out.map(function(x){return x.v;}).join("");
}
SSF._eval = eval_fmt;
function choose_fmt(fmt, v, o) {
	if(typeof fmt === 'number') fmt = table_fmt[fmt];
	if(typeof fmt === "string") fmt = split_fmt(fmt);
	var l = fmt.length;
	switch(fmt.length) {
		case 1: fmt = [fmt[0], fmt[0], fmt[0], "@"]; break;
		case 2: fmt = [fmt[0], fmt[fmt[1] === "@"?0:1], fmt[0], "@"]; break;
		case 4: break;
		default: throw "cannot find right format for |" + fmt + "|";
	}
	if(typeof v !== "number") return [fmt.length, fmt[3]];
	return [l, v > 0 ? fmt[0] : v < 0 ? fmt[1] : fmt[2]];
}

var format = function format(fmt,v,o) {
	fixopts(o = (o||{}));
	if(fmt === 0) return general_fmt(v, o);
	if(typeof fmt === 'number') fmt = table_fmt[fmt];
	var f = choose_fmt(fmt, v, o);
	return eval_fmt(f[1], v, o, f[0]);
};

SSF._choose = choose_fmt;
SSF._table = table_fmt;
SSF.load = function(fmt, idx) { table_fmt[idx] = fmt; };
SSF.format = format;
};
make_ssf(SSF);
