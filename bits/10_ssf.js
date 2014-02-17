/* Spreadsheet Format -- jump to XLSX for the XLSX code */
/* ssf.js (C) 2013-2014 SheetJS -- http://sheetjs.com */
var SSF = {};
var make_ssf = function(SSF){
var _strrev = function(x) { return String(x).split("").reverse().join("");};
function fill(c,l) { return new Array(l+1).join(c); }
function pad(v,d,c){var t=String(v);return t.length>=d?t:(fill(c||0,d-t.length)+t);}
function rpad(v,d,c){var t=String(v);return t.length>=d?t:(t+fill(c||0,d-t.length));}
SSF.version = '0.5.8';
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
	if((r = fmt.match(/^#*0+\.(0+)/))) {
		o = Math.round(val * Math.pow(10,r[1].length));
		return String(o/Math.pow(10,r[1].length)).replace(/^([^\.]+)$/,"$1."+r[1]).replace(/\.$/,"."+r[1]).replace(/\.([0-9]*)$/,function($$, $1) { return "." + $1 + fill("0", r[1].length-$1.length); });
	}
	if((r = fmt.match(/^(0*)\.(#*)$/))) {
		o = Math.round(val*Math.pow(10,r[2].length));
		return String(o * Math.pow(10,-r[2].length)).replace(/\.(\d*[1-9])0*$/,".$1").replace(/^([-]?\d*)$/,"$1.").replace(/^0\./,r[1].length?"0.":".");
	}
	if((r = fmt.match(/^#,##0([.]?)$/))) return sign + commaify(String(Math.round(aval)));
	if((r = fmt.match(/^#,##0\.([#0]*0)$/))) {
		rr = Math.round((val-Math.floor(val))*Math.pow(10,r[1].length));
		return val < 0 ? "-" + write_num(type, fmt, -val) : commaify(String(Math.floor(val))) + "." + pad(rr,r[1].length,0);
	}
	if((r = fmt.match(/^# ([?]+)([ ]?)\/([ ]?)([?]+)/))) {
		rr = Math.min(Math.max(r[1].length, r[4].length),7);
		ff = frac(aval, Math.pow(10,rr)-1, true);
		return sign + (ff[0]||(ff[1] ? "" : "0")) + " " + (ff[1] ? pad(ff[1],rr," ") + r[2] + "/" + r[3] + rpad(ff[2],rr," "): fill(" ", 2*rr+1 + r[2].length + r[3].length));
	}
	switch(fmt) {
		case "0": case "#0": return Math.round(val);
		case "#.##": o = Math.round(val*100);
			return String(o/100).replace(/^([^\.]+)$/,"$1.").replace(/^0\.$/,".");
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
			case '*': ++i; if(fmt[i] == ' ') ++i; break; // **
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
			case 't': case 'T': case ' ': case 'D': break;
			case 'd': case 'm': case 'y': case 'h': case 'H': case 'M': case 's': case 'e': case 'Z':
				out[i].v = write_date(out[i].t, out[i].v, dt);
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
