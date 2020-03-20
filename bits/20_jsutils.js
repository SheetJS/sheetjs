function keys(o/*:any*/)/*:Array<any>*/ {
	var ks = Object.keys(o), o2 = [];
	for(var i = 0; i < ks.length; ++i) if(Object.prototype.hasOwnProperty.call(o, ks[i])) o2.push(ks[i]);
	return o2;
}

function evert_key(obj/*:any*/, key/*:string*/)/*:EvertType*/ {
	var o = ([]/*:any*/), K = keys(obj);
	for(var i = 0; i !== K.length; ++i) if(o[obj[K[i]][key]] == null) o[obj[K[i]][key]] = K[i];
	return o;
}

function evert(obj/*:any*/)/*:EvertType*/ {
	var o = ([]/*:any*/), K = keys(obj);
	for(var i = 0; i !== K.length; ++i) o[obj[K[i]]] = K[i];
	return o;
}

function evert_num(obj/*:any*/)/*:EvertNumType*/ {
	var o = ([]/*:any*/), K = keys(obj);
	for(var i = 0; i !== K.length; ++i) o[obj[K[i]]] = parseInt(K[i],10);
	return o;
}

function evert_arr(obj/*:any*/)/*:EvertArrType*/ {
	var o/*:EvertArrType*/ = ([]/*:any*/), K = keys(obj);
	for(var i = 0; i !== K.length; ++i) {
		if(o[obj[K[i]]] == null) o[obj[K[i]]] = [];
		o[obj[K[i]]].push(K[i]);
	}
	return o;
}

var basedate = new Date(1899, 11, 30, 0, 0, 0); // 2209161600000
function datenum(v/*:Date*/, date1904/*:?boolean*/)/*:number*/ {
	var epoch = v.getTime();
	if(date1904) epoch -= 1462*24*60*60*1000;
	var dnthresh = basedate.getTime() + (v.getTimezoneOffset() - basedate.getTimezoneOffset()) * 60000;
	return (epoch - dnthresh) / (24 * 60 * 60 * 1000);
}
var refdate = new Date();
var dnthresh = basedate.getTime() + (refdate.getTimezoneOffset() - basedate.getTimezoneOffset()) * 60000;
var refoffset = refdate.getTimezoneOffset();
function numdate(v/*:number*/)/*:Date*/ {
	var out = new Date();
	out.setTime(v * 24 * 60 * 60 * 1000 + dnthresh);
	if (out.getTimezoneOffset() !== refoffset) {
		out.setTime(out.getTime() + (out.getTimezoneOffset() - refoffset) * 60000);
	}
	return out;
}

/* ISO 8601 Duration */
function parse_isodur(s) {
	var sec = 0, mt = 0, time = false;
	var m = s.match(/P([0-9\.]+Y)?([0-9\.]+M)?([0-9\.]+D)?T([0-9\.]+H)?([0-9\.]+M)?([0-9\.]+S)?/);
	if(!m) throw new Error("|" + s + "| is not an ISO8601 Duration");
	for(var i = 1; i != m.length; ++i) {
		if(!m[i]) continue;
		mt = 1;
		if(i > 3) time = true;
		switch(m[i].slice(m[i].length-1)) {
			case 'Y':
				throw new Error("Unsupported ISO Duration Field: " + m[i].slice(m[i].length-1));
			case 'D': mt *= 24;
				/* falls through */
			case 'H': mt *= 60;
				/* falls through */
			case 'M':
				if(!time) throw new Error("Unsupported ISO Duration Field: M");
				else mt *= 60;
				/* falls through */
			case 'S': break;
		}
		sec += mt * parseInt(m[i], 10);
	}
	return sec;
}

var good_pd_date = new Date('2017-02-19T19:06:09.000Z');
if(isNaN(good_pd_date.getFullYear())) good_pd_date = new Date('2/19/17');
var good_pd = good_pd_date.getFullYear() == 2017;
/* parses a date as a local date */
function parseDate(str/*:string|Date*/, fixdate/*:?number*/)/*:Date*/ {
	var d = new Date(str);
	if(good_pd) {
		/*:: if(fixdate == null) fixdate = 0; */
		if(fixdate > 0) d.setTime(d.getTime() + d.getTimezoneOffset() * 60 * 1000);
		else if(fixdate < 0) d.setTime(d.getTime() - d.getTimezoneOffset() * 60 * 1000);
		return d;
	}
	if(str instanceof Date) return str;
	if(good_pd_date.getFullYear() == 1917 && !isNaN(d.getFullYear())) {
		var s = d.getFullYear();
		if(str.indexOf("" + s) > -1) return d;
		d.setFullYear(d.getFullYear() + 100); return d;
	}
	var n = str.match(/\d+/g)||["2017","2","19","0","0","0"];
	var out = new Date(+n[0], +n[1] - 1, +n[2], (+n[3]||0), (+n[4]||0), (+n[5]||0));
	if(str.indexOf("Z") > -1) out = new Date(out.getTime() - out.getTimezoneOffset() * 60 * 1000);
	return out;
}

function cc2str(arr/*:Array<number>*/)/*:string*/ {
	var o = "";
	for(var i = 0; i != arr.length; ++i) o += String.fromCharCode(arr[i]);
	return o;
}

function dup(o/*:any*/)/*:any*/ {
	if(typeof JSON != 'undefined' && !Array.isArray(o)) return JSON.parse(JSON.stringify(o));
	if(typeof o != 'object' || o == null) return o;
	if(o instanceof Date) return new Date(o.getTime());
	var out = {};
	for(var k in o) if(Object.prototype.hasOwnProperty.call(o, k)) out[k] = dup(o[k]);
	return out;
}

function fill(c/*:string*/,l/*:number*/)/*:string*/ { var o = ""; while(o.length < l) o+=c; return o; }

/* TODO: stress test */
function fuzzynum(s/*:string*/)/*:number*/ {
	var v/*:number*/ = Number(s);
	if(!isNaN(v)) return v;
	var wt = 1;
	var ss = s.replace(/([\d]),([\d])/g,"$1$2").replace(/[$]/g,"").replace(/[%]/g, function() { wt *= 100; return "";});
	if(!isNaN(v = Number(ss))) return v / wt;
	ss = ss.replace(/[(](.*)[)]/,function($$, $1) { wt = -wt; return $1;});
	if(!isNaN(v = Number(ss))) return v / wt;
	return v;
}
function fuzzydate(s/*:string*/)/*:Date*/ {
	var o = new Date(s), n = new Date(NaN);
	var y = o.getYear(), m = o.getMonth(), d = o.getDate();
	if(isNaN(d)) return n;
	if(y < 0 || y > 8099) return n;
	if((m > 0 || d > 1) && y != 101) return o;
	if(s.toLowerCase().match(/jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec/)) return o;
	if(s.match(/[^-0-9:,\/\\]/)) return n;
	return o;
}

var safe_split_regex = "abacaba".split(/(:?b)/i).length == 5;
function split_regex(str/*:string*/, re, def/*:string*/)/*:Array<string>*/ {
	if(safe_split_regex || typeof re == "string") return str.split(re);
	var p = str.split(re), o = [p[0]];
	for(var i = 1; i < p.length; ++i) { o.push(def); o.push(p[i]); }
	return o;
}
