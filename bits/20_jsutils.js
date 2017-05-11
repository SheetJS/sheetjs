function isval(x/*:?any*/)/*:boolean*/ { return x !== undefined && x !== null; }

function keys(o/*:any*/)/*:Array<any>*/ { return Object.keys(o); }

function evert_key(obj/*:any*/, key/*:string*/)/*:EvertType*/ {
	var o = ([]/*:any*/), K = keys(obj);
	for(var i = 0; i !== K.length; ++i) o[obj[K[i]][key]] = K[i];
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

function datenum(v/*:Date*/, date1904/*:?boolean*/)/*:number*/ {
	var epoch = v.getTime();
	if(date1904) epoch += 1462*24*60*60*1000;
	return (epoch + 2209161600000) / (24 * 60 * 60 * 1000);
}
function numdate(v/*:number*/)/*:Date*/ {
	var date = SSF.parse_date_code(v);
	var val = new Date();
	if(date == null) throw new Error("Bad Date Code: " + v);
	val.setUTCDate(date.d);
	val.setUTCMonth(date.m-1);
	val.setUTCFullYear(date.y);
	val.setUTCHours(date.H);
	val.setUTCMinutes(date.M);
	val.setUTCSeconds(date.S);
	return val;
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
		switch(m[i].substr(m[i].length-1)) {
			case 'Y':
				throw new Error("Unsupported ISO Duration Field: " + m[i].substr(m[i].length-1));
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
var good_pd = good_pd_date.getFullYear() == 2017;
function parseDate(str/*:string|Date*/)/*:Date*/ {
	if(good_pd) return new Date(str);
	if(str instanceof Date) return str;
	var n = str.match(/\d+/g)||["2017","2","19","0","0","0"];
	return new Date(Date.UTC(+n[0], +n[1] - 1, +n[2], +n[3], +n[4], +n[5]));
}

function cc2str(arr/*:Array<number>*/)/*:string*/ {
	var o = "";
	for(var i = 0; i != arr.length; ++i) o += String.fromCharCode(arr[i]);
	return o;
}

function str2cc(str) {
	var o = [];
	for(var i = 0; i != str.length; ++i) o.push(str.charCodeAt(i));
	return o;
}

function dup(o/*:any*/)/*:any*/ {
	if(typeof JSON != 'undefined' && !Array.isArray(o)) return JSON.parse(JSON.stringify(o));
	if(typeof o != 'object' || o == null) return o;
	var out = {};
	for(var k in o) if(o.hasOwnProperty(k)) out[k] = dup(o[k]);
	return out;
}

function fill(c/*:string*/,l/*:number*/)/*:string*/ { var o = ""; while(o.length < l) o+=c; return o; }

/* TODO: stress test */
function fuzzydate(s/*:string*/)/*:Date*/ {
	var o = new Date(s), n = new Date(NaN);
	var y = o.getYear(), m = o.getMonth(), d = o.getDate();
	if(isNaN(d)) return n;
	if(y < 0 || y > 8099) return n;
	if((m > 0 || d > 1) && y != 101) return o;
	if(s.toLowerCase().match(/jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec/)) return o;
	if(!s.match(/[a-zA-Z]/)) return o;
	return n;
}

