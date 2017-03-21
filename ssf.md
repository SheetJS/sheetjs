# SSF

SpreadSheet Format (SSF) is a pure-JS library to format data using ECMA-376
spreadsheet format codes.

## Options

The various API functions take an `opts` argument which control parsing.  The
default options are described below:

```js>bits/10_opts.js
/* Options */
var opts_fmt/*:Array<Array<any> >*/ = [
```

There are two commonly-recognized date code formats:
 - 1900 mode (where date=0 is 1899-12-31)
 - 1904 mode (where date=0 is 1904-01-01)

The difference between the the 1900 and 1904 date modes is 1462 days.  Since
the 1904 date mode was only default in a few Mac variants of Excel (2011 uses
1900 mode), the default is 1900 mode.  Consistent with ECMA-376 the name is
`date1904`:

```
  ["date1904", 0],
```

The default output is a text representation (no effort to capture colors).  To
control the output, set the `output` variable:

- `text`: no color (default)
- `html`: html output using
- `ansi`: ansi color codes (requires `colors` module)

```
  ["output", ""],
```

These options are made available via the `opts` field:

```
  ["WTF", false]
];
function fixopts(o){
  for(var y = 0; y != opts_fmt.length; ++y) if(o[opts_fmt[y][0]]===undefined) o[opts_fmt[y][0]]=opts_fmt[y][1];
}
SSF.opts = opts_fmt;
```

## Conditional Format Codes

The specification is a bit unclear here.  It initially claims in §18.3.1:

> Up to four sections of format codes can be specified. The format codes,
separated by semicolons, define the formats for positive numbers, negative
numbers, zero values, and text, in that order.

Semicolons can be escaped with the `\` character, so we need to split on those
semicolons that aren't prefaced by a slash or within a quoted string:

```js>bits/80_split.js
function split_fmt(fmt/*:string*/)/*:Array<string>*/ {
  var out/*:Array<string>*/ = [];
  var in_str = false, cc;
  for(var i = 0, j = 0; i < fmt.length; ++i) switch((cc=fmt.charCodeAt(i))) {
    case 34: /* '"' */
      in_str = !in_str; break;
    case 95: case 42: case 92: /* '_' '*' '\\' */
      ++i; break;
    case 59: /* ';' */
      out[out.length] = fmt.substr(j,i-j);
      j = i+1;
  }
  out[out.length] = fmt.substr(j);
  if(in_str === true) throw new Error("Format |" + fmt + "| unterminated string ");
  return out;
}
SSF._split = split_fmt;
```

But it also allows for conditional formatting:

> To set number formats that are applied only if a number meets a specified
> condition, enclose the condition in square brackets. The condition consists
> of a comparison operator and a value. Comparison operators include:
> `=` Equal to;
> `>` Greater than;
> `<` Less than;
> `>=` Greater than or equal to,
> `<=` Less than or equal to,
> and `<>` Not equal to.

One problem is that Excel doesn't support three conditionals.  For example:

```>
[Red][<-25]General;[Blue][>25]General;[Green][<>0]General;[Yellow]General
```

One would expect that the format code would color all numbers that are `< -25`
in red, all numbers `> 25` in blue, nonzero numbers between `-25` and `25` in
green, and color `0` and text in yellow.  Excel doesn't do that.

The two-conditional case works in an "expected" way if you interpret the third
clause as the case for numbers that don't fit the first two:

```>
[Red][<-25]General;[Blue][>25]General;[Green]General;[Yellow]General
```

will render values below `-25` as Red, above `25` as Blue, Green for other
numbers, and Yellow for text.

Only the text case is allowed to have the `@` text sigil.  Excel interprets it
as the last format.

## Utility Functions

```js>bits/02_utilities.js
function _strrev(x/*:string*/)/*:string*/ { var o = "", i = x.length-1; while(i>=0) o += x.charAt(i--); return o; }
function fill(c/*:string*/,l/*:number*/)/*:string*/ { var o = ""; while(o.length < l) o+=c; return o; }
```

The next few helpers break up the general `pad` function into special cases:

```
function pad0(v/*:any*/,d/*:number*/)/*:string*/{var t=""+v; return t.length>=d?t:fill('0',d-t.length)+t;}
function pad_(v/*:any*/,d/*:number*/)/*:string*/{var t=""+v;return t.length>=d?t:fill(' ',d-t.length)+t;}
function rpad_(v/*:any*/,d/*:number*/)/*:string*/{var t=""+v; return t.length>=d?t:t+fill(' ',d-t.length);}
function pad0r1(v/*:any*/,d/*:number*/)/*:string*/{var t=""+Math.round(v); return t.length>=d?t:fill('0',d-t.length)+t;}
function pad0r2(v/*:any*/,d/*:number*/)/*:string*/{var t=""+v; return t.length>=d?t:fill('0',d-t.length)+t;}
var p2_32 = Math.pow(2,32);
function pad0r(v/*:any*/,d/*:number*/)/*:string*/{if(v>p2_32||v<-p2_32) return pad0r1(v,d); var i = Math.round(v); return pad0r2(i,d); }
```

Comparing against the string "general" is faster via char codes:

```
function isgeneral(s/*:string*/, i/*:?number*/)/*:boolean*/ { i = i || 0; return s.length >= 7 + i && (s.charCodeAt(i)|32) === 103 && (s.charCodeAt(i+1)|32) === 101 && (s.charCodeAt(i+2)|32) === 110 && (s.charCodeAt(i+3)|32) === 101 && (s.charCodeAt(i+4)|32) === 114 && (s.charCodeAt(i+5)|32) === 97 && (s.charCodeAt(i+6)|32) === 108; }
```

## General Number Format

The 'general' format for spreadsheets (identified by format code 0) is highly
context-sensitive and the implementation tries to follow the format to the best
of its abilities given the knowledge.

First: 32-bit integers in base 10 are shorter than 11 characters, so they will
always be written in full:

```js>bits/40_general.js
function general_fmt_int(v/*:number*/, opts/*:?any*/)/*:string*/ { return ""+v; }
SSF._general_int = general_fmt_int;
```

Next: other numbers require some finessing:

```
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
  //for(var i = 0; i != o.length; ++i) if(o.charCodeAt(i) === 46) return o.replace(gnr2,"").replace(gnr1,".$1");
  //return o;
  return o.indexOf(".") > -1 ? o.replace(gnr2,"").replace(gnr1,".$1") : o;
}
return function general_fmt_num(v/*:number*/, opts/*:?any*/)/*:string*/ {
  var V = Math.floor(Math.log(Math.abs(v))*Math.LOG10E), o;
  if(V >= -4 && V <= -1) o = v.toPrecision(10+V);
  else if(Math.abs(V) <= 9) o = gfn2(v);
  else if(V === 10) o = v.toFixed(10).substr(0,12);
  else o = gfn3(v);
  return gfn5(gfn4(o));
};})();
SSF._general_num = general_fmt_num;
```

Finally

```js>bits/40_general.js
function general_fmt(v/*:any*/, opts/*:?any*/) {
  switch(typeof v) {
```

For strings, just return the text as-is:

```
    case 'string': return v;
```

Booleans are serialized in upper case:

```
    case 'boolean': return v ? "TRUE" : "FALSE";
```

For numbers, call the relevant function (int or num) based on the value:

```
    case 'number': return (v|0) === v ? general_fmt_int(v, opts) : general_fmt_num(v, opts);
  }
```

Anything else is bad:

```
  throw new Error("unsupported value in General format: " + v);
}
SSF._general = general_fmt;
```

## Implied Number Formats

These are the commonly-used formats that have a special implied code.
None of the international formats are included here.

```js>bits/20_consts.js
var table_fmt = {
  /*::[*/0/*::]*/:  'General',
  /*::[*/1/*::]*/:  '0',
  /*::[*/2/*::]*/:  '0.00',
  /*::[*/3/*::]*/:  '#,##0',
  /*::[*/4/*::]*/:  '#,##0.00',
  /*::[*/9/*::]*/:  '0%',
  /*::[*/10/*::]*/: '0.00%',
  /*::[*/11/*::]*/: '0.00E+00',
  /*::[*/12/*::]*/: '# ?/?',
  /*::[*/13/*::]*/: '# ??/??',
```

Now Excel and other formats treat code 14 as `m/d/yy` (with slashes).  Given
that the spec gives no internationalization considerations, erring on the side
of the applications makes sense here:

```
  /*::[*/14/*::]*/: 'm/d/yy',
  /*::[*/15/*::]*/: 'd-mmm-yy',
  /*::[*/16/*::]*/: 'd-mmm',
  /*::[*/17/*::]*/: 'mmm-yy',
  /*::[*/18/*::]*/: 'h:mm AM/PM',
  /*::[*/19/*::]*/: 'h:mm:ss AM/PM',
  /*::[*/20/*::]*/: 'h:mm',
  /*::[*/21/*::]*/: 'h:mm:ss',
  /*::[*/22/*::]*/: 'm/d/yy h:mm',
  /*::[*/37/*::]*/: '#,##0 ;(#,##0)',
  /*::[*/38/*::]*/: '#,##0 ;[Red](#,##0)',
  /*::[*/39/*::]*/: '#,##0.00;(#,##0.00)',
  /*::[*/40/*::]*/: '#,##0.00;[Red](#,##0.00)',
  /*::[*/45/*::]*/: 'mm:ss',
  /*::[*/46/*::]*/: '[h]:mm:ss',
  /*::[*/47/*::]*/: 'mmss.0',
  /*::[*/48/*::]*/: '##0.0E+0',
  /*::[*/49/*::]*/: '@',
```

There are special implicit format codes identified in [ECMA-376] 18.8.30.
Assuming zh-tw is the default:

```
  /*::[*/56/*::]*/: '"上午/下午 "hh"時"mm"分"ss"秒 "',
```

some writers erroneously emit 65535 for general:

```
  /*::[*/65535/*::]*/: 'General'
};
```

## Dates and Time

The code `ddd` displays short day-of-week and `dddd` shows long day-of-week:

```js>bits/20_consts.js
var days = [
  ['Sun', 'Sunday'],
  ['Mon', 'Monday'],
  ['Tue', 'Tuesday'],
  ['Wed', 'Wednesday'],
  ['Thu', 'Thursday'],
  ['Fri', 'Friday'],
  ['Sat', 'Saturday']
];

```

`mmm` shows short month, `mmmm` shows long month, and `mmmmm` shows one char:

```
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
```

## Parsing Date and Time Codes

Most spreadsheet formats store dates and times as floating point numbers (where
the integer part is a day code based on a format and the fractional part is the
portion of a 24 hour day).

Excel supports the alternative Hijri calendar (indicated with `b2`):

```js>bits/50_date.js
function parse_date_code(v/*:number*/,opts/*:?any*/,b2/*:?boolean*/) {
```

Date codes beyond 12/31/9999 are invalid:

```
  if(v > 2958465 || v < 0) return null;
```

Now we can parse!

```
  var date = (v|0), time = Math.floor(86400 * (v - date)), dow=0;
  var dout=[];
  var out={D:date, T:time, u:86400*(v-date)-time,y:0,m:0,d:0,H:0,M:0,S:0,q:0};
  if(Math.abs(out.u) < 1e-6) out.u = 0;
  fixopts(opts != null ? opts : (opts=[]));
```

Excel help actually recommends treating the 1904 date codes as 1900 date codes
shifted by 1462 days.

```
  if(opts.date1904) date += 1462;
```

Due to floating point issues, correct for subseconds:

```
  if(out.u > 0.999) {
    out.u = 0;
    if(++time == 86400) { time = 0; ++date; }
  }
```

Due to a bug in Lotus 1-2-3 which was propagated by Excel and other variants,
the year 1900 is recognized as a leap year.  JS has no way of representing that
abomination as a `Date`, so the easiest way is to store the data as a tuple.

February 29, 1900 (date `60`) is recognized as a Wednesday.  Date `0` is treated
as January 0, 1900 rather than December 31, 1899.

```
  if(date === 60) {dout = b2 ? [1317,10,29] : [1900,2,29]; dow=3;}
  else if(date === 0) {dout = b2 ? [1317,8,29] : [1900,1,0]; dow=6;}
```

For the other dates, using the JS date mechanism suffices.

```
  else {
    if(date > 60) --date;
    /* 1 = Jan 1 1900 in Gregorian */
    var d = new Date(1900, 0, 1);
    d.setDate(d.getDate() + date - 1);
    dout = [d.getFullYear(), d.getMonth()+1,d.getDate()];
    dow = d.getDay();
```

Note that Excel opted to keep the day-of-week metric consistent with the extra
day.  In practice, that means the days before the fake leap day are off.  For
example, date code `55` is "Friday, February 24, 1900" when in fact it was a
Saturday.  The "right" thing to do is to keep the DOW consistent and just break
the fact that there are two Wednesdays in that "week".

```
    if(date < 60) dow = (dow + 6) % 7;
```

For the hijri calendar, the date needs to be fixed

```
    if(b2) dow = fix_hijri(d, dout);
  }
```

Because JS dates cannot represent the bad leap day, this returns an object:

```
  out.y = dout[0]; out.m = dout[1]; out.d = dout[2];
  out.S = time % 60; time = Math.floor(time / 60);
  out.M = time % 60; time = Math.floor(time / 60);
  out.H = time;
  out.q = dow;
  return out;
}
SSF.parse_date_code = parse_date_code;
```

TODO: suitable hijri correction

```js>bits/45_hijri.js
function fix_hijri(date, o) { return 0; }
```

## Evaluating Number Formats

The utility `commaify` adds commas to integers:

```js>bits/56_commaify.js
function commaify(s/*:string*/)/*:string*/ {
  if(s.length <= 3) return s;
  var j = (s.length % 3), o = s.substr(0,j);
  for(; j!=s.length; j+=3) o+=(o.length > 0 ? "," : "") + s.substr(j,3);
  return o;
}
```

`write_num` is broken into sub-functions to help with optimization:

```js>bits/57_numhead.js
var write_num = (function make_write_num(){
```

### Percentages

The underlying number for the percentages should be physically shifted:

```js>bits/59_numhelp.js
var pct1 = /%/g;
function write_num_pct(type/*:string*/, fmt/*:string*/, val/*:number*/)/*:string*/{
  var sfmt = fmt.replace(pct1,""), mul = fmt.length - sfmt.length;
  return write_num(type, sfmt, val * Math.pow(10,2*mul)) + fill("%",mul);
}
```

### Trailing Commas

Formats with multiple commas after the decimal point should be shifted by the
appropiate multiple of 1000 (more magic):

```js>bits/60_number.js
function write_num_cm(type/*:string*/, fmt/*:string*/, val/*:number*/)/*:string*/{
  var idx = fmt.length - 1;
  while(fmt.charCodeAt(idx-1) === 44) --idx;
  return write_num(type, fmt.substr(0,idx), val / Math.pow(10,3*(fmt.length-idx)));
}
```

### Exponential

For exponents, get the exponent and mantissa and format them separately:

```
function write_num_exp(fmt/*:string*/, val/*:number*/)/*:string*/{
  var o/*:string*/;
  var idx = fmt.indexOf("E") - fmt.indexOf(".") - 1;
```

For the special case of engineering notation, "shift" the decimal:

```
  if(fmt.match(/^#+0.0E\+0$/)) {
    var period = fmt.indexOf("."); if(period === -1) period=fmt.indexOf('E');
    var ee = Math.floor(Math.log(Math.abs(val))*Math.LOG10E)%period;
    if(ee < 0) ee += period;
    o = (val/Math.pow(10,ee)).toPrecision(idx+1+(period+ee)%period);
    if(o.indexOf("e") === -1) {
```

TODO: something cleaner

```
      var fakee = Math.floor(Math.log(Math.abs(val))*Math.LOG10E);
      if(o.indexOf(".") === -1) o = o.charAt(0) + "." + o.substr(1) + "E+" + (fakee - o.length+ee);
      else o += "E+" + (fakee - ee);
      while(o.substr(0,2) === "0.") {
        o = o.charAt(0) + o.substr(2,period) + "." + o.substr(2+period);
        o = o.replace(/^0+([1-9])/,"$1").replace(/^0+\./,"0.");
      }
      o = o.replace(/\+-/,"-");
    }
    o = o.replace(/^([+-]?)(\d*)\.(\d*)[Ee]/,function($$,$1,$2,$3) { return $1 + $2 + $3.substr(0,(period+ee)%period) + "." + $3.substr(ee) + "E"; });
  } else o = val.toExponential(idx);
  if(fmt.match(/E\+00$/) && o.match(/e[+-]\d$/)) o = o.substr(0,o.length-1) + "0" + o.charAt(o.length-1);
  if(fmt.match(/E\-/) && o.match(/e\+/)) o = o.replace(/e\+/,"e");
  return o.replace("e","E");
}
```

### Fractions

```
var frac1 = /# (\?+)( ?)\/( ?)(\d+)/;
function write_num_f1(r/*:Array<string>*/, aval/*:number*/, sign/*:string*/)/*:string*/ {
  var den = parseInt(r[4],10), rr = Math.round(aval * den), base = Math.floor(rr/den);
  var myn = (rr - base*den), myd = den;
  return sign + (base === 0 ? "" : ""+base) + " " + (myn === 0 ? fill(" ", r[1].length + 1 + r[4].length) : pad_(myn,r[1].length) + r[2] + "/" + r[3] + pad0(myd,r[4].length));
}
function write_num_f2(r/*:Array<string>*/, aval/*:number*/, sign/*:string*/)/*:string*/ {
  return sign + (aval === 0 ? "" : ""+aval) + fill(" ", r[1].length + 2 + r[4].length);
}
var dec1 = /^#*0*\.(0+)/;
var closeparen = /\).*[0#]/;
var phone = /\(###\) ###\\?-####/;
function hashq(str/*:string*/)/*:string*/ {
  var o = "", cc;
  for(var i = 0; i != str.length; ++i) switch((cc=str.charCodeAt(i))) {
    case 35: break;
    case 63: o+= " "; break;
    case 48: o+= "0"; break;
    default: o+= String.fromCharCode(cc);
  }
  return o;
}
```

V8 has an annoying habit of deoptimizing round and floor

```
function rnd(val/*:number*/, d/*:number*/)/*:string*/ { var dd = Math.pow(10,d); return ""+(Math.round(val * dd)/dd); }
function dec(val/*:number*/, d/*:number*/)/*:number*/ {
	if (d < ('' + Math.round((val-Math.floor(val))*Math.pow(10,d))).length) {
		return 0;
	}
	return Math.round((val-Math.floor(val))*Math.pow(10,d));
}
function carry(val/*:number*/, d/*:number*/)/*:number*/ {
	if (d < ('' + Math.round((val-Math.floor(val))*Math.pow(10,d))).length) {
		return 1;
	}
	return 0;
}
function flr(val/*:number*/)/*:string*/ { if(val < 2147483647 && val > -2147483648) return ""+(val >= 0 ? (val|0) : (val-1|0)); return ""+Math.floor(val); }
```

### Main Number Writing Function

Finally the body:

```
function write_num_flt(type/*:string*/, fmt/*:string*/, val/*:number*/)/*:string*/ {
```

For parentheses, explicitly resolve the sign issue:

```js>bits/60_number.js
  if(type.charCodeAt(0) === 40 && !fmt.match(closeparen)) {
    var ffmt = fmt.replace(/\( */,"").replace(/ \)/,"").replace(/\)/,"");
    if(val >= 0) return write_num_flt('n', ffmt, val);
    return '(' + write_num_flt('n', ffmt, -val) + ')';
  }
```

Helpers are used for:
- Percentage values
- Trailing commas
- Exponentials

```js>bits/60_number.js
  if(fmt.charCodeAt(fmt.length - 1) === 44) return write_num_cm(type, fmt, val);
  if(fmt.indexOf('%') !== -1) return write_num_pct(type, fmt, val);
  if(fmt.indexOf('E') !== -1) return write_num_exp(fmt, val);
```

TODO: localize the currency:

```
  if(fmt.charCodeAt(0) === 36) return "$"+write_num_flt(type,fmt.substr(fmt[1]==' '?2:1),val);
```

Some simple cases should be resolved first:

```
  var o;
  var r/*:?Array<string>*/, ri, ff, aval = Math.abs(val), sign = val < 0 ? "-" : "";
  if(fmt.match(/^00+$/)) return sign + pad0r(aval,fmt.length);
  if(fmt.match(/^[#?]+$/)) {
    o = pad0r(val,0); if(o === "0") o = "";
    return o.length > fmt.length ? o : hashq(fmt.substr(0,fmt.length-o.length)) + o;
  }
```

Fractions with known denominator are resolved by rounding:

```
  if((r = fmt.match(frac1))) return write_num_f1(r, aval, sign);
```

A few special general cases can be handled in a very dumb manner:

```
  if(fmt.match(/^#+0+$/)) return sign + pad0r(aval,fmt.length - fmt.indexOf("0"));
  if((r = fmt.match(dec1))) {
    // $FlowIgnore
    o = rnd(val, r[1].length).replace(/^([^\.]+)$/,"$1."+r[1]).replace(/\.$/,"."+r[1]).replace(/\.(\d*)$/,function($$, $1) { return "." + $1 + fill("0", r[1].length-$1.length); });
    return fmt.indexOf("0.") !== -1 ? o : o.replace(/^0\./,".");
  }
```

The next few simplifications ignore leading optional sigils (`#`):

```
  fmt = fmt.replace(/^#+([0.])/, "$1");
  if((r = fmt.match(/^(0*)\.(#*)$/))) {
    return sign + rnd(aval, r[2].length).replace(/\.(\d*[1-9])0*$/,".$1").replace(/^(-?\d*)$/,"$1.").replace(/^0\./,r[1].length?"0.":".");
  }
  if((r = fmt.match(/^#,##0(\.?)$/))) return sign + commaify(pad0r(aval,0));
  if((r = fmt.match(/^#,##0\.([#0]*0)$/))) {
    return val < 0 ? "-" + write_num_flt(type, fmt, -val) : commaify(""+(Math.floor(val) + carry(val, r[1].length))) + "." + pad0(dec(val, r[1].length),r[1].length);
  }
  if((r = fmt.match(/^#,#*,#0/))) return write_num_flt(type,fmt.replace(/^#,#*,/,""),val);
```

The `Zip Code + 4` format needs to treat an interstitial hyphen as a character:

```
  if((r = fmt.match(/^([0#]+)(\\?-([0#]+))+$/))) {
    o = _strrev(write_num_flt(type, fmt.replace(/[\\-]/g,""), val));
    ri = 0;
    return _strrev(_strrev(fmt.replace(/\\/g,"")).replace(/[0#]/g,function(x){return ri<o.length?o[ri++]:x==='0'?'0':"";}));
  }
```

There's a better way to generalize the phone number and other formats in terms
of first drawing the digits, but this selection allows for more nuance:

```
  if(fmt.match(phone)) {
    o = write_num_flt(type, "##########", val);
    return "(" + o.substr(0,3) + ") " + o.substr(3, 3) + "-" + o.substr(6);
  }
```

The frac helper function is used for fraction formats (defined below).

```
  var oa = "";
  if((r = fmt.match(/^([#0?]+)( ?)\/( ?)([#0?]+)/))) {
    ri = Math.min(/*::String(*/r[4]/*::)*/.length,7);
    ff = frac(aval, Math.pow(10,ri)-1, false);
    o = "" + sign;
    oa = write_num("n", /*::String(*/r[1]/*::)*/, ff[1]);
    if(oa[oa.length-1] == " ") oa = oa.substr(0,oa.length-1) + "0";
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
```

The general class `/^[#0?]+$/` treats the '0' as literal, '#' as noop, '?' as space:

```
  if((r = fmt.match(/^[#0?]+$/))) {
    o = pad0r(val, 0);
    if(fmt.length <= o.length) return o;
    return hashq(fmt.substr(0,fmt.length-o.length)) + o;
  }
  if((r = fmt.match(/^([#0?]+)\.([#0]+)$/))) {
    o = "" + val.toFixed(Math.min(r[2].length,10)).replace(/([^0])0+$/,"$1");
    ri = o.indexOf(".");
    var lres = fmt.indexOf(".") - ri, rres = fmt.length - o.length - lres;
    return hashq(fmt.substr(0,lres) + o + fmt.substr(fmt.length-rres));
  }
```

The default cases are hard-coded.  TODO: actually parse them

```js>bits/60_number.js
  if((r = fmt.match(/^00,000\.([#0]*0)$/))) {
    ri = dec(val, r[1].length);
```

Note that this is technically incorrect

```
    return val < 0 ? "-" + write_num_flt(type, fmt, -val) : commaify(flr(val)).replace(/^\d,\d{3}$/,"0$&").replace(/^\d*$/,function($$) { return "00," + ($$.length < 3 ? pad0(0,3-$$.length) : "") + $$; }) + "." + pad0(ri,r[1].length);
  }
  switch(fmt) {
    case "#,###": var x = commaify(pad0r(aval,0)); return x !== "0" ? sign + x : "";
```

For now, the default case is an error:

```js>bits/60_number.js
    default:
  }
  throw new Error("unsupported format |" + fmt + "|");
}

```

### Integer Optimizations

```
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
    var period = fmt.indexOf("."); if(period === -1) period=fmt.indexOf('E');
    var ee = Math.floor(Math.log(Math.abs(val))*Math.LOG10E)%period;
    if(ee < 0) ee += period;
    o = (val/Math.pow(10,ee)).toPrecision(idx+1+(period+ee)%period);
    if(!o.match(/[Ee]/)) {
      var fakee = Math.floor(Math.log(Math.abs(val))*Math.LOG10E);
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
function write_num_int(type/*:string*/, fmt/*:string*/, val/*:number*/)/*:string*/ {
  if(type.charCodeAt(0) === 40 && !fmt.match(closeparen)) {
    var ffmt = fmt.replace(/\( */,"").replace(/ \)/,"").replace(/\)/,"");
    if(val >= 0) return write_num_int('n', ffmt, val);
    return '(' + write_num_int('n', ffmt, -val) + ')';
  }
  if(fmt.charCodeAt(fmt.length - 1) === 44) return write_num_cm2(type, fmt, val);
  if(fmt.indexOf('%') !== -1) return write_num_pct2(type, fmt, val);
  if(fmt.indexOf('E') !== -1) return write_num_exp2(fmt, val);
  if(fmt.charCodeAt(0) === 36) return "$"+write_num_int(type,fmt.substr(fmt[1]==' '?2:1),val);
  var o;
  var r, ri, ff, aval = Math.abs(val), sign = val < 0 ? "-" : "";
  if(fmt.match(/^00+$/)) return sign + pad0(aval,fmt.length);
  if(fmt.match(/^[#?]+$/)) {
    o = (""+val); if(val === 0) o = "";
    return o.length > fmt.length ? o : hashq(fmt.substr(0,fmt.length-o.length)) + o;
  }
  if((r = fmt.match(frac1))) return write_num_f2(r, aval, sign);
  if(fmt.match(/^#+0+$/)) return sign + pad0(aval,fmt.length - fmt.indexOf("0"));
  if((r = fmt.match(dec1))) {
    // $FlowIgnore
    o = (""+val).replace(/^([^\.]+)$/,"$1."+r[1]).replace(/\.$/,"."+r[1]).replace(/\.(\d*)$/,function($$, $1) { return "." + $1 + fill("0", r[1].length-$1.length); });
    return fmt.indexOf("0.") !== -1 ? o : o.replace(/^0\./,".");
  }
  fmt = fmt.replace(/^#+([0.])/, "$1");
  if((r = fmt.match(/^(0*)\.(#*)$/))) {
    return sign + (""+aval).replace(/\.(\d*[1-9])0*$/,".$1").replace(/^(-?\d*)$/,"$1.").replace(/^0\./,r[1].length?"0.":".");
  }
  if((r = fmt.match(/^#,##0(\.?)$/))) return sign + commaify((""+aval));
  if((r = fmt.match(/^#,##0\.([#0]*0)$/))) {
    return val < 0 ? "-" + write_num_int(type, fmt, -val) : commaify((""+val)) + "." + fill('0',r[1].length);
  }
  if((r = fmt.match(/^#,#*,#0/))) return write_num_int(type,fmt.replace(/^#,#*,/,""),val);
  if((r = fmt.match(/^([0#]+)(\\?-([0#]+))+$/))) {
    o = _strrev(write_num_int(type, fmt.replace(/[\\-]/g,""), val));
    ri = 0;
    return _strrev(_strrev(fmt.replace(/\\/g,"")).replace(/[0#]/g,function(x){return ri<o.length?o[ri++]:x==='0'?'0':"";}));
  }
  if(fmt.match(phone)) {
    o = write_num_int(type, "##########", val);
    return "(" + o.substr(0,3) + ") " + o.substr(3, 3) + "-" + o.substr(6);
  }
  var oa = "";
  if((r = fmt.match(/^([#0?]+)( ?)\/( ?)([#0?]+)/))) {
    ri = Math.min(r[4].length,7);
    ff = frac(aval, Math.pow(10,ri)-1, false);
    o = "" + sign;
    oa = write_num("n", r[1], ff[1]);
    if(oa[oa.length-1] == " ") oa = oa.substr(0,oa.length-1) + "0";
    o += oa + r[2] + "/" + r[3];
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
    case "#,###": var x = commaify(""+aval); return x !== "0" ? sign + x : "";
    default:
  }
  throw new Error("unsupported format |" + fmt + "|");
}
```

The final function simply dispatches:

```
return function write_num(type/*:string*/, fmt/*:string*/, val/*:number*/)/*:string*/ {
  return (val|0) === val ? write_num_int(type, fmt, val) : write_num_flt(type, fmt, val);
};})();
```

## Evaluating Format Strings

```js>bits/82_eval.js
var abstime = /\[[HhMmSs]*\]/;
function eval_fmt(fmt/*:string*/, v/*:any*/, opts/*:any*/, flen/*:number*/) {
  var out = [], o = "", i = 0, c = "", lst='t', q, dt, j, cc;
  var hr='H';
  /* Tokenize */
  while(i < fmt.length) {
    switch((c = fmt.charAt(i))) {
```

LO Formats sometimes leak "GENERAL" or "General" to stand for general format:

```
      case 'G': /* General */
        if(!isgeneral(fmt, i)) throw new Error('unrecognized character ' + c + ' in ' +fmt);
        out[out.length] = {t:'G', v:'General'}; i+=7; break;
```

Text between double-quotes are treated literally, and individual characters are
literal if they are preceded by a slash.

The additional `i < fmt.length` guard was added due to potentially unterminated
strings generated by LO:

```
      case '"': /* Literal text */
        for(o="";(cc=fmt.charCodeAt(++i)) !== 34 && i < fmt.length;) o += String.fromCharCode(cc);
        out[out.length] = {t:'t', v:o}; ++i; break;
      case '\\': var w = fmt[++i], t = (w === "(" || w === ")") ? w : 't';
        out[out.length] = {t:t, v:w}; ++i; break;
```

The underscore character represents a literal space.  Apparently, it also marks
that the next character is junk.  Hence the read pointer is moved by 2:

```
      case '_': out[out.length] = {t:'t', v:" "}; i+=2; break;
```

The '@' symbol refers to the original text.  The ECMA spec is not complete, but
Excel does not allow for '@' and non-literal text to appear in the same format.
It seems as if they only support one mode.  (clearly this is a TODO for excel
mode but I'm not convinced that's the right approach)

```
      case '@': /* Text Placeholder */
        out[out.length] = {t:'T', v:v}; ++i; break;
```

`B1` and `B2` specify which calendar to use, while `b` is the buddhist year.  It
acts just like `y` except the year is shifted:

```
      case 'B': case 'b':
        if(fmt[i+1] === "1" || fmt[i+1] === "2") {
          if(dt==null) { dt=parse_date_code(v, opts, fmt[i+1] === "2"); if(dt==null) return ""; }
          out[out.length] = {t:'X', v:fmt.substr(i,2)}; lst = c; i+=2; break;
        }
        /* falls through */
```

The date codes `m,d,y,h,s` are standard.  There are some special formats like
`e / g` (era year) that have different behaviors in Japanese/Chinese locales.

```
      case 'M': case 'D': case 'Y': case 'H': case 'S': case 'E':
        c = c.toLowerCase();
        /* falls through */
      case 'm': case 'd': case 'y': case 'h': case 's': case 'e': case 'g':
```

Negative dates are immediately thrown out:

```
        if(v < 0) return "";
```

Merge strings like "mmmmm" or "hh" into one block:

```
        if(dt==null) { dt=parse_date_code(v, opts); if(dt==null) return ""; }
        o = c; while(++i<fmt.length && fmt[i].toLowerCase() === c) o+=c;
```

Only the forward corrections are made here.  The reverse corrections are made later:

```
        if(c === 'm' && lst.toLowerCase() === 'h') c = 'M'; /* m = minute */
        if(c === 'h') c = hr;
        out[out.length] = {t:c, v:o}; lst = c; break;
```

The (poorly documented) rule regarding `A/P` and `AM/PM` is that if they show up
in the format then _all_ instances of `h` are considered 12-hour and not 24-hour
format (even in cases like `hh AM/PM hh hh hh`).

However, the undocumented `H` and `HH` do appear to reset the `AM/PM` indicator.
It is not implemented at the moment because I am not 100% sure of the rules with
the HH/hh jazz.  TODO: investigate this further.

```
      case 'A':
        q={t:c, v:"A"};
        if(dt==null) dt=parse_date_code(v, opts);
        if(fmt.substr(i, 3) === "A/P") { if(dt!=null) q.v = dt.H >= 12 ? "P" : "A"; q.t = 'T'; hr='h';i+=3;}
        else if(fmt.substr(i,5) === "AM/PM") { if(dt!=null) q.v = dt.H >= 12 ? "PM" : "AM"; q.t = 'T'; i+=5; hr='h'; }
        else { q.t = "t"; ++i; }
        if(dt==null && q.t === 'T') return "";
        out[out.length] = q; lst = c; break;
```

Conditional and color blocks should be handled at one point (TODO).  The
pseudo-type `Z` is used to capture absolute time blocks:

```
      case '[':
        o = c;
        while(fmt[i++] !== ']' && i < fmt.length) o += fmt[i];
        if(o.substr(-1) !== ']') throw 'unterminated "[" block: |' + o + '|';
        if(o.match(abstime)) {
          if(dt==null) { dt=parse_date_code(v, opts); if(dt==null) return ""; }
          out[out.length] = {t:'Z', v:o.toLowerCase()};
        } else { o=""; }
        break;
```

Number blocks (following the general pattern `[0#?][0#?.,E+-%]*`) are grouped
together.  Literal hyphens are swallowed as well.  Since `.000` is a valid
term (for tenths/hundredths/thousandths of a second), it must be handled
separately:

```
      /* Numbers */
      case '.':
        if(dt != null) {
          o = c; while((c=fmt[++i]) === "0") o += c;
          out[out.length] = {t:'s', v:o}; break;
        }
        /* falls through */
      case '0': case '#':
        o = c; while("0#?.,E+-%".indexOf(c=fmt[++i]) > -1 || c=='\\' && fmt[i+1] == "-" && "0#".indexOf(fmt[i+2])>-1) o += c;
        out[out.length] = {t:'n', v:o}; break;

```

The fraction question mark characters present their own challenges.  For example, the
number 123.456 under format `|??| /  |???| |???| foo` is `|15432| /  |125| |   | foo`:

```
      case '?':
        o = c; while(fmt[++i] === c) o+=c;
        q={t:c, v:o}; out[out.length] = q; lst = c; break;
```

Due to how the CSV generation works, asterisk characters are discarded.  TODO:
communicate this somehow, possibly with an option

```
      case '*': ++i; if(fmt[i] == ' ' || fmt[i] == '*') ++i; break; // **
```


The open and close parens `()` also has special meaning (for negative numbers):

```
      case '(': case ')': out[out.length] = {t:(flen===1?'t':c), v:c}; ++i; break;
```

The nonzero digits show up in fraction denominators:

```
      case '1': case '2': case '3': case '4': case '5': case '6': case '7': case '8': case '9':
        o = c; while("0123456789".indexOf(fmt[++i]) > -1) o+=fmt[i];
        out[out.length] = {t:'D', v:o}; break;
```

The default magic characters are listed in subsubsections 18.8.30-31 of ECMA376:

```
      case ' ': out[out.length] = {t:c, v:c}; ++i; break;
      default:
        if(",$-+/():!^&'~{}<>=€acfijklopqrtuvwxz".indexOf(c) === -1) throw new Error('unrecognized character ' + c + ' in ' + fmt);
        out[out.length] = {t:'t', v:c}; ++i; break;
    }
  }
```

In order to identify cases like `MMSS`, where the fact that this is a minute
appears after the minute itself, scan backwards.  At the same time, we can
identify the smallest time unit (0 = no time, 1 = hour, 2 = minute, 3 = second)
and the required number of digits for the sub-seconds:

```
  var bt = 0, ss0 = 0, ssm;
  for(i=out.length-1, lst='t'; i >= 0; --i) {
    switch(out[i].t) {
      case 'h': case 'H': out[i].t = hr; lst='h'; if(bt < 1) bt = 1; break;
      case 's':
        if((ssm=out[i].v.match(/\.0+$/))) ss0=Math.max(ss0,ssm[0].length-1);
        if(bt < 3) bt = 3;
      /* falls through */
      case 'd': case 'y': case 'M': case 'e': lst=out[i].t; break;
      case 'm': if(lst === 's') { out[i].t = 'M'; if(bt < 2) bt = 2; } break;
      case 'X': if(out[i].v === "B2");
        break;
      case 'Z':
        if(bt < 1 && out[i].v.match(/[Hh]/)) bt = 1;
        if(bt < 2 && out[i].v.match(/[Mm]/)) bt = 2;
        if(bt < 3 && out[i].v.match(/[Ss]/)) bt = 3;
    }
  }
```

Having determined the smallest time unit, round appropriately:

```
  switch(bt) {
    case 0: break;
    case 1:
      /*::if(!dt) break;*/
      if(dt.u >= 0.5) { dt.u = 0; ++dt.S; }
      if(dt.S >=  60) { dt.S = 0; ++dt.M; }
      if(dt.M >=  60) { dt.M = 0; ++dt.H; }
      break;
    case 2:
      /*::if(!dt) break;*/
      if(dt.u >= 0.5) { dt.u = 0; ++dt.S; }
      if(dt.S >=  60) { dt.S = 0; ++dt.M; }
      break;
  }
```

Since number groups in a string should be treated as part of the same whole,
group them together to construct the real number string:

```
  /* replace fields */
  var nstr = "", jj;
  for(i=0; i < out.length; ++i) {
    switch(out[i].t) {
      case 't': case 'T': case ' ': case 'D': break;
      case 'X': out[i].v = ""; out[i].t = ";"; break;
      case 'd': case 'm': case 'y': case 'h': case 'H': case 'M': case 's': case 'e': case 'b': case 'Z':
        /*::if(!dt) throw "unreachable"; */
        out[i].v = write_date(out[i].t.charCodeAt(0), out[i].v, dt, ss0);
        out[i].t = 't'; break;
      case 'n': case '(': case '?':
        jj = i+1;
        while(out[jj] != null && (
          (c=out[jj].t) === "?" || c === "D" ||
          (c === " " || c === "t") && out[jj+1] != null && (out[jj+1].t === '?' || out[jj+1].t === "t" && out[jj+1].v === '/') ||
          out[i].t === '(' && (c === ' ' || c === 'n' || c === ')') ||
          c === 't' && (out[jj].v === '/' || '$€'.indexOf(out[jj].v) > -1 || out[jj].v === ' ' && out[jj+1] != null && out[jj+1].t == '?')
        )) {
          out[i].v += out[jj].v;
          out[jj] = {v:"", t:";"}; ++jj;
        }
        nstr += out[i].v;
        i = jj-1; break;
      case 'G': out[i].t = 't'; out[i].v = general_fmt(v,opts); break;
    }
  }
```

Next, process the complete number string:

```
  var vv = "", myv, ostr;
  if(nstr.length > 0) {
    myv = (v<0&&nstr.charCodeAt(0) === 45 ? -v : v); /* '-' */
    ostr = write_num(nstr.charCodeAt(0) === 40 ? '(' : 'n', nstr, myv); /* '(' */
    jj=ostr.length-1;
```

Find the first decimal point:

```
    var decpt = out.length;
    for(i=0; i < out.length; ++i) if(out[i] != null && out[i].v.indexOf(".") > -1) { decpt = i; break; }
    var lasti=out.length;
```

If there is no decimal point or exponential, the algorithm is straightforward:

```
    if(decpt === out.length && ostr.indexOf("E") === -1) {
      for(i=out.length-1; i>= 0;--i) {
        if(out[i] == null || 'n?('.indexOf(out[i].t) === -1) continue;
        if(jj>=out[i].v.length-1) { jj -= out[i].v.length; out[i].v = ostr.substr(jj+1, out[i].v.length); }
        else if(jj < 0) out[i].v = "";
        else { out[i].v = ostr.substr(0, jj+1); jj = -1; }
        out[i].t = 't';
        lasti = i;
      }
      if(jj>=0 && lasti<out.length) out[lasti].v = ostr.substr(0,jj+1) + out[lasti].v;
    }
```
Otherwise we have to do something a bit trickier:

```
    else if(decpt !== out.length && ostr.indexOf("E") === -1) {
      jj = ostr.indexOf(".")-1;
      for(i=decpt; i>= 0; --i) {
        if(out[i] == null || 'n?('.indexOf(out[i].t) === -1) continue;
        j=out[i].v.indexOf(".")>-1&&i===decpt?out[i].v.indexOf(".")-1:out[i].v.length-1;
        vv = out[i].v.substr(j+1);
        for(; j>=0; --j) {
          if(jj>=0 && (out[i].v[j] === "0" || out[i].v[j] === "#")) vv = ostr[jj--] + vv;
        }
        out[i].v = vv;
        out[i].t = 't';
        lasti = i;
      }
      if(jj>=0 && lasti<out.length) out[lasti].v = ostr.substr(0,jj+1) + out[lasti].v;
      jj = ostr.indexOf(".")+1;
      for(i=decpt; i<out.length; ++i) {
        if(out[i] == null || 'n?('.indexOf(out[i].t) === -1 && i !== decpt ) continue;
        j=out[i].v.indexOf(".")>-1&&i===decpt?out[i].v.indexOf(".")+1:0;
        vv = out[i].v.substr(0,j);
        for(; j<out[i].v.length; ++j) {
          if(jj<ostr.length) vv += ostr[jj++];
        }
        out[i].v = vv;
        out[i].t = 't';
        lasti = i;
      }
    }
  }
```

The magic in the next line is to ensure that the negative number is passed as
positive when there is an explicit hyphen before it (e.g. `#,##0.0;-#,##0.0`):

```
  for(i=0; i<out.length; ++i) if(out[i] != null && 'n(?'.indexOf(out[i].t)>-1) {
    myv = (flen >1 && v < 0 && i>0 && out[i-1].v === "-" ? -v:v);
    out[i].v = write_num(out[i].t, out[i].v, myv);
    out[i].t = 't';
  }
```

Now we just need to combine the elements

```
  var retval = "";
  for(i=0; i !== out.length; ++i) if(out[i] != null) retval += out[i].v;
  return retval;
}
SSF._eval = eval_fmt;
```


There is some overloading of the `m` character.  According to the spec:

> If "m" or "mm" code is used immediately after the "h" or "hh" code (for
hours) or immediately before the "ss" code (for seconds), the application shall
display minutes instead of the month.

```js>bits/50_date.js
/*jshint -W086 */
function write_date(type/*:number*/, fmt/*:string*/, val, ss0/*:?number*/)/*:string*/ {
  var o="", ss=0, tt=0, y = val.y, out, outl = 0;
  switch(type) {
```

`b` years are shifted by 543 (`y` 1900 == `b` 2443):

```
    case 98: /* 'b' buddhist year */
      y = val.y + 543;
      /* falls through */
```

`yyyyyyyyyyyyyyyyyyyy` is a 4 digit year

```
    case 121: /* 'y' year */
    switch(fmt.length) {
      case 1: case 2: out = y % 100; outl = 2; break;
      default: out = y % 10000; outl = 4; break;
    } break;
```

`mmmmmmmmmmmmmmmmmmmm` is treated as the full month name:

```
    case 109: /* 'm' month */
    switch(fmt.length) {
      case 1: case 2: out = val.m; outl = fmt.length; break;
      case 3: return months[val.m-1][1];
      case 5: return months[val.m-1][0];
      default: return months[val.m-1][2];
    } break;
```

`dddddddddddddddddddd` is treated as the full day name:

```
    case 100: /* 'd' day */
    switch(fmt.length) {
      case 1: case 2: out = val.d; outl = fmt.length; break;
      case 3: return days[val.q][0];
      default: return days[val.q][1];
    } break;
```

Abnormal hours and minutes are rejected:

```
    case 104: /* 'h' 12-hour */
    switch(fmt.length) {
      case 1: case 2: out = 1+(val.H+11)%12; outl = fmt.length; break;
      default: throw 'bad hour format: ' + fmt;
    } break;
    case 72: /* 'H' 24-hour */
    switch(fmt.length) {
      case 1: case 2: out = val.H; outl = fmt.length; break;
      default: throw 'bad hour format: ' + fmt;
    } break;
    case 77: /* 'M' minutes */
    switch(fmt.length) {
      case 1: case 2: out = val.M; outl = fmt.length; break;
      default: throw 'bad minute format: ' + fmt;
    } break;
```

Unfortunately, the actual subsecond string is based on the presence of other
terms.  That is passed via the `ss0` parameter:

```
    case 115: /* 's' seconds */
    if(val.u === 0) switch(fmt) {
      case 's': case 'ss': return pad0(val.S, fmt.length);
      case '.0': case '.00': case '.000':
    }
    switch(fmt) {
      case 's': case 'ss': case '.0': case '.00': case '.000':
        /*::if(!ss0) ss0 = 0; */
        if(ss0 >= 2) tt = ss0 === 3 ? 1000 : 100;
        else tt = ss0 === 1 ? 10 : 1;
        ss = Math.round((tt)*(val.S + val.u));
        if(ss >= 60*tt) ss = 0;
        if(fmt === 's') return ss === 0 ? "0" : ""+ss/tt;
        o = pad0(ss,2 + ss0);
        if(fmt === 'ss') return o.substr(0,2);
        return "." + o.substr(2,fmt.length-1);
      default: throw 'bad second format: ' + fmt;
    }
```

The `Z` type refers to absolute time measures:

```
    case 90: /* 'Z' absolute time */
    switch(fmt) {
      case '[h]': case '[hh]': out = val.D*24+val.H; break;
      case '[m]': case '[mm]': out = (val.D*24+val.H)*60+val.M; break;
      case '[s]': case '[ss]': out = ((val.D*24+val.H)*60+val.M)*60+Math.round(val.S+val.u); break;
      default: throw 'bad abstime format: ' + fmt;
    } outl = fmt.length === 3 ? 1 : 2; break;
```

The `e` format behavior in excel diverges from the spec.  It claims that `ee`
should be a two-digit year, but `ee` in excel is actually the four-digit year:

```
    case 101: /* 'e' era */
      out = y; outl = 1;
```

There is no input to the function that ends up triggering the default behavior:
it is not exported and is only called when the type is in `ymdhHMsZe`

```
  }
  if(outl > 0) return pad0(out, outl); else return "";
}
/*jshint +W086 */
```

Based on the value, `choose_fmt` picks the right format string.  If formats have
explicit negative specifications, those values should be passed as positive:

```js>bits/90_main.js
function choose_fmt(f/*:string*/, v) {
  var fmt = split_fmt(f);
  var l = fmt.length, lat = fmt[l-1].indexOf("@");
  if(l<4 && lat>-1) --l;
  if(fmt.length > 4) throw new Error("cannot find right format for |" + fmt.join("|") + "|");
```

Short-circuit the string case by using the last format if it has "@":

```
  if(typeof v !== "number") return [4, fmt.length === 4 || lat>-1?fmt[fmt.length-1]:"@"];
  switch(fmt.length) {
```

In the case of one format, if it contains an "@" then it is a text format.
There is a big TODO here regarding how to best handle this case.

```
    case 1: fmt = lat>-1 ? ["General", "General", "General", fmt[0]] : [fmt[0], fmt[0], fmt[0], "@"]; break;
```

In the case of 2 or 3 formats, if an `@` appears in the last field of the format
it is treated as the text format

```
    case 2: fmt = lat>-1 ? [fmt[0], fmt[0], fmt[0], fmt[1]] : [fmt[0], fmt[1], fmt[0], "@"]; break;
    case 3: fmt = lat>-1 ? [fmt[0], fmt[1], fmt[0], fmt[2]] : [fmt[0], fmt[1], fmt[2], "@"]; break;
    case 4: break;
  }
```

Here we have to scan for conditions.  Note that the grammar precludes decimals
but in practice they are fair game:

```js>bits/88_cond.js
var cfregex = /\[[=<>]/;
var cfregex2 = /\[([=<>]*)(-?\d+\.?\d*)\]/;
function chkcond(v, rr) {
  if(rr == null) return false;
  var thresh = parseFloat(rr[2]);
  switch(rr[1]) {
    case "=":  if(v == thresh) return true; break;
    case ">":  if(v >  thresh) return true; break;
    case "<":  if(v <  thresh) return true; break;
    case "<>": if(v != thresh) return true; break;
    case ">=": if(v >= thresh) return true; break;
    case "<=": if(v <= thresh) return true; break;
  }
  return false;
}
```

The main function checks for conditional operators and acts accordingly:

```js>bits/90_main.js
  var ff = v > 0 ? fmt[0] : v < 0 ? fmt[1] : fmt[2];
  if(fmt[0].indexOf("[") === -1 && fmt[1].indexOf("[") === -1) return [l, ff];
  if(fmt[0].match(cfregex) != null || fmt[1].match(cfregex) != null) {
    var m1 = fmt[0].match(cfregex2);
    var m2 = fmt[1].match(cfregex2);
    return chkcond(v, m1) ? [l, fmt[0]] : chkcond(v, m2) ? [l, fmt[1]] : [l, fmt[m1 != null && m2 != null ? 2 : 1]];
  }
  return [l, ff];
}
```

Finally, the format wrapper brings everything together:

```
function format(fmt/*:string|number*/,v/*:any*/,o/*:?any*/) {
  fixopts(o != null ? o : (o=[]));
```

The string format is saved to a different variable:

```
  var sfmt = "";
  switch(typeof fmt) {
    case "string": sfmt = fmt; break;
    case "number": sfmt = (o.table != null ? (o.table/*:any*/) : table_fmt)[fmt]; break;
  }
```

LibreOffice appears to emit the format "GENERAL" for general:

```
  if(isgeneral(sfmt,0)) return general_fmt(v, o);
  var f = choose_fmt(sfmt, v);
  if(isgeneral(f[1])) return general_fmt(v, o);
```

The boolean TRUE and FALSE are formatted as if they are the uppercase text:

```
  if(v === true) v = "TRUE"; else if(v === false) v = "FALSE";
```

Empty string should always emit empty, even if there are other characters:

```
  else if(v === "" || v == null) return "";
  return eval_fmt(f[1], v, o, f[0]);
}
```

The methods beginning with an underscore are subject to change and should not be
used directly in programs.

```js>bits/98_exports.js
SSF._table = table_fmt;
SSF.load = function load_entry(fmt/*:string*/, idx/*:number*/) { table_fmt[idx] = fmt; };
SSF.format = format;
```

To support multiple SSF tables:

```
SSF.get_table = function get_table() { return table_fmt; };
SSF.load_table = function load_table(tbl/*:{[n:number]:string}*/) { for(var i=0; i!=0x0188; ++i) if(tbl[i] !== undefined) SSF.load(tbl[i], i); };
```

## Fraction Library

The implementation is from [our frac library](https://github.com/SheetJS/frac/):

```js>bits/30_frac.js
function frac(x, D, mixed) {
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
}
```

## JS Boilerplate

```js>bits/00_header.js
/* ssf.js (C) 2013-present SheetJS -- http://sheetjs.com */
/*jshint -W041 */
var SSF = {};
var make_ssf = function make_ssf(SSF){
```

```js>bits/99_footer.js
};
make_ssf(SSF);
/*global module */
/*:: declare var DO_NOT_EXPORT_SSF: any; */
if(typeof module !== 'undefined' && typeof DO_NOT_EXPORT_SSF === 'undefined') module.exports = SSF;
```

## .vocrc and post-commands

```bash>tmp/post.sh
#!/bin/bash
npm install
echo "SSF.version = '"`grep version package.json | awk '{gsub(/[^0-9\.]/,"",$2); print $2}'`"';" > tmp/01_version.js
cat tmp/*.js > ssf.js
```

```json>.vocrc
{
  "post": "bash tmp/post.sh"
}
```

