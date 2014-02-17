# SSF

SpreadSheet Format (SSF) is a pure-JS library to format data using ECMA-376
spreadsheet format codes.

## Options

The various API functions take an `opts` argument which control parsing.  The
default options are described below:

```js>tmp/10_opts.js
/* Options */
var opts_fmt = {};
function fixopts(o){for(var y in opts_fmt) if(o[y]===undefined) o[y]=opts_fmt[y];}
SSF.opts = opts_fmt;
```

There are two commonly-recognized date code formats:
 - 1900 mode (where date=0 is 1899-12-31)
 - 1904 mode (where date=0 is 1904-01-01)

The difference between the the 1900 and 1904 date modes is 1462 days.  Since
the 1904 date mode was only default in a few Mac variants of Excel (2011 uses
1900 mode), the default is 1900 mode.  Consistent with ECMA-376 the name is
`date1904`:

```
opts_fmt.date1904 = 0;
```

The default output is a text representation (no effort to capture colors).  To
control the output, set the `output` variable:

- `text`: no color (default)
- `html`: html output using
- `ansi`: ansi color codes (requires `colors` module)

```
opts_fmt.output = "";
```

There are a few places where the specification is ambiguous or where Excel does
not follow the spec.  They are noted in the document.

The `mode` option controls compatibility:

- `ssf`: options that the author believes makes the most sense (default)
- `ecma`: compatibility with ECMA-376
- `excel`: compatibility with MS-XLSX

```
opts_fmt.mode = "";
```

## Conditional Format Codes

The specification is a bit unclear here.  It initially claims in §18.3.1:

> Up to four sections of format codes can be specified. The format codes,
separated by semicolons, define the formats for positive numbers, negative
numbers, zero values, and text, in that order.

Semicolons can be escaped with the `\` character, so we need to split on those
semicolons that aren't prefaced by a slash or within a quoted string:

```js>tmp/90_main.js
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


## General Number Format

The 'general' format for spreadsheets (identified by format code 0) is highly
context-sensitive and the implementation tries to follow the format to the best
of its abilities given the knowledge.

```js>tmp/40_general.js
var general_fmt = function(v) {
```

Booleans are serialized in upper case:

```
  if(typeof v === 'boolean') return v ? "TRUE" : "FALSE";
```

For numbers, try to display up to 11 digits of the number (the original code
`return v.toString().substr(0,11);` was not satisfactory in the case of 11 2/3)

```
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
```

For strings, just return the text as-is:

```
  if(typeof v === 'string') return v;
```

Anything else is bad:

```
  throw new Error("unsupported value in General format: " + v);
};
SSF._general = general_fmt;
```

## Implied Number Formats

These are the commonly-used formats that have a special implied code.
None of the international formats are included here.

```js>tmp/20_consts.js
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
```

Now Excel and other formats treat code 14 as `m/d/yy` (with slashes).  Given
that the spec gives no internationalization considerations, erring on the side
of the applications makes sense here:

```
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
```

There are special implicit format codes identified in [ECMA-376] 18.8.30.
Assuming zh-tw is the default:

```
  56: '"上午/下午 "hh"時"mm"分"ss"秒 "',
```

some writers erroneously emit 65535 for general:

```
  65535: 'General'
};
```

## Dates and Time

The code `ddd` displays short day-of-week and `dddd` shows long day-of-week:

```js>tmp/20_consts.js
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


```js>tmp/50_date.js
var parse_date_code = function parse_date_code(v,opts) {
  var date = Math.floor(v), time = Math.floor(86400 * (v - date)+1e-6), dow=0;
  var dout=[], out={D:date, T:time, u:86400*(v-date)-time}; fixopts(opts = (opts||{}));
```

Excel help actually recommends treating the 1904 date codes as 1900 date codes
shifted by 1462 days.

```
  if(opts.date1904) date += 1462;
```

Date codes beyond 12/31/9999 are invalid:

```
  if(date > 2958465) return null;
```

Due to a bug in Lotus 1-2-3 which was propagated by Excel and other variants,
the year 1900 is recognized as a leap year.  JS has no way of representing that
abomination as a `Date`, so the easiest way is to store the data as a tuple.

February 29, 1900 (date `60`) is recognized as a Wednesday.  Date `0` is treated
as January 0, 1900 rather than December 31, 1899.

```
  if(date === 60) {dout = [1900,2,29]; dow=3;}
  else if(date === 0) {dout = [1900,1,0]; dow=6;}
```

For the other dates, using the JS date mechanism suffices.

```
  else {
    if(date > 60) --date;
    /* 1 = Jan 1 1900 */
    var d = new Date(1900,0,1);
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
    if(/* opts.mode === 'excel' && */ date < 60) dow = (dow + 6) % 7;
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
};
SSF.parse_date_code = parse_date_code;
```

## Evaluating Number Formats

```js>tmp/60_number.js
var commaify = function(s) { return _strrev(_strrev(s).replace(/.../g,"$&,")).replace(/^,/,""); };
var write_num = function(type, fmt, val) {
```

For parentheses, explicitly resolve the sign issue:

```js>tmp/60_number.js
  if(type === '(') {
    var ffmt = fmt.replace(/\( */,"").replace(/ \)/,"").replace(/\)/,"");
    if(val >= 0) return write_num('n', ffmt, val);
    return '(' + write_num('n', ffmt, -val) + ')';
  }
```

Percentage values should be physically shifted:

```js>tmp/60_number.js
  var mul = 0, o;
  fmt = fmt.replace(/%/g,function(x) { mul++; return ""; });
  if(mul !== 0) return write_num(type, fmt, val * Math.pow(10,2*mul)) + fill("%",mul);
```

For exponents, get the exponent and mantissa and format them separately:

```
  if(fmt.indexOf("E") > -1) {
    var idx = fmt.indexOf("E") - fmt.indexOf(".") - 1;
```

For the special case of engineering notation, "shift" the decimal:

```
    if(fmt.match(/^#+0.0E\+0$/)) {
      var period = fmt.indexOf("."); if(period === -1) period=fmt.indexOf('E');
      var ee = (Number(val.toExponential(0).substr(2+(val<0))))%period;
      if(ee < 0) ee += period;
      o = (val/Math.pow(10,ee)).toPrecision(idx+1+(period+ee)%period);
      if(!o.match(/[Ee]/)) {
```

TODO: something reasonable

```
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
```

TODO: localize the currency:

```
  if(fmt[0] === "$") return "$"+write_num(type,fmt.substr(fmt[1]==' '?2:1),val);
```

Fractions with known denominator are resolved by rounding:

```js>tmp/60_number.js
  var r, rr, ff, aval = val < 0 ? -val : val, sign = val < 0 ? "-" : "";
  if((r = fmt.match(/# (\?+)([ ]?)\/([ ]?)(\d+)/))) {
    var den = Number(r[4]), rnd = Math.round(aval * den), base = Math.floor(rnd/den);
    var myn = (rnd - base*den), myd = den;
    return sign + (base?base:"") + " " + (myn === 0 ? fill(" ", r[1].length + 1 + r[4].length) : pad(myn,r[1].length," ") + r[2] + "/" + r[3] + pad(myd,r[4].length));
  }
```

A few special general cases can be handled in a very dumb manner:

```
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
```

The frac helper function is used for fraction formats (defined below).

```
  if((r = fmt.match(/^# ([?]+)([ ]?)\/([ ]?)([?]+)/))) {
    rr = Math.min(Math.max(r[1].length, r[4].length),7);
    ff = frac(aval, Math.pow(10,rr)-1, true);
    return sign + (ff[0]||(ff[1] ? "" : "0")) + " " + (ff[1] ? pad(ff[1],rr," ") + r[2] + "/" + r[3] + rpad(ff[2],rr," "): fill(" ", 2*rr+1 + r[2].length + r[3].length));
  }
```

The default cases are hard-coded.  TODO: actually parse them

```js>tmp/60_number.js
  switch(fmt) {
    case "0": case "#0": return Math.round(val);
    case "#.##": o = Math.round(val*100);
      return String(o/100).replace(/^([^\.]+)$/,"$1.").replace(/^0\.$/,".");
    case "#,###": var x = commaify(String(Math.round(aval))); return x !== "0" ? sign + x : "";
```

For now, the default case is an error:

```js>tmp/60_number.js
    default:
  }
  throw new Error("unsupported format |" + fmt + "|");
};
```

## Evaluating Format Strings

```js>tmp/90_main.js
function eval_fmt(fmt, v, opts, flen) {
  var out = [], o = "", i = 0, c = "", lst='t', q = {}, dt;
  fixopts(opts = (opts || {}));
  var hr='H';
  /* Tokenize */
  while(i < fmt.length) {
    switch((c = fmt[i])) {
```

LO Formats sometimes leak "GENERAL" or "General" to stand for general format:

```
      case 'G': /* General */
        if(fmt.substr(i, i+6).toLowerCase() !== "general")
          throw new Error('unrecognized character ' + fmt[i] + ' in ' +fmt);
        out.push({t:'G',v:'General'}); i+=7; break;
```

Text between double-quotes are treated literally, and individual characters are
literal if they are preceded by a slash.

The additional `i < fmt.length` guard was added due to potentially unterminated
strings generated by LO:

```
      case '"': /* Literal text */
        for(o="";fmt[++i] !== '"' && i < fmt.length;) o += fmt[i];
        out.push({t:'t', v:o}); ++i; break;
      case '\\': var w = fmt[++i], t = "()".indexOf(w) === -1 ? 't' : w;
        out.push({t:t, v:w}); ++i; break;
```

The underscore character represents a literal space.  Apparently, it also marks
that the next character is junk.  Hence the read pointer is moved by 2:

```
      case '_': out.push({t:'t', v:" "}); i+=2; break;
```

The '@' symbol refers to the original text.  The ECMA spec is not complete, but
Excel does not allow for '@' and non-literal text to appear in the same format.
It seems as if they only support one mode.  (clearly this is a TODO for excel
mode but I'm not convinced that's the right approach)

```
      case '@': /* Text Placeholder */
        out.push({t:'T', v:v}); ++i; break;
```

The date codes `m,d,y,h,s` are standard.  There are some special formats like
`e` (era year) that have different behaviors in Japanese/Chinese locales.

```
      /* Dates */
      case 'M': case 'D': case 'Y': case 'H': case 'S': case 'E':
        c = c.toLowerCase();
        /* falls through */
      case 'm': case 'd': case 'y': case 'h': case 's': case 'e':
```

Negative dates are immediately thrown out:

```
        if(v < 0) return "";
```

Merge strings like "mmmmm" or "hh" into one block:

```
        if(!dt) dt = parse_date_code(v, opts);
        if(!dt) return "";
        o = fmt[i]; while((fmt[++i]||"").toLowerCase() === c) o+=c;
```

For the special case of s.00, the suffix should be swallowed with the s:

```
        if(c === 's' && fmt[i] === '.' && fmt[i+1] === '0') { o+='.'; while(fmt[++i] === '0') o+= '0'; }
```

Only the forward corrections are made here.  The reverse corrections are made later:

```
        if(c === 'm' && lst.toLowerCase() === 'h') c = 'M'; /* m = minute */
        if(c === 'h') c = hr;
        o = o.toLowerCase();
        q={t:c, v:o}; out.push(q); lst = c; break;
```

The (poorly documented) rule regarding `A/P` and `AM/PM` is that if they show up
in the format then _all_ instances of `h` are considered 12-hour and not 24-hour
format (even in cases like `hh AM/PM hh hh hh`).

However, the undocumented `H` and `HH` do appear to reset the `AM/PM` indicator.
It is not implemented at the moment because I am not 100% sure of the rules with
the HH/hh jazz.  TODO: investigate this further.

```
      case 'A':
        if(!dt) dt = parse_date_code(v, opts);
        if(!dt) return "";
        q={t:c,v:"A"};
        if(fmt.substr(i, 3) === "A/P") {q.v = dt.H >= 12 ? "P" : "A"; q.t = 'T'; hr='h';i+=3;}
        else if(fmt.substr(i,5) === "AM/PM") { q.v = dt.H >= 12 ? "PM" : "AM"; q.t = 'T'; i+=5; hr='h'; }
        else { q.t = "t"; i++; }
        out.push(q); lst = c; break;
```

Conditional and color blocks should be handled at one point (TODO).  The
pseudo-type `Z` is used to capture absolute time blocks:

```
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
```

Number blocks (following the general pattern `[0#?][0#?.,E+-%]*`) are grouped together:

```
      /* Numbers */
      case '0': case '#': case '.':
        o = c; while("0#?.,E+-%".indexOf(c=fmt[++i]) > -1) o += c;
        out.push({t:'n', v:o}); break;

```

The fraction question mark characters present their own challenges.  For example, the
number 123.456 under format `|??| /  |???| |???| foo` is `|15432| /  |125| |   | foo`:

```
      case '?':
        o = fmt[i]; while(fmt[++i] === c) o+=c;
        q={t:c, v:o}; out.push(q); lst = c; break;
```

Due to how the CSV generation works, asterisk characters are discarded.  TODO:
communicate this somehow, possibly with an option

```
      case '*': ++i; if(fmt[i] == ' ') ++i; break; // **
```


The open and close parens `()` also has special meaning (for negative numbers):

```
      case '(': case ')': out.push({t:(flen===1?'t':c),v:c}); ++i; break;
```

The nonzero digits show up in fraction denominators:

```
      case '1': case '2': case '3': case '4': case '5': case '6': case '7': case '8': case '9':
        o = fmt[i]; while("0123456789".indexOf(fmt[++i]) > -1) o+=fmt[i];
        out.push({t:'D', v:o}); break;
```

The default magic characters are listed in subsubsections 18.8.30-31 of ECMA376:


```
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
```

The magic in the next line is to ensure that the negative number is passed as
positive when there is an explicit hyphen before it (e.g. `#,##0.0;-#,##0.0`):

```
        out[i].v = write_num(out[i].t, out[i].v, (flen >1 && v < 0 && i>0 && out[i-1].v == "-" ? -v:v));
        out[i].t = 't';
        i = jj-1; break;
      case 'G': out[i].t = 't'; out[i].v = general_fmt(v,opts); break;
```

The default case should not be reachable.  In testing, add the line
`default: console.error(out); throw "unrecognized type " + out[i].t;`

```
    }
  }
  return out.map(function(x){return x.v;}).join("");
}
SSF._eval = eval_fmt;
```


There is some overloading of the `m` character.  According to the spec:

> If "m" or "mm" code is used immediately after the "h" or "hh" code (for
hours) or immediately before the "ss" code (for seconds), the application shall
display minutes instead of the month.

```js>tmp/50_date.js
/*jshint -W086 */
var write_date = function(type, fmt, val) {
  if(val < 0) return "";
  var o;
  switch(type) {
    case 'y': switch(fmt) { /* year */
      case 'y': case 'yy': return pad(val.y % 100,2);
```

Apparently, even `yyyyyyyyyyyyyyyyyyyy` is a 4 digit year

```
      default: return pad(val.y % 10000,4);
    }
    case 'm': switch(fmt) { /* month */
      case 'm': return val.m;
      case 'mm': return pad(val.m,2);
      case 'mmm': return months[val.m-1][1];
      case 'mmmmm': return months[val.m-1][0];
```

Strangely enough, `mmmmmmmmmmmmmmmmmmmm` is treated as the full month name:

```
      default: return months[val.m-1][2];
    }
    case 'd': switch(fmt) { /* day */
      case 'd': return val.d;
      case 'dd': return pad(val.d,2);
      case 'ddd': return days[val.q][0];
```

Strangely enough, `dddddddddddddddddddd` is treated as the full day name:

```
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
```

The `Z` type refers to absolute time measures:

```
    case 'Z': switch(fmt) {
      case '[h]': case '[hh]': o = val.D*24+val.H; break;
      case '[m]': case '[mm]': o = (val.D*24+val.H)*60+val.M; break;
      case '[s]': case '[ss]': o = ((val.D*24+val.H)*60+val.M)*60+Math.round(val.S+val.u); break;
      default: throw 'bad abstime format: ' + fmt;
    } return fmt.length === 3 ? o : pad(o, 2);
```

The `e` format behavior in excel diverges from the spec.  It claims that `ee`
should be a two-digit year, but `ee` in excel is actually the four-digit year:

```
    /* TODO: handle the ECMA spec format ee -> yy */
    case 'e': { return val.y; } break;
```

There is no input to the function that ends up triggering the default behavior:
it is not exported and is only called when the type is in `ymdhHMsZe`

```
  }
};
/*jshint +W086 */
```

Based on the value, `choose_fmt` picks the right format string.  If formats have
explicit negative specifications, those values should be passed as positive:

```js>tmp/90_main.js
function choose_fmt(fmt, v, o) {
  if(typeof fmt === 'number') fmt = ((o&&o.table) ? o.table : table_fmt)[fmt];
  if(typeof fmt === "string") fmt = split_fmt(fmt);
  var l = fmt.length;
  if(l<4 && fmt[l-1].indexOf("@")>-1) --l;
  switch(fmt.length) {
```

In the case of one format, if it contains an "@" then it is a text format.
There is a big TODO here regarding how to best handle this case.

```
    case 1: fmt = fmt[0].indexOf("@")>-1 ? ["General", "General", "General", fmt[0]] : [fmt[0], fmt[0], fmt[0], "@"]; break;
```

In the case of 2 or 3 formats, if an `@` appears in the last field of the format
it is treated as the text format

```
    case 2: fmt = fmt[1].indexOf("@")>-1 ? [fmt[0], fmt[0], fmt[0], fmt[1]] : [fmt[0], fmt[1], fmt[0], "@"]; break;
    case 3: fmt = fmt[2].indexOf("@")>-1 ? [fmt[0], fmt[1], fmt[0], fmt[2]] : [fmt[0], fmt[1], fmt[2], "@"]; break;
    case 4: break;
    default: throw "cannot find right format for |" + fmt + "|";
  }
  if(typeof v !== "number") return [fmt.length, fmt[3]];
  return [l, v > 0 ? fmt[0] : v < 0 ? fmt[1] : fmt[2]];
}
```

Finally, the format wrapper brings everything together:

```
var format = function format(fmt,v,o) {
  fixopts(o = (o||{}));
```

LibreOffice appears to emit the format "GENERAL" for general:

```
  if(typeof fmt === "string" && fmt.toLowerCase() === "general") return general_fmt(v, o);
  if(typeof fmt === 'number') fmt = (o.table || table_fmt)[fmt];
  var f = choose_fmt(fmt, v, o);
  if(f[1].toLowerCase() === "general") return general_fmt(v,o);
```

The boolean TRUE and FALSE are formatted as if they are the uppercase text:

```
  if(v === true) v = "TRUE"; if(v === false) v = "FALSE";
```

Empty string should always emit empty, even if there are other characters:

```
  if(v === "" || typeof v === "undefined") return "";
  return eval_fmt(f[1], v, o, f[0]);
};

```

The methods beginning with an underscore are subject to change and should not be
used directly in programs.

```js>tmp/90_main.js

SSF._choose = choose_fmt;
SSF._table = table_fmt;
SSF.load = function(fmt, idx) { table_fmt[idx] = fmt; };
SSF.format = format;
```

To support multiple SSF tables:

```
SSF.get_table = function() { return table_fmt; };
SSF.load_table = function(tbl) { for(var i=0; i!=0x0188; ++i) if(tbl[i]) SSF.load(tbl[i], i); };
```

## Fraction Library

The implementation is from [our frac library](https://github.com/SheetJS/frac/):

```js>tmp/30_frac.js
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
```

## JS Boilerplate

```js>tmp/00_header.js
/* ssf.js (C) 2013-2014 SheetJS -- http://sheetjs.com */
var SSF = {};
var make_ssf = function(SSF){
var _strrev = function(x) { return String(x).split("").reverse().join("");};
function fill(c,l) { return new Array(l+1).join(c); }
function pad(v,d,c){var t=String(v);return t.length>=d?t:(fill(c||0,d-t.length)+t);}
function rpad(v,d,c){var t=String(v);return t.length>=d?t:(t+fill(c||0,d-t.length));}
```

```js>tmp/99_footer.js
};
make_ssf(SSF);
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

```>.gitignore
.gitignore
tmp/
node_modules/
.vocrc
```

```>.npmignore
test/*.tsv
node_modules/
tmp/
.gitignore
.vocrc
```

```make>Makefile
.PHONY: test ssf
ssf: ssf.md
        voc ssf.md

test:
        npm test

test_min:
        MINTEST=1 npm test

.PHONY: lint
lint:
        jshint ssf.js test/
```

Coverage tests use [blanket](http://npm.im/blanket):

```

.PHONY: cov
cov: tmp/coverage.html

tmp/coverage.html: ssf
        mocha --require blanket -R html-cov > tmp/coverage.html

.PHONY: cov_min
cov_min:
        MINTEST=1 make cov
```

Coveralls.io support

```

.PHONY: coveralls
coveralls:
        mocha --require blanket --reporter mocha-lcov-reporter | ./node_modules/coveralls/bin/coveralls.js
```

```json>package.json
{
  "name": "ssf",
  "version": "0.5.8",
  "author": "SheetJS",
  "description": "pure-JS library to format data using ECMA-376 spreadsheet Format Codes",
  "keywords": [ "format", "sprintf", "spreadsheet" ],
  "main": "ssf.js",
  "dependencies": {
    "voc":"",
    "colors":"",
    "frac":"0.3.1"
  },
  "devDependencies": {
    "mocha":""
  },
  "repository": { "type":"git", "url":"git://github.com/SheetJS/ssf.git" },
  "scripts": {
    "test": "mocha -R spec"
  },
  "bin": {
    "ssf": "./bin/ssf.njs"
  },
  "config": {
    "blanket": {
      "pattern": "ssf.js"
    }
  },
  "bugs": { "url": "https://github.com/SheetJS/ssf/issues" },
  "license": "Apache-2.0",
  "engines": { "node": ">=0.8" }
}
```

# Test Driver

Travis CI is used for node testing:

```>.travis.yml
language: node_js
node_js:
  - "0.10"
  - "0.8"
before_install:
  - "npm install -g mocha"
  - "npm install blanket"
  - "npm install coveralls mocha-lcov-reporter"
after_success:
  - "make coveralls"
```

The mocha test driver tests the implied formats:

```js>test/implied.js
var SSF = require('../');
var fs = require('fs'), assert = require('assert');
var data = JSON.parse(fs.readFileSync('./test/implied.json','utf8'));
var skip = [];
function doit(d) {
  d[1].forEach(function(r){if(!r[2])assert.equal(SSF.format(r[0],d[0]),r[1]);});
}
describe('implied formats', function() {
  data.forEach(function(d) {
    if(d.length == 2) it(d[0], function() { doit(d); });
    else it(d[1]+" for "+d[0], skip.indexOf(d[1]) > -1 ? null : function(){
      assert.equal(SSF.format(d[1], d[0], {}), d[2]);
    });
  });
});
```

The general test driver tests the General format:

```js>test/general.js
/* vim: set ts=2: */
var SSF = require('../');
var fs = require('fs'), assert = require('assert');
var data = JSON.parse(fs.readFileSync('./test/general.json','utf8'));
var skip = [];
describe('General format', function() {
  data.forEach(function(d) {
    it(d[1]+" for "+d[0], skip.indexOf(d[1]) > -1 ? null : function(){
      assert.equal(SSF.format(d[1], d[0], {}), d[2]);
    });
  });
  it('should fail for undefined and null', function() {
    assert.throws(function() {
      SSF.format("General", undefined);
      SSF.format("General", null);
    });
  });
});
```

The fraction test driver tests fractional formats:

```js>test/fraction.js
/* vim: set ts=2: */
var SSF = require('../');
var fs = require('fs'), assert = require('assert');
var data = JSON.parse(fs.readFileSync('./test/fraction.json','utf8'));
var skip = [];
describe('fractional formats', function() {
  data.forEach(function(d) {
    it(d[1]+" for "+d[0], skip.indexOf(d[1]) > -1 ? null : function(){
      var expected = d[2], actual = SSF.format(d[1], d[0], {});
      //var r = actual.match(/(-?)\d* *\d+\/\d+/);
      assert.equal(actual, expected);
    });
  });
});
```

The dates test driver tests the date and time formats:

```js>test/date.js
/* vim: set ts=2: */
/*jshint loopfunc:true */
var SSF = require('../');
var fs = require('fs'), assert = require('assert');
var dates = fs.readFileSync('./test/dates.tsv','utf8').split("\n");
var times = fs.readFileSync('./test/times.tsv','utf8').split("\n");

function doit(data) {
  var step = Math.ceil(data.length/100), i = 1;
  var headers = data[0].split("\t");
  for(j=0;j<=100;++j) it(j, function() {
    for(var k = 0; k <= step; ++k,++i) {
      if(!data[i]) return;
      var d = data[i].replace(/#{255}/g,"").split("\t");
      for(var w = 1; w < headers.length; ++w) {
        var expected = d[w], actual = SSF.format(headers[w], Number(d[0]), {});
        if(actual != expected) throw [actual, expected, w, headers[w],d[0],d].join("|");
        actual = SSF.format(headers[w].toUpperCase(), Number(d[0]), {});
        if(actual != expected) throw [actual, expected, w, headers[w],d[0],d].join("|");
      }
    }
  });
}
describe('time formats', function() { doit(times.slice(0,1000)); });
describe('date formats', function() {
  doit(process.env.MINTEST ? dates.slice(0,1000) : dates);
  it('should fail for bad formats', function() {
    var bad = [];
    var chk = function(fmt){ return function(){ SSF.format(fmt,0); }; };
    bad.forEach(function(fmt){assert.throws(chk(fmt));});
  });
});
```

The exponential test driver tests exponential formats (pipe denotes fails)

```js>test/exp.js
/* vim: set ts=2: */
/*jshint loopfunc:true */
var SSF = require('../');
var fs = require('fs'), assert = require('assert');
var data = fs.readFileSync('./test/exp.tsv','utf8').split("\n");
function doit(d, headers) {
  it(d[0], function() {
    for(var w = 1; w < headers.length; ++w) {
      var expected = d[w].replace("|", ""), actual;
      try { actual = SSF.format(headers[w], Number(d[0]), {}); } catch(e) { }
      if(actual != expected && d[w][0] !== "|") throw [actual, expected, w, headers[w],d[0],d].join("|");
    }
  });
}
describe('exponential formats', function() {
  var headers = data[0].split("\t");
  for(var j=1;j<data.length;++j) {
    if(!data[j]) return;
    doit(data[j].replace(/#{255}/g,"").split("\t"), headers);
  }
});
```

The oddities test driver tests random odd formats

```js>test/oddities.js
/* vim: set ts=2: */
/*jshint loopfunc:true */
var SSF = require('../');
var fs = require('fs'), assert = require('assert');
var data = JSON.parse(fs.readFileSync('./test/oddities.json','utf8'));
describe('oddities', function() {
  data.forEach(function(d) {
    it(d[0], function(){
      for(j=1;j<d.length;++j) {
        if(d[j].length == 2) {
          var expected = d[j][1], actual = SSF.format(d[0], d[j][0], {});
          assert.equal(actual, expected);
        } else assert.throws(function() { SSF.format(d[0], d[j][0]); });
      }
    });
  });
  it('should fail for bad formats', function() {
    var bad = ['##,##'];
    var chk = function(fmt){ return function(){ SSF.format(fmt,0); }; };
    bad.forEach(function(fmt){assert.throws(chk(fmt));});
  });
});
```

# LICENSE

```>LICENSE
Copyright (C) 2013-2014   SheetJS

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
```
