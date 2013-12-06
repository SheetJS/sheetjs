# SSF

SpreadSheet Format (SSF) is a pure-JS library to format data using ECMA-376
spreadsheet format codes.

## Options

The various API functions take an `opts` argument which control parsing.  The
default options are described below:

```js>tmp/opts.js
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

```js>tmp/main.js
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

```js>tmp/general.js
var general_fmt = function(v) {
```

Booleans are serialized in upper case:

```
  if(typeof v === 'boolean') return v ? "TRUE" : "FALSE";
```



```
};
SSF._general = general_fmt;
```

## Implied Number Formats

These are the commonly-used formats that have a special implied code.
None of the international formats are included here.

```js>tmp/consts.js
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
```

Now Excel and other formats treat code 14 as `mm/dd/yy` (with slashes).  Given
that the spec gives no internationalization considerations, erring on the side
of the applications makes sense here:

```
  14: 'mm/dd/yy',
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
```

These test cases were manually generated in Excel 2011 [value, code, result]:

```json>test/implied.json
[
  [12345.6789, 0, "12345.6789"],
  [12345.6789, 1, "12346"],
  [12345.6789, 2, "12345.68"],
  [12345.6789, 3, "12,346"],
  [12345.6789, 4, "12,345.68"],
  [12345.6789, 9, "1234568%"],
  [12345.6789, 10, "1234567.89%"],
  [12345.6789, 11, "1.23E+04"],
  [12345.6789, 12, "12345 2/3"],
  [12345.6789, 13, "12345 55/81"],
  [12345.6789, 14, "10/18/33"],
  [12345.6789, 15, "18-Oct-33"],
  [12345.6789, 16, "18-Oct"],
  [12345.6789, 17, "Oct-33"],
  [12345.6789, 18, "4:17 PM"],
  [12345.6789, 19, "4:17:37 PM"],
  [12345.6789, 20, "16:17"],
  [12345.6789, 21, "16:17:37"],
  [12345.6789, 22, "10/18/33 16:17"],
  [12345.6789, 37, "12,346"],
  [12345.6789, 38, "12,346"],
  [12345.6789, 39, "12,345.68"],
  [12345.6789, 40, "12,345.68"],
  [12345.6789, 45, "17:37"],
  [12345.6789, 46, "296296:17:37"],
  [12345.6789, 47, "1737.0"],
  [12345.6789, 48, "12.3E+3"],
  [12345.6789, 49, "12345.6789"]
]
```

## Dates and Time

The code `ddd` displays short day-of-week and `dddd` shows long day-of-week:

```js>tmp/consts.js
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


```js>tmp/date.js
var parse_date_code = function parse_date_code(v,opts) {
  var date = Math.floor(v), time = Math.round(86400 * (v - date)), dow=0;
  var dout=[], out={D:date, T:time}; fixopts(opts = (opts||{}));
```

Excel help actually recommends treating the 1904 date codes as 1900 date codes
shifted by 1462 days.

```
  if(opts.date1904) date += 1462;
```

Due to a bug in Lotus 1-2-3 which was propagated by Excel and other variants,
the year 1900 is recognized as a leap year.  JS has no way of representing that
abomination as a `Date`, so the easiest way is to store the data as a tuple.

February 29, 1900 (date `60`) is recognized as a Wednesday.

```
  if(date === 60) {dout = [1900,2,29]; dow=3;}
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
    if(opts.mode === 'excel' && date < 60) dow = (dow + 6) % 7;
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

## Evaluating Format Strings

```js>tmp/main.js
function eval_fmt(fmt, v, opts) {
  var out = [], o = "", i = 0, c = "", lst='t', q = {}, dt;
  fixopts(opts = (opts || {}));
  var hr='H';
  /* Tokenize */
  while(i < fmt.length) {
    switch((c = fmt[i])) {
```

Text between double-quotes are treated literally, and individual characters are
literal if they are preceded by a slash:

```
      case '"': /* Literal text */
        for(o="";fmt[++i] !== '"';) o += fmt[(fmt[i] === '\\' ? ++i : i)];
        out.push({t:'t', v:o}); break;
      case '\\': out.push({t:'t', v:fmt[++i]}); ++i; break;
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
      case 'm': case 'd': case 'y': case 'h': case 's': case 'e':
        if(!dt) dt = parse_date_code(v, opts);
        o = fmt[i]; while(fmt[++i] === c) o+=c;
        if(c === 'm' && lst.toLowerCase() === 'h') c = 'M'; /* m = minute */
        if(c === 'h') c = hr;
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
        q={t:c,v:"A"};
        if(fmt.substr(i, 3) === "A/P") {q.v = dt.H >= 12 ? "P" : "A"; q.t = 'T'; hr='h';i+=3;}
        else if(fmt.substr(i,5) === "AM/PM") { q.v = dt.H >= 12 ? "PM" : "AM"; q.t = 'T'; i+=5; hr='h'; }
        else q.t = "t";
        out.push(q); lst = c; break;
      case '[': /* TODO: Fix this -- ignore all conditionals and formatting */
        while(fmt[i++] !== ']'); break;
      default:
        if("$-+/():!^&'~{}<>= ".indexOf(c) === -1)
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
      case 't': case 'T': break;
      case 'd': case 'm': case 'y': case 'h': case 'H': case 'M': case 's': case 'A': case 'e':
        out[i].v = write_date(out[i].t, out[i].v, dt);
        out[i].t = 't'; break;
      default: throw "unrecognized type " + out[i].t;
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


```js>tmp/date.js
var write_date = function(type, fmt, val) {
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
      default: throw 'bad second format: ' + fmt;
    } break;
```

The `e` format behavior in excel diverges from the spec.  It claims that `ee`
should be a two-digit year, but `ee` in excel is actually the four-digit year:

```
    /* TODO: handle the ECMA spec format ee -> yy */
    case 'e': { return val.y; } break;
    case 'A': return (val.h>=12 ? 'P' : 'A') + fmt.substr(1);
    default: throw 'bad format type ' + type + ' in ' + fmt;
  }
};
```



```js>tmp/main.js
function choose_fmt(fmt, v) {
  if(typeof fmt === "string") fmt = split_fmt(fmt);
  if(typeof v !== "number") return fmt[3];
  return v > 0 ? fmt[0] : v < 0 ? fmt[1] : fmt[2];
}

var format = function format(fmt,v,o) {
  fixopts(o = (o||{}));
  if(fmt === 0) return general_fmt(v, o);
  if(typeof fmt === 'number') fmt = table_fmt[fmt];
  var f = choose_fmt(fmt, v, o);
  return eval_fmt(f, v, o);
};

```





```js>tmp/main.js

SSF._choose = choose_fmt;
SSF._table = table_fmt;
SSF.load = function(fmt, idx) { table_fmt[idx] = fmt; };
SSF.format = format;
```

## JS Boilerplate

```js>tmp/00_header.js
var SSF;
(function(SSF){
String.prototype.reverse=function(){return this.split("").reverse().join("");};
var _strrev = function(x) { return String(x).reverse(); };
function fill(c,l) { return new Array(l+1).join(c); }
function pad(v,d){var t=String(v);return t.length>=d?t:(fill(0,d-t.length)+t);}
```

```js>tmp/zz_footer_n.js
})(typeof exports !== 'undefined' ? exports : SSF);
```

```js>tmp/zz_footer.js
})(SSF);
```

## .vocrc and post-commands

```bash>tmp/post.sh
#!/bin/bash
npm install
cat tmp/{00_header,opts,consts,general,date,main,zz_footer_n}.js > ssf_node.js
cat tmp/{00_header,opts,consts,general,date,main,zz_footer}.js > ssf.js
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

```json>package.json
{
  "name": "ssf",
  "version": "0.1.0",
  "author": "SheetJS",
  "description": "pure-JS library to format data using ECMA-376 spreadsheet Format Codes",
  "keywords": [ "format", "sprintf", "spreadsheet" ],
  "main": "ssf_node.js",
  "dependencies": {
    "voc":"",
    "colors":""
  },
  "devDependencies": {
    "mocha":""
  },
  "repository": { "type":"git", "url":"git://github.com/SheetJS/ssf.git" },
  "scripts": {
    "test": "mocha -R spec"
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
```

The mocha test driver tests the implied formats:

```js>test/implied.js
/* vim: set ts=2: */
var SSF = require('../');
var fs = require('fs'), assert = require('assert');
var data = JSON.parse(fs.readFileSync('./test/implied.json','utf8'));
describe('implied formats', function() {
  data.forEach(function(d) {
    it(d[1]+" for "+d[0], (d[1]<14||d[1]>22)?null:function(){
      assert.equal(SSF.format(d[1], d[0], {}), d[2]);
    });
  });
});
```

The old test driver was manual:

```js>tmp/test.njs
var SSF = require('../ssf_node');
var x = 'd\\-mmm\\-yy\\ yyyy\\ dd\\ \\;\\ yy\\ mm\\ dd';
var y = 'd\\-mmm\\-yy\\ yyyy\\ dd\\ ;\\ yy\\ mm\\ dd';
var z = 'd\\ dd\\ ddd\\ dddd\\ m\\ mm\\ mmm\\ mmmm\\ mmmmm\\ yy\\ yyyy';
console.error(SSF.parse_date_code(65.9));
console.error(SSF.format(x, 65.9));
console.error(SSF.format(y, 65.9));
console.error()
console.error(SSF.format(z, 55.9));
console.error(SSF.format(z, 55.9, {mode:"excel"}));
console.error(SSF.format(z, 55.9));
console.error()
console.error(SSF.format(z, 65.9));
console.error(SSF.format(z, 65.9, {mode:"excel"}));
console.error(SSF.format(z, 65.9));
console.error()
console.error(SSF.format(19, 65.9));
console.error(SSF.format(20, 65.9));
```

# LICENSE

```>LICENSE
Copyright 2013   SheetJS

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
