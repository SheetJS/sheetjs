# [SheetJS SSF](http://sheetjs.com)

ssf (SpreadSheet Format) is a pure-JS library to format data using ECMA-376
spreadsheet format codes (used in popular spreadsheet software packages).


## Installation

With [npm](https://www.npmjs.org/package/ssf):

```bash
$ npm install ssf
```

In the browser:

```html
<script src="ssf.js"></script>
```

The browser exposes a variable `SSF`

When installed globally, npm installs a script `ssf` that renders the format
string with the given arguments.  Running the script with `-h` displays help.

The script will manipulate `module.exports` if available (e.g. in a CommonJS
`require` context).  This is not always desirable.  To prevent the behavior,
define `DO_NOT_EXPORT_SSF`.

## Usage

`SSF.format(fmt, val, opts)` formats `val` using the format `fmt`.

If `fmt` is a string, it will be parsed and evaluated.  If `fmt` is a `number`,
the actual format will be the corresponding entry in the internal format table.
For a raw numeric format like `000`, the value should be passed as a string.

Date arguments are interpreted in the local time of the JS client.

The options argument may contain the following keys:

| Option Name | Default | Description                                          |
| :---------- | :-----: | :--------------------------------------------------- |
| date1904    | false   | Use 1904 date system if true, 1900 system if false   |

### Manipulating the Internal Format Table

Binary spreadsheet formats store cell formats in a table and reference by index.
This library uses a global table:

`SSF._table` is the underlying object, mapping numeric keys to format strings.

`SSF.load(fmt:string, idx:?number):number` assigns the format to the specified
index and returns the index.  If the index is not specified, SSF will search the
space for an available format slot pick an unused slot.  For compatibility with
the XLS and XLSB file formats, custom indices should be in the valid ranges
`5-8`, `23-26`, `41-44`, `63-66`, `164-382` (see `[MS-XLSB] 2.4.655 BrtFmt`)

`SSF.get_table()` gets the internal format table (number to format mapping).

`SSF.load_table(table)` sets the internal format table.

### Other Utilities

`SSF.parse_date_code(val:number, opts:?any)` parses `val`, returning an object:

```typescript
type SSFDate = {
  D:number; /* number of whole days since relevant epoch, 0 <= D */
  y:number; /* integral year portion, epoch_year <= y */
  m:number; /* integral month portion, 1 <= m <= 12 */
  d:number; /* integral day portion, subject to gregorian YMD constraints */
  q:number; /* integral day of week (0=Sunday .. 6=Saturday) 0 <= q <= 6 */

  T:number; /* number of seconds since midnight, 0 <= T < 86400 */
  H:number; /* integral number of hours since midnight, 0 <= H < 24 */
  M:number; /* integral number of minutes since the last hour, 0 <= M < 60 */
  S:number; /* integral number of seconds since the last minute, 0 <= S < 60 */
  u:number; /* sub-second part of time, 0 <= u < 1 */
}
```

`SSF.is_date(fmt:string):boolean` returns `true` if `fmt` encodes a date format.

## License

Please consult the attached LICENSE file for details.  All rights not explicitly
granted by the Apache 2.0 license are reserved by the Original Author.

## References

- [ECMA-376] Office Open XML File Formats
- [MS-XLSB] Excel (.xlsb) Binary File Format

## Badges

[![Sauce Test Status](https://saucelabs.com/browser-matrix/ssfjs.svg)](https://saucelabs.com/u/ssfjs)

[![Build Status](https://travis-ci.org/SheetJS/ssf.svg?branch=master)](https://travis-ci.org/SheetJS/ssf)

[![Coverage Status](http://img.shields.io/coveralls/SheetJS/ssf/master.svg)](https://coveralls.io/r/SheetJS/ssf?branch=master)

[![NPM Downloads](https://img.shields.io/npm/dt/ssf.svg)](https://npmjs.org/package/ssf)

[![Dependencies Status](https://david-dm.org/sheetjs/ssf/status.svg)](https://david-dm.org/sheetjs/ssf)

[![ghit.me](https://ghit.me/badge.svg?repo=sheetjs/js-xlsx)](https://ghit.me/repo/sheetjs/js-xlsx)

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/ssf?pixel)](https://github.com/SheetJS/ssf)
