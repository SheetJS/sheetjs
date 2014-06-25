# SSF

SpreadSheet Format (SSF) is a pure-JS library to format data using ECMA-376 
spreadsheet format codes (like those used in Microsoft Excel)

This is written in [voc](https://npmjs.org/package/voc) -- see ssf.md for code.

To build: `voc ssf.md`

## Setup

In the browser:

    <script src="ssf.js"></script>

In node:

    var SSF = require('ssf');

The script will manipulate `module.exports` if available (e.g. in a CommonJS 
`require` context).  This is not always desirable.  To prevent the behavior, 
define `DO_NOT_EXPORT_SSF`:

## Usage

`.load(fmt, idx)` sets custom formats (generally indices above `164`).

`.format(fmt, val, opts)` formats `val` using the format `fmt`.  If `fmt` is of 
type `number`, the internal table (and custom formats) will be used.  If `fmt` 
is a literal format, then it will be parsed and evaluated.

`.parse_date_code(val, opts)` parses `val` as date code and returns object:

- `D,T`: Date (`[val]`) Time (`{val}`)
- `y,m,d`: Year, Month, Day
- `H,M,S,u`: (0-23)Hour, Minute, Second, Sub-second
- `q`: Day of Week (0=Sunday, 1=Monday, ..., 5=Friday, 6=Saturday)

`.get_table()` gets the internal format table (number to format mapping).

`.load_table(table)` sets the internal format table.

## Notes

Format code 14 in the spec is broken; the correct format is 'mm/dd/yy' (dashes,
not spaces)

## License

Apache 2.0

## Tests

[![Build Status](https://travis-ci.org/SheetJS/ssf.svg?branch=master)](https://travis-ci.org/SheetJS/ssf)

[![Coverage Status](https://coveralls.io/repos/SheetJS/ssf/badge.png?branch=master)](https://coveralls.io/r/SheetJS/ssf?branch=master)

[![githalytics.com alpha](https://cruel-carlota.pagodabox.com/c1dac903f4b43f82a529bc8df145d085 "githalytics.com")](http://githalytics.com/SheetJS/ssf)

