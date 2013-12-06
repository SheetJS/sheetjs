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

## Usage

`.load(fmt, idx)` sets custom formats (generally indices above `164`)

`.format(fmt, val)` formats `val` using the format `fmt`.  If `fmt` is of type
`number`, the internal table (and custom formats) will be used.  If `fmt` is a
literal format, then it will be parsed and evaluated.

## Notes

Format code 14 in the spec is broken; the correct format is 'mm/dd/yy' (dashes,
not spaces)

## License

Apache 2.0
