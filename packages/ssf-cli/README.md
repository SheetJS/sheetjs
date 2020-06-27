# [SSF Command-Line Interface](http://sheetjs.com)

ssf (SpreadSheet Format) is a pure JS library to format data using ECMA-376
spreadsheet format codes (used in popular spreadsheet software packages).

This CLI tool formats numbers from shell scripts and other command-line tools.

## Installation

With [npm](https://www.npmjs.org/package/ssf-cli):

```bash
$ npm install -g ssf-cli
```

## Usage

`ssf-cli` takes two arguments: the format string and the value to be formatted.

The value is formatted twice, once interpreting the value as a string and once
interpreting the value as a number, and both results are printed to standard
output, with a pipe character `|` after each value:

```bash
$ bin/ssf.njs "#,##0.00" 12345
12345|12,345.00|
$ bin/ssf.njs "0;0;0;:@:" 12345
:12345:|12345|
```

Extracting the values in a pipeline is straightforward with AWK:

```bash
$ bin/ssf.njs "#,##0.00" 12345 | awk -F\| '{print $2}'
12,345.00
```

## License

Please consult the attached LICENSE file for details.  All rights not explicitly
granted by the Apache 2.0 license are reserved by the Original Author.

## Credits

Special thanks to [Garrett Luu](https://garrettluu.com/) for spinning off the
command from the SSF module.

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/ssf?pixel)](https://github.com/SheetJS/ssf)
