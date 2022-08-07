# s2c

`mod.ts` exports the following methods:

- `parse_book_from_request` reads a field from a request, parses the field with
  the Drash body parser, and returns a SheetJS workbook object. This does not
  use the filesystem, so it supports Deno Deploy and other restricted services.

`s2c.ts` is the script that powers <https://s2c.sheetjs.com>

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
