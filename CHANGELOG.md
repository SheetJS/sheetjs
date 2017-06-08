# CHANGELOG

This log is intended to keep track of backwards-incompatible changes, including
but not limited to API changes and file location changes.  Minor behavioral
changes may not be included if they are not expected to break existing code.


## Unreleased (2017-06-??)

* HTML Table output header/footer should not include `<table>` tag

## 0.10.2 (2017-05-16)

* Dates are converted to numbers by default (set `cellDates:true` to emit Dates)
* Module does not export CFB

## 0.9.10 (2017-04-08)

* `--perf` renamed to `--read-only`

## 0.9.9 (2017-04-03)

* default output format changed to XLSB
* comment text line endings are now normalized
* errors thrown on write when worksheets have invalid names

## 0.9.7 (2017-03-28)

* XLS legacy `!range` field removed
* Hyperlink tooltip is stored in the `Tooltip` field

## 0.9.6 (2017-03-25)

* `sheet_to_json` now passes `null` values when `raw` is set to `true`
* `sheet_to_json` treats `null` stub cells as values in conjunction with `raw`

## 0.9.5 (2017-03-22)

* `cellDates` affects parsing in non-XLSX formats

## 0.9.3 (2017-03-15)

* XLML property names are more closely mapped to the XLSX equivalent
* Stub cells are now cell type `z`

## 0.9.2 (2017-03-13)

* Removed stale TypeScript definition files.  Flowtype comments are used in the
  `xlsx.flow.js` source and stripped to produce `xlsx.js`.
* sed usage reworked to support GNU sed in-place form.  BSD sed seems to work,
  but the build script has not been tested on other sed variants:

```bash
$ sed -i.ext  [...] # GNU
$ sed -i .ext [...] # bsd
```

## 0.9.0 (2017-03-09)

* Removed ods.js source.  The xlsx.js source absorbed the ODS logic and exposes
  the ODS variable, so projects should remove references to ods.js

