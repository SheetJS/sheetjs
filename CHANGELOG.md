# CHANGELOG

This log is intended to keep track of backwards-incompatible changes, including
but not limited to API changes and file location changes.  Minor behavioral
changes may not be included if they are not expected to break existing code.


## Unreleased

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

