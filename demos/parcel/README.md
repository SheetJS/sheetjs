# Parcel

Parcel Bundler starting from version 1.5.0 should play nice with this library
out of the box.  The standard import form can be used in JS files:

```js
import XLSX from 'xlsx'
```

Errors of the form `Could not statically evaluate fs call` stem from a parcel
[bug](https://github.com/parcel-bundler/parcel/pull/523#issuecomment-357486164).
Upgrade to version 1.5.0 or later.

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
