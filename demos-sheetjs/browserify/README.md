# Browserify

The library is compatible with Browserify and should just work out of the box.

This demo uses the `require` form to expose the whole library, enabling client
code to access the library with `var XLSX = require('xlsx')`.  The JS code from
the root demo was moved to a separate `app.js` script.  That script is bundled:

```bash
browserify app.js > browserify.js
uglifyjs browserify.js > browserify.min.js
```

### Worker Scripts

Browserify can also bundle worker scripts!  Instead of using `importScripts`,
the worker script should require the module:

```diff
-importScripts('dist/xlsx.full.min.js');
+var XLSX = require('xlsx');
```

The same process generates the worker script:

```bash
browserify xlsxworker.js > worker.js
uglifyjs worker.js > worker.min.js
```

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
