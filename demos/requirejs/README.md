# RequireJS

The minified dist files trip up the RequireJS mechanism.  To bypass, the scripts
automatically expose an `XLSX` variable that can be used if the require callback
argument is `_XLSX` rather than `XLSX`:

```js
require(["xlsx.full.min"], function(_XLSX) { /* ... */ });
```

This demo uses the `r.js` optimizer to build a source file.
