# Rollup

This library presents itself as a CommonJS library, so some configuration is
required.  The examples at <https://rollupjs.org> can be followed pretty much in
verbatim.  This sample demonstrates a bundle for browser as well as for node.

This demo uses the `import` form to expose the whole library, enabling client
code to access the library with `import XLSX from 'xlsx'`.  The JS code from
the root demo was moved to a separate `app.js` script.

## Required Plugins

The `rollup-plugin-node-resolve` and `rollup-plugin-commonjs` plugins are used:

```js
import resolve from 'rollup-plugin-node-resolve';
import commonjs from 'rollup-plugin-commonjs';
export default {
	/* ... */
	plugins: [
		resolve({
			module: false, // <-- this library is not an ES6 module
			browser: true, // <-- suppress node-specific features
		}),
		commonjs()
	],
	/* ... */
};
```

For the browser deployments, the output format is `'iife'`.  For node, the
output format is `'cjs'`.

### Worker Scripts

Rollup can also bundle worker scripts!  Instead of using `importScripts`, the
worker script should import the module:

```diff
-importScripts('dist/xlsx.full.min.js');
+import XLSX from 'xlsx';
```

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
