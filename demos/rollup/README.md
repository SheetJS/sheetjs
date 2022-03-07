# Rollup

This library has a proper ESM build that is enabled by default:

```js
import { read, utils } from 'xlsx';
```

This sample demonstrates a bundle for browser as well as for node.

This demo uses the `import` form to expose the whole library, enabling client
code to access the library with `import XLSX from 'xlsx'`.  The JS code from
the root demo was moved to a separate `app.js` script.

## Required Plugins

The `rollup-plugin-node-resolve` plugin is used:

```js
import resolve from 'rollup-plugin-node-resolve';
export default {
	/* ... */
	plugins: [
		resolve({
			module: false, // <-- this library is not an ES6 module
			browser: true, // <-- suppress node-specific features
		})
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
+import * as XLSX from 'xlsx'; // or do named imports
```

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
