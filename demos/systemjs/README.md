# SystemJS Demos

SystemJS supports both browser and nodejs deployments.  It does not recognize
browser environments and automatically suppress node core modules, but with some
configuration magic SystemJS can load the library.

## Browser

SystemJS fails by default because the library does not export anything in the
web browser.  This is easily addressed in the config:

```js
SystemJS.config({
	meta: {
		'xlsx': {
			exports: 'XLSX' // <-- tell SystemJS to expose the XLSX variable
		}
	},
	map: {
		'xlsx': 'xlsx.full.min.js', // <-- make sure xlsx.full.min.js is in same dir
		'fs': '',     // <--|
		'crypto': '', // <--| suppress native node modules
		'stream': ''  // <--|
	}
});
SystemJS.import('main.js')
```

In your main JS script, just use require:

```js
var XLSX = require('xlsx');
var w = XLSX.read('abc,def\nghi,jkl', {type:'binary'});
var j = XLSX.utils.sheet_to_json(w.Sheets[w.SheetNames[0]], {header:1});
console.log(j);
```

The file functions `readFile` and `writeFile` are not available in the browser.

## Node

The node core modules should be mapped to their `@node` equivalents:

```js
var SystemJS = require('systemjs');
SystemJS.config({
	map: {
		'xlsx': 'node_modules/xlsx/xlsx.js',
		'fs': '@node/fs',
		'crypto': '@node/crypto',
		'stream': '@node/stream'
	}
});
```

And use is pretty straightforward:

```js
SystemJS.import('xlsx').then(function(XLSX) {
	/* XLSX is available here */
	var w = XLSX.readFile('test.xlsx');
	var j = XLSX.utils.sheet_to_json(w.Sheets[w.SheetNames[0]], {header:1});
	console.log(j);
});
```

