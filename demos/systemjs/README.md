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

Note: The `readFile` and `writeFile` functions are not available in the browser.

## Web Workers

Web Workers can load the SystemJS library with `importScripts`, but the imported
code cannot assign the original worker's `onmessage` callback.  This demo works
around the limitation by exposing the desired function as a global:

```js
/* main worker script */
importScripts('system.js');

SystemJS.config({ /* ... browser config ... */ });

onmessage = function(evt) {
	SystemJS.import('xlsxworker.js').then(function() { _cb(evt); });
};

/* xlsxworker.js */
var XLSX = require('xlsx');

_cb = function(evt) { /* ... do work here ... */ };
```

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

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
