# Angular 2+

The library can be imported directly from TS code with:

```typescript
import * as XLSX from 'xlsx';
```

This demo uses an array of arrays (type `Array<Array<any>>`) as the core state.
The component template includes a file input element, a table that updates with
the data, and a button to export the data.

## Switching between Angular versions

Modules that work with Angular 2 largely work as-is with Angular 4.  Switching
between versions is mostly a matter of installing the correct version of the
core and associated modules.  This demo includes a `package.json` for Angular 2
and another `package.json` for Angular 4.

Switching to Angular 2 is as simple as:

```bash
$ cp package.json-angular2 package.json
$ npm install
$ ng serve
```

Switching to Angular 4 is as simple as:

```bash
$ cp package.json-angular4 package.json
$ npm install
$ ng serve
```

## XLSX Symlink

In this tree, `node_modules/xlsx` is a symlink pointing back to the root.  This
enables testing the development version of the library.  In order to use this
demo in other applications, add the `xlsx` dependency:

```bash
$ npm install --save xlsx

```

## SystemJS Configuration

The default angular-cli configuration requires no additional configuration.

Some deployments use the SystemJS loader, which does require configuration.  The
SystemJS example shows the required meta and map settings:

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
```
