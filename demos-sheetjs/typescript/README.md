# TypeScript

The library exports can be imported directly from TS code with:

```typescript
import * as XLSX from 'xlsx';
```

The library type definitions are available in the repo at `types/index.d.ts` and
in the node module.  The definitions are also available in places that serve the
node module, like [unpkg](https://unpkg.com/xlsx/types/index.d.ts).

This demo shows a small utility function that reads the first worksheet and
converts to an array of arrays.  The utility function is designed to be used in
the browser and server.  This project shows a complete deployment as a simple
browser script and as a node module.

This demo is intended to illustrate simple and direct use of the `tsc` command
line utility.  The Angular 2+ demo shows a more advanced TypeScript deployment.


## Named Exports

Newer TypeScript versions (2.6+) support named exports:

```typescript
import { read, write, utils } from 'xlsx'
```

However, since this is not supported in all deployments, it is generally easier
to use the glob import form and destructuring assignment:

```typescript
import * as XLSX from 'xlsx';
const { read, write, utils } = XLSX;
```


## Library Type Definitions

Types are exposed in the node module directly in the path `/types/index.d.ts`.
[unpkg CDN includes the definitions](https://unpkg.com/xlsx/types/index.d.ts).
The named `@types/xlsx` module should not be installed!

Using the glob import, types must be explicitly scoped:

```typescript
import * as XLSX from 'xlsx';
/* the workbook type is accessible as XLSX.WorkBook */
const wb: XLSX.WorkBook = XLSX.read(data, options);
```

Using named imports, the explicit type name should be imported:

```typescript
import { read, WorkBook } from 'xlsx'
const wb: WorkBook = read(data, options);
```


## Demo Project Structure

`lib/index.ts` is the TS library that will be transpiled to `dist/index.js` and
`dist/index.d.ts`.

`demo.js` is a node script that uses the generated library.

`src/index.js` is the browser entry point.  The `browserify` bundle tool is used
to generate `dist/browser.js`, a browser script loaded by `index.html`.

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
