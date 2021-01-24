# FuseBox

This library is built with some dynamic logic to determine if it is invoked in a
script tag or in nodejs.  FuseBox does not understand those feature tests, so by
default it will do some strange things.

## TypeScript Support

As with most TS modules in FuseBox, the glob import form should be used:

```typescript
import * as XLSX from 'xlsx';
```

The included `sheetjs.ts` script will be transpiled and bundled to `server.js`
for the `"node"` target and `client.js` for the `"browser"` target.

## Proper Target Detection

Out of the box, FuseBox will automatically provide shims to browser globals like
`process` and `Browser`.  The proper way to detect `node` uses `process`:

```typescript
if(typeof process != 'undefined' && process.versions && process.versions.node) {
  /* Script is running in nodejs */
} else {
  /* Script is running in a browser environment */
}
```

## Server Target

The FuseBox documentation configuration can be used as-is:

```js
const fuse = FuseBox.init({
  homeDir: ".",
  target: "node",
  output: "$name.js"
});
fuse.bundle("server").instructions(">sheetjs.ts"); fuse.run();
```

## Browser Target

The native shims must be suppressed for browser usage:

```js
const fuse = FuseBox.init({
  homeDir: ".",
  target: "browser",
  natives: {
    Buffer: false,
    stream: false,
    process: false
  },
  output: "$name.js"
});
fuse.bundle("client").instructions(">sheetjs.ts"); fuse.run();
```

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
