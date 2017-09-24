## Installation

In the browser, just add a script tag:

```html
<script lang="javascript" src="dist/xlsx.full.min.js"></script>
```

<details>
  <summary><b>CDN Availability</b> (click to show)</summary>

|    CDN     | URL                                      |
|-----------:|:-----------------------------------------|
|    `unpkg` | <https://unpkg.com/xlsx/>                |
| `jsDelivr` | <https://jsdelivr.com/package/npm/xlsx>  |
|    `CDNjs` | <http://cdnjs.com/libraries/xlsx>        |

`unpkg` makes the latest version available at:

```html
<script src="https://unpkg.com/xlsx/dist/xlsx.full.min.js"></script>
```

</details>


With [npm](https://www.npmjs.org/package/xlsx):

```bash
$ npm install xlsx
```

With [bower](http://bower.io/search/?q=js-xlsx):

```bash
$ bower install js-xlsx
```

### JS Ecosystem Demos

The [`demos` directory](demos/) includes sample projects for:

**Frameworks and APIs**
- [`angular 1.x`](demos/angular/)
- [`angular 2.x / 4.x`](demos/angular2/)
- [`meteor`](demos/meteor/)
- [`react and react-native`](demos/react/)
- [`vue 2.x and weex`](demos/vue/)
- [`XMLHttpRequest and fetch`](demos/xhr/)
- [`nodejs server`](demos/server/)

**Bundlers and Tooling**
- [`browserify`](demos/browserify/)
- [`requirejs`](demos/requirejs/)
- [`rollup`](demos/rollup/)
- [`systemjs`](demos/systemjs/)
- [`webpack 2.x`](demos/webpack/)

**Platforms and Integrations**
- [`electron application`](demos/electron/)
- [`nw.js application`](demos/nwjs/)
- [`Adobe ExtendScript`](demos/extendscript/)
- [`Headless Browsers`](demos/headless/)
- [`canvas-datagrid`](demos/datagrid/)
- [`Swift JSC and other engines`](demos/altjs/)

### Optional Modules

<details>
  <summary><b>Optional features</b> (click to show)</summary>

The node version automatically requires modules for additional features.  Some
of these modules are rather large in size and are only needed in special
circumstances, so they do not ship with the core.  For browser use, they must
be included directly:

```html
<!-- international support from js-codepage -->
<script src="dist/cpexcel.js"></script>
```

An appropriate version for each dependency is included in the dist/ directory.

The complete single-file version is generated at `dist/xlsx.full.min.js`

Webpack and Browserify builds include optional modules by default.  Webpack can
be configured to remove support with `resolve.alias`:

```js
  /* uncomment the lines below to remove support */
  resolve: {
    alias: { "./dist/cpexcel.js": "" } // <-- omit international support
  }
```

</details>

### ECMAScript 5 Compatibility

Since the library uses functions like `Array#forEach`, older browsers require
[shims to provide missing functions](http://oss.sheetjs.com/js-xlsx/shim.js).

To use the shim, add the shim before the script tag that loads `xlsx.js`:

```html
<!-- add the shim first -->
<script type="text/javascript" src="shim.js"></script>
<!-- after the shim is referenced, add the library -->
<script type="text/javascript" src="xlsx.full.min.js"></script>
```

