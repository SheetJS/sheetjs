## Installation

In the browser, just add a script tag:

```html
<script lang="javascript" src="dist/xlsx.full.min.js"></script>
```

With [npm](https://www.npmjs.org/package/xlsx):

```bash
$ npm install xlsx
```

With [bower](http://bower.io/search/?q=js-xlsx):

```bash
$ bower install js-xlsx
```

CDNjs automatically pulls the latest version and makes all versions available at
<http://cdnjs.com/libraries/xlsx>

### JS Ecosystem Demos

The `demos` directory includes sample projects for:

**JS Frameworks and APIs**
- [`angular 1.x`](demos/angular/)
- [`angular 2.x / 4.x`](demos/angular2/)
- [`meteor`](demos/meteor/)
- [`vue 2.x`](demos/vue/)
- [`XMLHttpRequest and fetch`](demos/xhr/)

**JS Bundlers and Tooling**
- [`browserify`](demos/browserify/)
- [`requirejs`](demos/requirejs/)
- [`rollup`](demos/rollup/)
- [`systemjs`](demos/systemjs/)
- [`webpack 2.x`](demos/webpack/)

**JS Platforms and Integrations**
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

Webpack and browserify builds include optional modules by default.  Webpack can
be configured to remove support with `resolve.alias`:

```js
  /* uncomment the lines below to remove support */
  resolve: {
    alias: { "./dist/cpexcel.js": "" } // <-- omit international support
  }
```

</details>

### ECMAScript 5 Compatibility

Since xlsx.js uses ES5 functions like `Array#forEach`, older browsers require
[Polyfills](http://git.io/QVh77g).  This repo and the gh-pages branch include
[a shim](https://github.com/SheetJS/js-xlsx/blob/master/shim.js)

To use the shim, add the shim before the script tag that loads xlsx.js:

```html
<script type="text/javascript" src="/path/to/shim.js"></script>
```

