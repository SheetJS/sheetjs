## Getting Started

### Installation

In the browser, just add a script tag:

```html
<script lang="javascript" src="dist/xlsx.full.min.js"></script>
```

<details>
  <summary><b>CDN Availability</b> (click to show)</summary>

|    CDN     | URL                                        |
|-----------:|:-------------------------------------------|
|    `unpkg` | <https://unpkg.com/xlsx/>                  |
| `jsDelivr` | <https://jsdelivr.com/package/npm/xlsx>    |
|    `CDNjs` | <https://cdnjs.com/libraries/xlsx>         |
|    `packd` | <https://bundle.run/xlsx@latest?name=XLSX> |

`unpkg` makes the latest version available at:

```html
<script src="https://unpkg.com/xlsx/dist/xlsx.full.min.js"></script>
```

</details>


With [npm](https://www.npmjs.org/package/xlsx):

```bash
$ npm install xlsx
```

With [bower](https://bower.io/search/?q=js-xlsx):

```bash
$ bower install js-xlsx
```

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

A slimmer build is generated at `dist/xlsx.mini.min.js`. Compared to full build:
- codepage library skipped (no support for XLS encodings)
- XLSX compression option not currently available
- no support for XLSB / XLS / Lotus 1-2-3 / SpreadsheetML 2003
- node stream utils removed

Webpack and Browserify builds include optional modules by default.  Webpack can
be configured to remove support with `resolve.alias`:

```js
  /* uncomment the lines below to remove support */
  resolve: {
    alias: { "./dist/cpexcel.js": "" } // <-- omit international support
  }
```

</details>

<details>
  <summary><b>ECMAScript 3 Compatibility</b> (click to show)</summary>

For broad compatibility with JavaScript engines, the library is written using
ECMAScript 3 language dialect as well as some ES5 features like `Array#forEach`.
Older browsers require shims to provide missing functions.

To use the shim, add the shim before the script tag that loads `xlsx.js`:

```html
<!-- add the shim first -->
<script type="text/javascript" src="shim.min.js"></script>
<!-- after the shim is referenced, add the library -->
<script type="text/javascript" src="xlsx.full.min.js"></script>
```

The script also includes `IE_LoadFile` and `IE_SaveFile` for loading and saving
files in Internet Explorer versions 6-9.  The `xlsx.extendscript.js` script
bundles the shim in a format suitable for Photoshop and other Adobe products.

</details>

