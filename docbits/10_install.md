## Getting Started

### Installation

**Standalone Browser Scripts**

The complete browser standalone build is saved to `dist/xlsx.full.min.js` and
can be directly added to a page with a `script` tag:

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

For example, `unpkg` makes the latest version available at:

```html
<script src="https://unpkg.com/xlsx/dist/xlsx.full.min.js"></script>
```

</details>

<details>
  <summary><b>Browser builds</b> (click to show)</summary>

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


With [bower](https://bower.io/search/?q=js-xlsx):

```bash
$ bower install js-xlsx
```

**Deno**

The [`sheetjs`](https://deno.land/x/sheetjs) package is hosted by Deno:

```ts
// @deno-types="https://deno.land/x/sheetjs/types/index.d.ts"
import * as XLSX from 'https://deno.land/x/sheetjs/xlsx.mjs'

/* load the codepage support library for extended support with older formats  */
import * as cptable from 'https://deno.land/x/sheetjs/dist/cpexcel.full.mjs';
XLSX.set_cptable(cptable);
```

**NodeJS**

With [npm](https://www.npmjs.org/package/xlsx):

```bash
$ npm install xlsx
```

By default, the module supports `require`:

```js
var XLSX = require("xlsx");
```

The module also ships with `xlsx.mjs` for use with `import`:

```js
import * as XLSX from 'xlsx/xlsx.mjs';

/* load 'fs' for readFile and writeFile support */
import * as fs from 'fs';
XLSX.set_fs(fs);

/* load the codepage support library for extended support with older formats  */
import * as cpexcel from 'xlsx/dist/cpexcel.full.mjs';
XLSX.set_cptable(cpexcel);
```

**Photoshop and InDesign**

`dist/xlsx.extendscript.js` is an ExtendScript build for Photoshop and InDesign
that is included in the `npm` package.  It can be directly referenced with a
`#include` directive:

```extendscript
#include "xlsx.extendscript.js"
```


<details>
  <summary><b>Internet Explorer and ECMAScript 3 Compatibility</b> (click to show)</summary>

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

