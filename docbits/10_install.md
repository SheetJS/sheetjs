## Getting Started

### Installation

**Standalone Browser Scripts**

The complete browser standalone build is saved to `dist/xlsx.full.min.js` and
can be directly added to a page with a `script` tag:

```html
<script lang="javascript" src="dist/xlsx.full.min.js"></script>
```

Each standalone release script is available at <https://cdn.sheetjs.com/>.  The
latest version uses the `latest` tag:

```html
<!-- use the latest version -->
<script lang="javascript" src="https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js"></script>
```

A specific release can be referenced by version:

```html
<!-- use version 0.18.5 -->
<script lang="javascript" src="https://cdn.sheetjs.com/xlsx-0.18.5/package/dist/xlsx.full.min.js"></script>
```

</details>

<details>
  <summary><b>Browser builds</b> (click to show)</summary>

The complete single-file version is generated at `dist/xlsx.full.min.js`

`dist/xlsx.core.min.js` omits codepage library (no support for XLS encodings)

A slimmer build is generated at `dist/xlsx.mini.min.js`. Compared to full build:
- codepage library skipped (no support for XLS encodings)
- no support for XLSB / XLS / Lotus 1-2-3 / SpreadsheetML 2003 / Numbers
- node stream utils removed

These scripts are also available on the CDN:

```html
<!-- use xlsx.core.min.js from the latest version -->
<script lang="javascript" src="https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.core.min.js"></script>
```

</details>


With [bower](https://bower.io/search/?q=js-xlsx):

```bash
$ bower install js-xlsx
```

**ECMAScript Modules**

The ECMAScript Module build is saved to `xlsx.mjs` and can be directly added to
a page with a `script` tag using `type=module`:

```html
<script type="module">
import { read, writeFileXLSX } from "./xlsx.mjs";

/* load the codepage support library for extended support with older formats  */
import { set_cptable } from "./xlsx.mjs";
import * as cptable from './dist/cpexcel.full.mjs';
set_cptable(cptable);
</script>
```

The [npm package](https://www.npmjs.org/package/xlsx) also exposes the module
with the `module` parameter, supported in Angular and other projects:

```ts
import { read, writeFileXLSX } from "xlsx";

/* load the codepage support library for extended support with older formats  */
import { set_cptable } from "xlsx";
import * as cptable from 'xlsx/dist/cpexcel.full.mjs';
set_cptable(cptable);
```

**Deno**

`xlsx.mjs` can be imported in Deno:

```ts
// @deno-types="https://cdn.sheetjs.com/xlsx-latest/package/types/index.d.ts"
import * as XLSX from 'https://cdn.sheetjs.com/xlsx-latest/package/xlsx.mjs';

/* load the codepage support library for extended support with older formats  */
import * as cptable from 'https://cdn.sheetjs.com/xlsx-latest/package/dist/cpexcel.full.mjs';
XLSX.set_cptable(cptable);
```

**NodeJS**

Modules are available on [the public npm registry](https://www.npmjs.org/package/xlsx):

```bash
$ pnpm install xlsx  # using pnpm
$ yarn add xlsx      # using yarn
$ npm install xlsx   # using npm
```

Tarballs are also available on <https://cdn.sheetjs.com>:

```bash
$ npm install https://cdn.sheetjs.com/xlsx-latest/xlsx-latest.tgz
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

/* load 'stream' for stream support */
import { Readable } from 'stream';
XLSX.stream.set_readable(Readable);

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

