## Getting Started

### Installation

#### Standalone Browser Scripts

Each standalone release script is available at <https://cdn.sheetjs.com/>.

The current version is `0.18.6` and can be referenced as follows:

```html
<!-- use version 0.18.6 -->
<script lang="javascript" src="https://cdn.sheetjs.com/xlsx-0.18.6/package/dist/xlsx.full.min.js"></script>
```

The `latest` tag references the latest version and updates with each release:

```html
<!-- use the latest version -->
<script lang="javascript" src="https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js"></script>
```

**For production use, scripts should be downloaded and added to a public folder
alongside other scripts.**

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
<!-- use xlsx.mini.min.js from version 0.18.6 -->
<script lang="javascript" src="https://cdn.sheetjs.com/xlsx-0.18.6/package/dist/xlsx.mini.min.js"></script>
```

</details>


[Bower](https://bower.io/search/?q=js-xlsx) will pull the entire repo:

```bash
$ bower install js-xlsx
```

Bower will place the standalone scripts in `bower_components/js-xlsx/dist/`

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

Due to SSL certificate compatibility issues, it is highly recommended to save
the Standalone and Shim scripts from <https://cdn.sheetjs.com/> and add to a
public directory in the site.

The script also includes `IE_LoadFile` and `IE_SaveFile` for loading and saving
files in Internet Explorer versions 6-9.  The `xlsx.extendscript.js` script
bundles the shim in a format suitable for Photoshop and other Adobe products.

</details>


#### ECMAScript Modules

_Browser ESM_

The ECMAScript Module build is saved to `xlsx.mjs` and can be directly added to
a page with a `script` tag using `type="module"`:

```html
<script type="module">
import { read, writeFileXLSX } from "https://cdn.sheetjs.com/xlsx-0.18.6/package/xlsx.mjs";

/* load the codepage support library for extended support with older formats  */
import { set_cptable } from "https://cdn.sheetjs.com/xlsx-0.18.6/package/xlsx.mjs";
import * as cptable from 'https://cdn.sheetjs.com/xlsx-0.18.6/package/dist/cpexcel.full.mjs';
set_cptable(cptable);
</script>
```

_Frameworks (Angular, VueJS, React) and Bundlers (webpack, etc)_

The NodeJS package is readily installed from the tarballs:

```bash
$ npm  install --save https://cdn.sheetjs.com/xlsx-0.18.6/xlsx-0.18.6.tgz # npm
$ pnpm install --save https://cdn.sheetjs.com/xlsx-0.18.6/xlsx-0.18.6.tgz # pnpm
$ yarn add     --save https://cdn.sheetjs.com/xlsx-0.18.6/xlsx-0.18.6.tgz # yarn
```

Once installed, the library can be imported under the name `xlsx`:

```ts
import { read, writeFileXLSX } from "xlsx";

/* load the codepage support library for extended support with older formats  */
import { set_cptable } from "xlsx";
import * as cptable from 'xlsx/dist/cpexcel.full.mjs';
set_cptable(cptable);
```

#### Deno

`xlsx.mjs` can be imported in Deno:

```ts
// @deno-types="https://cdn.sheetjs.com/xlsx-0.18.6/package/types/index.d.ts"
import * as XLSX from 'https://cdn.sheetjs.com/xlsx-0.18.6/package/xlsx.mjs';

/* load the codepage support library for extended support with older formats  */
import * as cptable from 'https://cdn.sheetjs.com/xlsx-0.18.6/package/dist/cpexcel.full.mjs';
XLSX.set_cptable(cptable);
```

#### NodeJS

Tarballs are available on <https://cdn.sheetjs.com>.

Each individual version can be referenced using a similar URL pattern.
<https://cdn.sheetjs.com/xlsx-0.18.6/xlsx-0.18.6.tgz> is the URL for `0.18.6`

<https://cdn.sheetjs.com/xlsx-latest/xlsx-latest.tgz> is a link to the latest
version and will refresh on each release.

_Installation_

Tarballs can be directly installed using a package manager:

```bash
$ npm  install https://cdn.sheetjs.com/xlsx-0.18.6/xlsx-0.18.6.tgz # npm
$ pnpm install https://cdn.sheetjs.com/xlsx-0.18.6/xlsx-0.18.6.tgz # pnpm
$ yarn add     https://cdn.sheetjs.com/xlsx-0.18.6/xlsx-0.18.6.tgz # yarn
```

For general stability, "vendoring" modules is the recommended approach:

1) Download the tarball (`xlsx-0.18.6.tgz`) for the desired version. The current
   version is available at <https://cdn.sheetjs.com/xlsx-0.18.6/xlsx-0.18.6.tgz>

2) Create a `vendor` subdirectory at the root of your project and move the
   tarball to that folder.  Add it to your project repository.

3) Install the tarball using a package manager:

```bash
$ npm  install --save file:vendor/xlsx-0.18.6.tgz # npm
$ pnpm install --save file:vendor/xlsx-0.18.6.tgz # pnpm
$ yarn add            file:vendor/xlsx-0.18.6.tgz # yarn
```

The package will be installed and accessible as `xlsx`.

_Usage_

By default, the module supports `require` and it will automatically add support
for streams and filesystem access:

```js
var XLSX = require("xlsx");
```

The module also ships with `xlsx.mjs` for use with `import`.  The `mjs` version
does not automatically load native node modules:

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

#### Photoshop and InDesign

`dist/xlsx.extendscript.js` is an ExtendScript build for Photoshop and InDesign.
<https://cdn.sheetjs.com/xlsx-0.18.6/package/dist/xlsx.extendscript.js> is the
current version.  After downloading the script, it can be directly referenced
with a `#include` directive:

```extendscript
#include "xlsx.extendscript.js"
```


