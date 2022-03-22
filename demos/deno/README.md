# Deno

Deno is a runtime capable of running JS code including this library.  There are
a few different builds and recommended use cases as covered in this demo.

Due to ongoing stability and sync issues with the Deno registry, scripts should
use [the `unpkg` CDN build](https://unpkg.com/xlsx/xlsx.mjs):

```js
// @deno-types="https://unpkg.com/xlsx/types/index.d.ts"
import * as XLSX from 'https://unpkg.com/xlsx/xlsx.mjs';

/* load the codepage support library for extended support with older formats  */
import * as cptable from 'https://unpkg.com/xlsx/dist/cpexcel.full.mjs';
XLSX.set_cptable(cptable);
```


## Reading and Writing Files

In general, the command-line flag `--allow-read` must be passed to enable file
reading.  The flag `--allow-write` must be passed to enable file writing.

_Reading a File_

```ts
const workbook = XLSX.readFile("test.xlsx");
```

_Writing a File_

Older versions of the library did not properly detect features from Deno, so the
`buffer` export would return an array of bytes.  Since `Deno.writeFileSync` does
not handle byte arrays, user code must generate a `Uint8Array` first:

```ts
XLSX.writeFile(workbook, "test.xlsb");
```

## Demos

**Complete Examples**

`sheet2csv.ts` is a complete command-line tool for generating CSV text from
workbooks.  Building the application is incredibly straightforward:

```bash
$ deno compile -r --allow-read sheet2csv.ts  # build the sheet2csv binary
$ ./sheet2csv test.xlsx                      # print the first worksheet as CSV
$ ./sheet2csv test.xlsx s5s                  # print worksheet "s5s" as CSV
```

The [`server` demo](../server) includes a sample Deno server for parsing uploads
and generating HTML TABLE previews.


**Module Import Scenarios**

All demos attempt to read a file and write a new file.  [`doit.ts`](./doit.ts)
accepts the `XLSX` module as an argument.

- `x` imports the ESM build without the codepage library:

```ts
// @deno-types="https://unpkg.com/xlsx/types/index.d.ts"
import * as XLSX from 'https://unpkg.com/xlsx/xlsx.mjs';
```

- `mjs` imports the ESM build and the associated codepage library:

```ts
import * as XLSX from '../../xlsx.mjs';
/* recommended for reading XLS files */
import * as cptable from '../../dist/cptable.full.mjs';
XLSX.set_cptable(cptable);
```

- `node` uses the node compatibility layer:

```ts
import { createRequire } from 'https://deno.land/std/node/module.ts';
const require = createRequire(import.meta.url);
const XLSX = require('../../');
```



[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
