# Deno

Deno is a runtime capable of running JS code including this library.  There are
a few different builds and recommended use cases as covered in this demo.

For user code, [the `sheetjs` module](https://deno.land/x/sheetjs) can be used.

## Reading and Writing Files

In general, the command-line flag `--allow-read` must be passed to enable file
reading.  The flag `--allow-write` must be passed to enable file writing.

Starting in version 0.18.1, this library will check for the `Deno` global and
use `Deno.readFileSync` and `Deno.writeFileSync` behind the scenes.

For older versions, the API functions must be called from user code.

_Reading a File_

```ts
const filedata = Deno.readFileSync("test.xlsx");
const workbook = XLSX.read(filedata, {type: "buffer"});
/* DO SOMETHING WITH workbook HERE */
```

_Writing a File_

Older versions of the library did not properly detect features from Deno, so the
`buffer` export would return an array of bytes.  Since `Deno.writeFileSync` does
not handle byte arrays, user code must generate a `Uint8Array` first:

```ts
const buf = XLSX.write(workbook, {type: "buffer", bookType: "xlsb"});
const u8: Uint8Array = new Uint8Array(buf);
Deno.writeFileSync("test.xlsb", u8);
```

## Demos

**Complete Example**

`sheet2csv.ts` is a complete command-line tool for generating CSV text from
workbooks.  Building the application is incredibly straightforward:

```bash
$ deno compile -r --allow-read sheet2csv.ts  # build the sheet2csv binary
$ ./sheet2csv test.xlsx                      # print the first worksheet as CSV
$ ./sheet2csv test.xlsx s5s                  # print worksheet "s5s" as CSV
```

**Module Import Scenarios**

All demos attempt to read a file and write a new file.  [`doit.ts`](./doit.ts)
accepts the `XLSX` module as an argument.

- `x` imports the ESM build without the codepage library:

```ts
// @deno-types="https://deno.land/x/sheetjs/types/index.d.ts"
import * as XLSX from 'https://deno.land/x/sheetjs/xlsx.mjs';
```

- `mjs` imports the ESM build and the associated codepage library:

```ts
import * as XLSX from '../../xlsx.mjs';
/* recommended for reading XLS files */
import * as cptable from '../../dist/cptable.full.mjs';
XLSX.set_cptable(cptable);
```

- `jspm` imports the browser standalone script using JSPM:

```ts
import * as XLSX from 'https://jspm.dev/npm:xlsx!cjs';
```

- `node` uses the node compatibility layer:

```ts
import { createRequire } from 'https://deno.land/std/node/module.ts';
const require = createRequire(import.meta.url);
const XLSX = require('../../');
```



[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
