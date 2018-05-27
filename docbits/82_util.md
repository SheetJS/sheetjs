## Utility Functions

The `sheet_to_*` functions accept a worksheet and an optional options object.

The `*_to_sheet` functions accept a data object and an optional options object.

The examples are based on the following worksheet:

```
XXX| A | B | C | D | E | F | G |
---+---+---+---+---+---+---+---+
 1 | S | h | e | e | t | J | S |
 2 | 1 | 2 | 3 | 4 | 5 | 6 | 7 |
 3 | 2 | 3 | 4 | 5 | 6 | 7 | 8 |
```

### Array of Arrays Input

`XLSX.utils.aoa_to_sheet` takes an array of arrays of JS values and returns a
worksheet resembling the input data.  Numbers, Booleans and Strings are stored
as the corresponding styles.  Dates are stored as date or numbers.  Array holes
and explicit `undefined` values are skipped.  `null` values may be stubbed. All
other values are stored as strings.  The function takes an options argument:

| Option Name |  Default | Description                                         |
| :---------- | :------: | :-------------------------------------------------- |
|`dateNF`     |  FMT 14  | Use specified date format in string output          |
|`cellDates`  |  false   | Store dates as type `d` (default is `n`)            |
|`sheetStubs` |  false   | Create cell objects of type `z` for `null` values   |

<details>
  <summary><b>Examples</b> (click to show)</summary>

To generate the example sheet:

```js
var ws = XLSX.utils.aoa_to_sheet([
  "SheetJS".split(""),
  [1,2,3,4,5,6,7],
  [2,3,4,5,6,7,8]
]);
```
</details>

`XLSX.utils.sheet_add_aoa` takes an array of arrays of JS values and updates an
existing worksheet object.  It follows the same process as `aoa_to_sheet` and
accepts an options argument:

| Option Name |  Default | Description                                         |
| :---------- | :------: | :-------------------------------------------------- |
|`dateNF`     |  FMT 14  | Use specified date format in string output          |
|`cellDates`  |  false   | Store dates as type `d` (default is `n`)            |
|`sheetStubs` |  false   | Create cell objects of type `z` for `null` values   |
|`origin`     |          | Use specified cell as starting point (see below)    |

`origin` is expected to be one of:

| `origin`         | Description                                               |
| :--------------- | :-------------------------------------------------------- |
| (cell object)    | Use specified cell (cell object)                          |
| (string)         | Use specified cell (A1-style cell)                        |
| (number >= 0)    | Start from the first column at specified row (0-indexed)  |
| -1               | Append to bottom of worksheet starting on first column    |
| (default)        | Start from cell A1                                        |


<details>
  <summary><b>Examples</b> (click to show)</summary>

Consider the worksheet:

```
XXX| A | B | C | D | E | F | G |
---+---+---+---+---+---+---+---+
 1 | S | h | e | e | t | J | S |
 2 | 1 | 2 |   |   | 5 | 6 | 7 |
 3 | 2 | 3 |   |   | 6 | 7 | 8 |
 4 | 3 | 4 |   |   | 7 | 8 | 9 |
 5 | 4 | 5 | 6 | 7 | 8 | 9 | 0 |
```

This worksheet can be built up in the order `A1:G1, A2:B4, E2:G4, A5:G5`:

```js
/* Initial row */
var ws = XLSX.utils.aoa_to_sheet([ "SheetJS".split("") ]);

/* Write data starting at A2 */
XLSX.utils.sheet_add_aoa(ws, [[1,2], [2,3], [3,4]], {origin: "A2"});

/* Write data starting at E2 */
XLSX.utils.sheet_add_aoa(ws, [[5,6,7], [6,7,8], [7,8,9]], {origin:{r:1, c:4}});

/* Append row */
XLSX.utils.sheet_add_aoa(ws, [[4,5,6,7,8,9,0]], {origin: -1});
```

</details>

### Array of Objects Input

`XLSX.utils.json_to_sheet` takes an array of objects and returns a worksheet
with automatically-generated "headers" based on the keys of the objects.  The
default column order is determined by the first appearance of the field using
`Object.keys`, but can be overridden using the options argument:

| Option Name |  Default | Description                                         |
| :---------- | :------: | :-------------------------------------------------- |
|`header`     |          | Use specified column order (default `Object.keys`)  |
|`dateNF`     |  FMT 14  | Use specified date format in string output          |
|`cellDates`  |  false   | Store dates as type `d` (default is `n`)            |
|`skipHeader` |  false   | If true, do not include header row in output        |

<details>
  <summary><b>Examples</b> (click to show)</summary>

The original sheet cannot be reproduced in the obvious way since JS object keys
must be unique. After replacing the second `e` and `S` with `e_1` and `S_1`:

```js
var ws = XLSX.utils.json_to_sheet([
  { S:1, h:2, e:3, e_1:4, t:5, J:6, S_1:7 },
  { S:2, h:3, e:4, e_1:5, t:6, J:7, S_1:8 }
], {header:["S","h","e","e_1","t","J","S_1"]});
```

Alternatively, the header row can be skipped:

```js
var ws = XLSX.utils.json_to_sheet([
  { A:"S", B:"h", C:"e", D:"e", E:"t", F:"J", G:"S" },
  { A: 1,  B: 2,  C: 3,  D: 4,  E: 5,  F: 6,  G: 7  },
  { A: 2,  B: 3,  C: 4,  D: 5,  E: 6,  F: 7,  G: 8  }
], {header:["A","B","C","D","E","F","G"], skipHeader:true});
```

</details>

`XLSX.utils.sheet_add_json` takes an array of objects and updates an existing
worksheet object.  It follows the same process as `json_to_sheet` and accepts
an options argument:

| Option Name |  Default | Description                                         |
| :---------- | :------: | :-------------------------------------------------- |
|`header`     |          | Use specified column order (default `Object.keys`)  |
|`dateNF`     |  FMT 14  | Use specified date format in string output          |
|`cellDates`  |  false   | Store dates as type `d` (default is `n`)            |
|`skipHeader` |  false   | If true, do not include header row in output        |
|`origin`     |          | Use specified cell as starting point (see below)    |

`origin` is expected to be one of:

| `origin`         | Description                                               |
| :--------------- | :-------------------------------------------------------- |
| (cell object)    | Use specified cell (cell object)                          |
| (string)         | Use specified cell (A1-style cell)                        |
| (number >= 0)    | Start from the first column at specified row (0-indexed)  |
| -1               | Append to bottom of worksheet starting on first column    |
| (default)        | Start from cell A1                                        |


<details>
  <summary><b>Examples</b> (click to show)</summary>

Consider the worksheet:

```
XXX| A | B | C | D | E | F | G |
---+---+---+---+---+---+---+---+
 1 | S | h | e | e | t | J | S |
 2 | 1 | 2 |   |   | 5 | 6 | 7 |
 3 | 2 | 3 |   |   | 6 | 7 | 8 |
 4 | 3 | 4 |   |   | 7 | 8 | 9 |
 5 | 4 | 5 | 6 | 7 | 8 | 9 | 0 |
```

This worksheet can be built up in the order `A1:G1, A2:B4, E2:G4, A5:G5`:

```js
/* Initial row */
var ws = XLSX.utils.json_to_sheet([
  { A: "S", B: "h", C: "e", D: "e", E: "t", F: "J", G: "S" }
], {header: ["A", "B", "C", "D", "E", "F", "G"], skipHeader: true});

/* Write data starting at A2 */
XLSX.utils.sheet_add_json(ws, [
  { A: 1, B: 2 }, { A: 2, B: 3 }, { A: 3, B: 4 }
], {skipHeader: true, origin: "A2"});

/* Write data starting at E2 */
XLSX.utils.sheet_add_json(ws, [
  { A: 5, B: 6, C: 7 }, { A: 6, B: 7, C: 8 }, { A: 7, B: 8, C: 9 }
], {skipHeader: true, origin: { r: 1, c: 4 }, header: [ "A", "B", "C" ]});

/* Append row */
XLSX.utils.sheet_add_json(ws, [
  { A: 4, B: 5, C: 6, D: 7, E: 8, F: 9, G: 0 }
], {header: ["A", "B", "C", "D", "E", "F", "G"], skipHeader: true, origin: -1});
```

</details>

### HTML Table Input

`XLSX.utils.table_to_sheet` takes a table DOM element and returns a worksheet
resembling the input table.  Numbers are parsed.  All other data will be stored
as strings.

`XLSX.utils.table_to_book` produces a minimal workbook based on the worksheet.

Both functions accept options arguments:

| Option Name |  Default | Description                                         |
| :---------- | :------: | :-------------------------------------------------- |
|`raw`        |          | If true, every cell will hold raw strings           |
|`dateNF`     |  FMT 14  | Use specified date format in string output          |
|`cellDates`  |  false   | Store dates as type `d` (default is `n`)            |
|`sheetRows`  |    0     | If >0, read the first `sheetRows` rows of the table |
|`display`    |  false   | If true, hidden rows and cells will not be parsed   |


<details>
  <summary><b>Examples</b> (click to show)</summary>

To generate the example sheet, start with the HTML table:

```html
<table id="sheetjs">
<tr><td>S</td><td>h</td><td>e</td><td>e</td><td>t</td><td>J</td><td>S</td></tr>
<tr><td>1</td><td>2</td><td>3</td><td>4</td><td>5</td><td>6</td><td>7</td></tr>
<tr><td>2</td><td>3</td><td>4</td><td>5</td><td>6</td><td>7</td><td>8</td></tr>
</table>
```

To process the table:

```js
var tbl = document.getElementById('sheetjs');
var wb = XLSX.utils.table_to_book(tbl);
```
</details>

Note: `XLSX.read` can handle HTML represented as strings.

### Formulae Output

`XLSX.utils.sheet_to_formulae` generates an array of commands that represent
how a person would enter data into an application.  Each entry is of the form
`A1-cell-address=formula-or-value`.  String literals are prefixed with a `'` in
accordance with Excel.

<details>
  <summary><b>Examples</b> (click to show)</summary>

For the example sheet:

```js
> var o = XLSX.utils.sheet_to_formulae(ws);
> [o[0], o[5], o[10], o[15], o[20]];
[ 'A1=\'S', 'F1=\'J', 'D2=4', 'B3=3', 'G3=8' ]
```
</details>

### Delimiter-Separated Output

As an alternative to the `writeFile` CSV type, `XLSX.utils.sheet_to_csv` also
produces CSV output.  The function takes an options argument:

| Option Name |  Default | Description                                         |
| :---------- | :------: | :-------------------------------------------------- |
|`FS`         |  `","`   | "Field Separator"  delimiter between fields         |
|`RS`         |  `"\n"`  | "Record Separator" delimiter between rows           |
|`dateNF`     |  FMT 14  | Use specified date format in string output          |
|`strip`      |  false   | Remove trailing field separators in each record **  |
|`blankrows`  |  true    | Include blank lines in the CSV output               |
|`skipHidden` |  false   | Skips hidden rows/columns in the CSV output         |

- `strip` will remove trailing commas from each line under default `FS/RS`
- `blankrows` must be set to `false` to skip blank lines.

<details>
  <summary><b>Examples</b> (click to show)</summary>

For the example sheet:

```js
> console.log(XLSX.utils.sheet_to_csv(ws));
S,h,e,e,t,J,S
1,2,3,4,5,6,7
2,3,4,5,6,7,8
> console.log(XLSX.utils.sheet_to_csv(ws, {FS:"\t"}));
S	h	e	e	t	J	S
1	2	3	4	5	6	7
2	3	4	5	6	7	8
> console.log(XLSX.utils.sheet_to_csv(ws,{FS:":",RS:"|"}));
S:h:e:e:t:J:S|1:2:3:4:5:6:7|2:3:4:5:6:7:8|
```
</details>

#### UTF-16 Unicode Text

The `txt` output type uses the tab character as the field separator.  If the
`codepage` library is available (included in full distribution but not core),
the output will be encoded in `CP1200` and the BOM will be prepended.

`XLSX.utils.sheet_to_txt` takes the same arguments as `sheet_to_csv`.

### HTML Output

As an alternative to the `writeFile` HTML type, `XLSX.utils.sheet_to_html` also
produces HTML output.  The function takes an options argument:

| Option Name |  Default | Description                                         |
| :---------- | :------: | :-------------------------------------------------- |
|`id`         |          | Specify the `id` attribute for the `TABLE` element  |
|`editable`   |  false   | If true, set `contenteditable="true"` for every TD  |
|`header`     |          | Override header (default `html body`)               |
|`footer`     |          | Override footer (default `/body /html`)             |

<details>
  <summary><b>Examples</b> (click to show)</summary>

For the example sheet:

```js
> console.log(XLSX.utils.sheet_to_html(ws));
// ...
```
</details>

### JSON

`XLSX.utils.sheet_to_json` generates different types of JS objects. The function
takes an options argument:

| Option Name |  Default | Description                                         |
| :---------- | :------: | :-------------------------------------------------- |
|`raw`        | `false`  | Use raw values (true) or formatted strings (false)  |
|`range`      | from WS  | Override Range (see table below)                    |
|`header`     |          | Control output format (see table below)             |
|`dateNF`     |  FMT 14  | Use specified date format in string output          |
|`defval`     |          | Use specified value in place of null or undefined   |
|`blankrows`  |    **    | Include blank lines in the output **                |

- `raw` only affects cells which have a format code (`.z`) field or a formatted
  text (`.w`) field.
- If `header` is specified, the first row is considered a data row; if `header`
  is not specified, the first row is the header row and not considered data.
- When `header` is not specified, the conversion will automatically disambiguate
  header entries by affixing `_` and a count starting at `1`.  For example, if
  three columns have header `foo` the output fields are `foo`, `foo_1`, `foo_2`
- `null` values are returned when `raw` is true but are skipped when false.
- If `defval` is not specified, null and undefined values are skipped normally.
  If specified, all null and undefined points will be filled with `defval`
- When `header` is `1`, the default is to generate blank rows.  `blankrows` must
  be set to `false` to skip blank rows.
- When `header` is not `1`, the default is to skip blank rows.  `blankrows` must
  be true to generate blank rows

`range` is expected to be one of:

| `range`          | Description                                               |
| :--------------- | :-------------------------------------------------------- |
| (number)         | Use worksheet range but set starting row to the value     |
| (string)         | Use specified range (A1-style bounded range string)       |
| (default)        | Use worksheet range (`ws['!ref']`)                        |

`header` is expected to be one of:

| `header`         | Description                                               |
| :--------------- | :-------------------------------------------------------- |
| `1`              | Generate an array of arrays ("2D Array")                  |
| `"A"`            | Row object keys are literal column labels                 |
| array of strings | Use specified strings as keys in row objects              |
| (default)        | Read and disambiguate first row as keys                   |

If header is not `1`, the row object will contain the non-enumerable property
`__rowNum__` that represents the row of the sheet corresponding to the entry.

<details>
  <summary><b>Examples</b> (click to show)</summary>

For the example sheet:

```js
> XLSX.utils.sheet_to_json(ws);
[ { S: 1, h: 2, e: 3, e_1: 4, t: 5, J: 6, S_1: 7 },
  { S: 2, h: 3, e: 4, e_1: 5, t: 6, J: 7, S_1: 8 } ]

> XLSX.utils.sheet_to_json(ws, {header:"A"});
[ { A: 'S', B: 'h', C: 'e', D: 'e', E: 't', F: 'J', G: 'S' },
  { A: '1', B: '2', C: '3', D: '4', E: '5', F: '6', G: '7' },
  { A: '2', B: '3', C: '4', D: '5', E: '6', F: '7', G: '8' } ]

> XLSX.utils.sheet_to_json(ws, {header:["A","E","I","O","U","6","9"]});
[ { '6': 'J', '9': 'S', A: 'S', E: 'h', I: 'e', O: 'e', U: 't' },
  { '6': '6', '9': '7', A: '1', E: '2', I: '3', O: '4', U: '5' },
  { '6': '7', '9': '8', A: '2', E: '3', I: '4', O: '5', U: '6' } ]

> XLSX.utils.sheet_to_json(ws, {header:1});
[ [ 'S', 'h', 'e', 'e', 't', 'J', 'S' ],
  [ '1', '2', '3', '4', '5', '6', '7' ],
  [ '2', '3', '4', '5', '6', '7', '8' ] ]
```

Example showing the effect of `raw`:

```js
> ws['A2'].w = "3";                          // set A2 formatted string value

> XLSX.utils.sheet_to_json(ws, {header:1});
[ [ 'S', 'h', 'e', 'e', 't', 'J', 'S' ],
  [ '3', '2', '3', '4', '5', '6', '7' ],     // <-- A2 uses the formatted string
  [ '2', '3', '4', '5', '6', '7', '8' ] ]

> XLSX.utils.sheet_to_json(ws, {header:1, raw:true});
[ [ 'S', 'h', 'e', 'e', 't', 'J', 'S' ],
  [ 1, 2, 3, 4, 5, 6, 7 ],                   // <-- A2 uses the raw value
  [ 2, 3, 4, 5, 6, 7, 8 ] ]
```
</details>

