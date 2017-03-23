## Utility Functions

The `sheet_to_*` functions accept a worksheet and an optional options object.

The examples are based on the following worksheet:

```
XXX| A | B | C | D | E | F | G |
---+---+---+---+---+---+---+---+
 1 | S | h | e | e | t | J | S |
 2 | 1 | 2 | 3 | 4 | 5 | 6 | 7 |
 3 | 2 | 3 | 4 | 5 | 6 | 7 | 8 |
```

### Formulae Output

`XLSX.utils.sheet_to_formulae` generates an array of commands that represent
how a person would enter data into an application.  Each entry is of the form
`A1-cell-address=formula-or-value`.  String literals are prefixed with a `'` in
accordance with Excel.  For the example sheet:

```js
> var o = XLSX.utils.sheet_to_formulae(ws);
> o.filter(function(v, i) { return i % 5 === 0; });
[ 'A1=\'S', 'F1=\'J', 'D2=4', 'B3=3', 'G3=8' ]
```

### CSV and general DSV Output

As an alternative to the `writeFile` CSV type, `XLSX.utils.sheet_to_csv` also
produces CSV output.  The function takes an options argument:

| Option Name |  Default | Description                                         |
| :---------- | :------: | :-------------------------------------------------- |
| FS          |  `","`   | "Field Separator"  delimiter between fields         |
| RS          |  `"\n"`  | "Record Separator" delimiter between rows           |
| dateNF      |  fmt 14  | Use specified date format in string output          |
| strip       |  false   | Remove trailing field separators in each record **  |

- `strip` will remove trailing commas from each line under default `FS/RS`

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
> console.log(X.utils.sheet_to_csv(_ws,{FS:":",RS:"|"}));
S:h:e:e:t:J:S|1:2:3:4:5:6:7|2:3:4:5:6:7:8|
```

### JSON

`XLSX.utils.sheet_to_json` and the alias `XLSX.utils.sheet_to_row_object_array`
generate different types of JS objects.  The function takes an options argument:

| Option Name |  Default | Description                                         |
| :---------- | :------: | :-------------------------------------------------- |
| raw         | `false`  | Use raw values (true) or formatted strings (false)  |
| range       | from WS  | Override Range (see table below)                    |
| header      |          | Control output format (see table below)             |
| dateNF      |  fmt 14  | Use specified date format in string output          |
| defval      |          | Use specified value in place of null or undefined   |

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

`range` is expected to be one of:

| `range`          | Description                                               |
| :--------------- | :-------------------------------------------------------- |
| (number)         | Use worksheet range but set starting row to the value     |
| (string)         | Use specified range (A1-style bounded range string)       |
| (default)        | Use worksheet range (`ws['!ref']`)                        |

`header` is expected to be one of:

| `header`         | Description                                               |
| :--------------- | :-------------------------------------------------------- |
| `1`              | Generate an array of arrays                               |
| `"A"`            | Row object keys are literal column labels                 |
| array of strings | Use specified strings as keys in row objects              |
| (default)        | Read and disambiguate first row as keys                   |

For the example sheet:

```js
> console.log(X.utils.sheet_to_json(_ws));
[ { S: 1, h: 2, e: 3, e_1: 4, t: 5, J: 6, S_1: 7 },
  { S: 2, h: 3, e: 4, e_1: 5, t: 6, J: 7, S_1: 8 } ]

> console.log(X.utils.sheet_to_json(_ws, {header:1}));
[ [ 'S', 'h', 'e', 'e', 't', 'J', 'S' ],
  [ 1, 2, 3, 4, 5, 6, 7 ],
  [ 2, 3, 4, 5, 6, 7, 8 ] ]

> console.log(X.utils.sheet_to_json(_ws, {header:"A"}));
[ { A: 'S', B: 'h', C: 'e', D: 'e', E: 't', F: 'J', G: 'S' },
  { A: 1, B: 2, C: 3, D: 4, E: 5, F: 6, G: 7 },
  { A: 2, B: 3, C: 4, D: 5, E: 6, F: 7, G: 8 } ]
> console.log(X.utils.sheet_to_json(_ws, {header:["A","E","I","O","U","6","9"]}));
[ { '6': 'J', '9': 'S', A: 'S', E: 'h', I: 'e', O: 'e', U: 't' },
  { '6': 6, '9': 7, A: 1, E: 2, I: 3, O: 4, U: 5 },
  { '6': 7, '9': 8, A: 2, E: 3, I: 4, O: 5, U: 6 } ]
```

Example showing the effect of `raw`:

```js
> _ws['A2'].w = "1";                         // set A2 formatted string value
> console.log(X.utils.sheet_to_json(_ws, {header:1}));
[ [ 'S', 'h', 'e', 'e', 't', 'J', 'S' ],
  [ '1', 2, 3, 4, 5, 6, 7 ],                 // <-- A2 uses the formatted string
  [ 2, 3, 4, 5, 6, 7, 8 ] ]
> console.log(X.utils.sheet_to_json(_ws, {header:1, raw:true}));
[ [ 'S', 'h', 'e', 'e', 't', 'J', 'S' ],
  [ 1, 2, 3, 4, 5, 6, 7 ],                   // <-- A2 uses the raw value
  [ 2, 3, 4, 5, 6, 7, 8 ] ]
```

