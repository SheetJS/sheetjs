#### Number Formats

The `cell.w` formatted text for each cell is produced from `cell.v` and `cell.z`
format.  If the format is not specified, the Excel `General` format is used.
The format can either be specified as a string or as an index into the format
table.  Parsers are expected to populate `workbook.SSF` with the number format
table.  Writers are expected to serialize the table.

Custom tools should ensure that the local table has each used format string
somewhere in the table.  Excel convention mandates that the custom formats start
at index 164.  The following example creates a custom format from scratch:

<details>
	<summary><b>New worksheet with custom format</b> (click to show)</summary>

```js
var tbl = {};
XLSX.SSF.init_table(tbl); // <-- load builtin formats
tbl[164] = "\"T\"\ #0.00";
var wb = {
	SSF: tbl,
	SheetNames: ["Sheet1"],
	Sheets: {
		Sheet1: {
			"!ref":"A1:C1",
			A1: { t:"n", v:10000 },             // <-- General format
			B1: { t:"n", v:10000, z: tbl[4] },  // <-- Builtin format
			C1: { t:"n", v:10000, z: tbl[164] } // <-- Custom format
		}
	}
}
```
</details>

The rules are slightly different from how Excel displays custom number formats.
In particular, literal characters must be wrapped in double quotes or preceded
by a backslash. For more info, see the Excel documentation article
[`Create or delete a custom number format`](https://support.office.com/en-us/article/78f2a361-936b-4c03-8772-09fab54be7f4)
or ECMA-376 18.8.31 (Number Formats)

<details>
	<summary><b>Default Number Formats</b> (click to show)</summary>

The default formats are listed in ECMA-376 18.8.30:

| ID | Format                     |
|---:|:---------------------------|
|  0 | `General`                  |
|  1 | `0`                        |
|  2 | `0.00`                     |
|  3 | `#,##0`                    |
|  4 | `#,##0.00`                 |
|  9 | `0%`                       |
| 10 | `0.00%`                    |
| 11 | `0.00E+00`                 |
| 12 | `# ?/?`                    |
| 13 | `# ??/??`                  |
| 14 | `m/d/yy` (see below)       |
| 15 | `d-mmm-yy`                 |
| 16 | `d-mmm`                    |
| 17 | `mmm-yy`                   |
| 18 | `h:mm AM/PM`               |
| 19 | `h:mm:ss AM/PM`            |
| 20 | `h:mm`                     |
| 21 | `h:mm:ss`                  |
| 22 | `m/d/yy h:mm`              |
| 37 | `#,##0 ;(#,##0)`           |
| 38 | `#,##0 ;[Red](#,##0)`      |
| 39 | `#,##0.00;(#,##0.00)`      |
| 40 | `#,##0.00;[Red](#,##0.00)` |
| 45 | `mm:ss`                    |
| 46 | `[h]:mm:ss`                |
| 47 | `mmss.0`                   |
| 48 | `##0.0E+0`                 |
| 49 | `@`                        |

</details>

Format 14 (`m/d/yy`) is localized by Excel: even though the file specifies that
number format, it will be drawn differently based on system settings.  It makes
sense when the producer and consumer of files are in the same locale, but that
is not always the case over the Internet.  To get around this ambiguity, parse
functions accept the `dateNF` option to override the interpretation of that
specific format string.
